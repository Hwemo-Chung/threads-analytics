"""
게시글 텍스트 저장함 생성기

- 전체 게시 본문 정리 저장 (md / jsonl / csv)
- 월별·미디어타입별 분류
- analyze.py 연동 또는 단독 실행

사용법:
  python3 archive.py
  python3 archive.py -i output/analysis_YYYYMMDD_HHMMSS.json
  python3 archive.py -i path.json -o output/저장함/custom
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional, Tuple

JST = timezone(timedelta(hours=9))


def parse_ts(ts_str: Optional[str]) -> Optional[datetime]:
    if not ts_str:
        return None
    try:
        return datetime.fromisoformat(ts_str.replace("+0000", "+00:00")).astimezone(JST)
    except ValueError:
        return None


def load_analysis(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _insight(post: dict, key: str, default: int = 0) -> int:
    ins = post.get("insights") or {}
    val = ins.get(key, default)
    try:
        return int(val or 0)
    except (TypeError, ValueError):
        return default


def _active_posts(posts: List[dict]) -> List[dict]:
    return [p for p in posts if p.get("media_type") != "REPOST_FACADE"]


def _post_rows(posts: List[dict]) -> List[dict]:
    rows = []
    for p in posts:
        dt = parse_ts(p.get("timestamp"))
        text = p.get("text") or ""
        views = _insight(p, "views")
        likes = _insight(p, "likes")
        replies = _insight(p, "replies")
        reposts = _insight(p, "reposts")
        quotes = _insight(p, "quotes")
        shares = _insight(p, "shares")
        engagement = likes + replies + reposts + quotes
        rows.append({
            "id": p.get("id", ""),
            "timestamp": p.get("timestamp", ""),
            "datetime_jst": dt.strftime("%Y-%m-%d %H:%M") if dt else "",
            "date": dt.strftime("%Y-%m-%d") if dt else "",
            "month": dt.strftime("%Y-%m") if dt else "unknown",
            "media_type": p.get("media_type") or "UNKNOWN",
            "text": text,
            "text_len": len(text.replace("\n", " ").strip()),
            "permalink": p.get("permalink") or "",
            "views": views,
            "likes": likes,
            "replies": replies,
            "reposts": reposts,
            "quotes": quotes,
            "shares": shares,
            "engagement": engagement,
            "like_rate": round((likes / views * 100), 4) if views > 0 else 0.0,
            "is_repost": p.get("media_type") == "REPOST_FACADE",
        })
    # newest first
    rows.sort(key=lambda r: r["timestamp"] or "", reverse=True)
    return rows


def _write_md_entry(f, idx: int, row: dict) -> None:
    title_date = row["datetime_jst"] or row["timestamp"] or "unknown"
    f.write(f"## {idx}. [{title_date}] {row['media_type']}\n\n")
    f.write(
        f"- id: `{row['id']}`\n"
        f"- 조회 {row['views']:,} · 좋아요 {row['likes']:,} · 답글 {row['replies']:,} · "
        f"리포스트 {row['reposts']:,} · 인용 {row['quotes']:,} · "
        f"인게이지먼트 {row['engagement']:,} · 좋아요율 {row['like_rate']}%\n"
    )
    if row["permalink"]:
        f.write(f"- 링크: {row['permalink']}\n")
    f.write("\n")
    body = row["text"].strip() if row["text"] else "_(텍스트 없음 — 미디어 전용)_"
    f.write(body)
    f.write("\n\n---\n\n")


def save_text_archive(
    profile: dict,
    posts: List[dict],
    user_insights: Optional[dict] = None,
    analyzed_at: Optional[str] = None,
    archive_root: Optional[str] = None,
    source_path: Optional[str] = None,
) -> str:
    """
    Write organized text archive under output/저장함/.

    Returns absolute path to the snapshot directory.
    """
    base_dir = archive_root or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "output", "저장함"
    )
    run_ts = datetime.now(JST).strftime("%Y%m%d_%H%M%S")
    snap_dir = os.path.join(base_dir, "snapshots", run_ts)
    by_month_dir = os.path.join(snap_dir, "by_month")
    by_type_dir = os.path.join(snap_dir, "by_type")
    os.makedirs(by_month_dir, exist_ok=True)
    os.makedirs(by_type_dir, exist_ok=True)

    all_rows = _post_rows(posts)
    active_rows = [r for r in all_rows if not r["is_repost"]]
    username = (profile or {}).get("username") or "unknown"
    followers = (user_insights or {}).get("followers_count", 0)
    analyzed_at = analyzed_at or datetime.now(JST).isoformat()

    # --- index.md ---
    type_counter = Counter(r["media_type"] for r in active_rows)
    month_counter = Counter(r["month"] for r in active_rows)
    with_text = sum(1 for r in active_rows if r["text"].strip())
    total_chars = sum(r["text_len"] for r in active_rows)

    index_path = os.path.join(snap_dir, "README.md")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(f"# Threads 저장함 — @{username}\n\n")
        f.write(f"- 생성 시각: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S')} JST\n")
        f.write(f"- 원본 분석 시각: {analyzed_at}\n")
        if source_path:
            f.write(f"- 소스 JSON: `{os.path.basename(source_path)}`\n")
        f.write(f"- 팔로워: {followers:,}\n")
        f.write(f"- 전체 게시물: {len(all_rows)} (리포스트 제외 {len(active_rows)})\n")
        f.write(f"- 텍스트 있는 게시물: {with_text}\n")
        f.write(f"- 총 글자 수(공백 정리 후): {total_chars:,}\n\n")
        f.write("## 파일 안내\n\n")
        f.write("| 파일 | 내용 |\n|---|---|\n")
        f.write("| `all_texts.md` | 전체 본문 (최신순, 리포스트 제외) |\n")
        f.write("| `all_posts.jsonl` | 게시물 1줄 1JSON (메트릭 포함) |\n")
        f.write("| `index.csv` | 표 형태 목록 (엑셀 열기 가능) |\n")
        f.write("| `by_month/*.md` | 월별 본문 묶음 |\n")
        f.write("| `by_type/*.md` | 미디어 타입별 본문 묶음 |\n")
        f.write("| `text_only.txt` | 본문만 연속 텍스트 |\n\n")
        f.write("## 미디어 타입 분포\n\n")
        for mt, n in type_counter.most_common():
            f.write(f"- {mt}: {n}\n")
        f.write("\n## 월별 게시 수\n\n")
        for month in sorted(month_counter.keys()):
            f.write(f"- {month}: {month_counter[month]}\n")

    # --- all_texts.md ---
    with open(os.path.join(snap_dir, "all_texts.md"), "w", encoding="utf-8") as f:
        f.write(f"# @{username} 전체 게시 텍스트\n\n")
        f.write(f"총 {len(active_rows)}개 (REPOST 제외) · 최신순\n\n---\n\n")
        for i, row in enumerate(active_rows, 1):
            _write_md_entry(f, i, row)

    # --- text_only.txt ---
    with open(os.path.join(snap_dir, "text_only.txt"), "w", encoding="utf-8") as f:
        for row in active_rows:
            body = (row["text"] or "").strip()
            if not body:
                continue
            f.write(body)
            f.write("\n\n====\n\n")

    # --- all_posts.jsonl ---
    with open(os.path.join(snap_dir, "all_posts.jsonl"), "w", encoding="utf-8") as f:
        for row in all_rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")

    # --- index.csv ---
    csv_path = os.path.join(snap_dir, "index.csv")
    fieldnames = [
        "id", "datetime_jst", "month", "media_type", "text_len",
        "views", "likes", "replies", "reposts", "quotes", "shares",
        "engagement", "like_rate", "permalink", "text_preview",
    ]
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in all_rows:
            preview = (row["text"] or "").replace("\n", " ")[:120]
            writer.writerow({**row, "text_preview": preview})

    # --- by_month ---
    by_month: Dict[str, List[dict]] = defaultdict(list)
    for row in active_rows:
        by_month[row["month"]].append(row)
    for month, rows in sorted(by_month.items()):
        path = os.path.join(by_month_dir, f"{month}.md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"# {month} · {len(rows)}개\n\n")
            for i, row in enumerate(rows, 1):
                _write_md_entry(f, i, row)

    # --- by_type ---
    by_type: Dict[str, List[dict]] = defaultdict(list)
    for row in active_rows:
        by_type[row["media_type"]].append(row)
    for mtype, rows in sorted(by_type.items(), key=lambda x: -len(x[1])):
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in mtype)
        path = os.path.join(by_type_dir, f"{safe}.md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"# {mtype} · {len(rows)}개\n\n")
            for i, row in enumerate(rows, 1):
                _write_md_entry(f, i, row)

    # --- manifest.json ---
    manifest = {
        "created_at": datetime.now(JST).isoformat(),
        "analyzed_at": analyzed_at,
        "source_path": source_path,
        "username": username,
        "followers": followers,
        "post_count": len(all_rows),
        "active_count": len(active_rows),
        "with_text": with_text,
        "total_chars": total_chars,
        "type_counts": dict(type_counter),
        "month_counts": dict(sorted(month_counter.items())),
        "snapshot_dir": snap_dir,
    }
    with open(os.path.join(snap_dir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    # --- latest pointer (copy key files via rewrite path file) ---
    latest_dir = os.path.join(base_dir, "latest")
    os.makedirs(latest_dir, exist_ok=True)
    # Write pointer instead of full copy (cheap + clear)
    pointer = {
        "snapshot": snap_dir,
        "created_at": manifest["created_at"],
        "username": username,
        "active_count": len(active_rows),
    }
    with open(os.path.join(latest_dir, "POINTER.json"), "w", encoding="utf-8") as f:
        json.dump(pointer, f, ensure_ascii=False, indent=2)

    # Also materialize latest copies of primary human-readable files
    for name in ("README.md", "all_texts.md", "text_only.txt", "index.csv", "all_posts.jsonl", "manifest.json"):
        src = os.path.join(snap_dir, name)
        dst = os.path.join(latest_dir, name)
        if os.path.isfile(src):
            with open(src, "r", encoding="utf-8") as rf:
                content = rf.read()
            with open(dst, "w", encoding="utf-8") as wf:
                wf.write(content)

    # latest by_month / by_type mirrors
    for sub in ("by_month", "by_type"):
        src_sub = os.path.join(snap_dir, sub)
        dst_sub = os.path.join(latest_dir, sub)
        os.makedirs(dst_sub, exist_ok=True)
        # clear old md files in latest subdirs
        for old in os.listdir(dst_sub):
            if old.endswith(".md"):
                try:
                    os.remove(os.path.join(dst_sub, old))
                except OSError:
                    pass
        for fname in os.listdir(src_sub):
            if not fname.endswith(".md"):
                continue
            with open(os.path.join(src_sub, fname), "r", encoding="utf-8") as rf:
                content = rf.read()
            with open(os.path.join(dst_sub, fname), "w", encoding="utf-8") as wf:
                wf.write(content)

    return snap_dir


def resolve_input(input_arg: Optional[str]) -> str:
    import glob as _glob

    if input_arg:
        if not os.path.isfile(input_arg):
            print(f"[ERROR] 입력 파일 없음: {input_arg}")
            sys.exit(1)
        return input_arg

    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    files = sorted(_glob.glob(os.path.join(output_dir, "analysis_*.json")))
    if not files:
        print("[ERROR] output/analysis_*.json 없음. 먼저 analyze.py를 실행하세요.")
        sys.exit(1)
    return files[-1]


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Threads 게시글 텍스트 저장함 생성")
    p.add_argument("-i", "--input", default=None, help="analysis_*.json 경로")
    p.add_argument(
        "-o", "--output",
        default=None,
        help="저장함 루트 (기본: output/저장함)",
    )
    p.add_argument(
        "--include-reposts-in-md",
        action="store_true",
        help="(예약) md에 리포스트 포함 — 현재 기본 제외, jsonl/csv에는 항상 포함",
    )
    return p.parse_args(argv)


def main(argv=None) -> str:
    args = parse_args(argv)
    input_path = resolve_input(args.input)
    data = load_analysis(input_path)

    snap = save_text_archive(
        profile=data.get("profile") or {},
        posts=data.get("posts") or [],
        user_insights=data.get("user_insights") or {},
        analyzed_at=data.get("analyzed_at"),
        archive_root=args.output,
        source_path=input_path,
    )

    print(f"입력: {input_path}")
    print(f"저장함 스냅샷: {snap}")
    print(f"최신 포인터: {os.path.join(os.path.dirname(snap), '..', 'latest')}")
    # normalize latest path display
    latest = os.path.normpath(os.path.join(snap, "..", "..", "latest"))
    print(f"latest 경로: {latest}")
    return snap


if __name__ == "__main__":
    main()
