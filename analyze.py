"""
Threads 계정 종합 분석 스크립트

사용법:
  python analyze.py
  python analyze.py --refresh-insights --fail-on-api-error
  python analyze.py --ttl-days 30 --max-posts 100
사전 조건: auth.py 실행하여 .env에 ACCESS_TOKEN 저장 완료
"""

import argparse
import os
import sys
import json
import time
from datetime import datetime, timedelta, timezone
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from dotenv import load_dotenv
from tabulate import tabulate

load_dotenv()

ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "")
USER_ID = os.getenv("USER_ID", "me")
API_BASE = "https://graph.threads.net/v1.0"
# 0 = never expire; default 7 days (stale views/likes refresh)
INSIGHTS_CACHE_TTL_DAYS = int(os.getenv("INSIGHTS_CACHE_TTL_DAYS", "7"))
API_MAX_RETRIES = int(os.getenv("API_MAX_RETRIES", "3"))
TOKEN_WARN_DAYS = int(os.getenv("TOKEN_WARN_DAYS", "7"))

JST = timezone(timedelta(hours=9))
UTC = timezone.utc

# Runtime API counters (reset per run)
_api_stats = {"ok": 0, "error": 0, "timeout": 0, "retry": 0}


def reset_api_stats():
    _api_stats.update({"ok": 0, "error": 0, "timeout": 0, "retry": 0})


def api_get(endpoint: str, params: dict = None, max_retries: int = None) -> dict:
    """GET with timeout, 429/5xx backoff, and error counters."""
    params = dict(params or {})
    params["access_token"] = ACCESS_TOKEN
    retries = API_MAX_RETRIES if max_retries is None else max_retries

    for attempt in range(retries + 1):
        try:
            resp = requests.get(f"{API_BASE}/{endpoint}", params=params, timeout=30)
        except requests.exceptions.Timeout:
            _api_stats["timeout"] += 1
            if attempt < retries:
                wait = 2 ** attempt
                _api_stats["retry"] += 1
                print(f"  [TIMEOUT] {endpoint}: 재시도 {attempt + 1}/{retries} ({wait}s)")
                time.sleep(wait)
                continue
            print(f"  [TIMEOUT] {endpoint}: 30초 초과 (재시도 소진)")
            return {}
        except requests.exceptions.RequestException as exc:
            _api_stats["error"] += 1
            if attempt < retries:
                wait = 2 ** attempt
                _api_stats["retry"] += 1
                print(f"  [NETWORK] {endpoint}: {exc} — 재시도 {attempt + 1}/{retries}")
                time.sleep(wait)
                continue
            print(f"  [NETWORK] {endpoint}: {exc}")
            return {}

        if resp.status_code == 429 or resp.status_code >= 500:
            _api_stats["retry"] += 1
            if attempt >= retries:
                _api_stats["error"] += 1
                print(
                    f"  [API ERROR] {endpoint}: {resp.status_code} - {resp.text[:200]} "
                    f"(재시도 소진)"
                )
                return {}
            retry_after = resp.headers.get("Retry-After", "")
            if retry_after.isdigit():
                wait = max(int(retry_after), 1)
            else:
                wait = min(2 ** attempt * 2, 60)
            print(
                f"  [RATE/SERVER] {endpoint}: {resp.status_code} — "
                f"{wait}s 후 재시도 ({attempt + 1}/{retries})"
            )
            time.sleep(wait)
            continue

        if resp.status_code != 200:
            _api_stats["error"] += 1
            print(f"  [API ERROR] {endpoint}: {resp.status_code} - {resp.text[:200]}")
            return {}

        _api_stats["ok"] += 1
        try:
            return resp.json()
        except ValueError:
            _api_stats["error"] += 1
            print(f"  [API ERROR] {endpoint}: JSON 파싱 실패")
            return {}

    return {}


def validate_token():
    if not ACCESS_TOKEN or ACCESS_TOKEN == "":
        print("[ERROR] ACCESS_TOKEN이 없습니다. 먼저 auth.py를 실행해주세요.")
        sys.exit(1)


def warn_token_expiry(warn_days: int = TOKEN_WARN_DAYS) -> None:
    """Print token expiry status from TOKEN_EXPIRES_AT in .env."""
    raw = os.getenv("TOKEN_EXPIRES_AT", "").strip()
    if not raw:
        print(
            "[WARN] TOKEN_EXPIRES_AT 없음. auth.py/refresh_token.py 재실행 시 기록됩니다. "
            "만료 전 refresh_token.py 권장."
        )
        return
    try:
        expires_at = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        if expires_at.tzinfo is None:
            expires_at = expires_at.replace(tzinfo=UTC)
    except ValueError:
        print(f"[WARN] TOKEN_EXPIRES_AT 형식 오류: {raw}")
        return

    now = datetime.now(UTC)
    remaining = expires_at.astimezone(UTC) - now
    days_left = remaining.total_seconds() / 86400
    local = expires_at.astimezone(JST).strftime("%Y-%m-%d %H:%M JST")

    if days_left <= 0:
        print(f"[ERROR] 토큰 만료됨 ({local}). auth.py를 다시 실행하세요.")
        sys.exit(1)
    if days_left <= warn_days:
        print(
            f"[WARN] 토큰 {days_left:.1f}일 후 만료 예정 ({local}). "
            f"python refresh_token.py 실행을 권장합니다."
        )
    else:
        print(f"토큰 유효: 약 {days_left:.0f}일 남음 (만료 {local})")


def fetch_profile() -> dict:
    print("\n1. 프로필 정보 조회 중...")
    fields = "id,username,name,threads_profile_picture_url,threads_biography,is_verified"
    data = api_get(USER_ID, {"fields": fields})
    if not data:
        return {}

    rows = [
        ["유저 ID", data.get("id", "-")],
        ["유저네임", f"@{data.get('username', '-')}"],
        ["이름", data.get("name", "-")],
        ["바이오", data.get("threads_biography", "-")],
        ["인증 계정", "O" if data.get("is_verified") else "X"],
    ]
    print(tabulate(rows, headers=["항목", "값"], tablefmt="simple_outline"))
    return data


def fetch_all_posts(max_posts: int = None) -> list:
    print("\n2. 게시물 전체 조회 중...")
    all_posts = []
    # topic_tag는 토픽 없는 게시물에서 키 자체가 빠져서 옴 (누락 = 토픽 없음)
    fields = "id,media_type,text,timestamp,permalink,like_count,replies_count,topic_tag"

    params = {"fields": fields, "limit": 100}
    page = 1

    while True:
        data = api_get(f"{USER_ID}/threads", params)
        posts = data.get("data", [])
        if not posts:
            break

        all_posts.extend(posts)
        print(f"  페이지 {page}: {len(posts)}개 로드 (총 {len(all_posts)}개)")

        if max_posts is not None and len(all_posts) >= max_posts:
            all_posts = all_posts[:max_posts]
            print(f"  --max-posts={max_posts} 적용, 수집 중단")
            break

        cursors = data.get("paging", {}).get("cursors", {})
        after = cursors.get("after")
        if not after or len(posts) < 100:
            break

        params = {"fields": fields, "limit": 100, "after": after}
        page += 1
        time.sleep(0.5)

    print(f"  총 {len(all_posts)}개 게시물 조회 완료")
    return all_posts


def _load_insights_cache(cache_path: str) -> dict:
    """Load insights cache. Migrate v1 {id: metrics} → v2 {id: {fetched_at, metrics}}."""
    if not os.path.exists(cache_path):
        return {}
    with open(cache_path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if not isinstance(raw, dict):
        return {}

    if raw.get("version") == 2 and isinstance(raw.get("entries"), dict):
        return raw["entries"]

    entries = {}
    for media_id, value in raw.items():
        if media_id == "version":
            continue
        if not isinstance(value, dict):
            continue
        if "metrics" in value and isinstance(value["metrics"], dict):
            entries[media_id] = {
                "fetched_at": value.get("fetched_at", "1970-01-01T00:00:00+00:00"),
                "metrics": value["metrics"],
            }
        else:
            # v1 flat metrics: no timestamp → force refresh once under TTL
            entries[media_id] = {
                "fetched_at": "1970-01-01T00:00:00+00:00",
                "metrics": value,
            }
    return entries


def _is_cache_fresh(entry: dict, ttl_days: int) -> bool:
    if ttl_days <= 0:
        return True
    fetched_at = entry.get("fetched_at")
    if not fetched_at:
        return False
    try:
        ts = datetime.fromisoformat(str(fetched_at).replace("Z", "+00:00"))
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=UTC)
        age = datetime.now(UTC) - ts.astimezone(UTC)
        return age <= timedelta(days=ttl_days)
    except ValueError:
        return False


def _save_insights_cache(cache_path: str, entries: dict) -> None:
    """Atomic write so crash mid-save does not corrupt the cache file."""
    payload = {"version": 2, "entries": entries}
    tmp_path = cache_path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f)
    os.replace(tmp_path, cache_path)


def fetch_post_insights(
    posts: list,
    ttl_days: int = None,
    force_refresh: bool = False,
    max_workers: int = 5,
) -> list:
    print("\n3. 게시물별 인사이트 수집 중 (5 병렬, 배치 저장)...")
    total = len(posts)
    ttl = INSIGHTS_CACHE_TTL_DAYS if ttl_days is None else ttl_days
    metrics = "views,likes,replies,reposts,quotes,shares"
    cache_path = os.path.join(os.path.dirname(__file__), "output", "insights_cache.json")
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)

    cached = _load_insights_cache(cache_path)
    if force_refresh:
        fresh_count = 0
        print(f"  캐시 로드: {len(cached)}개 (--refresh-insights: 전체 재조회)")
        to_fetch = list(posts)
    else:
        fresh_count = sum(1 for e in cached.values() if _is_cache_fresh(e, ttl))
        print(f"  캐시 로드: {len(cached)}개 (TTL {ttl}일 이내 유효: {fresh_count}개)")
        to_fetch = [
            p for p in posts
            if p["id"] not in cached or not _is_cache_fresh(cached[p["id"]], ttl)
        ]
    print(f"  신규/만료 재조회: {len(to_fetch)}개 / 총 {total}개")

    def fetch_single(post):
        media_id = post["id"]
        insights_data = api_get(f"{media_id}/insights", {"metric": metrics})
        insight_map = {}
        for item in insights_data.get("data", []):
            name = item["name"]
            value = item.get("values", [{}])[0].get("value", 0)
            insight_map[name] = value
        return media_id, insight_map

    BATCH_SIZE = 100
    done = 0
    fetched_at = datetime.now(UTC).isoformat()
    for batch_start in range(0, len(to_fetch), BATCH_SIZE):
        batch = to_fetch[batch_start:batch_start + BATCH_SIZE]
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(fetch_single, p): p for p in batch}
            for future in as_completed(futures):
                try:
                    media_id, insight_map = future.result()
                except Exception as exc:
                    post = futures[future]
                    print(f"  [ERROR] insights {post.get('id')}: {exc}")
                    continue
                # Keep previous metrics if API returned empty (transient failure)
                if not insight_map and media_id in cached:
                    continue
                cached[media_id] = {
                    "fetched_at": fetched_at,
                    "metrics": insight_map,
                }
                done += 1
                if done % 50 == 0:
                    print(f"  {done}/{len(to_fetch)} 완료")

        _save_insights_cache(cache_path, cached)
        print(f"  배치 저장 ({batch_start + len(batch)}/{len(to_fetch)})")

    print(f"  전체 인사이트 수집 완료: {len(cached)}개")

    enriched = []
    for p in posts:
        entry = cached.get(p["id"], {})
        p["insights"] = entry.get("metrics", {}) if isinstance(entry, dict) else {}
        enriched.append(p)
    return enriched


def _insight_total(item: dict):
    """값이 없으면 None. 진짜 0과 '데이터 없음'을 구분해야 결측 스냅샷이 안 섞인다."""
    total_value = item.get("total_value")
    if isinstance(total_value, dict):
        value = total_value.get("value")
        if isinstance(value, (int, float)):
            return int(value)

    values = item.get("values", [])
    if isinstance(values, list) and values:
        return sum(v.get("value", 0) for v in values if isinstance(v, dict))
    return None


def fetch_user_insights() -> dict:
    print("\n4. 계정 인사이트 조회 중...")
    result = {}

    followers_data = api_get(
        f"{USER_ID}/threads_insights",
        {"metric": "followers_count"},
    )
    for item in followers_data.get("data", []):
        if item["name"] == "followers_count":
            total = _insight_total(item)
            # 0/None을 그대로 저장하면 스냅샷 간 비교에서 가짜 급감이 된다
            if total is not None and total > 0:
                result["followers_count"] = total
    if "followers_count" not in result:
        print("  [WARN] 팔로워 수를 가져오지 못했습니다. 이번 스냅샷에는 기록하지 않습니다.")

    now = datetime.now(JST)
    since = int((now - timedelta(days=30)).timestamp())
    until = int(now.timestamp())

    period_data = api_get(
        f"{USER_ID}/threads_insights",
        {
            "metric": "views,likes,replies,reposts,quotes",
            "since": since,
            "until": until,
        },
    )
    for item in period_data.get("data", []):
        name = item["name"]
        total = _insight_total(item)
        if total is None:
            print(f"  [WARN] 30일 {name} 지표가 비어 있습니다. 기록하지 않습니다.")
            continue
        result[f"30d_{name}"] = total

    # 일별 조회수 시계열(730일). _insight_total()은 일 버킷을 합쳐버리므로 직접 파싱한다.
    try:
        daily_data = api_get(
            f"{USER_ID}/threads_insights",
            {
                "metric": "views",
                "since": int((now - timedelta(days=729)).timestamp()),
                "until": until,
            },
        )
        for item in daily_data.get("data", []):
            if item.get("name") != "views":
                continue
            daily = [
                {"date": v["end_time"][:10], "value": int(v.get("value", 0) or 0)}
                for v in item.get("values", [])
                if isinstance(v, dict) and v.get("end_time")
            ]
            if daily:
                result["daily_views"] = daily
    except Exception as exc:
        print(f"  [WARN] 일별 조회수 시계열 조회 실패: {exc}")

    def fmt(key):
        value = result.get(key)
        return f"{value:,}" if isinstance(value, int) else "-"

    rows = [
        ["팔로워 수", fmt("followers_count")],
        ["조회수 (30일)", fmt("30d_views")],
        ["좋아요 (30일)", fmt("30d_likes")],
        ["답글 (30일)", fmt("30d_replies")],
        ["리포스트 (30일)", fmt("30d_reposts")],
        ["인용 (30일)", fmt("30d_quotes")],
    ]
    print(tabulate(rows, headers=["지표", "값"], tablefmt="simple_outline"))
    return result


def fetch_follower_demographics() -> dict:
    print("\n5. 팔로워 인구통계 조회 중...")
    demographics = {}

    for breakdown in ["country", "city", "gender", "age"]:
        data = api_get(
            f"{USER_ID}/threads_insights",
            {"metric": "follower_demographics", "breakdown": breakdown},
        )
        items = data.get("data", [])
        if not items:
            print(f"  {breakdown}: 데이터 없음 (팔로워 100명 미만일 수 있음)")
            continue

        breakdown_results = {}
        for item in items:
            for val_obj in item.get("total_value", {}).get("breakdowns", [{}]):
                for result in val_obj.get("results", []):
                    key = result.get("dimension_values", ["-"])[0]
                    value = result.get("value", 0)
                    breakdown_results[key] = value

        if not breakdown_results:
            for item in items:
                for val_obj in item.get("values", [{}]):
                    if isinstance(val_obj.get("value"), dict):
                        breakdown_results = val_obj["value"]

        demographics[breakdown] = breakdown_results

        if breakdown_results:
            sorted_items = sorted(breakdown_results.items(), key=lambda x: x[1], reverse=True)[:10]
            label_map = {"country": "국가", "city": "도시", "gender": "성별", "age": "연령대"}
            print(f"\n  [{label_map.get(breakdown, breakdown)}] Top 10:")
            rows = [[k, f"{v:,}명"] for k, v in sorted_items]
            print(tabulate(rows, headers=[label_map.get(breakdown, breakdown), "팔로워"], tablefmt="simple_outline"))

    return demographics


def print_post_ranking(posts: list):
    print("\n" + "=" * 60)
    print("6. 게시물 성과 랭킹")
    print("=" * 60)

    active = [p for p in posts if p.get("media_type") != "REPOST_FACADE"]

    def get_engagement(p):
        ins = p.get("insights") or {}
        return ins.get("likes", 0) + ins.get("replies", 0) + ins.get("reposts", 0) + ins.get("quotes", 0)

    sorted_by_engagement = sorted(active, key=get_engagement, reverse=True)[:15]
    sorted_by_views = sorted(active, key=lambda p: (p.get("insights") or {}).get("views", 0), reverse=True)[:15]

    print("\n[인게이지먼트 TOP 15]")
    rows = []
    for i, p in enumerate(sorted_by_engagement, 1):
        ins = p.get("insights", {})
        text_preview = (p.get("text") or "")[:40].replace("\n", " ")
        engagement = get_engagement(p)
        rows.append([
            i,
            text_preview or "(미디어)",
            f"{ins.get('views', 0):,}",
            ins.get("likes", 0),
            ins.get("replies", 0),
            ins.get("reposts", 0),
            engagement,
        ])
    print(tabulate(rows, headers=["#", "내용", "조회", "좋아요", "답글", "리포스트", "총 인게이지먼트"], tablefmt="simple_outline"))

    print("\n[조회수 TOP 15]")
    rows = []
    for i, p in enumerate(sorted_by_views, 1):
        ins = p.get("insights", {})
        text_preview = (p.get("text") or "")[:40].replace("\n", " ")
        rows.append([
            i,
            text_preview or "(미디어)",
            f"{ins.get('views', 0):,}",
            ins.get("likes", 0),
            ins.get("replies", 0),
        ])
    print(tabulate(rows, headers=["#", "내용", "조회수", "좋아요", "답글"], tablefmt="simple_outline"))


def print_content_analysis(posts: list):
    print("\n" + "=" * 60)
    print("7. 콘텐츠 분석")
    print("=" * 60)

    active = [p for p in posts if p.get("media_type") != "REPOST_FACADE"]
    type_counter = Counter()
    type_engagement = {}

    for p in active:
        media_type = p.get("media_type", "UNKNOWN")
        type_counter[media_type] += 1

        ins = p.get("insights") or {}
        engagement = (
            ins.get("likes", 0)
            + ins.get("replies", 0)
            + ins.get("reposts", 0)
            + ins.get("quotes", 0)
        )
        if media_type not in type_engagement:
            type_engagement[media_type] = []
        type_engagement[media_type].append(engagement)

    print("\n[미디어 타입별 분석]")
    rows = []
    for mtype, count in type_counter.most_common():
        eng_list = type_engagement.get(mtype, [])
        avg_eng = sum(eng_list) / len(eng_list) if eng_list else 0
        rows.append([mtype, count, f"{avg_eng:.1f}"])
    print(tabulate(rows, headers=["타입", "게시물 수", "평균 인게이지먼트"], tablefmt="simple_outline"))

    if active:
        timestamps = []
        for p in active:
            ts = p.get("timestamp", "")
            if ts:
                try:
                    dt = datetime.fromisoformat(ts.replace("+0000", "+00:00"))
                    timestamps.append(dt)
                except ValueError:
                    pass

        if timestamps:
            hour_counter = Counter(dt.astimezone(JST).hour for dt in timestamps)
            # Locale-independent weekday keys (match export_excel)
            weekday_counter = Counter(
                ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][dt.astimezone(JST).weekday()]
                for dt in timestamps
            )

            print("\n[시간대별 게시 빈도 (JST)]")
            rows = []
            for hour in range(24):
                count = hour_counter.get(hour, 0)
                bar = "█" * count
                rows.append([f"{hour:02d}시", count, bar])
            print(tabulate(rows, headers=["시간", "게시 수", ""], tablefmt="simple_outline"))

            print("\n[요일별 게시 빈도]")
            day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            day_kr = {"Monday": "월", "Tuesday": "화", "Wednesday": "수", "Thursday": "목", "Friday": "금", "Saturday": "토", "Sunday": "일"}
            rows = []
            for day in day_order:
                count = weekday_counter.get(day, 0)
                bar = "█" * count
                rows.append([day_kr.get(day, day), count, bar])
            print(tabulate(rows, headers=["요일", "게시 수", ""], tablefmt="simple_outline"))


def print_summary(profile: dict, posts: list, user_insights: dict):
    print("\n" + "=" * 60)
    print("8. 종합 요약")
    print("=" * 60)

    active = [p for p in posts if p.get("media_type") != "REPOST_FACADE"]
    total_posts = len(active)
    followers = user_insights.get("followers_count", 0)

    total_likes = sum((p.get("insights") or {}).get("likes", 0) for p in active)
    total_views = sum((p.get("insights") or {}).get("views", 0) for p in active)
    total_replies = sum((p.get("insights") or {}).get("replies", 0) for p in active)
    total_reposts = sum((p.get("insights") or {}).get("reposts", 0) for p in active)
    total_quotes = sum((p.get("insights") or {}).get("quotes", 0) for p in active)

    avg_likes = total_likes / total_posts if total_posts else 0
    avg_views = total_views / total_posts if total_posts else 0
    engagement_rate = (
        (total_likes + total_replies + total_reposts + total_quotes) / total_views * 100
        if total_views else 0
    )

    rows = [
        ["계정", f"@{profile.get('username', '-')}"],
        ["팔로워", f"{followers:,}명" if followers else "-"],
        ["총 게시물", f"{total_posts}개"],
        ["총 조회수", f"{total_views:,}"],
        ["총 좋아요", f"{total_likes:,}"],
        ["총 답글", f"{total_replies:,}"],
        ["총 리포스트", f"{total_reposts:,}"],
        ["평균 조회수/게시물", f"{avg_views:,.0f}"],
        ["평균 좋아요/게시물", f"{avg_likes:.1f}"],
        ["인게이지먼트율", f"{engagement_rate:.2f}%"],
    ]
    print(tabulate(rows, headers=["지표", "값"], tablefmt="simple_outline"))


def save_raw_data(profile, posts, user_insights, demographics, partial: bool = False) -> str:
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now(JST).strftime("%Y%m%d_%H%M%S")
    filepath = os.path.join(output_dir, f"analysis_{timestamp}.json")

    export = {
        "analyzed_at": datetime.now(JST).isoformat(),
        # --max-posts로 일부만 받은 스냅샷은 완전 수집본과 섞이면 안 된다. 종단 분석이
        # 게시물 급감으로 오독하고, export_excel의 기본 입력으로도 잡히기 때문이다.
        "partial": bool(partial),
        "profile": profile,
        "posts": posts,
        "user_insights": user_insights,
        "follower_demographics": demographics,
    }

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(export, f, ensure_ascii=False, indent=2)

    print(f"\n원본 데이터 저장: {filepath}")
    return filepath


def save_archive(profile, posts, user_insights, source_path: str = None) -> str:
    """Build text 저장함 under output/저장함/."""
    from archive import save_text_archive

    snap = save_text_archive(
        profile=profile or {},
        posts=posts or [],
        user_insights=user_insights or {},
        source_path=source_path,
    )
    print(f"텍스트 저장함: {snap}")
    print(f"  → 최신: {os.path.join(os.path.dirname(__file__), 'output', '저장함', 'latest')}")
    return snap


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Threads 계정 종합 분석")
    parser.add_argument(
        "--refresh-insights",
        action="store_true",
        help="인사이트 캐시 TTL 무시, 전체 재조회",
    )
    parser.add_argument(
        "--ttl-days",
        type=int,
        default=None,
        help="캐시 TTL(일). 미지정 시 env INSIGHTS_CACHE_TTL_DAYS 또는 7",
    )
    parser.add_argument(
        "--max-posts",
        type=int,
        default=None,
        help="수집 게시물 상한 (테스트·빠른 샘플용)",
    )
    parser.add_argument(
        "--skip-demographics",
        action="store_true",
        help="팔로워 인구통계 조회 생략",
    )
    parser.add_argument(
        "--fail-on-api-error",
        action="store_true",
        help="API 오류/타임아웃이 1건 이상이면 exit code 2",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=5,
        help="인사이트 병렬 워커 수 (기본 5)",
    )
    parser.add_argument(
        "--skip-archive",
        action="store_true",
        help="텍스트 저장함(output/저장함) 생성 생략",
    )
    parser.add_argument(
        "--export-excel",
        action="store_true",
        help="분석 직후 export_excel.py 자동 실행",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    reset_api_stats()
    validate_token()
    warn_token_expiry()

    print("=" * 60)
    print("  Threads 계정 종합 분석")
    print(f"  실행 시각: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S')} (JST)")
    if args.refresh_insights:
        print("  옵션: --refresh-insights")
    if args.ttl_days is not None:
        print(f"  옵션: --ttl-days={args.ttl_days}")
    if args.max_posts is not None:
        print(f"  옵션: --max-posts={args.max_posts}")
    print("=" * 60)

    profile = fetch_profile()
    if not profile:
        print("[ERROR] 프로필 조회 실패. 토큰/권한을 확인하세요.")
        sys.exit(1)

    posts = fetch_all_posts(max_posts=args.max_posts)
    if not posts:
        print("[ERROR] 게시물이 없거나 조회에 실패했습니다.")
        if args.fail_on_api_error or _api_stats["error"] or _api_stats["timeout"]:
            sys.exit(2)
        sys.exit(1)

    enriched_posts = fetch_post_insights(
        posts,
        ttl_days=args.ttl_days,
        force_refresh=args.refresh_insights,
        max_workers=max(1, args.workers),
    )
    user_insights = fetch_user_insights()
    demographics = {} if args.skip_demographics else fetch_follower_demographics()

    print_post_ranking(enriched_posts)
    print_content_analysis(enriched_posts)
    print_summary(profile, enriched_posts, user_insights)
    json_path = save_raw_data(
        profile, enriched_posts, user_insights, demographics,
        partial=args.max_posts is not None,
    )
    if args.max_posts is not None:
        print(f"  [주의] --max-posts={args.max_posts} 부분 수집본입니다. "
              "종단 분석과 export_excel 기본 입력에서 제외됩니다.")

    if not args.skip_archive:
        save_archive(profile, enriched_posts, user_insights, source_path=json_path)

    if args.export_excel:
        try:
            from export_excel import main as export_main
            export_main(["-i", json_path])
        except Exception as exc:
            print(f"[WARN] Excel 자동 생성 실패: {exc}")

    print(
        f"\nAPI 통계: ok={_api_stats['ok']} error={_api_stats['error']} "
        f"timeout={_api_stats['timeout']} retry={_api_stats['retry']}"
    )
    print("\n분석 완료!")

    if args.fail_on_api_error and (_api_stats["error"] or _api_stats["timeout"]):
        print("[ERROR] API 오류가 발생했습니다 (--fail-on-api-error).")
        sys.exit(2)


if __name__ == "__main__":
    main()
