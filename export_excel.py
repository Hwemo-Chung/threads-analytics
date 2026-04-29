import json
import os
import sys
from datetime import datetime, timedelta, timezone
from collections import Counter

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

JST = timezone(timedelta(hours=9))

# Locale-independent weekday names (strftime %A is locale-dependent → KeyError on Korean/Japanese systems)
_WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _weekday_en(dt):
    """Return English weekday name regardless of system locale."""
    return _WEEKDAY_NAMES[dt.weekday()]
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
SUB_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
HIGHLIGHT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def style_header(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def auto_width(ws, min_width=10, max_width=60):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        length = max((len(str(c.value or "")) for c in col_cells), default=0)
        ws.column_dimensions[col_letter].width = min(max(length + 2, min_width), max_width)


def parse_ts(ts_str):
    if not ts_str:
        return None
    try:
        return datetime.fromisoformat(ts_str.replace("+0000", "+00:00")).astimezone(JST)
    except ValueError:
        return None


def load_data(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def sheet_all_posts(wb, posts):
    ws = wb.create_sheet("전체 게시물")
    headers = [
        "번호", "작성일시(JST)", "미디어타입", "내용(100자)",
        "조회수", "좋아요", "답글", "리포스트", "인용", "공유",
        "인게이지먼트", "좋아요율(%)", "링크"
    ]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    sorted_posts = sorted(posts, key=lambda p: p.get("timestamp") or "", reverse=True)

    for i, p in enumerate(sorted_posts, 1):
        ins = (p.get("insights") or {})
        views = ins.get("views", 0)
        likes = ins.get("likes", 0)
        replies = ins.get("replies", 0)
        reposts = ins.get("reposts", 0)
        quotes = ins.get("quotes", 0)
        shares = ins.get("shares", 0)
        engagement = likes + replies + reposts + quotes
        like_rate = (likes / views * 100) if views > 0 else 0

        dt = parse_ts(p.get("timestamp"))
        dt_str = dt.strftime("%Y-%m-%d %H:%M") if dt else ""
        text = (p.get("text") or "").replace("\n", " ")[:100]

        ws.append([
            i, dt_str, p.get("media_type", ""), text,
            views, likes, replies, reposts, quotes, shares,
            engagement, round(like_rate, 2), p.get("permalink", "")
        ])

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    auto_width(ws)


def sheet_ranking(wb, posts, followers):
    ws = wb.create_sheet("인사이트 랭킹")

    def add_section(title, sorted_list, start_row):
        ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=13)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=10)

        headers = ["순위", "내용(50자)", "작성일", "조회수", "좋아요", "답글", "리포스트", "인게이지먼트", "좋아요율(%)", "링크"]
        for c, h in enumerate(headers, 1):
            ws.cell(row=start_row + 1, column=c, value=h)
        style_header(ws, start_row + 1, len(headers))

        for i, p in enumerate(sorted_list[:30], 1):
            ins = (p.get("insights") or {})
            views = ins.get("views", 0)
            likes = ins.get("likes", 0)
            replies = ins.get("replies", 0)
            reposts = ins.get("reposts", 0)
            engagement = likes + replies + reposts + ins.get("quotes", 0)
            like_rate = (likes / views * 100) if views > 0 else 0
            dt = parse_ts(p.get("timestamp"))
            dt_str = dt.strftime("%Y-%m-%d") if dt else ""
            text = (p.get("text") or "").replace("\n", " ")[:50]
            row = start_row + 2 + i - 1
            ws.append([])
            for c, v in enumerate([i, text, dt_str, views, likes, replies, reposts, engagement, round(like_rate, 2), p.get("permalink", "")], 1):
                ws.cell(row=row, column=c, value=v)
            if i <= 3:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=row, column=c).fill = HIGHLIGHT_FILL

        return start_row + 2 + min(len(sorted_list), 30) + 2

    active_posts = [p for p in posts if p.get("media_type") != "REPOST_FACADE"]

    by_views = sorted(active_posts, key=lambda p: (p.get("insights") or {}).get("views", 0), reverse=True)
    row = add_section("📊 조회수 TOP 30", by_views, 1)

    by_engagement = sorted(active_posts, key=lambda p: sum([
        (p.get("insights") or {}).get("likes", 0),
        (p.get("insights") or {}).get("replies", 0),
        (p.get("insights") or {}).get("reposts", 0),
        (p.get("insights") or {}).get("quotes", 0),
    ]), reverse=True)
    row = add_section("❤️ 인게이지먼트 TOP 30", by_engagement, row)

    min_views = 500
    qualified = [p for p in active_posts if (p.get("insights") or {}).get("views", 0) >= min_views]
    by_like_rate = sorted(qualified, key=lambda p: (p.get("insights") or {}).get("likes", 0) / max((p.get("insights") or {}).get("views", 1), 1), reverse=True)
    row = add_section(f"🎯 좋아요율 TOP 30 (조회 {min_views}+ 기준)", by_like_rate, row)

    by_viral = sorted(active_posts, key=lambda p: (p.get("insights") or {}).get("reposts", 0) + (p.get("insights") or {}).get("quotes", 0), reverse=True)
    add_section("🔄 바이럴(리포스트+인용) TOP 30", by_viral, row)

    auto_width(ws)


def sheet_time_analysis(wb, posts):
    ws = wb.create_sheet("시간대 분석")
    active_posts = [p for p in posts if p.get("media_type") != "REPOST_FACADE"]

    hour_data = {}
    weekday_data = {}
    day_kr = {"Monday": "월", "Tuesday": "화", "Wednesday": "수", "Thursday": "목", "Friday": "금", "Saturday": "토", "Sunday": "일"}

    for p in active_posts:
        dt = parse_ts(p.get("timestamp"))
        if not dt:
            continue
        ins = (p.get("insights") or {})
        views = ins.get("views", 0)
        likes = ins.get("likes", 0)
        engagement = likes + ins.get("replies", 0) + ins.get("reposts", 0)

        h = dt.hour
        if h not in hour_data:
            hour_data[h] = {"count": 0, "views": 0, "likes": 0, "engagement": 0}
        hour_data[h]["count"] += 1
        hour_data[h]["views"] += views
        hour_data[h]["likes"] += likes
        hour_data[h]["engagement"] += engagement

        wd = _weekday_en(dt)
        if wd not in weekday_data:
            weekday_data[wd] = {"count": 0, "views": 0, "likes": 0, "engagement": 0}
        weekday_data[wd]["count"] += 1
        weekday_data[wd]["views"] += views
        weekday_data[wd]["likes"] += likes
        weekday_data[wd]["engagement"] += engagement

    ws.cell(row=1, column=1, value="시간대별 성과 (JST)").font = Font(bold=True, size=13)
    headers = ["시간", "게시수", "총조회수", "평균조회수", "총좋아요", "평균좋아요", "평균인게이지먼트"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c, value=h)
    style_header(ws, 2, len(headers))

    best_hour_avg_views = 0
    for h in range(24):
        d = hour_data.get(h, {"count": 0, "views": 0, "likes": 0, "engagement": 0})
        cnt = d["count"] or 1
        avg_views = d["views"] / cnt
        if avg_views > best_hour_avg_views:
            best_hour_avg_views = avg_views
        ws.cell(row=3 + h, column=1, value=f"{h:02d}:00")
        ws.cell(row=3 + h, column=2, value=d["count"])
        ws.cell(row=3 + h, column=3, value=d["views"])
        ws.cell(row=3 + h, column=4, value=round(avg_views))
        ws.cell(row=3 + h, column=5, value=d["likes"])
        ws.cell(row=3 + h, column=6, value=round(d["likes"] / cnt, 1))
        ws.cell(row=3 + h, column=7, value=round(d["engagement"] / cnt, 1))

    for h in range(24):
        d = hour_data.get(h, {"count": 0, "views": 0})
        cnt = d["count"] or 1
        if d["views"] / cnt >= best_hour_avg_views * 0.8:
            for c in range(1, len(headers) + 1):
                ws.cell(row=3 + h, column=c).fill = HIGHLIGHT_FILL

    row_start = 28
    ws.cell(row=row_start, column=1, value="요일별 성과").font = Font(bold=True, size=13)
    for c, h in enumerate(headers, 1):
        ws.cell(row=row_start + 1, column=c, value=h.replace("시간", "요일"))
    style_header(ws, row_start + 1, len(headers))

    for i, day in enumerate(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]):
        d = weekday_data.get(day, {"count": 0, "views": 0, "likes": 0, "engagement": 0})
        cnt = d["count"] or 1
        r = row_start + 2 + i
        ws.cell(row=r, column=1, value=day_kr.get(day, day))
        ws.cell(row=r, column=2, value=d["count"])
        ws.cell(row=r, column=3, value=d["views"])
        ws.cell(row=r, column=4, value=round(d["views"] / cnt))
        ws.cell(row=r, column=5, value=d["likes"])
        ws.cell(row=r, column=6, value=round(d["likes"] / cnt, 1))
        ws.cell(row=r, column=7, value=round(d["engagement"] / cnt, 1))

    auto_width(ws)


def sheet_demographics(wb, demographics, followers):
    ws = wb.create_sheet("팔로워 인구통계")
    label_map = {"country": "국가", "city": "도시", "gender": "성별", "age": "연령대"}
    row = 1

    ws.cell(row=row, column=1, value=f"총 팔로워: {followers:,}명").font = Font(bold=True, size=13)
    row += 2

    for key in ["country", "city", "gender", "age"]:
        data = demographics.get(key, {})
        if not data or not isinstance(data, dict):
            continue
        ws.cell(row=row, column=1, value=label_map.get(key, key)).font = Font(bold=True, size=12)
        row += 1
        headers = [label_map.get(key, key), "팔로워수", "비율(%)"]
        for c, h in enumerate(headers, 1):
            ws.cell(row=row, column=c, value=h)
        style_header(ws, row, len(headers))
        row += 1

        total = sum(data.values())
        for k, v in sorted(data.items(), key=lambda x: x[1], reverse=True):
            pct = (v / total * 100) if total > 0 else 0
            ws.cell(row=row, column=1, value=k)
            ws.cell(row=row, column=2, value=v)
            ws.cell(row=row, column=3, value=round(pct, 1))
            row += 1
        row += 2

    auto_width(ws)


def sheet_insights_report(wb, posts, user_insights, followers, username=""):
    ws = wb.create_sheet("성장 인사이트")
    active = [p for p in posts if p.get("media_type") != "REPOST_FACADE"]

    total_views = sum((p.get("insights") or {}).get("views", 0) for p in active)
    total_likes = sum((p.get("insights") or {}).get("likes", 0) for p in active)
    total_replies = sum((p.get("insights") or {}).get("replies", 0) for p in active)
    total_reposts = sum((p.get("insights") or {}).get("reposts", 0) for p in active)

    type_stats = {}
    for p in active:
        mt = p.get("media_type", "UNKNOWN")
        ins = (p.get("insights") or {})
        if mt not in type_stats:
            type_stats[mt] = {"count": 0, "views": 0, "likes": 0, "engagement": 0}
        type_stats[mt]["count"] += 1
        type_stats[mt]["views"] += ins.get("views", 0)
        eng = ins.get("likes", 0) + ins.get("replies", 0) + ins.get("reposts", 0) + ins.get("quotes", 0)
        type_stats[mt]["engagement"] += eng

    viral_posts = [p for p in active if (p.get("insights") or {}).get("views", 0) >= 10000]
    high_engage = [p for p in active if (p.get("insights") or {}).get("views", 0) >= 500]
    high_engage_sorted = sorted(high_engage,
        key=lambda p: ((p.get("insights") or {}).get("likes", 0) + (p.get("insights") or {}).get("replies", 0)) / max((p.get("insights") or {}).get("views", 1), 1),
        reverse=True)

    row = 1
    ws.cell(row=row, column=1, value=f"@{username} 성장 인사이트 리포트").font = Font(bold=True, size=14)
    ws.cell(row=row + 1, column=1, value=f"분석일: {datetime.now(JST).strftime('%Y-%m-%d')} / 팔로워: {followers:,}명 / 게시물: {len(active)}개 (리포스트 제외)")
    row += 3

    ws.cell(row=row, column=1, value="1. 핵심 지표").font = Font(bold=True, size=12)
    row += 1
    kpi = [
        ["총 조회수", f"{total_views:,}"],
        ["총 좋아요", f"{total_likes:,}"],
        ["총 답글", f"{total_replies:,}"],
        ["총 리포스트", f"{total_reposts:,}"],
        ["평균 조회수/게시물", f"{total_views // max(len(active), 1):,}"],
        ["평균 좋아요/게시물", f"{total_likes / max(len(active), 1):.1f}"],
        ["전체 인게이지먼트율", f"{(total_likes + total_replies + total_reposts) / max(total_views, 1) * 100:.2f}%"],
        ["바이럴 게시물 (1만+조회)", f"{len(viral_posts)}개 ({len(viral_posts)/max(len(active), 1)*100:.1f}%)"],
        ["30일 조회수", f"{user_insights.get('30d_views', 0):,}"],
    ]
    for label, val in kpi:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=val)
        row += 1
    row += 2

    ws.cell(row=row, column=1, value="2. 미디어 타입별 성과").font = Font(bold=True, size=12)
    row += 1
    headers = ["타입", "게시물수", "총조회수", "평균조회수", "총인게이지먼트", "평균인게이지먼트"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=row, column=c, value=h)
    style_header(ws, row, len(headers))
    row += 1
    for mt, s in sorted(type_stats.items(), key=lambda x: x[1]["engagement"] / max(x[1]["count"], 1), reverse=True):
        cnt = s["count"]
        ws.cell(row=row, column=1, value=mt)
        ws.cell(row=row, column=2, value=cnt)
        ws.cell(row=row, column=3, value=s["views"])
        ws.cell(row=row, column=4, value=round(s["views"] / cnt))
        ws.cell(row=row, column=5, value=s["engagement"])
        ws.cell(row=row, column=6, value=round(s["engagement"] / cnt, 1))
        row += 1
    row += 2

    ws.cell(row=row, column=1, value="3. 좋아요율 TOP 20 (조회 500+ 기준) — 팬이 반응하는 콘텐츠").font = Font(bold=True, size=12)
    row += 1
    headers = ["순위", "내용(60자)", "조회수", "좋아요", "좋아요율(%)", "작성일"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=row, column=c, value=h)
    style_header(ws, row, len(headers))
    row += 1
    for i, p in enumerate(high_engage_sorted[:20], 1):
        ins = (p.get("insights") or {})
        views = ins.get("views", 1)
        likes = ins.get("likes", 0)
        dt = parse_ts(p.get("timestamp"))
        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=(p.get("text") or "").replace("\n", " ")[:60])
        ws.cell(row=row, column=3, value=views)
        ws.cell(row=row, column=4, value=likes)
        ws.cell(row=row, column=5, value=round(likes / max(views, 1) * 100, 2))
        ws.cell(row=row, column=6, value=dt.strftime("%Y-%m-%d") if dt else "")
        row += 1
    row += 2

    ws.cell(row=row, column=1, value="4. 바이럴 게시물 패턴 분석 (1만+ 조회)").font = Font(bold=True, size=12)
    row += 1
    if viral_posts:
        viral_types = Counter(p.get("media_type") for p in viral_posts)
        viral_hours = Counter()
        viral_weekdays = Counter()
        for p in viral_posts:
            dt = parse_ts(p.get("timestamp"))
            if dt:
                viral_hours[dt.hour] += 1
                viral_weekdays[_weekday_en(dt)] += 1
        top_hours = viral_hours.most_common(3)
        top_days = viral_weekdays.most_common(3)
        day_kr = {"Monday": "월", "Tuesday": "화", "Wednesday": "수", "Thursday": "목", "Friday": "금", "Saturday": "토", "Sunday": "일"}

        patterns = [
            ["바이럴 게시물 수", f"{len(viral_posts)}개"],
            ["주요 미디어 타입", ", ".join(f"{k}({v})" for k, v in viral_types.most_common())],
            ["바이럴 잘 되는 시간", ", ".join(f"{h}시({c}건)" for h, c in top_hours)],
            ["바이럴 잘 되는 요일", ", ".join(f"{day_kr.get(d,d)}({c}건)" for d, c in top_days)],
        ]
        for label, val in patterns:
            ws.cell(row=row, column=1, value=label)
            ws.cell(row=row, column=2, value=val)
            row += 1
    row += 2

    ws.cell(row=row, column=1, value="5. 성장 전략 제안").font = Font(bold=True, size=12)
    row += 1
    strategies = [
        "캐러셀 활용 확대 — 평균 인게이지먼트 105.8 (텍스트 22.8의 4.6배). 현재 16개 → 주 1-2회로 늘릴 것",
        "골든타임 집중 — 오전 8-11시(JST)가 게시량·도달 모두 피크. 핵심 콘텐츠는 이 시간대에 게시",
        "바이럴 공식 반복 — 1만+ 조회 게시물의 공통점: 공감형 유머/사회이슈+개인경험. IT·일상·시사 크로스오버",
        "좋아요율 높은 주제 강화 — 좋아요율 TOP 게시물 패턴을 분석해 반복 (좋아요율 = 팬 충성도 지표)",
        "팔로워 인구통계 활용 — 35-44세 남성(한국) 핵심층. 이 타겟에 맞는 콘텐츠 톤 유지",
        "답글 유도 — 인게이지먼트율 1.01%는 양호하나, 질문형 마무리로 답글 비중 높이면 알고리즘 가중치 증가",
        "주말 활동 강화 — 토·일 게시량 감소하나, 경쟁 콘텐츠도 줄어 도달률 높을 가능성",
    ]
    for s in strategies:
        ws.cell(row=row, column=1, value=f"• {s}")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
        row += 1

    auto_width(ws)


def _active_posts(posts):
    return [p for p in posts if p.get("media_type") != "REPOST_FACADE"]


def _post_metrics(post):
    ins = post.get("insights") or {}
    views = ins.get("views", 0)
    likes = ins.get("likes", 0)
    replies = ins.get("replies", 0)
    reposts = ins.get("reposts", 0)
    quotes = ins.get("quotes", 0)
    shares = ins.get("shares", 0)
    return {
        "views": views,
        "likes": likes,
        "replies": replies,
        "reposts": reposts,
        "quotes": quotes,
        "shares": shares,
        "engagement": likes + replies + reposts + quotes,
        "like_rate": (likes / views * 100) if views > 0 else 0,
        "dt": parse_ts(post.get("timestamp")),
        "text": (post.get("text") or "").replace("\n", " "),
        "media_type": post.get("media_type", ""),
        "permalink": post.get("permalink", ""),
    }


def _avg(values):
    return sum(values) / len(values) if values else 0


def _median(values):
    if not values:
        return 0
    sorted_values = sorted(values)
    mid = len(sorted_values) // 2
    if len(sorted_values) % 2 == 1:
        return sorted_values[mid]
    return (sorted_values[mid - 1] + sorted_values[mid]) / 2


def _set_section_title(ws, row, title, end_col=10):
    ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=13)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    return row + 1


def _set_headers(ws, row, headers):
    for c, header in enumerate(headers, 1):
        ws.cell(row=row, column=c, value=header)
    style_header(ws, row, len(headers))
    return row + 1


def _fill_row(ws, row, max_col, fill):
    for c in range(1, max_col + 1):
        ws.cell(row=row, column=c).fill = fill


def sheet_viral_analysis(wb, posts):
    ws = wb.create_sheet("바이럴 심층분석")
    active_posts = _active_posts(posts)
    annotated = [(p, _post_metrics(p)) for p in active_posts]
    viral = [(p, m) for p, m in annotated if m["views"] >= 10000]
    normal = [(p, m) for p, m in annotated if m["views"] < 10000]
    total_views = sum(m["views"] for _, m in annotated)
    overall_viral_ratio = (len(viral) / len(annotated) * 100) if annotated else 0
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    day_kr = {"Monday": "월", "Tuesday": "화", "Wednesday": "수", "Thursday": "목", "Friday": "금", "Saturday": "토", "Sunday": "일"}
    row = 1

    row = _set_section_title(ws, row, "1. 바이럴 게시물 목록")
    headers = ["순위", "작성일", "미디어타입", "내용(80자)", "조회수", "좋아요", "답글", "리포스트", "좋아요율(%)", "링크"]
    row = _set_headers(ws, row, headers)
    for i, (post, m) in enumerate(sorted(viral, key=lambda x: x[1]["views"], reverse=True), 1):
        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=m["dt"].strftime("%Y-%m-%d") if m["dt"] else "")
        ws.cell(row=row, column=3, value=m["media_type"])
        ws.cell(row=row, column=4, value=m["text"][:80])
        ws.cell(row=row, column=5, value=m["views"])
        ws.cell(row=row, column=6, value=m["likes"])
        ws.cell(row=row, column=7, value=m["replies"])
        ws.cell(row=row, column=8, value=m["reposts"])
        ws.cell(row=row, column=9, value=round(m["like_rate"], 2))
        ws.cell(row=row, column=10, value=m["permalink"])
        if i <= 3:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    if not viral:
        ws.cell(row=row, column=1, value="해당 없음")
        row += 1
    row += 1

    row = _set_section_title(ws, row, "2. 바이럴 vs 일반 비교", end_col=8)
    headers = ["구분", "게시물수", "평균조회", "평균좋아요", "평균답글", "평균리포스트", "평균좋아요율(%)", "전체조회비중(%)"]
    row = _set_headers(ws, row, headers)
    compare_groups = [("바이럴(1만+)", viral), ("일반(<1만)", normal)]
    for label, group in compare_groups:
        views_sum = sum(m["views"] for _, m in group)
        likes = [m["likes"] for _, m in group]
        replies = [m["replies"] for _, m in group]
        reposts = [m["reposts"] for _, m in group]
        like_rates = [m["like_rate"] for _, m in group]
        ws.append([
            label,
            len(group),
            round(views_sum / len(group)) if group else 0,
            round(_avg(likes), 1),
            round(_avg(replies), 1),
            round(_avg(reposts), 1),
            round(_avg(like_rates), 2),
            round((views_sum / total_views * 100), 2) if total_views > 0 else 0,
        ])
        if label == "바이럴(1만+)":
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "3. 바이럴 시간대 히트맵", end_col=4)
    headers = ["시간", "바이럴수", "일반수", "바이럴비율(%)"]
    row = _set_headers(ws, row, headers)
    hour_stats = {h: {"viral": 0, "normal": 0} for h in range(24)}
    for _, m in viral:
        if m["dt"]:
            hour_stats[m["dt"].hour]["viral"] += 1
    for _, m in normal:
        if m["dt"]:
            hour_stats[m["dt"].hour]["normal"] += 1
    for hour in range(24):
        v = hour_stats[hour]["viral"]
        n = hour_stats[hour]["normal"]
        ratio = (v / (v + n) * 100) if (v + n) > 0 else 0
        ws.append([f"{hour:02d}시", v, n, round(ratio, 2)])
        if ratio > overall_viral_ratio:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "4. 바이럴 요일분포", end_col=4)
    headers = ["요일", "바이럴수", "일반수", "바이럴비율(%)"]
    row = _set_headers(ws, row, headers)
    weekday_stats = {day: {"viral": 0, "normal": 0} for day in day_order}
    for _, m in viral:
        if m["dt"]:
            weekday_stats[_weekday_en(m["dt"])]["viral"] += 1
    for _, m in normal:
        if m["dt"]:
            weekday_stats[_weekday_en(m["dt"])]["normal"] += 1
    for day in day_order:
        v = weekday_stats[day]["viral"]
        n = weekday_stats[day]["normal"]
        ratio = (v / (v + n) * 100) if (v + n) > 0 else 0
        ws.append([day_kr[day], v, n, round(ratio, 2)])
        if ratio > overall_viral_ratio:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "5. 바이럴 미디어타입", end_col=4)
    headers = ["타입", "바이럴수", "전체수", "바이럴비율(%)"]
    row = _set_headers(ws, row, headers)
    type_stats = {}
    for _, m in annotated:
        mt = m["media_type"]
        if mt not in type_stats:
            type_stats[mt] = {"viral": 0, "total": 0}
        type_stats[mt]["total"] += 1
        if m["views"] >= 10000:
            type_stats[mt]["viral"] += 1
    for media_type, stats in sorted(type_stats.items(), key=lambda x: ((x[1]["viral"] / x[1]["total"]) if x[1]["total"] else 0, x[1]["total"]), reverse=True):
        ratio = (stats["viral"] / stats["total"] * 100) if stats["total"] else 0
        ws.append([media_type, stats["viral"], stats["total"], round(ratio, 2)])
        row += 1

    ws.freeze_panes = "A3"
    auto_width(ws)


def sheet_like_rate_analysis(wb, posts):
    ws = wb.create_sheet("좋아요율 심층분석")
    active_posts = _active_posts(posts)
    qualified = [(p, _post_metrics(p)) for p in active_posts if (p.get("insights") or {}).get("views", 0) >= 500]
    row = 1

    row = _set_section_title(ws, row, "1. 좋아요율 분포", end_col=5)
    headers = ["구간", "게시물수", "비율(%)", "평균조회수", "평균좋아요수"]
    row = _set_headers(ws, row, headers)
    buckets = [
        ("0-1%", 0, 1),
        ("1-2%", 1, 2),
        ("2-3%", 2, 3),
        ("3-5%", 3, 5),
        ("5-10%", 5, 10),
        ("10%+", 10, None),
    ]
    total_qualified = len(qualified)
    for label, low, high in buckets:
        bucket_posts = []
        for post, m in qualified:
            rate = m["like_rate"]
            if high is None:
                matched = rate >= low
            else:
                matched = low <= rate < high
            if matched:
                bucket_posts.append((post, m))
        ws.append([
            label,
            len(bucket_posts),
            round((len(bucket_posts) / total_qualified * 100), 2) if total_qualified else 0,
            round(_avg([m["views"] for _, m in bucket_posts])) if bucket_posts else 0,
            round(_avg([m["likes"] for _, m in bucket_posts]), 1) if bucket_posts else 0,
        ])
        row += 1
    row += 1

    top_posts = sorted(qualified, key=lambda x: x[1]["like_rate"], reverse=True)[:30]
    row = _set_section_title(ws, row, "2. 좋아요율 TOP 30", end_col=9)
    headers = ["순위", "내용(80자)", "작성일", "미디어타입", "조회수", "좋아요", "답글", "좋아요율(%)", "링크"]
    row = _set_headers(ws, row, headers)
    for i, (post, m) in enumerate(top_posts, 1):
        ws.append([
            i,
            m["text"][:80],
            m["dt"].strftime("%Y-%m-%d") if m["dt"] else "",
            m["media_type"],
            m["views"],
            m["likes"],
            m["replies"],
            round(m["like_rate"], 2),
            m["permalink"],
        ])
        if i <= 3:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    row += 1

    worst_posts = sorted(qualified, key=lambda x: x[1]["like_rate"])[:30]
    row = _set_section_title(ws, row, "3. 좋아요율 WORST 30", end_col=9)
    row = _set_headers(ws, row, headers)
    for i, (post, m) in enumerate(worst_posts, 1):
        ws.append([
            i,
            m["text"][:80],
            m["dt"].strftime("%Y-%m-%d") if m["dt"] else "",
            m["media_type"],
            m["views"],
            m["likes"],
            m["replies"],
            round(m["like_rate"], 2),
            m["permalink"],
        ])
        row += 1
    row += 1

    row = _set_section_title(ws, row, "4. 미디어타입별 좋아요율", end_col=4)
    headers = ["타입", "게시물수", "평균좋아요율(%)", "중앙값좋아요율(%)"]
    row = _set_headers(ws, row, headers)
    type_stats = {}
    for _, m in qualified:
        mt = m["media_type"]
        if mt not in type_stats:
            type_stats[mt] = []
        type_stats[mt].append(m["like_rate"])
    for media_type, rates in sorted(type_stats.items(), key=lambda x: _avg(x[1]), reverse=True):
        ws.append([media_type, len(rates), round(_avg(rates), 2), round(_median(rates), 2)])
        row += 1

    ws.freeze_panes = "A3"
    auto_width(ws)


def sheet_content_optimization(wb, posts):
    ws = wb.create_sheet("콘텐츠 최적화")
    active_posts = _active_posts(posts)
    annotated = [(p, _post_metrics(p)) for p in active_posts]
    row = 1

    row = _set_section_title(ws, row, "1. 텍스트 길이별 성과", end_col=7)
    headers = ["구간", "게시물수", "평균조회", "평균좋아요", "평균답글", "평균좋아요율(%)", "판정"]
    row = _set_headers(ws, row, headers)
    text_posts = [(p, m) for p, m in annotated if m["media_type"] == "TEXT_POST"]
    length_buckets = [
        ("~50자", 0, 50),
        ("51-100자", 51, 100),
        ("101-200자", 101, 200),
        ("201-300자", 201, 300),
        ("300자+", 301, None),
    ]
    bucket_rows = []
    for label, low, high in length_buckets:
        bucket = []
        for post, m in text_posts:
            text_len = len((post.get("text") or "").replace("\n", " ").strip())
            if high is None:
                matched = text_len >= low
            else:
                matched = low <= text_len <= high
            if matched:
                bucket.append((post, m))
        bucket_rows.append({
            "label": label,
            "count": len(bucket),
            "avg_views": round(_avg([m["views"] for _, m in bucket])) if bucket else 0,
            "avg_likes": round(_avg([m["likes"] for _, m in bucket]), 1) if bucket else 0,
            "avg_replies": round(_avg([m["replies"] for _, m in bucket]), 1) if bucket else 0,
            "avg_like_rate": round(_avg([m["like_rate"] for _, m in bucket]), 2) if bucket else 0,
        })
    non_zero_rows = [r for r in bucket_rows if r["count"] > 0]
    best_label = max(non_zero_rows, key=lambda x: x["avg_like_rate"])["label"] if non_zero_rows else ""
    worst_label = min(non_zero_rows, key=lambda x: x["avg_like_rate"])["label"] if non_zero_rows else ""
    for bucket in bucket_rows:
        if bucket["label"] == best_label:
            verdict = "🏆 최적"
        elif bucket["label"] == worst_label and bucket["label"] != best_label:
            verdict = "⚠️ 데드존"
        else:
            verdict = "✅ 양호"
        ws.append([
            bucket["label"],
            bucket["count"],
            bucket["avg_views"],
            bucket["avg_likes"],
            bucket["avg_replies"],
            bucket["avg_like_rate"],
            verdict,
        ])
        row += 1
    row += 1

    row = _set_section_title(ws, row, "2. 미디어타입별 상세 비교", end_col=9)
    headers = ["타입", "게시물수", "비중(%)", "평균조회", "평균좋아요", "평균답글", "평균리포스트", "평균인게이지먼트", "평균좋아요율(%)"]
    row = _set_headers(ws, row, headers)
    type_stats = {}
    for _, m in annotated:
        mt = m["media_type"]
        if mt not in type_stats:
            type_stats[mt] = []
        type_stats[mt].append(m)
    sorted_types = sorted(type_stats.items(), key=lambda x: _avg([m["engagement"] for m in x[1]]), reverse=True)
    for idx, (media_type, items) in enumerate(sorted_types, 1):
        ws.append([
            media_type,
            len(items),
            round(len(items) / len(annotated) * 100, 2) if annotated else 0,
            round(_avg([m["views"] for m in items])),
            round(_avg([m["likes"] for m in items]), 1),
            round(_avg([m["replies"] for m in items]), 1),
            round(_avg([m["reposts"] for m in items]), 1),
            round(_avg([m["engagement"] for m in items]), 1),
            round(_avg([m["like_rate"] for m in items]), 2),
        ])
        if idx == 1:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "3. 캐러셀 vs 텍스트 직접 비교", end_col=4)
    headers = ["지표", "캐러셀", "텍스트", "배율(캐러셀/텍스트)"]
    row = _set_headers(ws, row, headers)
    carousel_items = [m for _, m in annotated if m["media_type"] == "CAROUSEL_ALBUM"]
    text_items = [m for _, m in annotated if m["media_type"] == "TEXT_POST"]

    def metric_value(items, field):
        if field == "count":
            return len(items)
        return _avg([m[field] for m in items]) if items else 0

    comparisons = [
        ("게시물수", metric_value(carousel_items, "count"), metric_value(text_items, "count")),
        ("평균조회", metric_value(carousel_items, "views"), metric_value(text_items, "views")),
        ("평균좋아요", metric_value(carousel_items, "likes"), metric_value(text_items, "likes")),
        ("평균답글", metric_value(carousel_items, "replies"), metric_value(text_items, "replies")),
        ("평균리포스트", metric_value(carousel_items, "reposts"), metric_value(text_items, "reposts")),
        ("평균인게이지먼트", metric_value(carousel_items, "engagement"), metric_value(text_items, "engagement")),
        ("평균좋아요율(%)", metric_value(carousel_items, "like_rate"), metric_value(text_items, "like_rate")),
    ]
    for label, carousel_value, text_value in comparisons:
        ratio = "-"
        if text_value:
            ratio = f"{(carousel_value / text_value):.2f}x"
        elif carousel_value:
            ratio = "∞"
        ws.append([
            label,
            round(carousel_value, 2) if label != "게시물수" else int(carousel_value),
            round(text_value, 2) if label != "게시물수" else int(text_value),
            ratio,
        ])
        row += 1
    row += 1

    row = _set_section_title(ws, row, "4. 시간대 × 요일 매트릭스", end_col=8)
    headers = ["시간", "월", "화", "수", "목", "금", "토", "일"]
    row = _set_headers(ws, row, headers)
    matrix = {(hour, day): [] for hour in range(24) for day in range(7)}
    for _, m in annotated:
        if not m["dt"]:
            continue
        matrix[(m["dt"].hour, m["dt"].weekday())].append(m["views"])
    matrix_values = []
    for hour in range(24):
        row_values = [f"{hour:02d}시"]
        for day in range(7):
            avg_views = round(_avg(matrix[(hour, day)])) if matrix[(hour, day)] else 0
            row_values.append(avg_views)
            matrix_values.append(avg_views)
        ws.append(row_values)
        row += 1
    non_zero_matrix = sorted(v for v in matrix_values if v > 0)
    percentile_80 = non_zero_matrix[int((len(non_zero_matrix) - 1) * 0.8)] if non_zero_matrix else 0
    matrix_start_row = row - 24
    for r in range(matrix_start_row, matrix_start_row + 24):
        for c in range(2, 9):
            value = ws.cell(row=r, column=c).value or 0
            if value >= percentile_80 and value > 0:
                ws.cell(row=r, column=c).fill = HIGHLIGHT_FILL

    ws.freeze_panes = "A3"
    auto_width(ws)


def sheet_monthly_trend(wb, posts):
    ws = wb.create_sheet("월별 트렌드")
    active_posts = _active_posts(posts)
    annotated = [(p, _post_metrics(p)) for p in active_posts]
    monthly = {}
    for _, m in annotated:
        if not m["dt"]:
            continue
        month_key = m["dt"].strftime("%Y-%m")
        if month_key not in monthly:
            monthly[month_key] = []
        monthly[month_key].append(m)
    overall_avg_views = _avg([m["views"] for _, m in annotated]) if annotated else 0
    row = 1

    row = _set_section_title(ws, row, "1. 월별 성과 추이", end_col=10)
    headers = ["월", "게시물수", "총조회", "평균조회", "총좋아요", "평균좋아요", "총답글", "인게이지먼트율(%)", "전월대비조회증감(%)", "전월대비좋아요증감(%)"]
    row = _set_headers(ws, row, headers)
    prev_views = None
    prev_likes = None
    month_rows = []
    for month in sorted(monthly):
        items = monthly[month]
        total_views = sum(m["views"] for m in items)
        total_likes = sum(m["likes"] for m in items)
        total_replies = sum(m["replies"] for m in items)
        total_engagement = sum(m["engagement"] for m in items)
        avg_views = round(total_views / len(items)) if items else 0
        avg_likes = round(total_likes / len(items), 1) if items else 0
        view_growth = "-" if prev_views in (None, 0) else round((total_views - prev_views) / prev_views * 100, 2)
        like_growth = "-" if prev_likes in (None, 0) else round((total_likes - prev_likes) / prev_likes * 100, 2)
        month_rows.append((month, avg_views))
        ws.append([
            month,
            len(items),
            total_views,
            avg_views,
            total_likes,
            avg_likes,
            total_replies,
            round((total_engagement / total_views * 100), 2) if total_views else 0,
            view_growth,
            like_growth,
        ])
        if avg_views > overall_avg_views:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        prev_views = total_views
        prev_likes = total_likes
        row += 1
    row += 1

    row = _set_section_title(ws, row, "2. 분기별 요약", end_col=5)
    headers = ["분기", "게시물수", "평균조회", "평균좋아요", "인게이지먼트율(%)"]
    row = _set_headers(ws, row, headers)
    quarterly = {}
    for _, m in annotated:
        if not m["dt"]:
            continue
        quarter = (m["dt"].month - 1) // 3 + 1
        key = f"{m['dt'].year} Q{quarter}"
        if key not in quarterly:
            quarterly[key] = []
        quarterly[key].append(m)
    for quarter_key in sorted(quarterly):
        items = quarterly[quarter_key]
        total_views = sum(m["views"] for m in items)
        total_engagement = sum(m["engagement"] for m in items)
        ws.append([
            quarter_key,
            len(items),
            round(_avg([m["views"] for m in items])),
            round(_avg([m["likes"] for m in items]), 1),
            round((total_engagement / total_views * 100), 2) if total_views else 0,
        ])
        row += 1
    row += 1

    row = _set_section_title(ws, row, "3. 성장 구간 분석", end_col=5)
    headers = ["구간명", "기간", "게시물수", "평균조회", "평균좋아요", "인게이지먼트율(%)"]
    row = _set_headers(ws, row, headers)
    phases = [
        ("초기", "2024-09~2024-12", "2024-09", "2024-12"),
        ("성장기", "2025-01~2025-06", "2025-01", "2025-06"),
        ("변동기", "2025-07~2025-12", "2025-07", "2025-12"),
        ("현재", "2026-01~", "2026-01", None),
    ]
    for phase_name, label, start_month, end_month in phases:
        items = []
        for _, m in annotated:
            if not m["dt"]:
                continue
            month_key = m["dt"].strftime("%Y-%m")
            if month_key < start_month:
                continue
            if end_month and month_key > end_month:
                continue
            items.append(m)
        total_views = sum(m["views"] for m in items)
        total_engagement = sum(m["engagement"] for m in items)
        ws.append([
            phase_name,
            label,
            len(items),
            round(_avg([m["views"] for m in items])) if items else 0,
            round(_avg([m["likes"] for m in items]), 1) if items else 0,
            round((total_engagement / total_views * 100), 2) if total_views else 0,
        ])
        row += 1

    ws.freeze_panes = "A3"
    auto_width(ws)


def sheet_growth_strategy(wb, posts, demographics, followers):
    ws = wb.create_sheet("10만 성장전략")
    active_posts = _active_posts(posts)
    annotated = [(p, _post_metrics(p)) for p in active_posts]
    total_views = sum(m["views"] for _, m in annotated)
    total_engagement = sum(m["engagement"] for _, m in annotated)
    engagement_rate = (total_engagement / total_views * 100) if total_views else 0
    viral_ratio = (sum(1 for _, m in annotated if m["views"] >= 10000) / len(annotated) * 100) if annotated else 0
    carousel_ratio = (sum(1 for _, m in annotated if m["media_type"] == "CAROUSEL_ALBUM") / len(annotated) * 100) if annotated else 0
    reply_post_ratio = (sum(1 for _, m in annotated if m["replies"] > 0) / len(annotated) * 100) if annotated else 0
    avg_like_rate = _avg([m["like_rate"] for _, m in annotated]) if annotated else 0
    hour_counter = Counter(m["dt"].hour for _, m in annotated if m["dt"])
    top_hours = sorted(hour_counter.items(), key=lambda x: x[1], reverse=True)[:2]
    crowded_hours = ",".join(f"{h:02d}시" for h, _ in top_hours) if top_hours else "데이터 부족"

    value_keywords = ["팁", "인사이트", "노하우", "정리", "가이드", "분석", "조언", "배운", "방법", "공식", "리뷰", "교육"]
    engagement_keywords = ["?", "어떻게", "어떤", "여러분", "생각", "의견", "알려", "추천", "맞나요", "인가요", "질문", "고견"]
    cta_keywords = ["링크", "프로필", "팔로우", "구독", "블로그", "뉴스레터", "DM", "문의", "보러", "공유"]
    mix_counter = Counter()
    for post, m in annotated:
        text = (post.get("text") or "").replace("\n", " ")
        if any(keyword in text for keyword in cta_keywords):
            mix_counter["소프트CTA"] += 1
        elif any(keyword in text for keyword in engagement_keywords):
            mix_counter["참여유도"] += 1
        elif m["media_type"] == "CAROUSEL_ALBUM" or any(keyword in text for keyword in value_keywords):
            mix_counter["가치콘텐츠"] += 1
        else:
            mix_counter["개인스토리"] += 1

    def judge_range(value, low, high=None, unit="%"):
        if high is None:
            if value >= low:
                return "🟢 적정"
            if value >= low * 0.7:
                return "🟡 근접"
            return "🔴 부족"
        if low <= value <= high:
            return "🟢 적정"
        margin = max((high - low) * 0.25, 0.5)
        if (low - margin) <= value <= (high + margin):
            return "🟡 근접"
        return "🔴 개선필요"

    deadzone_bucket = "201-300자 데드존"
    text_posts = [(p, m) for p, m in annotated if m["media_type"] == "TEXT_POST"]
    if text_posts:
        candidate_buckets = [
            ("~50자", [m["like_rate"] for p, m in text_posts if len((p.get("text") or "").replace("\n", " ").strip()) <= 50]),
            ("51-100자", [m["like_rate"] for p, m in text_posts if 51 <= len((p.get("text") or "").replace("\n", " ").strip()) <= 100]),
            ("101-200자", [m["like_rate"] for p, m in text_posts if 101 <= len((p.get("text") or "").replace("\n", " ").strip()) <= 200]),
            ("201-300자", [m["like_rate"] for p, m in text_posts if 201 <= len((p.get("text") or "").replace("\n", " ").strip()) <= 300]),
            ("300자+", [m["like_rate"] for p, m in text_posts if len((p.get("text") or "").replace("\n", " ").strip()) >= 301]),
        ]
        non_empty = [(label, _avg(rates)) for label, rates in candidate_buckets if rates]
        if non_empty:
            deadzone_bucket = f"{min(non_empty, key=lambda x: x[1])[0]} 데드존"

    row = 1
    row = _set_section_title(ws, row, "1. 현재 위치 진단", end_col=4)
    headers = ["지표", "현재값", "10만계정기준", "판정"]
    row = _set_headers(ws, row, headers)
    diagnosis_rows = [
        ["팔로워", f"{followers:,}", "100,000", f"🔴 {followers / 100000 * 100:.1f}%"],
        ["인게이지먼트율", f"{engagement_rate:.2f}%", "2-5%", judge_range(engagement_rate, 2, 5)],
        ["바이럴비율(1만+조회)", f"{viral_ratio:.2f}%", "5-10%", judge_range(viral_ratio, 5, 10)],
        ["캐러셀비중", f"{carousel_ratio:.2f}%", "10-20%", judge_range(carousel_ratio, 10, 20)],
        ["답글있는게시물", f"{reply_post_ratio:.2f}%", "70%+", judge_range(reply_post_ratio, 70)],
        ["평균좋아요율", f"{avg_like_rate:.2f}%", "3-5%", judge_range(avg_like_rate, 3, 5)],
    ]
    for values in diagnosis_rows:
        ws.append(values)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "2. 즉시 실행 TOP 5", end_col=5)
    headers = ["우선순위", "액션", "현재값", "목표값", "예상효과"]
    row = _set_headers(ws, row, headers)
    actions = [
        [1, "캐러셀 주2회", f"현재 {carousel_ratio:.1f}%", "10-15%", "인게이지먼트 4.9배"],
        [2, "게시시간 13-14시,21시 분산", f"{crowded_hours} 과밀", "13-14시,21시", "평균조회 2-8배"],
        [3, "100-200자+질문마무리", deadzone_bucket, "100-200자", "좋아요율 4배"],
        [4, "일일10개 사려깊은 답글", f"답글 {reply_post_ratio:.1f}%", "70%+", "팔로워전환 가속"],
        [5, "바이오 리뉴얼", "불명확", "뭘얻는지명확", "전환율 2-3배"],
    ]
    for values in actions:
        ws.append(values)
        if values[0] == 1:
            _fill_row(ws, row, len(headers), HIGHLIGHT_FILL)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "3. 10만 로드맵", end_col=5)
    headers = ["단계", "기간", "목표팔로워", "핵심액션", "KPI"]
    row = _set_headers(ws, row, headers)
    roadmap = [
        ["Phase 1", "1-2개월", "3000", "프로필최적화+캐러셀주2+시간조정", "팔로워주100+증가"],
        ["Phase 2", "3-4개월", "8000", "질문형마무리+답글루틴+참여팟", "인게이지먼트율2%+"],
        ["Phase 3", "5-8개월", "25000", "콘텐츠기둥3개+바이럴공식반복+인스타크로스", "월평균조회5000+"],
        ["Phase 4", "9-18개월", "100000", "권위구축+협업+데이터기반재활용", "바이럴비율5%+"],
    ]
    for values in roadmap:
        ws.append(values)
        row += 1
    row += 1

    row = _set_section_title(ws, row, "4. 콘텐츠 믹스 공식", end_col=5)
    headers = ["비중", "유형", "설명", "현재", "목표"]
    row = _set_headers(ws, row, headers)
    mix_rows = [
        ["40%", "가치콘텐츠", "팁/인사이트/교육", f"{(mix_counter['가치콘텐츠'] / len(annotated) * 100):.1f}%" if annotated else "0%", "40%"],
        ["30%", "개인스토리", "여정/비하인드", f"{(mix_counter['개인스토리'] / len(annotated) * 100):.1f}%" if annotated else "0%", "30%"],
        ["20%", "참여유도", "질문/투표/오픈스레드", f"{(mix_counter['참여유도'] / len(annotated) * 100):.1f}%" if annotated else "0%", "20%"],
        ["10%", "소프트CTA", "링크/전환", f"{(mix_counter['소프트CTA'] / len(annotated) * 100):.1f}%" if annotated else "0%", "10%"],
    ]
    for values in mix_rows:
        ws.append(values)
        row += 1

    ws.freeze_panes = "A3"
    auto_width(ws)


def main():
    import glob as _glob

    output_dir = os.path.join(os.path.dirname(__file__), "output")
    pattern = os.path.join(output_dir, "analysis_*.json")
    json_files = sorted(_glob.glob(pattern))

    if not json_files:
        print("[ERROR] output/ 디렉토리에 analysis_*.json 파일이 없습니다.")
        print("  먼저 analyze.py를 실행해주세요.")
        sys.exit(1)

    input_path = json_files[-1]  # 가장 최신 파일 사용
    timestamp = datetime.now(JST).strftime("%Y%m%d")
    output_path = os.path.join(output_dir, f"threads_analysis_{timestamp}.xlsx")

    data = load_data(input_path)
    posts = data.get("posts", [])
    user_insights = data.get("user_insights", {})
    demographics = data.get("follower_demographics", {})
    followers = user_insights.get("followers_count", 0)

    username = data.get("profile", {}).get("username", "unknown")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_all_posts(wb, posts)
    sheet_ranking(wb, posts, followers)
    sheet_time_analysis(wb, posts)
    sheet_demographics(wb, demographics, followers)
    sheet_insights_report(wb, posts, user_insights, followers, username)
    sheet_viral_analysis(wb, posts)
    sheet_like_rate_analysis(wb, posts)
    sheet_content_optimization(wb, posts)
    sheet_monthly_trend(wb, posts)
    sheet_growth_strategy(wb, posts, demographics, followers)

    wb.save(output_path)
    print(f"Excel 저장 완료: {output_path}")
    print(f"시트: {wb.sheetnames}")


if __name__ == "__main__":
    main()
