"""
Threads 계정 종합 분석 스크립트

사용법: python analyze.py
사전 조건: auth.py 실행하여 .env에 ACCESS_TOKEN 저장 완료
"""

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

JST = timezone(timedelta(hours=9))


def api_get(endpoint: str, params: dict = None) -> dict:
    params = params or {}
    params["access_token"] = ACCESS_TOKEN
    try:
        resp = requests.get(f"{API_BASE}/{endpoint}", params=params, timeout=30)
    except requests.exceptions.Timeout:
        print(f"  [TIMEOUT] {endpoint}: 30초 초과")
        return {}
    if resp.status_code != 200:
        print(f"  [API ERROR] {endpoint}: {resp.status_code} - {resp.text[:200]}")
        return {}
    return resp.json()


def validate_token():
    if not ACCESS_TOKEN or ACCESS_TOKEN == "":
        print("[ERROR] ACCESS_TOKEN이 없습니다. 먼저 auth.py를 실행해주세요.")
        sys.exit(1)


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


def fetch_all_posts() -> list:
    print("\n2. 게시물 전체 조회 중...")
    all_posts = []
    fields = "id,media_type,text,timestamp,permalink,like_count,replies_count"

    params = {"fields": fields, "limit": 100}
    page = 1

    while True:
        data = api_get(f"{USER_ID}/threads", params)
        posts = data.get("data", [])
        if not posts:
            break

        all_posts.extend(posts)
        print(f"  페이지 {page}: {len(posts)}개 로드 (총 {len(all_posts)}개)")

        cursors = data.get("paging", {}).get("cursors", {})
        after = cursors.get("after")
        if not after or len(posts) < 100:
            break

        params = {"fields": fields, "limit": 100, "after": after}
        page += 1
        time.sleep(0.5)

    print(f"  총 {len(all_posts)}개 게시물 조회 완료")
    return all_posts


def fetch_post_insights(posts: list) -> list:
    print("\n3. 게시물별 인사이트 수집 중 (5 병렬, 배치 저장)...")
    total = len(posts)
    metrics = "views,likes,replies,reposts,quotes,shares"
    cache_path = os.path.join(os.path.dirname(__file__), "output", "insights_cache.json")
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)

    cached = {}
    if os.path.exists(cache_path):
        with open(cache_path, "r") as f:
            cached = json.load(f)
        print(f"  캐시에서 {len(cached)}개 로드됨")

    to_fetch = [p for p in posts if p["id"] not in cached]
    print(f"  신규 조회 필요: {len(to_fetch)}개 / 총 {total}개")

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
    for batch_start in range(0, len(to_fetch), BATCH_SIZE):
        batch = to_fetch[batch_start:batch_start + BATCH_SIZE]
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(fetch_single, p): p for p in batch}
            for future in as_completed(futures):
                media_id, insight_map = future.result()
                cached[media_id] = insight_map
                done += 1
                if done % 50 == 0:
                    print(f"  {done}/{len(to_fetch)} 완료")

        with open(cache_path, "w") as f:
            json.dump(cached, f)
        print(f"  배치 저장 ({batch_start + len(batch)}/{len(to_fetch)})")

    print(f"  전체 인사이트 수집 완료: {len(cached)}개")

    enriched = []
    for p in posts:
        p["insights"] = cached.get(p["id"], {})
        enriched.append(p)
    return enriched


def fetch_user_insights() -> dict:
    print("\n4. 계정 인사이트 조회 중...")
    result = {}

    followers_data = api_get(
        f"{USER_ID}/threads_insights",
        {"metric": "followers_count"},
    )
    for item in followers_data.get("data", []):
        if item["name"] == "followers_count":
            result["followers_count"] = item.get("values", [{}])[0].get("value", 0)

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
        values = item.get("values", [])
        total = sum(v.get("value", 0) for v in values)
        result[f"30d_{name}"] = total

    rows = [
        ["팔로워 수", result.get("followers_count", "-")],
        ["조회수 (30일)", f"{result.get('30d_views', 0):,}"],
        ["좋아요 (30일)", f"{result.get('30d_likes', 0):,}"],
        ["답글 (30일)", f"{result.get('30d_replies', 0):,}"],
        ["리포스트 (30일)", f"{result.get('30d_reposts', 0):,}"],
        ["인용 (30일)", f"{result.get('30d_quotes', 0):,}"],
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

    def get_engagement(p):
        ins = p.get("insights", {})
        return ins.get("likes", 0) + ins.get("replies", 0) + ins.get("reposts", 0) + ins.get("quotes", 0)

    sorted_by_engagement = sorted(posts, key=get_engagement, reverse=True)[:15]
    sorted_by_views = sorted(posts, key=lambda p: p.get("insights", {}).get("views", 0), reverse=True)[:15]

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

    type_counter = Counter()
    type_engagement = {}

    for p in posts:
        media_type = p.get("media_type", "UNKNOWN")
        type_counter[media_type] += 1

        ins = p.get("insights", {})
        engagement = ins.get("likes", 0) + ins.get("replies", 0) + ins.get("reposts", 0)
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

    if posts:
        timestamps = []
        for p in posts:
            ts = p.get("timestamp", "")
            if ts:
                try:
                    dt = datetime.fromisoformat(ts.replace("+0000", "+00:00"))
                    timestamps.append(dt)
                except ValueError:
                    pass

        if timestamps:
            hour_counter = Counter(dt.astimezone(JST).hour for dt in timestamps)
            weekday_counter = Counter(dt.astimezone(JST).strftime("%A") for dt in timestamps)

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

    total_posts = len(posts)
    followers = user_insights.get("followers_count", 0)

    total_likes = sum(p.get("insights", {}).get("likes", 0) for p in posts)
    total_views = sum(p.get("insights", {}).get("views", 0) for p in posts)
    total_replies = sum(p.get("insights", {}).get("replies", 0) for p in posts)
    total_reposts = sum(p.get("insights", {}).get("reposts", 0) for p in posts)

    avg_likes = total_likes / total_posts if total_posts else 0
    avg_views = total_views / total_posts if total_posts else 0
    engagement_rate = (total_likes + total_replies + total_reposts) / total_views * 100 if total_views else 0

    rows = [
        ["계정", f"@{profile.get('username', '-')}"],
        ["팔로워", f"{followers:,}명"],
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


def save_raw_data(profile, posts, user_insights, demographics):
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now(JST).strftime("%Y%m%d_%H%M%S")
    filepath = os.path.join(output_dir, f"analysis_{timestamp}.json")

    export = {
        "analyzed_at": datetime.now(JST).isoformat(),
        "profile": profile,
        "posts": posts,
        "user_insights": user_insights,
        "follower_demographics": demographics,
    }

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(export, f, ensure_ascii=False, indent=2)

    print(f"\n원본 데이터 저장: {filepath}")


def main():
    validate_token()

    print("=" * 60)
    print(f"  Threads 계정 종합 분석")
    print(f"  실행 시각: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S')} (JST)")
    print("=" * 60)

    profile = fetch_profile()
    posts = fetch_all_posts()
    enriched_posts = fetch_post_insights(posts)
    user_insights = fetch_user_insights()
    demographics = fetch_follower_demographics()

    print_post_ranking(enriched_posts)
    print_content_analysis(enriched_posts)
    print_summary(profile, enriched_posts, user_insights)
    save_raw_data(profile, enriched_posts, user_insights, demographics)

    print("\n분석 완료!")


if __name__ == "__main__":
    main()
