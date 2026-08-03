"""
Excel 워크북 출력 언어 전환 (ko / en / ja).

호출부를 t()로 감싸지 않는다. 완성된 워크북을 저장 직전에 훑어서
UI 문자열만 사전 정확 일치로 치환한다. 사전에 없는 문자열(게시물 본문,
토픽 태그값 등)은 손대지 않으므로 데이터가 훼손될 수 없다.

동적으로 조립된 문자열(임계값이 박힌 섹션 제목 등)은 정확 일치가 안 되므로
PATTERNS의 정규식으로 따로 처리한다.
"""

import os
import re

LANGS = ("ko", "en", "ja")
EXCEL_SHEET_MAX = 31  # Excel 시트명 상한. 넘기면 openpyxl이 조용히 잘라낸다.


def resolve_lang(explicit=None):
    """--lang > THREADS_LANG > ko"""
    for candidate in (explicit, os.getenv("THREADS_LANG")):
        if candidate:
            value = candidate.strip().lower()
            if value in LANGS:
                return value
            raise ValueError(f"지원하지 않는 언어: {candidate} (ko, en, ja 중 하나)")
    return "ko"


# 시트명은 Excel 31자 제한이 있어 짧게 잡는다.
UI = {
    "en": {
        # --- 시트명 ---
        "전체 게시물": "All Posts",
        "인사이트 랭킹": "Rankings",
        "시간대 분석": "Time of Day",
        "팔로워 인구통계": "Follower Demographics",
        "성장 인사이트": "Growth Insights",
        "바이럴 심층분석": "Viral Deep Dive",
        "좋아요율 심층분석": "Like Rate Deep Dive",
        "콘텐츠 최적화": "Content Optimization",
        "월별 트렌드": "Monthly Trend",
        "일별 조회수 추이": "Daily Views",
        "스냅샷 성장추이": "Snapshot Growth",
        "10만 성장전략": "Path to 100K",

        # --- 섹션 제목 ---
        "1. 핵심 지표": "1. Key Metrics",
        "1. 현재 위치 진단": "1. Where You Stand",
        "1. 좋아요율 분포": "1. Like Rate Distribution",
        "1. 월별 성과 추이": "1. Monthly Performance",
        "1. 스냅샷 이력": "1. Snapshot History",
        "1. 최근 90일 일별 조회수": "1. Daily Views, Last 90 Days",
        "1. 글자수 구간별 성과 (조회수 vs 인게이지먼트율)":
            "1. Performance by Length (Views vs Engagement Rate)",
        "2. 미디어 타입별 성과": "2. Performance by Media Type",
        "2. 미디어타입별 상세 비교": "2. Media Type Comparison",
        "2. 바이럴 vs 일반 비교": "2. Viral vs Normal",
        "2. 분기별 요약": "2. Quarterly Summary",
        "2. 좋아요율 TOP 30": "2. Top 30 by Like Rate",
        "2. 즉시 실행 TOP 5": "2. Top 5 Actions to Take Now",
        "2. 월별 실제 노출 조회수 (노출일 기준)":
            "2. Monthly Delivered Views (by Delivery Date)",
        "2. 구간별 성장률 (결측 스냅샷 구간은 팔로워 계산 제외)":
            "2. Growth by Interval (intervals with missing data excluded)",
        "3. 좋아요율 WORST 30": "3. Bottom 30 by Like Rate",
        "3. 캐러셀 vs 텍스트 직접 비교": "3. Carousel vs Text",
        "3. 바이럴 시간대 히트맵": "3. Viral Heatmap by Hour",
        "3. 요일별 노출량": "3. Delivered Views by Weekday",
        "3. 성장 구간 분석 (시간 분위, 계정 독립)":
            "3. Growth by Era (time quartiles)",
        "3. 성장 로드맵 (현재 팔로워 기준 스케일)":
            "3. Growth Roadmap (scaled to current followers)",
        "3. 좋아요율 TOP 20 (조회 500+ 기준) — 팬이 반응하는 콘텐츠":
            "3. Top 20 by Like Rate (500+ views) — What Fans React To",
        "4. 미디어타입별 좋아요율": "4. Like Rate by Media Type",
        "4. 시간대 × 요일 매트릭스": "4. Hour x Weekday Matrix",
        "4. 바이럴 요일분포": "4. Viral Posts by Weekday",
        "4. 콘텐츠 믹스 공식": "4. Content Mix Formula",
        "4. 요약": "4. Summary",
        "4. 리포스트 비중 추이 (REPOST_FACADE)": "4. Repost Share Over Time",
        "5. 바이럴 미디어타입": "5. Viral Posts by Media Type",
        "5. 성장 전략 제안": "5. Growth Strategy",
        "5. 토픽 태그별 성과": "5. Performance by Topic Tag",
        "6. 조회수 집중도": "6. View Concentration",

        # --- 공통 헤더 ---
        "번호": "No.", "순위": "Rank", "날짜": "Date", "작성일": "Posted",
        "작성일시(JST)": "Posted (JST)", "구분": "Category", "구간": "Range",
        "구간명": "Interval", "기간": "Period", "단계": "Stage", "지표": "Metric",
        "값": "Value", "비중": "Share", "비고": "Note", "연월": "Month",
        "설명": "Description", "해석": "Reading", "판정": "Verdict",
        "유형": "Type", "타입": "Type", "액션": "Action", "목표": "Target",
        "우선순위": "Priority", "링크": "Link", "토픽": "Topic",
        "요일": "Weekday", "시간": "Hour", "일수": "Days", "분기": "Quarter",
        "국가": "Country", "도시": "City", "성별": "Gender", "연령대": "Age",
        "현재": "Current", "현재값": "Current", "목표값": "Target",
        "핵심액션": "Key Action", "데이터 근거": "Evidence",
        "미디어타입": "Media Type", "게시수": "Posts", "게시물수": "Posts",
        "글자수구간": "Length Range", "스냅샷": "Snapshot",
        "관측주수": "Weeks Observed", "게시일후경과": "Days Since Posting",
        "대상게시물수": "Posts Measured", "전체수": "Total",

        # --- 지표 ---
        "조회수": "Views", "총조회": "Total Views", "총조회수": "Total Views",
        "총 조회수": "Total Views", "좋아요": "Likes", "총좋아요": "Total Likes",
        "총 좋아요": "Total Likes", "답글": "Replies", "총답글": "Total Replies",
        "총 답글": "Total Replies", "리포스트": "Reposts",
        "리포스트수": "Reposts", "총 리포스트": "Total Reposts",
        "인용": "Quotes", "공유": "Shares", "팔로워": "Followers",
        "팔로워수": "Followers", "팔로워증가": "Followers Gained",
        "인게이지먼트": "Engagement", "총인게이지먼트": "Total Engagement",
        "인게이지먼트율": "Engagement Rate",
        "인게이지먼트율(%)": "Engagement Rate (%)",
        "전체 인게이지먼트율": "Overall Engagement Rate",
        "좋아요율(%)": "Like Rate (%)",
        "원본게시물": "Original Posts", "원본게시물수": "Original Posts",
        "전체게시물": "All Posts", "답글있는게시물": "Posts With Replies",

        # --- 평균 / 중앙값 ---
        "평균조회": "Avg Views", "평균조회수": "Avg Views",
        "평균좋아요": "Avg Likes", "평균좋아요수": "Avg Likes",
        "평균답글": "Avg Replies", "평균리포스트": "Avg Reposts",
        "평균인게이지먼트": "Avg Engagement", "평균일조회": "Avg Daily Views",
        "평균좋아요율": "Avg Like Rate", "평균좋아요율(%)": "Avg Like Rate (%)",
        "평균 조회수/게시물": "Avg Views / Post",
        "평균 좋아요/게시물": "Avg Likes / Post",
        "중앙값조회": "Median Views", "중앙값좋아요": "Median Likes",
        "중앙값ER(%)": "Median ER (%)",
        "중앙값좋아요율(%)": "Median Like Rate (%)",
        "중앙값일조회": "Median Daily Views",
        "중앙값조회증가": "Median View Gain",
        "평균÷중앙값 배수": "Mean / Median Ratio",
        "1위 제외 평균": "Avg Excluding Top 1",
        "7일이동평균": "7-Day Moving Avg",

        # --- 비율 / 증감 ---
        "비율(%)": "Share (%)", "비중(%)": "Share (%)", "증감(%)": "Change (%)",
        "증가율(%)": "Growth (%)", "전월대비(%)": "MoM (%)",
        "전월대비조회증감(%)": "MoM Views (%)",
        "전월대비좋아요증감(%)": "MoM Likes (%)",
        "전주동요일대비(%)": "vs Same Day Last Week (%)",
        "전체조회비중(%)": "Share of Total Views (%)",
        "총조회비중(%)": "Share of Total Views (%)",
        "상위1% 조회비중(%)": "Top 1% Share of Views (%)",
        "상위5% 조회비중(%)": "Top 5% Share of Views (%)",
        "상위10% 조회비중(%)": "Top 10% Share of Views (%)",
        "최다조회 1건 비중(%)": "Single Best Post Share (%)",
        "데드율(%)": "Dead Rate (%)",
        "바이럴비율(%)": "Viral Rate (%)", "바이럴수": "Viral Posts",
        "바이럴 게시물 수": "Viral Posts", "일반수": "Normal Posts",
        "바이럴기준(적용값)": "Viral Threshold (applied)",
        "캐러셀비중": "Carousel Share",
        "배율(캐러셀/텍스트)": "Ratio (Carousel / Text)",
        "30일조회": "30d Views", "30일좋아요": "30d Likes",
        "30일 조회수": "30d Views",
        "30일조회증감(%)": "30d Views Change (%)",
        "원본1건당팔로워": "Followers per Post",
        "원본게시물증가": "Original Posts Added",
        "일평균팔로워": "Followers / Day",
        "일평균조회": "Views / Day",
        "최근 30일": "Last 30 Days", "직전 30일": "Previous 30 Days",
        "730일 총조회": "Total Views, 730 Days",
        "최고일": "Best Day", "최저일": "Worst Day",
        "스냅샷일시": "Snapshot Time",
        "주요 미디어 타입": "Primary Media Type",
        "성과 상위 길이대": "Best Performing Length",
        "목표팔로워": "Target Followers",
        "10만계정기준": "vs 100K Benchmark",

        # --- 라벨 / 판정 ---
        "🟢우수": "🟢 Strong", "⚪보통": "⚪ Average", "🔴회피": "🔴 Avoid",
        "표본부족": "Too Few Samples", "정상": "OK",
        "✅양호": "✅ Good", "🔴 부족": "🔴 Weak", "🔴 개선필요": "🔴 Needs Work",
        "🏆조회최적": "🏆 Best for Reach", "⚠️조회데드존": "⚠️ Reach Dead Zone",
        "📊 조회수 TOP 30": "📊 Top 30 by Views",
        "❤️ 인게이지먼트 TOP 30": "❤️ Top 30 by Engagement",
        "🔄 바이럴(리포스트+인용) TOP 30": "🔄 Top 30 by Viral (Reposts + Quotes)",
        "텍스트": "Text", "캐러셀": "Carousel",
        "토픽 없음": "No Topic", "기타(n<5)": "Other (n<5)",
        "요일별 성과": "Performance by Weekday",
        "데드율(<500조회,%)": "Dead Rate (<500 views, %)",
        "30일좋아요 데이터 결측": "30d likes missing",
        "팔로워 데이터 결측 / 30일좋아요 데이터 결측":
            "followers missing / 30d likes missing",
        "API 기준 일 경계(UTC-07:00), 게시물 시각(JST)과 다름":
            "API day boundary is UTC-07:00, which differs from post times (JST)",

        # --- 요일 ---
        "월": "Mon", "화": "Tue", "수": "Wed", "목": "Thu",
        "금": "Fri", "토": "Sat", "일": "Sun",

        # --- 글자수 구간 ---
        "0-50자": "0-50 chars", "50-100자": "50-100 chars",
        "100-150자": "100-150 chars", "150-200자": "150-200 chars",
        "200-300자": "200-300 chars", "300-400자": "300-400 chars",
        "400-500자": "400-500 chars", "500자+": "500+ chars",
        "타입 비교 데이터 부족": "Not enough data to compare media types",
        "데이터 없음 (일별 시계열 미수집)": "No data (daily series not collected)",
        "데이터 없음 (topic_tag 미수집 — analyze.py 재실행 필요)":
            "No data (topic_tag not collected — re-run analyze.py)",
        "데이터 없음": "No data",
        "해당 없음": "N/A",
        "🟡 근접": "🟡 Close",

        # --- 경과일 / 분위 / 기간 ---
        "0-7일": "0-7 days", "7-14일": "7-14 days", "14-30일": "14-30 days",
        "30-60일": "30-60 days", "60-90일": "60-90 days", "90일+": "90+ days",
        "1분위(초기)": "Q1 (earliest)", "2분위": "Q2", "3분위": "Q3",
        "4분위(최근)": "Q4 (latest)",
        "1-2개월": "1-2 months", "3-4개월": "3-4 months",
        "5-8개월": "5-8 months", "9-18개월": "9-18 months",

        # --- 전략 문구 ---
        "가치제안 1문장": "One-line value proposition",
        "가치콘텐츠": "Value content",
        "개인스토리": "Personal story",
        "여정/비하인드": "Journey / behind the scenes",
        "팁/인사이트/교육": "Tips / insights / teaching",
        "질문/투표/오픈스레드": "Questions / polls / open threads",
        "참여유도": "Invite participation",
        "소프트CTA": "Soft CTA",
        "링크/전환": "Link / conversion",
        "바이오 명확화": "Clarify bio",
        "프로필 점검": "Review profile",
        "게시시간 분산": "Spread out posting times",
        "캐러셀 주2회": "Carousel twice a week",
        "월평균조회 상위화": "Raise monthly average views",
        "전환 경로 명확화": "Clarify conversion path",
        "일일 사려깊은 답글 루틴": "Daily thoughtful-reply routine",
        "답글 비중↑ → 도달 가중 기대": "More replies -> expect reach weighting",
        "최적 길이+질문마무리": "Optimal length + question ending",
        "프로필최적화+캐러셀주2+시간조정":
            "Optimize profile + carousel 2x/wk + shift timing",
        "질문형마무리+답글루틴+참여팟":
            "Question endings + reply routine + engagement pods",
        "권위구축+협업+데이터기반재활용":
            "Build authority + collaborate + reuse what data says works",
        "콘텐츠기둥3개+바이럴공식반복+크로스채널":
            "3 content pillars + repeat the viral formula + cross-channel",
        "바이럴 잘 되는 시간": "Best hours for viral",
        "바이럴 잘 되는 요일": "Best weekdays for viral",
        "바이럴비율5%+": "Viral rate 5%+",
        "인게이지먼트율2%+": "Engagement rate 2%+",
        "• 좋아요율 높은 주제 강화 — 좋아요율 TOP 게시물 패턴을 분석해 반복 (좋아요율 = 팬 충성도 지표)":
            "• Double down on what fans like — study your top like-rate posts and repeat the pattern "
            "(like rate is a loyalty signal, not a reach signal)",
        "• 주말 활동 강화 — 토·일 게시량이 평일 대비 적음. 경쟁 콘텐츠가 줄어 도달 기회일 수 있음":
            "• Try weekends — you post less on Sat/Sun than on weekdays, and less competing content "
            "can mean more reach",
        "내용(50자)": "Text (50 chars)", "내용(60자)": "Text (60 chars)",
        "내용(80자)": "Text (80 chars)", "내용(100자)": "Text (100 chars)",
    },

    "ja": {
        # --- 시트명 ---
        "전체 게시물": "全投稿",
        "인사이트 랭킹": "ランキング",
        "시간대 분석": "時間帯分析",
        "팔로워 인구통계": "フォロワー属性",
        "성장 인사이트": "成長インサイト",
        "바이럴 심층분석": "バイラル詳細分析",
        "좋아요율 심층분석": "いいね率詳細分析",
        "콘텐츠 최적화": "コンテンツ最適化",
        "월별 트렌드": "月次トレンド",
        "일별 조회수 추이": "日次表示回数",
        "스냅샷 성장추이": "スナップショット推移",
        "10만 성장전략": "10万への戦略",

        # --- 섹션 제목 ---
        "1. 핵심 지표": "1. 主要指標",
        "1. 현재 위치 진단": "1. 現在地の診断",
        "1. 좋아요율 분포": "1. いいね率の分布",
        "1. 월별 성과 추이": "1. 月次パフォーマンス",
        "1. 스냅샷 이력": "1. スナップショット履歴",
        "1. 최근 90일 일별 조회수": "1. 直近90日の日次表示回数",
        "1. 글자수 구간별 성과 (조회수 vs 인게이지먼트율)":
            "1. 文字数別パフォーマンス（表示回数 vs エンゲージメント率）",
        "2. 미디어 타입별 성과": "2. メディアタイプ別パフォーマンス",
        "2. 미디어타입별 상세 비교": "2. メディアタイプ別の比較",
        "2. 바이럴 vs 일반 비교": "2. バイラル vs 通常",
        "2. 분기별 요약": "2. 四半期サマリー",
        "2. 좋아요율 TOP 30": "2. いいね率 TOP 30",
        "2. 즉시 실행 TOP 5": "2. 今すぐやること TOP 5",
        "2. 월별 실제 노출 조회수 (노출일 기준)":
            "2. 月別の実表示回数（表示日基準）",
        "2. 구간별 성장률 (결측 스냅샷 구간은 팔로워 계산 제외)":
            "2. 区間別の成長率（欠損スナップショットの区間は除外）",
        "3. 좋아요율 WORST 30": "3. いいね率 WORST 30",
        "3. 캐러셀 vs 텍스트 직접 비교": "3. カルーセル vs テキスト",
        "3. 바이럴 시간대 히트맵": "3. バイラルの時間帯ヒートマップ",
        "3. 요일별 노출량": "3. 曜日別の表示回数",
        "3. 성장 구간 분석 (시간 분위, 계정 독립)":
            "3. 時期別の成長分析（時間四分位）",
        "3. 성장 로드맵 (현재 팔로워 기준 스케일)":
            "3. 成長ロードマップ（現在のフォロワー基準）",
        "3. 좋아요율 TOP 20 (조회 500+ 기준) — 팬이 반응하는 콘텐츠":
            "3. いいね率 TOP 20（表示500以上）— ファンが反応する投稿",
        "4. 미디어타입별 좋아요율": "4. メディアタイプ別いいね率",
        "4. 시간대 × 요일 매트릭스": "4. 時間帯 × 曜日マトリクス",
        "4. 바이럴 요일분포": "4. バイラル投稿の曜日分布",
        "4. 콘텐츠 믹스 공식": "4. コンテンツ配分の型",
        "4. 요약": "4. サマリー",
        "4. 리포스트 비중 추이 (REPOST_FACADE)": "4. リポスト比率の推移",
        "5. 바이럴 미디어타입": "5. バイラル投稿のメディアタイプ",
        "5. 성장 전략 제안": "5. 成長戦略",
        "5. 토픽 태그별 성과": "5. トピックタグ別パフォーマンス",
        "6. 조회수 집중도": "6. 表示回数の集中度",

        # --- 공통 헤더 ---
        "번호": "No.", "순위": "順位", "날짜": "日付", "작성일": "投稿日",
        "작성일시(JST)": "投稿日時(JST)", "구분": "区分", "구간": "区間",
        "구간명": "区間", "기간": "期間", "단계": "段階", "지표": "指標",
        "값": "値", "비중": "比率", "비고": "備考", "연월": "年月",
        "설명": "説明", "해석": "読み方", "판정": "判定",
        "유형": "種別", "타입": "種別", "액션": "アクション", "목표": "目標",
        "우선순위": "優先度", "링크": "リンク", "토픽": "トピック",
        "요일": "曜日", "시간": "時間", "일수": "日数", "분기": "四半期",
        "국가": "国", "도시": "都市", "성별": "性別", "연령대": "年齢層",
        "현재": "現在", "현재값": "現在値", "목표값": "目標値",
        "핵심액션": "重点アクション", "데이터 근거": "根拠データ",
        "미디어타입": "メディアタイプ", "게시수": "投稿数", "게시물수": "投稿数",
        "글자수구간": "文字数帯", "스냅샷": "スナップショット",
        "관측주수": "観測週数", "게시일후경과": "投稿からの経過",
        "대상게시물수": "対象投稿数", "전체수": "全体",

        # --- 지표 ---
        "조회수": "表示回数", "총조회": "総表示回数", "총조회수": "総表示回数",
        "총 조회수": "総表示回数", "좋아요": "いいね", "총좋아요": "総いいね",
        "총 좋아요": "総いいね", "답글": "返信", "총답글": "総返信",
        "총 답글": "総返信", "리포스트": "リポスト",
        "리포스트수": "リポスト数", "총 리포스트": "総リポスト",
        "인용": "引用", "공유": "シェア", "팔로워": "フォロワー",
        "팔로워수": "フォロワー数", "팔로워증가": "フォロワー増加",
        "인게이지먼트": "エンゲージメント",
        "총인게이지먼트": "総エンゲージメント",
        "인게이지먼트율": "エンゲージメント率",
        "인게이지먼트율(%)": "エンゲージメント率(%)",
        "전체 인게이지먼트율": "全体エンゲージメント率",
        "좋아요율(%)": "いいね率(%)",
        "원본게시물": "オリジナル投稿", "원본게시물수": "オリジナル投稿数",
        "전체게시물": "全投稿", "답글있는게시물": "返信のある投稿",

        # --- 평균 / 중앙값 ---
        "평균조회": "平均表示", "평균조회수": "平均表示回数",
        "평균좋아요": "平均いいね", "평균좋아요수": "平均いいね数",
        "평균답글": "平均返信", "평균리포스트": "平均リポスト",
        "평균인게이지먼트": "平均エンゲージメント",
        "평균일조회": "1日平均表示", "평균좋아요율": "平均いいね率",
        "평균좋아요율(%)": "平均いいね率(%)",
        "평균 조회수/게시물": "平均表示 / 投稿",
        "평균 좋아요/게시물": "平均いいね / 投稿",
        "중앙값조회": "表示中央値", "중앙값좋아요": "いいね中央値",
        "중앙값ER(%)": "ER中央値(%)",
        "중앙값좋아요율(%)": "いいね率中央値(%)",
        "중앙값일조회": "日次表示の中央値",
        "중앙값조회증가": "表示増加の中央値",
        "평균÷중앙값 배수": "平均÷中央値",
        "1위 제외 평균": "1位を除く平均",
        "7일이동평균": "7日移動平均",

        # --- 비율 / 증감 ---
        "비율(%)": "比率(%)", "비중(%)": "比率(%)", "증감(%)": "増減(%)",
        "증가율(%)": "増加率(%)", "전월대비(%)": "前月比(%)",
        "전월대비조회증감(%)": "前月比 表示(%)",
        "전월대비좋아요증감(%)": "前月比 いいね(%)",
        "전주동요일대비(%)": "前週同曜日比(%)",
        "전체조회비중(%)": "総表示に占める割合(%)",
        "총조회비중(%)": "総表示に占める割合(%)",
        "상위1% 조회비중(%)": "上位1%の表示占有率(%)",
        "상위5% 조회비중(%)": "上位5%の表示占有率(%)",
        "상위10% 조회비중(%)": "上位10%の表示占有率(%)",
        "최다조회 1건 비중(%)": "最多表示1件の占有率(%)",
        "데드율(%)": "デッド率(%)",
        "바이럴비율(%)": "バイラル率(%)", "바이럴수": "バイラル投稿数",
        "바이럴 게시물 수": "バイラル投稿数", "일반수": "通常投稿数",
        "바이럴기준(적용값)": "バイラル基準(適用値)",
        "캐러셀비중": "カルーセル比率",
        "배율(캐러셀/텍스트)": "倍率(カルーセル/テキスト)",
        "30일조회": "30日表示", "30일좋아요": "30日いいね",
        "30일 조회수": "30日表示回数",
        "30일조회증감(%)": "30日表示の増減(%)",
        "원본1건당팔로워": "投稿1件あたりフォロワー",
        "원본게시물증가": "オリジナル投稿の増加",
        "일평균팔로워": "1日平均フォロワー",
        "일평균조회": "1日平均表示",
        "최근 30일": "直近30日", "직전 30일": "その前の30日",
        "730일 총조회": "730日の総表示",
        "최고일": "最高日", "최저일": "最低日",
        "스냅샷일시": "取得日時",
        "주요 미디어 타입": "主なメディアタイプ",
        "성과 상위 길이대": "成績の良い文字数帯",
        "목표팔로워": "目標フォロワー",
        "10만계정기준": "10万アカウント基準",

        # --- 라벨 / 판정 ---
        "🟢우수": "🟢 優秀", "⚪보통": "⚪ 普通", "🔴회피": "🔴 回避",
        "표본부족": "サンプル不足", "정상": "正常",
        "✅양호": "✅ 良好", "🔴 부족": "🔴 不足", "🔴 개선필요": "🔴 改善必要",
        "🏆조회최적": "🏆 表示最適", "⚠️조회데드존": "⚠️ 表示デッドゾーン",
        "📊 조회수 TOP 30": "📊 表示回数 TOP 30",
        "❤️ 인게이지먼트 TOP 30": "❤️ エンゲージメント TOP 30",
        "🔄 바이럴(리포스트+인용) TOP 30": "🔄 バイラル(リポスト+引用) TOP 30",
        "텍스트": "テキスト", "캐러셀": "カルーセル",
        "토픽 없음": "トピックなし", "기타(n<5)": "その他(n<5)",
        "요일별 성과": "曜日別パフォーマンス",
        "데드율(<500조회,%)": "デッド率(表示500未満, %)",
        "30일좋아요 데이터 결측": "30日いいねデータ欠損",
        "팔로워 데이터 결측 / 30일좋아요 데이터 결측":
            "フォロワーデータ欠損 / 30日いいねデータ欠損",
        "API 기준 일 경계(UTC-07:00), 게시물 시각(JST)과 다름":
            "APIの日境界はUTC-07:00で、投稿時刻(JST)とは異なる",

        # --- 요일 ---
        "월": "月", "화": "火", "수": "水", "목": "木",
        "금": "金", "토": "土", "일": "日",

        # --- 글자수 구간 ---
        "0-50자": "0-50字", "50-100자": "50-100字",
        "100-150자": "100-150字", "150-200자": "150-200字",
        "200-300자": "200-300字", "300-400자": "300-400字",
        "400-500자": "400-500字", "500자+": "500字以上",
        "타입 비교 데이터 부족": "メディアタイプを比較するにはデータが足りない",
        "데이터 없음 (일별 시계열 미수집)": "データなし（日次時系列は未取得）",
        "데이터 없음 (topic_tag 미수집 — analyze.py 재실행 필요)":
            "データなし（topic_tag 未取得 — analyze.py を再実行してください）",
        "데이터 없음": "データなし",
        "해당 없음": "該当なし",
        "🟡 근접": "🟡 目前",

        # --- 경과일 / 분위 / 기간 ---
        "0-7일": "0-7日", "7-14일": "7-14日", "14-30일": "14-30日",
        "30-60일": "30-60日", "60-90일": "60-90日", "90일+": "90日以上",
        "1분위(초기)": "第1四分位(初期)", "2분위": "第2四分位",
        "3분위": "第3四分位", "4분위(최근)": "第4四分位(直近)",
        "1-2개월": "1-2ヶ月", "3-4개월": "3-4ヶ月",
        "5-8개월": "5-8ヶ月", "9-18개월": "9-18ヶ月",

        # --- 전략 문구 ---
        "가치제안 1문장": "価値提案を1文で",
        "가치콘텐츠": "価値コンテンツ",
        "개인스토리": "個人のストーリー",
        "여정/비하인드": "過程・舞台裏",
        "팁/인사이트/교육": "ヒント・知見・解説",
        "질문/투표/오픈스레드": "質問・投票・オープンスレッド",
        "참여유도": "参加を促す",
        "소프트CTA": "ソフトCTA",
        "링크/전환": "リンク・転換",
        "바이오 명확화": "プロフィール文を明確に",
        "프로필 점검": "プロフィールの見直し",
        "게시시간 분산": "投稿時間を分散",
        "캐러셀 주2회": "カルーセルを週2回",
        "월평균조회 상위화": "月平均表示を引き上げる",
        "전환 경로 명확화": "転換経路を明確に",
        "일일 사려깊은 답글 루틴": "毎日の丁寧な返信ルーティン",
        "답글 비중↑ → 도달 가중 기대": "返信比率↑ → リーチ加重を期待",
        "최적 길이+질문마무리": "最適な文字数＋質問で締める",
        "프로필최적화+캐러셀주2+시간조정":
            "プロフィール最適化＋カルーセル週2＋時間調整",
        "질문형마무리+답글루틴+참여팟":
            "質問で締める＋返信ルーティン＋参加コミュニティ",
        "권위구축+협업+데이터기반재활용":
            "権威構築＋協業＋データに基づく再利用",
        "콘텐츠기둥3개+바이럴공식반복+크로스채널":
            "コンテンツの柱3本＋バイラル型の反復＋クロスチャネル",
        "바이럴 잘 되는 시간": "バイラルしやすい時間",
        "바이럴 잘 되는 요일": "バイラルしやすい曜日",
        "바이럴비율5%+": "バイラル率5%以上",
        "인게이지먼트율2%+": "エンゲージメント率2%以上",
        "• 좋아요율 높은 주제 강화 — 좋아요율 TOP 게시물 패턴을 분석해 반복 (좋아요율 = 팬 충성도 지표)":
            "• いいね率が高いテーマを伸ばす — いいね率 TOP の投稿のパターンを分析して繰り返す"
            "（いいね率はリーチではなくファンの熱量の指標）",
        "• 주말 활동 강화 — 토·일 게시량이 평일 대비 적음. 경쟁 콘텐츠가 줄어 도달 기회일 수 있음":
            "• 週末を試す — 土日の投稿量が平日より少ない。競合する投稿が減る分、リーチの機会になりうる",
        "내용(50자)": "本文(50字)", "내용(60자)": "本文(60字)",
        "내용(80자)": "本文(80字)", "내용(100자)": "本文(100字)",
    },
}

# 임계값이 박혀 동적으로 조립되는 문자열. 정확 일치가 불가능하므로 정규식으로 처리한다.
# 임계값·집계값이 박혀 동적으로 조립되는 문자열. 정확 일치가 불가능하므로 정규식으로 처리한다.
# repl 자리에 함수도 쓸 수 있다 (re.sub 규약). 목록형 문자열은 함수로 항목마다 치환한다.

_WD_EN = {"월": "Mon", "화": "Tue", "수": "Wed", "목": "Thu", "금": "Fri", "토": "Sat", "일": "Sun"}
_WD_JA = {"월": "月", "화": "火", "수": "水", "목": "木", "금": "金", "토": "土", "일": "日"}


def _wd_counts(mapping, unit):
    """'수(5건), 금(5건)' -> 'Wed (5), Fri (5)'"""
    def repl(m):
        parts = []
        for day, count in re.findall(r"([월화수목금토일])\((\d+)건\)", m.group(0)):
            parts.append(f"{mapping[day]}{unit.format(n=count)}")
        return ", ".join(parts)
    return repl


def _hour_counts(fmt):
    """'14시(3건), 17시(3건)' -> '14:00 (3), 17:00 (3)'"""
    def repl(m):
        return ", ".join(
            fmt.format(h=int(h), n=c)
            for h, c in re.findall(r"(\d+)시\((\d+)건\)", m.group(0))
        )
    return repl


def _hour_list(fmt, sep=", "):
    """'14시, 08시, 13시' 또는 '09시,10시' -> 시각 목록"""
    def repl(m):
        return sep.join(fmt.format(h=int(h)) for h in re.findall(r"(\d+)시", m.group(1)))
    return repl


def _type_counts(unit):
    """'TEXT_POST 1712건, IMAGE 215건' -> 'TEXT_POST 1712, IMAGE 215'"""
    def repl(m):
        return ", ".join(
            f"{name} {unit.format(n=count)}"
            for name, count in re.findall(r"([A-Z_]+) (\d+)건", m.group(0))
        )
    return repl


_WD_RUN = r"[월화수목금토일]\(\d+건\)(?:, [월화수목금토일]\(\d+건\))*"
_HR_RUN = r"\d+시\(\d+건\)(?:, \d+시\(\d+건\))*"
_TY_RUN = r"[A-Z_]+ \d+건(?:, [A-Z_]+ \d+건)*"

def _len_label(unit):
    """'400-500자' -> '400-500 chars' / '400-500字'. '500자+' 형태도 처리한다."""
    def conv(s):
        s = re.sub(r"(\d+)-(\d+)자", lambda m: f"{m.group(1)}-{m.group(2)}{unit}", s)
        return re.sub(r"(\d+)자\+", lambda m: f"{m.group(1)}{unit}+", s)
    return conv


def _length_tradeoff(template, unit):
    """글자수 상충 해석 문장. 구간 라벨이 문장 안에 박혀 있어 따로 변환한다."""
    label = _len_label(unit)
    def repl(m):
        return template.format(
            a=label(m.group(1)), er_a=m.group(2), v_a=m.group(3),
            b=label(m.group(4)), er_b=m.group(5), v_b=m.group(6), gap=m.group(7),
        )
    return repl


_TRADEOFF_RE = re.compile(
    r"^글이 길수록 인게이지먼트율은 올라가지만 조회수는 급감합니다 "
    r"\((.+): ER ([\d.]+)% / 조회 (\d+), (.+): ER ([\d.]+)% / 조회 (\d+)\) "
    r"— 조회 격차 ([\d.]+)배$"
)

PATTERNS = {
    "en": [
        # --- 바이럴 임계값이 박힌 제목/라벨 ---
        (re.compile(r"^(\d+)\. 바이럴 게시물 목록 \((.+)\+ 조회\)$"),
         r"\1. Viral Posts (\2+ views)"),
        (re.compile(r"^(\d+)\. 바이럴 게시물 패턴 분석 \((.+)\+ 조회\)$"),
         r"\1. Viral Post Patterns (\2+ views)"),
        (re.compile(r"^바이럴 게시물 \((.+)\+조회\)$"), r"Viral posts (\1+ views)"),
        (re.compile(r"^바이럴\((.+)\+\)$"), r"Viral (\1+)"),
        (re.compile(r"^일반\(<(.+)\)$"), r"Normal (<\1)"),
        (re.compile(r"^바이럴비율\((.+)\+조회\)$"), r"Viral rate (\1+ views)"),
        (re.compile(r"^(\d{4})-(\d{2}) \(부분\)$"), r"\1-\2 (partial)"),
        (re.compile(r"^@(\S+) 성장 인사이트 리포트$"), r"@\1 Growth Insights Report"),
        (re.compile(r"^총 팔로워: ([\d,]+)명$"), r"Total followers: \1"),
        (re.compile(r"^총 팔로워: -$"), "Total followers: -"),
        (re.compile(r"^주 ([+-][\d,]+) 팔로워$"), r"\1 followers / week"),

        # --- 목록형 (항목마다 치환) ---
        (re.compile(r"^" + _WD_RUN + r"$"), _wd_counts(_WD_EN, " ({n})")),
        (re.compile(r"^" + _HR_RUN + r"$"), _hour_counts("{h:02d}:00 ({n})")),

        # --- 시각·수량 라벨 ---
        (re.compile(r"^(\d{2})시$"), lambda m: f"{int(m.group(1)):02d}:00"),
        (re.compile(r"^(\d+)개$"), r"\1"),
        (re.compile(r"^(\d+)개 \(([\d.]+)%\)$"), r"\1 (\2%)"),
        (re.compile(r"^(\d{2})시 집중$"), lambda m: f"Focus on {int(m.group(1)):02d}:00"),
        (re.compile(r"^현재 ([\d.]+)%$"), r"Currently \1%"),
        (re.compile(r"^((?:\d+시,?)+) 과밀$"), _hour_list("{h:02d}:00", ", ")),
        (re.compile(r"^답글있는글 ([\d.]+)%$"), r"\1% of posts have replies"),
        (re.compile(r"^~?(\d+)(?:-(\d+))?자\+? 데드존$"),
         lambda m: (f"{m.group(1)}-{m.group(2)} chars" if m.group(2) else f"0-{m.group(1)} chars")
                   + " dead zone"),
        (re.compile(r"^• 게시 요일 분산 — 현재 최빈 요일 ([월화수목금토일])\. "
                    r"성과 상위 요일에 핵심 콘텐츠 배치$"),
         lambda m: f"• Spread out your weekdays — you post most on {_WD_EN[m.group(1)]}. "
                   "Move your best work to the days that actually perform"),
        (re.compile(r"^캐러셀/텍스트 인게이지먼트 ([\d.]+)배$"),
         r"Carousel vs text engagement \1x"),
        (re.compile(r"^길이구간 좋아요율 최대/최소 ([\d.]+)배$"),
         r"Like-rate spread across length buckets \1x"),
        (re.compile(r"^최고 (\d{2})시 평균조회 ([\d.]+)배\(전체대비\)$"),
         lambda m: f"Best hour {int(m.group(1)):02d}:00, {m.group(2)}x overall average"),

        # --- 섹션 제목·주석 ---
        (re.compile(r"^🎯 좋아요율 TOP (\d+) \(조회 (\d+)\+ 기준\)$"),
         r"🎯 Top \1 by Like Rate (\2+ views)"),
        (re.compile(r"^3\. 게시물 조회수 누적 곡선 \(id 조인, n<(\d+) 구간은 '-'\)$"),
         r"3. View Accrual Curve (joined by id; '-' where n<\1)"),
        (re.compile(r"^시간대별 성과 \(JST, 중앙값 기준 · 데드=(\d+)조회 미만 · 판정은 (\d+)건 이상\)$"),
         r"Performance by hour (JST, medians; dead = under \1 views; verdict needs \2+ posts)"),
        (re.compile(r"^분석일: (\S+) / 팔로워: ([\d,]+)명 / 게시물: (\d+)개 \(리포스트 제외\)$"),
         r"Analyzed \1 / Followers \2 / Posts \3 (reposts excluded)"),
        (re.compile(r"^분석일: (\S+) / 팔로워: - / 게시물: (\d+)개 \(리포스트 제외\)$"),
         r"Analyzed \1 / Followers - / Posts \2 (reposts excluded)"),
        (re.compile(r"^미디어타입 전체 포함: (" + _TY_RUN + r") / 데드율 = 조회 (\d+) 미만 비율$"),
         lambda m: "All media types included: "
                   + _type_counts("{n}")(re.match(_TY_RUN, m.group(1)))
                   + f" / dead rate = share under {m.group(2)} views"),

        # --- 해석 문장 ---
        (_TRADEOFF_RE, _length_tradeoff(
            "Longer posts raise engagement rate but collapse reach "
            "({a}: ER {er_a}% / {v_a} views, {b}: ER {er_b}% / {v_b} views) — {gap}x reach gap",
            " chars")),
        (re.compile(r"^상위 1%\((\d+)개\)가 전체 조회의 ([\d.]+)%, 최다 1건만으로 ([\d.]+)%를 차지합니다\. "
                    r"평균 ([\d,]+)은 중앙값 ([\d,]+)의 ([\d.]+)배이며, 1위 게시물을 빼면 평균이 ([\d,]+)로 내려갑니다\. "
                    r"다른 시트의 '평균' 지표는 소수 이상치가 만든 값이므로 중앙값과 함께 해석하세요\.$"),
         r"The top 1% (\1 posts) hold \2% of all views, and the single best post alone holds \3%. "
         r"The mean of \4 is \6x the median of \5; drop the top post and the mean falls to \7. "
         r"Treat every 'average' on the other sheets as an outlier artifact and read it next to the median."),

        # --- 전략 불릿 ---
        (re.compile(r"^• 캐러셀 활용 확대 — 평균 인게이지먼트 캐러셀 ([\d.]+) vs 텍스트 ([\d.]+) "
                    r"\(([\d.]+)배\)\. 현재 (\d+)개 → 주 1-2회 목표$"),
         r"• Post more carousels — avg engagement \1 for carousels vs \2 for text (\3x). "
         r"You have \4; aim for 1-2 per week"),
        (re.compile(r"^• 미디어 타입 실험 — 캐러셀 ([\d.]+) / 텍스트 ([\d.]+) 평균 인게이지먼트\. 상위 타입 비중 확대$"),
         r"• Test media types — avg engagement \1 carousel / \2 text. Shift toward the stronger one"),
        (re.compile(r"^• 골든타임 집중 — 평균 조회 상위 시간대\(JST\): ((?:\d+시(?:, )?)+)\. 핵심 콘텐츠를 이 시간대에 게시$"),
         lambda m: "• Use your golden hours — highest average views (JST): "
                   + ", ".join(f"{int(h):02d}:00" for h in re.findall(r"(\d+)시", m.group(1)))
                   + ". Post your best work then"),
        (re.compile(r"^• 바이럴 패턴 반복 — 상위 조회 바이럴 (\d+)개 중 최빈 타입 (\S+)\((\d+)건\)\. "
                    r"해당 포맷·주제 패턴 재사용$"),
         r"• Repeat what went viral — of your top \1 viral posts, \2 is the most common type (\3). "
         r"Reuse that format and subject"),
        (re.compile(r"^• 답글 유도 — 현재 인게이지먼트율 ([\d.]+)%\. 질문형 마무리로 답글 비중을 높이면 알고리즘 가중치에 유리$"),
         r"• Invite replies — engagement rate is \1%. Ending on a question lifts reply share, "
         r"which the ranking appears to weight"),
        (re.compile(r"^• 팔로워 인구통계 활용 — 핵심층: 연령 (\S+), 성별 (\S+), 국가 (\S+)\. 이 타겟에 맞는 콘텐츠 톤 유지$"),
         r"• Write for who actually follows you — core segment: age \1, gender \2, country \3"),
    ],

    "ja": [
        (re.compile(r"^(\d+)\. 바이럴 게시물 목록 \((.+)\+ 조회\)$"),
         r"\1. バイラル投稿一覧（表示\2以上）"),
        (re.compile(r"^(\d+)\. 바이럴 게시물 패턴 분석 \((.+)\+ 조회\)$"),
         r"\1. バイラル投稿のパターン分析（表示\2以上）"),
        (re.compile(r"^바이럴 게시물 \((.+)\+조회\)$"), r"バイラル投稿（表示\1以上）"),
        (re.compile(r"^바이럴\((.+)\+\)$"), r"バイラル（\1以上）"),
        (re.compile(r"^일반\(<(.+)\)$"), r"通常（\1未満）"),
        (re.compile(r"^바이럴비율\((.+)\+조회\)$"), r"バイラル率（表示\1以上）"),
        (re.compile(r"^(\d{4})-(\d{2}) \(부분\)$"), r"\1-\2（一部）"),
        (re.compile(r"^@(\S+) 성장 인사이트 리포트$"), r"@\1 成長インサイトレポート"),
        (re.compile(r"^총 팔로워: ([\d,]+)명$"), r"総フォロワー: \1"),
        (re.compile(r"^총 팔로워: -$"), "総フォロワー: -"),
        (re.compile(r"^주 ([+-][\d,]+) 팔로워$"), r"週 \1 フォロワー"),

        (re.compile(r"^" + _WD_RUN + r"$"), _wd_counts(_WD_JA, "（{n}件）")),
        (re.compile(r"^" + _HR_RUN + r"$"), _hour_counts("{h}時（{n}件）")),

        (re.compile(r"^(\d{2})시$"), lambda m: f"{int(m.group(1))}時"),
        (re.compile(r"^(\d+)개$"), r"\1件"),
        (re.compile(r"^(\d+)개 \(([\d.]+)%\)$"), r"\1件 (\2%)"),
        (re.compile(r"^(\d{2})시 집중$"), lambda m: f"{int(m.group(1))}時に集中"),
        (re.compile(r"^현재 ([\d.]+)%$"), r"現在 \1%"),
        (re.compile(r"^((?:\d+시,?)+) 과밀$"), _hour_list("{h}時", "・")),
        (re.compile(r"^답글있는글 ([\d.]+)%$"), r"返信のある投稿 \1%"),
        (re.compile(r"^~?(\d+)(?:-(\d+))?자\+? 데드존$"),
         lambda m: (f"{m.group(1)}-{m.group(2)}字" if m.group(2) else f"{m.group(1)}字未満")
                   + "のデッドゾーン"),
        (re.compile(r"^• 게시 요일 분산 — 현재 최빈 요일 ([월화수목금토일])\. "
                    r"성과 상위 요일에 핵심 콘텐츠 배치$"),
         lambda m: f"• 曜日を分散する — 現在は{_WD_JA[m.group(1)]}曜に偏っている。"
                   "成績の良い曜日に主力を回す"),
        (re.compile(r"^캐러셀/텍스트 인게이지먼트 ([\d.]+)배$"),
         r"カルーセル/テキストのエンゲージメント \1倍"),
        (re.compile(r"^길이구간 좋아요율 최대/최소 ([\d.]+)배$"),
         r"文字数帯のいいね率 最大/最小 \1倍"),
        (re.compile(r"^최고 (\d{2})시 평균조회 ([\d.]+)배\(전체대비\)$"),
         lambda m: f"最高は{int(m.group(1))}時、平均表示が全体の{m.group(2)}倍"),

        (re.compile(r"^🎯 좋아요율 TOP (\d+) \(조회 (\d+)\+ 기준\)$"),
         r"🎯 いいね率 TOP \1（表示\2以上）"),
        (re.compile(r"^3\. 게시물 조회수 누적 곡선 \(id 조인, n<(\d+) 구간은 '-'\)$"),
         r"3. 表示回数の累積カーブ（id結合、n<\1 の区間は '-'）"),
        (re.compile(r"^시간대별 성과 \(JST, 중앙값 기준 · 데드=(\d+)조회 미만 · 판정은 (\d+)건 이상\)$"),
         r"時間帯別の成績（JST・中央値基準・デッドは表示\1未満・判定は\2件以上）"),
        (re.compile(r"^분석일: (\S+) / 팔로워: ([\d,]+)명 / 게시물: (\d+)개 \(리포스트 제외\)$"),
         r"分析日: \1 / フォロワー: \2 / 投稿: \3件（リポスト除く）"),
        (re.compile(r"^분석일: (\S+) / 팔로워: - / 게시물: (\d+)개 \(리포스트 제외\)$"),
         r"分析日: \1 / フォロワー: - / 投稿: \2件（リポスト除く）"),
        (re.compile(r"^미디어타입 전체 포함: (" + _TY_RUN + r") / 데드율 = 조회 (\d+) 미만 비율$"),
         lambda m: "全メディアタイプを含む: "
                   + _type_counts("{n}件")(re.match(_TY_RUN, m.group(1)))
                   + f" / デッド率 = 表示{m.group(2)}未満の割合"),

        (_TRADEOFF_RE, _length_tradeoff(
            "長い投稿ほどエンゲージメント率は上がるが表示回数は激減する"
            "（{a}: ER {er_a}% / 表示 {v_a}、{b}: ER {er_b}% / 表示 {v_b}）— 表示の差 {gap}倍",
            "字")),
        (re.compile(r"^상위 1%\((\d+)개\)가 전체 조회의 ([\d.]+)%, 최다 1건만으로 ([\d.]+)%를 차지합니다\. "
                    r"평균 ([\d,]+)은 중앙값 ([\d,]+)의 ([\d.]+)배이며, 1위 게시물을 빼면 평균이 ([\d,]+)로 내려갑니다\. "
                    r"다른 시트의 '평균' 지표는 소수 이상치가 만든 값이므로 중앙값과 함께 해석하세요\.$"),
         r"上位1%（\1件）が全表示の\2%を占め、最多の1件だけで\3%を占めます。"
         r"平均\4は中央値\5の\6倍で、1位を除くと平均は\7まで下がります。"
         r"他シートの「平均」は少数の外れ値が作った値なので、必ず中央値と併せて読んでください。"),

        (re.compile(r"^• 캐러셀 활용 확대 — 평균 인게이지먼트 캐러셀 ([\d.]+) vs 텍스트 ([\d.]+) "
                    r"\(([\d.]+)배\)\. 현재 (\d+)개 → 주 1-2회 목표$"),
         r"• カルーセルを増やす — 平均エンゲージメントはカルーセル\1 対 テキスト\2（\3倍）。"
         r"現在\4件、週1〜2回を目標に"),
        (re.compile(r"^• 미디어 타입 실험 — 캐러셀 ([\d.]+) / 텍스트 ([\d.]+) 평균 인게이지먼트\. 상위 타입 비중 확대$"),
         r"• メディアタイプを試す — 平均エンゲージメント カルーセル\1 / テキスト\2。強い方の比率を上げる"),
        (re.compile(r"^• 골든타임 집중 — 평균 조회 상위 시간대\(JST\): ((?:\d+시(?:, )?)+)\. 핵심 콘텐츠를 이 시간대에 게시$"),
         lambda m: "• ゴールデンタイムに寄せる — 平均表示が高い時間帯(JST): "
                   + "、".join(f"{int(h)}時" for h in re.findall(r"(\d+)시", m.group(1)))
                   + "。主力の投稿はこの時間に"),
        (re.compile(r"^• 바이럴 패턴 반복 — 상위 조회 바이럴 (\d+)개 중 최빈 타입 (\S+)\((\d+)건\)\. "
                    r"해당 포맷·주제 패턴 재사용$"),
         r"• 当たった型を繰り返す — 上位バイラル\1件のうち最頻タイプは\2（\3件）。"
         r"その形式とテーマを再利用する"),
        (re.compile(r"^• 답글 유도 — 현재 인게이지먼트율 ([\d.]+)%\. 질문형 마무리로 답글 비중을 높이면 알고리즘 가중치에 유리$"),
         r"• 返信を誘う — 現在のエンゲージメント率は\1%。質問で締めると返信比率が上がり、"
         r"ランキング上も有利に働く"),
        (re.compile(r"^• 팔로워 인구통계 활용 — 핵심층: 연령 (\S+), 성별 (\S+), 국가 (\S+)\. 이 타겟에 맞는 콘텐츠 톤 유지$"),
         r"• 実際のフォロワーに向けて書く — 中核層: 年齢\1、性別\2、国\3"),
    ],
}

def _convert(text, table, patterns):
    if text in table:
        return table[text]
    for pattern, repl in patterns:
        if pattern.match(text):
            return pattern.sub(repl, text)
    return None


def translate_workbook(wb, lang):
    """워크북을 제자리에서 번역한다. lang이 'ko'면 아무것도 하지 않는다.

    본문 열을 따로 보호하지 않는다. 치환은 사전 정확 일치라서, 게시물 본문이
    바뀌려면 본문(50/60/80/100자로 절단된 값)이 UI 문자열과 글자 단위로 같아야
    한다. 2,292개 게시물 x 절단길이 4종 + 토픽 태그 전수 검사에서 일치 0건이었다.
    열 단위로 막으려 했더니 같은 시트의 다른 섹션 표가 같은 열 번호를 써서
    멀쩡한 헤더까지 번역이 막혔다. tests/test_core.py가 이 불변식을 지킨다.
    """
    if lang == "ko":
        return wb
    table = UI.get(lang)
    if not table:
        return wb
    patterns = PATTERNS.get(lang, [])

    for ws in wb.worksheets:
        title = _convert(ws.title, table, patterns)
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    converted = _convert(cell.value, table, patterns)
                    if converted is not None:
                        cell.value = converted
        if title:
            # 시트명은 마지막에 바꾼다. Excel 31자 제한을 넘기면 openpyxl이 조용히 자른다.
            ws.title = title[:EXCEL_SHEET_MAX]
    return wb
