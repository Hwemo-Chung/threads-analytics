# Threads Analyzer

Meta Threads API를 활용한 계정 종합 분석 도구. OAuth 인증부터 게시물 수집, 인사이트 분석, Excel 리포트 생성까지 자동화합니다.

## 기능

- **OAuth 2.0 인증** — 로컬 HTTPS 서버로 Threads 토큰 자동 발급
- **전체 게시물 수집** — 페이지네이션으로 모든 게시물 + 인사이트 병렬 수집
- **계정 인사이트** — 팔로워 수, 30일 조회수, 인구통계(국가/도시/성별/연령)
- **Excel 리포트** (12개 시트) — 랭킹, 시간대 분석, 일별 시계열, 바이럴 분석, 성장 전략 등
- **3개 언어 출력** — 한국어 / English / 日本語 (`--lang`)
- **토큰 갱신** — 60일 장기 토큰 만료 전 갱신

## 사전 준비

### 1. Meta Developer 앱 생성

1. [Meta for Developers](https://developers.facebook.com/)에서 앱 생성
2. **Use case** → "Other" → "Consumer" 선택
3. 앱 대시보드에서 **Threads API** 제품 추가
   > ⚠️ **주의**: 제품 검색창에 "Threads"를 직접 타이핑하거나 붙여넣기하면 검색이 안 될 수 있습니다. 반드시 **드롭다운 목록에서 "Threads API"를 직접 클릭**하여 선택하세요.
4. **App Settings > Basic**에서 확인:
   - `Threads App ID`
   - `Threads App Secret`
5. **Threads API > Settings**에서 콜백 URL 3개 모두 등록:
   - Callback URL: `https://localhost:8888/callback`
   - Deauthorize URL: `https://localhost:8888/deauthorize`
   - Data Deletion URL: `https://localhost:8888/delete`
6. **Threads API > Permissions**에서 테스터 추가 (본인 Threads 계정)
   - Threads 앱에서 테스터 초대 수락

> ⚠️ Meta는 3개 콜백 URL 모두 **HTTPS 필수**입니다.

### 2. Python 환경

- Python 3.8 이상 필요

## 설치

```bash
git clone https://github.com/Hwemo-Chung/threads-analytics.git
cd threads-analytics
python3 setup.py
```

`setup.py`가 의존성 설치, SSL 인증서 생성, `.env` 작성, 인증까지 한 번에 처리합니다.
Meta 앱 ID/Secret만 준비해서 붙여넣으면 됩니다. 이미 설정된 값은 건드리지 않으므로 다시 실행해도 안전합니다.

> 수동으로 하시려면 아래 "수동 설정"을 따르세요.

## 수동 설정

### 1. 환경변수 설정

```bash
cp .env.example .env
pip install -r requirements.txt
```

`.env` 파일을 열고 앱 정보 입력:

```
THREADS_APP_ID=your_app_id_here
THREADS_APP_SECRET=your_app_secret_here
REDIRECT_URI=https://localhost:8888/callback
INSIGHTS_CACHE_TTL_DAYS=7
```

> `INSIGHTS_CACHE_TTL_DAYS`: 게시물 인사이트 캐시 유효 기간(일). 만료된 항목만 재조회합니다. `0`이면 만료 없음.

### 2. SSL 인증서 생성

OAuth 콜백을 받기 위한 자체 서명 인증서가 필요합니다:

```bash
openssl req -x509 -newkey rsa:2048 \
  -keyout localhost.key -out localhost.crt \
  -days 365 -nodes -subj "/CN=localhost"
```

## 사용법

### Step 1: 인증

```bash
python auth.py
```

1. 브라우저가 자동으로 열립니다
2. Threads에서 로그인 & 권한 승인
3. 리디렉션 시 브라우저 보안 경고 → **"고급" → "localhost(안전하지 않음)으로 이동"** 클릭
4. 인증 성공 시 `.env`에 `ACCESS_TOKEN`과 `USER_ID`가 자동 저장됩니다

> 💡 브라우저가 안 열리면 터미널에 표시된 URL을 직접 붙여넣으세요.

### Step 2: 데이터 수집 & 분석

```bash
python3 analyze.py
```

- 전체 게시물을 페이지네이션으로 수집
- 게시물별 인사이트를 5개 병렬로 수집 (캐시 + TTL)
- 계정 인사이트 + 팔로워 인구통계 수집
- 토큰 만료 사전 경고 (`TOKEN_EXPIRES_AT`)
- 429/5xx 자동 재시도
- 터미널 요약 + `output/analysis_*.json` 저장

**주요 옵션:**

```bash
python3 analyze.py --refresh-insights          # 캐시 무시, 인사이트 전체 재조회
python3 analyze.py --ttl-days 30               # 이번 실행만 TTL 30일
python3 analyze.py --max-posts 100             # 최근 페이지 기준 상한 (빠른 샘플)
python3 analyze.py --skip-demographics         # 인구통계 생략
python3 analyze.py --fail-on-api-error         # API 오류 시 exit 2
python3 analyze.py --workers 3                 # 병렬 워커 수
```

> ⚠️ API 호출이 많아 **약 10분** 소요될 수 있습니다. tmux 등에서 실행을 권장합니다.

### Step 3: Excel 리포트 생성

```bash
python3 export_excel.py
python3 export_excel.py -i output/analysis_YYYYMMDD_HHMMSS.json
python3 export_excel.py -i path/to.json -o path/to/report.xlsx
```

기본: `output/` 최신 `analysis_*.json` → `output/threads_analysis_YYYYMMDD.xlsx`

**리포트 언어 (한국어 / 영어 / 일본어):**

```bash
python3 export_excel.py --lang en      # English
python3 export_excel.py --lang ja      # 日本語
python3 export_excel.py                # 한국어 (기본)
```

`.env`에 `THREADS_LANG=en`을 넣어 고정할 수도 있습니다. `--lang`이 우선합니다.
시트명·헤더·판정 라벨만 번역되고 **게시물 본문과 숫자는 원본 그대로** 유지됩니다.

**생성되는 시트 (12개):**

| 시트 | 내용 |
|---|---|
| 전체 게시물 | 전체 목록 (날짜/조회/좋아요/인게이지먼트, 필터 가능) |
| 인사이트 랭킹 | 조회수/인게이지먼트/좋아요율/바이럴 각 TOP 30 |
| 시간대 분석 | 24시간대별 + 요일별 **중앙값**과 데드율(500조회 미만 비율) |
| 팔로워 인구통계 | 국가/도시/성별/연령대 분포 |
| 성장 인사이트 | KPI 요약, 미디어타입별 성과, 바이럴 패턴 |
| 바이럴 심층분석 | 계정 규모에 맞춘 동적 임계값(P99), 조회 집중도 |
| 좋아요율 심층분석 | 좋아요율 분포, TOP/WORST 30, 미디어타입별 중앙값 |
| 콘텐츠 최적화 | 글자수별 조회수 vs 인게이지먼트율 상충, 토픽 태그별 성과 |
| 월별 트렌드 | 월별 추이(전월대비 증감%), 분기별 요약 |
| 일별 조회수 추이 | 최대 730일 일별 시계열, 7일 이동평균, 노출일 기준 월별 |
| 스냅샷 성장추이 | 실행할 때마다 쌓이는 팔로워 이력 (API가 제공하지 않는 데이터) |
| 10만 성장전략 | 현재 위치 진단, 즉시 실행 TOP 5, 로드맵 |

> **평균이 아니라 중앙값을 봅니다.** 바이럴 게시물 하나가 평균을 끌어올려 시간대 판정을 뒤집기 때문입니다.
> 실제로 제작자 계정에서는 평균 기준 최고였던 시간대가 중앙값 기준으로는 최악(데드율 70%)이었습니다.

### 토큰 갱신 (60일마다)

```bash
python3 refresh_token.py
```

토큰 만료 전에 실행하면 60일 연장되고 `TOKEN_EXPIRES_AT`이 갱신됩니다.  
`analyze.py`는 만료 7일 전(`TOKEN_WARN_DAYS`) 경고, 만료 시 중단합니다.  
이미 만료된 경우 `auth.py`를 다시 실행하세요.

### 텍스트 저장함

분석 JSON에서 전체 게시 본문을 정리해 `output/저장함/`에 둡니다.

```bash
python3 archive.py                              # 최신 analysis_*.json
python3 archive.py -i output/analysis_....json  # 특정 파일
```

또는 분석 시 자동 생성(기본 ON, `--skip-archive`로 끄기):

```bash
python3 analyze.py
python3 analyze.py --export-excel   # 분석 + 저장함 + Excel
```

**저장함 구조:**

```
output/저장함/
├── latest/                 # 항상 최신 스냅샷 복사본
│   ├── README.md
│   ├── all_texts.md        # 전체 본문 (최신순)
│   ├── text_only.txt       # 본문만
│   ├── all_posts.jsonl
│   ├── index.csv
│   ├── by_month/*.md
│   └── by_type/*.md
└── snapshots/YYYYMMDD_HHMMSS/   # 이력 보관
```

> Meta API에는 앱 북마크(타인 글 저장함) 조회 엔드포인트가 **없습니다**.  
> 이 저장함은 **본인 게시글 텍스트 아카이브**입니다.

### 테스트

```bash
python3 -m unittest discover -s tests -v
```

## 프로젝트 구조

```
threads-analytics/
├── setup.py             # 설치 도우미 (의존성·인증서·.env·인증 일괄)
├── i18n.py              # 리포트 출력 언어 (ko/en/ja)
├── auth.py              # OAuth 2.0 (HTTPS, state CSRF, TOKEN_EXPIRES_AT)
├── analyze.py           # 수집·분석 (재시도, TTL, CLI, 저장함)
├── archive.py           # 게시 텍스트 저장함
├── export_excel.py      # JSON → Excel 12시트 (-i/-o)
├── refresh_token.py     # 장기 토큰 갱신
├── tests/               # 단위 테스트 (네트워크 없음)
├── requirements.txt
├── .env.example
├── .gitignore
└── output/              # 분석 결과 (gitignore)
```

## API 참고사항

- **Rate Limit**: 약 4,800 calls/hour (Threads API). 429/5xx는 자동 백오프 재시도
- **토큰 유효기간**: 장기 토큰 60일, `TOKEN_EXPIRES_AT` 기록, `refresh_token.py`로 갱신
- **팔로워 인구통계**: 팔로워 100명 이상부터 제공
- **인사이트 캐시**: v2 + TTL(`INSIGHTS_CACHE_TTL_DAYS`, 기본 7일). `--refresh-insights`로 강제 재조회
- **저장소 범위**: 공개 대상은 분석 스크립트·문서. 마케팅/라이팅 킷(`release/` 등)은 로컬 전용(gitignore)

## 라이선스

MIT
