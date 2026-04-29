# Threads Analyzer

Meta Threads API를 활용한 계정 종합 분석 도구. OAuth 인증부터 게시물 수집, 인사이트 분석, Excel 리포트 생성까지 자동화합니다.

## 기능

- **OAuth 2.0 인증** — 로컬 HTTPS 서버로 Threads 토큰 자동 발급
- **전체 게시물 수집** — 페이지네이션으로 모든 게시물 + 인사이트 병렬 수집
- **계정 인사이트** — 팔로워 수, 30일 조회수, 인구통계(국가/도시/성별/연령)
- **Excel 리포트** (10개 시트) — 랭킹, 시간대 분석, 바이럴 분석, 성장 전략 등
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
git clone https://github.com/YOUR_USERNAME/threads-analytics.git
cd threads-analytics
pip install -r requirements.txt
```

## 설정

### 1. 환경변수 설정

```bash
cp .env.example .env
```

`.env` 파일을 열고 앱 정보 입력:

```
THREADS_APP_ID=your_app_id_here
THREADS_APP_SECRET=your_app_secret_here
REDIRECT_URI=https://localhost:8888/callback
```

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
python analyze.py
```

- 전체 게시물을 페이지네이션으로 수집
- 게시물별 인사이트를 5개 병렬로 수집 (캐시 지원)
- 계정 인사이트 + 팔로워 인구통계 수집
- 터미널에 분석 요약 출력
- `output/` 디렉토리에 JSON 원본 데이터 저장

> ⚠️ API 호출이 많아 **약 10분** 소요될 수 있습니다. tmux 등에서 실행을 권장합니다.

### Step 3: Excel 리포트 생성

```bash
python export_excel.py
```

`output/` 디렉토리에서 가장 최신 분석 JSON을 읽어 Excel 리포트를 생성합니다.

**생성되는 시트 (10개):**

| 시트 | 내용 |
|---|---|
| 전체 게시물 | 전체 목록 (날짜/조회/좋아요/인게이지먼트, 필터 가능) |
| 인사이트 랭킹 | 조회수/인게이지먼트/좋아요율/바이럴 각 TOP 30 |
| 시간대 분석 | 24시간대별 + 요일별 평균 성과 |
| 팔로워 인구통계 | 국가/도시/성별/연령대 분포 |
| 성장 인사이트 | KPI 요약, 미디어타입별 성과, 바이럴 패턴 |
| 바이럴 심층분석 | 1만+ 조회 전수 목록, 바이럴 vs 일반 비교 |
| 좋아요율 심층분석 | 좋아요율 분포, TOP/WORST 30, 미디어타입별 중앙값 |
| 콘텐츠 최적화 | 텍스트 길이별 성과, 캐러셀 vs 텍스트 비교 |
| 월별 트렌드 | 월별 추이(전월대비 증감%), 분기별 요약 |
| 10만 성장전략 | 현재 위치 진단, 즉시 실행 TOP 5, 로드맵 |

### 토큰 갱신 (60일마다)

```bash
python refresh_token.py
```

토큰 만료 전에 실행하면 60일 연장됩니다. 만료된 경우 `auth.py`를 다시 실행해주세요.

## 프로젝트 구조

```
threads-analytics/
├── auth.py              # OAuth 2.0 인증 (HTTPS 로컬서버)
├── analyze.py           # 게시물 수집 + 인사이트 분석
├── export_excel.py      # JSON → Excel 10시트 리포트
├── refresh_token.py     # 장기 토큰 갱신
├── requirements.txt     # Python 의존성
├── .env.example         # 환경변수 템플릿
├── .gitignore
└── output/              # 분석 결과 (gitignore됨)
    ├── analysis_*.json
    ├── insights_cache.json
    └── threads_analysis_*.xlsx
```

## API 참고사항

- **Rate Limit**: 약 4,800 calls/hour (Threads API)
- **토큰 유효기간**: 장기 토큰 60일, `refresh_token.py`로 갱신
- **팔로워 인구통계**: 팔로워 100명 이상부터 제공
- **인사이트 캐시**: `insights_cache.json`에 저장되어 재실행 시 중복 호출 방지

## 라이선스

MIT
