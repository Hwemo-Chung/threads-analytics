# Threads Analyzer

[한국어](./README.md) · **English** · [日本語](./README.ja.md)

Pull every post from your Meta Threads account and get a 12-sheet Excel report — rankings, hour-of-day performance, a 730-day daily view series, viral concentration, follower demographics.

**Averages lie on Threads.** One viral post drags the mean up and flips your conclusions. This tool reports medians and dead rates alongside the averages, so the hour that looks best on a mean often turns out to be the worst.

On the author's own account (1,954 original posts, 4.4M lifetime views):

| | |
|---|---|
| Median views per post | **412** |
| Mean views per post | **2,254** — 5.5x the median |
| Share of all views held by the top 1% of posts | **48.3%** |
| Posts that never reached 500 views | **56.5%** |
| 10:00 KST — 2nd most-used posting hour | median **278**, dead rate **70.0%** |
| 01:00 KST | median **666**, dead rate **43.8%** |

## Features

- **OAuth 2.0** — local HTTPS server issues your long-lived token
- **Full post collection** — paginated, with per-post insights fetched in parallel and cached
- **Account insights** — followers, 30-day views, demographics (country / city / gender / age)
- **12-sheet Excel report** — see the table below
- **Three output languages** — Korean / English / Japanese via `--lang`
- **Token refresh** — extend the 60-day long-lived token before it expires

## Requirements

### 1. A Meta developer app

You need your own app. There is no way around this — Meta only exposes insights through an app owned by the account holder. Roughly 10 minutes.

1. Create an app at [Meta for Developers](https://developers.facebook.com/)
2. **Use case** → "Other" → "Consumer"
3. Add the **Threads API** product
   > ⚠️ Typing or pasting "Threads" into the product search box often returns nothing. Click **"Threads API" in the dropdown list** instead.
4. **App Settings > Basic** — copy your `Threads App ID` and `Threads App Secret`
5. **Threads API > Settings** — register all three callback URLs:
   - Callback URL: `https://localhost:8888/callback`
   - Deauthorize URL: `https://localhost:8888/deauthorize`
   - Data Deletion URL: `https://localhost:8888/delete`
6. **Threads API > Permissions** — add your own Threads account as a tester, then accept the invite from the Threads app

> ⚠️ Meta requires **HTTPS** on all three callback URLs. That is why the setup generates a self-signed certificate.

Because you own the app, this runs in development mode and **no Meta App Review is required**.

### 2. Python 3.8+

## Install

```bash
git clone https://github.com/Hwemo-Chung/threads-analytics.git
cd threads-analytics
python3 setup.py
```

`setup.py` installs dependencies, generates the SSL certificate, writes `.env`, and runs the OAuth flow. You only paste the App ID and Secret. Re-running it is safe — existing tokens and settings are preserved.

<details>
<summary>Manual setup</summary>

```bash
cp .env.example .env
pip install -r requirements.txt

openssl req -x509 -newkey rsa:2048 \
  -keyout localhost.key -out localhost.crt \
  -days 365 -nodes -subj "/CN=localhost"

python3 auth.py
```

Fill in `.env`:

```
THREADS_APP_ID=your_app_id_here
THREADS_APP_SECRET=your_app_secret_here
REDIRECT_URI=https://localhost:8888/callback
INSIGHTS_CACHE_TTL_DAYS=7
```
</details>

## Usage

### Collect and analyze

```bash
python3 analyze.py
```

Paginates through every post, fetches per-post insights with 5 workers, collects account insights and follower demographics, warns before token expiry, retries 429/5xx with backoff, and writes `output/analysis_*.json`.

```bash
python3 analyze.py --refresh-insights     # ignore cache, refetch every post
python3 analyze.py --ttl-days 30          # override cache TTL for this run
python3 analyze.py --max-posts 100        # quick sample
python3 analyze.py --skip-demographics
python3 analyze.py --fail-on-api-error    # exit 2 on API error
python3 analyze.py --workers 3
python3 analyze.py --export-excel         # analyze, then build the workbook
```

> First run takes about **10 minutes** for ~2,000 posts. Later runs hit the cache and take 1-2 minutes.

### Build the Excel report

```bash
python3 export_excel.py
python3 export_excel.py -i output/analysis_YYYYMMDD_HHMMSS.json
python3 export_excel.py -o path/to/report.xlsx
```

**Output language:**

```bash
python3 export_excel.py --lang en      # English
python3 export_excel.py --lang ja      # 日本語
python3 export_excel.py                # Korean (default)
```

You can pin it with `THREADS_LANG=en` in `.env`; `--lang` wins.

Sheet names, headers, verdict labels **and the generated strategy commentary** are all translated.
**Post text and every number stay exactly as collected.**

### The 12 sheets

| Sheet | What it holds |
|---|---|
| All Posts | Every post — date, views, likes, engagement. Filterable |
| Rankings | Top 30 by views / engagement / like rate / viral |
| Time of Day | 24 hours x weekday, **median** views and dead rate (share under 500 views) |
| Follower Demographics | Country / city / gender / age |
| Growth Insights | KPI summary, performance by media type, viral patterns |
| Viral Deep Dive | Threshold scaled to your account (P99, not a fixed number), view concentration |
| Like Rate Deep Dive | Distribution, top/bottom 30, median by media type |
| Content Optimization | Views vs engagement rate by post length, performance by topic tag |
| Monthly Trend | Month over month, quarterly summary |
| Daily Views | Up to 730 days of daily data, 7-day moving average, delivered-view months |
| Snapshot Growth | Follower history accumulated across runs — **the API will not give you this** |
| Path to 100K | Where you stand, top 5 actions, roadmap |

Two sheets deserve a note:

- **Daily Views** uses *delivery date*, not post date. Monthly Trend bins a post's lifetime views by when it was published, which structurally understates recent months. These two sheets answer different questions on purpose.
- **Snapshot Growth** exists because `followers_count` returns a scalar even when you pass `since`/`until`. There is no follower time series in the API. Local snapshots are the only history that exists, so run `analyze.py` on a schedule if you want this sheet to be useful.

### Refresh the token (every 60 days)

```bash
python3 refresh_token.py
```

`analyze.py` warns `TOKEN_WARN_DAYS` (default 7) before expiry and stops once expired. If it already expired, run `auth.py` again.

### Text archive

```bash
python3 archive.py
```

Writes every post body to `output/저장함/` as Markdown, JSONL and CSV, split by month and media type. Generated automatically during `analyze.py` unless you pass `--skip-archive`.

> The Threads API has **no endpoint for saved posts** (other people's posts you bookmarked). This archive covers your own posts only.

### Tests

```bash
python3 -m unittest discover -s tests -v
```

No network, no secrets.

## Known gaps

- Long cells hit the 60-character column-width cap and appear clipped. Click the cell to see the full value.
- Follower demographics require **100+ followers** — a Meta restriction, not a bug.
- Hour-of-day and length analyses withhold a verdict below 30 samples rather than guessing.

## API notes

- **Rate limit** — roughly 4,800 calls/hour. 429 and 5xx are retried with backoff
- **Token** — long-lived tokens last 60 days; `TOKEN_EXPIRES_AT` is recorded in `.env`
- **Insight cache** — v2 with TTL (`INSIGHTS_CACHE_TTL_DAYS`, default 7 days)
- **Unknown fields are silently dropped** by `GET /{user-id}/threads`. A 200 response is not proof a field exists — verify by counting key presence across real posts. `/me` does error on unknown fields; the `/threads` and `/replies` edges do not
- **Not available in the API**, despite what you might expect: `saves`, `impressions`, `reach`, `profile_visits`, and any daily follower time series

## Privacy

There is no server. Everything runs on your machine, calls the Meta API directly, and writes to `output/`. Your token lives in a local `.env` with mode 600. `.env`, the certificates and `output/` are gitignored.

## License

MIT
