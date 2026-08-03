# Sample Report

[한국어](../README.md) · [English](../README.en.md) · [日本語](../README.ja.md)

Excerpts from a report built with the bundled sample dataset. **You do not need a Meta app to reproduce this:**

```bash
git clone https://github.com/Hwemo-Chung/threads-analytics.git
cd threads-analytics
pip install -r requirements.txt
python3 export_excel.py -i samples/sample_analysis.json --lang en -o sample.xlsx
```

`samples/sample_analysis.json` is synthetic — post text is generated, the username is fake, and IDs
and permalinks are placeholders. Only the *shape* of the numbers is drawn from a real account, so the
analysis reads the way a real one would. 960 posts, 12 sheets.

Everything below is truncated. The real workbook has more rows and more sections per sheet.


## Time of Day — only the hours the tool is willing to judge

Verdicts are withheld below 30 posts in an hour. Note that a high *average* and a high *median* are not the same hour.

| Hour | Posts | Median Views | Avg Views | Dead Rate (%) | Median Likes | Avg Engagement | Verdict |
|---|---|---|---|---|---|---|---|
| 08:00 | 39 | 521 | 6200 | 48.72 | 3 | 14 | 🟢 Strong |
| 09:00 | 95 | 312 | 2682 | 67.37 | 3 | 37 | 🔴 Avoid |
| 10:00 | 75 | 277 | 1608 | 69.33 | 3 | 8.8 | 🔴 Avoid |
| 11:00 | 68 | 272 | 3370 | 66.18 | 3 | 50 | 🔴 Avoid |
| 16:00 | 47 | 633 | 1640 | 42.55 | 4 | 12 | 🟢 Strong |
| 17:00 | 40 | 630 | 5021 | 45 | 3 | 38.4 | 🟢 Strong |

## Content Optimization — reach and engagement pull in opposite directions

| Length Range | Posts | Median Views | Avg Views | Median ER (%) | Dead Rate (%) |
|---|---|---|---|---|---|
| 0-50 chars | 216 | 392 | 2525 | 0.8 | 56.94 |
| 50-100 chars | 215 | 674 | 4617 | 0.86 | 40 |
| 100-150 chars | 125 | 494 | 4197 | 0.88 | 52.8 |
| 150-200 chars | 64 | 554 | 2942 | 1.07 | 46.88 |
| 200-300 chars | 96 | 301 | 615 | 1.5 | 77.08 |
| 300-400 chars | 71 | 389 | 14189 | 1.71 | 60.56 |
| 400-500 chars | 109 | 181 | 413 | 1.71 | 82.57 |
| 500+ chars | 4 | 144 | 200 | 2.39 | 100 |
| Longer posts raise engagement rate but collapse reach (300-400 chars: ER 1.71% / 389 views, 50-100 chars: ER 0.86% / 674 views) — 1.7x reach gap |  |  |  |  |  |

## Viral Deep Dive — how much of the reach is one lucky post

| 6. View Concentration |  |
|---|---|
| Metric | Value |
| Original Posts | 900 |
| Avg Views | 3737 |
| Median Views | 398 |
| Mean / Median Ratio | 9.38 |
| P50 | 398 |
| P75 | 1016 |
| P90 | 3490 |
| P99 | 54317 |
| Viral Threshold (applied) | 54317 |
| Top 1% Share of Views (%) | 49.52 |

## Daily Views — 730-day series, by delivery date

| Date | Views | 7-Day Moving Avg | vs Same Day Last Week (%) |
|---|---|---|---|
| 2026-05-05 | 3108 | 34098 | -88.81 |
| 2026-05-06 | 4872 | 20797 | -95.03 |
| 2026-05-07 | 5135 | 10210 | -93.52 |
| 2026-05-08 | 2481 | 7053 | -89.9 |
| 2026-05-09 | 2872 | 5240 | -81.55 |
| 2026-05-10 | 9124 | 4586 | -33.41 |

## Snapshot Growth — follower history the API refuses to give you

| Snapshot Time | All Posts | Original Posts | Followers | 30d Views | 30d Likes | Note |
|---|---|---|---|---|---|---|
| 2026-04-27 14:28 | 1822 | 1561 | 1493 | 429264 | - | 30d likes missing |
| 2026-04-27 14:28 | 1822 | 1561 | - | 429264 | - | followers missing / 30d likes missing |
| 2026-06-13 05:37 | 2021 | 1725 | - | 490835 | - | followers missing / 30d likes missing |
| 2026-06-13 10:07 | 2022 | 1726 | 1614 | 494309 | 5107 | OK |
| 2026-06-14 19:15 | 2035 | 1738 | 1616 | 501060 | 5171 | OK |
| 2026-07-16 12:15 | 2197 | 1869 | 1652 | 469292 | 1194 | OK |

## What each sheet answers

| Sheet | Question it answers |
|---|---|
| All Posts | What do I actually have? |
| Rankings | What were my best posts, under four different definitions of "best"? |
| Time of Day | When should I post — judged on medians, not averages? |
| Follower Demographics | Who is actually following me? |
| Growth Insights | What is the one-page summary? |
| Viral Deep Dive | How much of my reach is luck, and where is the real threshold? |
| Like Rate Deep Dive | What do my existing fans respond to? |
| Content Optimization | How long should a post be, and which topics work? |
| Monthly Trend | Am I getting better or worse? |
| Daily Views | When were views *delivered*, rather than when were posts published? |
| Snapshot Growth | How are followers trending? |
| Path to 100K | What do I do on Monday morning? |

> This page is generated from the sample workbook, not written by hand.
> `tests/test_core.py` builds all three languages from `samples/sample_analysis.json`,
> so a broken sheet fails the suite before this page can go stale.

