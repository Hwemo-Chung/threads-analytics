"""Core unit tests — no network, no secrets."""

import json
import os
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import analyze
import auth
import export_excel


class TestInsightTotal(unittest.TestCase):
    def test_total_value(self):
        item = {"name": "followers_count", "total_value": {"value": 1234}}
        self.assertEqual(analyze._insight_total(item), 1234)

    def test_values_sum(self):
        item = {"values": [{"value": 10}, {"value": 20}]}
        self.assertEqual(analyze._insight_total(item), 30)

    def test_empty(self):
        # 의도 변경: 예전엔 0을 돌려줘 결측이 진짜 0으로 저장됐다. 이제 결측은 None.
        self.assertIsNone(analyze._insight_total({}))
        self.assertIsNone(analyze._insight_total({"values": []}))
        self.assertEqual(analyze._insight_total({"total_value": {"value": 0}}), 0)


class TestInsightsCache(unittest.TestCase):
    def test_v1_migrate_and_ttl(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "insights_cache.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"p1": {"views": 1, "likes": 2}}, f)

            entries = analyze._load_insights_cache(path)
            self.assertIn("p1", entries)
            self.assertEqual(entries["p1"]["metrics"]["views"], 1)
            self.assertFalse(analyze._is_cache_fresh(entries["p1"], 7))

            now = datetime.now(timezone.utc).isoformat()
            entries["p1"]["fetched_at"] = now
            self.assertTrue(analyze._is_cache_fresh(entries["p1"], 7))

            analyze._save_insights_cache(path, entries)
            with open(path, encoding="utf-8") as f:
                raw = json.load(f)
            self.assertEqual(raw["version"], 2)
            self.assertIn("entries", raw)

    def test_ttl_zero_never_expires(self):
        entry = {"fetched_at": "1970-01-01T00:00:00+00:00", "metrics": {}}
        self.assertTrue(analyze._is_cache_fresh(entry, 0))


class TestApiRetry(unittest.TestCase):
    def setUp(self):
        analyze.reset_api_stats()
        analyze.ACCESS_TOKEN = "test-token"

    @patch("analyze.requests.get")
    def test_retry_on_429_then_ok(self, mock_get):
        bad = MagicMock()
        bad.status_code = 429
        bad.headers = {"Retry-After": "0"}
        bad.text = "rate"
        good = MagicMock()
        good.status_code = 200
        good.json.return_value = {"data": []}
        mock_get.side_effect = [bad, good]

        with patch("analyze.time.sleep"):
            result = analyze.api_get("me", max_retries=2)

        self.assertEqual(result, {"data": []})
        self.assertEqual(mock_get.call_count, 2)
        self.assertGreaterEqual(analyze._api_stats["retry"], 1)
        self.assertEqual(analyze._api_stats["ok"], 1)

    @patch("analyze.requests.get")
    def test_error_counted(self, mock_get):
        bad = MagicMock()
        bad.status_code = 400
        bad.text = "bad request"
        mock_get.return_value = bad

        result = analyze.api_get("me", max_retries=0)
        self.assertEqual(result, {})
        self.assertEqual(analyze._api_stats["error"], 1)


class TestOAuthState(unittest.TestCase):
    def test_auth_url_embeds_state(self):
        with patch.object(auth, "APP_ID", "app123"):
            url = auth.get_auth_url("secret-state-xyz")
        self.assertIn("state=secret-state-xyz", url)
        self.assertIn("client_id=app123", url)


class TestGrowthStrategies(unittest.TestCase):
    def test_data_driven_no_hardcoded_account(self):
        strats = export_excel._build_growth_strategies(
            active=[{
                "media_type": "TEXT_POST",
                "timestamp": "2026-01-01T00:00:00+0000",
                "insights": {"views": 1000, "likes": 10, "replies": 1, "reposts": 0, "quotes": 0},
                "text": "hello?",
            }],
            type_stats={
                "TEXT_POST": {"count": 1, "views": 1000, "likes": 10, "engagement": 11},
                "CAROUSEL_ALBUM": {"count": 1, "views": 5000, "likes": 50, "engagement": 80},
            },
            viral_posts=[{"media_type": "CAROUSEL_ALBUM", "insights": {"views": 20000}}],
            demographics={"age": {"25-34": 10}, "gender": {"F": 8}, "country": {"JP": 9}},
            total_views=6000,
            total_likes=60,
            total_replies=2,
            total_reposts=0,
            total_quotes=0,
        )
        joined = "\n".join(strats)
        self.assertIn("캐러셀", joined)
        self.assertIn("25-34", joined)
        self.assertNotIn("105.8", joined)
        self.assertNotIn("35-44", joined)


class TestExportCli(unittest.TestCase):
    def test_build_workbook_sheet_count(self):
        data = {
            "profile": {"username": "testuser"},
            "posts": [{
                "id": "1",
                "media_type": "TEXT_POST",
                "text": "hi",
                "timestamp": "2026-01-15T10:00:00+0000",
                "permalink": "https://example.com",
                "insights": {
                    "views": 100, "likes": 5, "replies": 1,
                    "reposts": 0, "quotes": 0, "shares": 0,
                },
            }],
            "user_insights": {
                "followers_count": 500,
                "30d_views": 1000,
                "daily_views": [
                    {"date": f"2026-0{1 + i // 28}-{i % 28 + 1:02d}", "value": i * 10}
                    for i in range(100)
                ],
            },
            "follower_demographics": {"country": {"KR": 100}},
        }
        # load_snapshots()는 실제 output/ 디렉터리를 읽으므로 고정해야 테스트가 결정적이다
        with patch.object(export_excel, "load_snapshots", return_value=[]):
            self.assertEqual(len(export_excel.build_workbook(data).sheetnames), 11)
        # 스냅샷 2개 이상일 때만 '스냅샷 성장추이' 시트가 붙는다
        snap = {
            "label": "2026-01-01 00:00", "dt": datetime(2026, 1, 1), "posts_total": 1,
            "reposts": 0, "active": 1, "followers": 10, "views_30d": 100,
            "likes_30d": 10, "flags": [], "by_id": {},
        }
        with patch.object(export_excel, "load_snapshots", return_value=[snap, dict(snap, dt=datetime(2026, 2, 1), label="2026-02-01 00:00")]):
            wb = export_excel.build_workbook(data)
        self.assertEqual(len(wb.sheetnames), 12)
        self.assertIn("스냅샷 성장추이", wb.sheetnames)
        self.assertIn("전체 게시물", wb.sheetnames)
        self.assertIn("일별 조회수 추이", wb.sheetnames)
        self.assertIn("10만 성장전략", wb.sheetnames)


class TestLongitudinal(unittest.TestCase):
    def test_missing_follower_snapshot_never_diffed(self):
        def snap(label, followers, dt):
            return {
                "label": label, "dt": dt, "posts_total": 10, "reposts": 1, "active": 9,
                "followers": followers, "views_30d": 100, "likes_30d": 0,
                "flags": [] if followers > 0 else ["팔로워 데이터 결측"], "by_id": {},
            }

        wb = export_excel.openpyxl.Workbook()
        wb.remove(wb.active)
        # 0일 구간은 이제 생략되므로 fixture에 실제 날짜를 준다 (검증 의도는 동일)
        export_excel.sheet_longitudinal(wb, [
            snap("A", 1493, datetime(2026, 4, 27)),
            snap("B", 0, datetime(2026, 6, 13)),
            snap("C", 1614, datetime(2026, 7, 16)),
        ])
        ws = wb["스냅샷 성장추이"]
        col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        hdr = col_a.index("구간") + 1  # 2. 구간별 성장률 헤더 행
        # 결측 스냅샷을 끼고 diff하면 -1493 같은 가짜 급감이 나온다 → 두 구간 모두 '-'
        self.assertEqual([ws.cell(row=hdr + 1, column=4).value, ws.cell(row=hdr + 2, column=4).value], ["-", "-"])
        self.assertIn("데이터 없음", [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)])


class TestTokenWarn(unittest.TestCase):
    def test_expired_exits(self):
        past = (datetime.now(timezone.utc) - timedelta(days=1)).isoformat()
        with patch.dict(os.environ, {"TOKEN_EXPIRES_AT": past}):
            with self.assertRaises(SystemExit) as ctx:
                analyze.warn_token_expiry(warn_days=7)
            self.assertEqual(ctx.exception.code, 1)


class TestArchive(unittest.TestCase):
    def test_save_text_archive(self):
        import archive

        data_posts = [
            {
                "id": "1",
                "media_type": "TEXT_POST",
                "text": "첫번째 글 본문입니다",
                "timestamp": "2026-01-15T10:00:00+0000",
                "permalink": "https://example.com/1",
                "insights": {"views": 100, "likes": 5, "replies": 1, "reposts": 0, "quotes": 0},
            },
            {
                "id": "2",
                "media_type": "TEXT_POST",
                "text": "두번째 글",
                "timestamp": "2026-02-01T08:00:00+0000",
                "permalink": "https://example.com/2",
                "insights": {"views": 50, "likes": 2, "replies": 0, "reposts": 0, "quotes": 0},
            },
            {
                "id": "3",
                "media_type": "REPOST_FACADE",
                "text": "리포스트",
                "timestamp": "2026-02-02T08:00:00+0000",
                "insights": {},
            },
        ]
        with tempfile.TemporaryDirectory() as tmp:
            snap = archive.save_text_archive(
                profile={"username": "tester"},
                posts=data_posts,
                user_insights={"followers_count": 10},
                archive_root=tmp,
                source_path="fake.json",
            )
            self.assertTrue(os.path.isdir(snap))
            self.assertTrue(os.path.isfile(os.path.join(snap, "all_texts.md")))
            self.assertTrue(os.path.isfile(os.path.join(snap, "all_posts.jsonl")))
            self.assertTrue(os.path.isfile(os.path.join(snap, "index.csv")))
            self.assertTrue(os.path.isfile(os.path.join(tmp, "latest", "all_texts.md")))
            with open(os.path.join(snap, "all_texts.md"), encoding="utf-8") as f:
                body = f.read()
            self.assertIn("첫번째 글 본문입니다", body)
            # REPOST_FACADE posts excluded from md (metric label "리포스트" may still appear)
            self.assertNotIn("] REPOST_FACADE", body)
            self.assertNotIn("\n리포스트\n", body)
            with open(os.path.join(snap, "all_posts.jsonl"), encoding="utf-8") as f:
                lines = f.readlines()
            self.assertEqual(len(lines), 3)


if __name__ == "__main__":
    unittest.main()
