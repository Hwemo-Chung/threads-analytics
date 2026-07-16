"""장기 토큰 갱신 스크립트 — 만료 전 실행하여 60일 연장"""

import os
import sys
from datetime import datetime, timedelta, timezone

import requests
from dotenv import load_dotenv, set_key

load_dotenv()

ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "")
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
API_BASE = "https://graph.threads.net"
UTC = timezone.utc


def save_token_expiry(expires_in) -> str:
    if not isinstance(expires_in, (int, float)):
        return ""
    expires_at = datetime.now(UTC) + timedelta(seconds=int(expires_in))
    iso = expires_at.isoformat()
    set_key(ENV_PATH, "TOKEN_EXPIRES_AT", iso)
    return iso


def refresh_token():
    if not ACCESS_TOKEN:
        print("[ERROR] ACCESS_TOKEN이 없습니다.")
        sys.exit(1)

    raw = os.getenv("TOKEN_EXPIRES_AT", "").strip()
    if raw:
        try:
            expires_at = datetime.fromisoformat(raw.replace("Z", "+00:00"))
            if expires_at.tzinfo is None:
                expires_at = expires_at.replace(tzinfo=UTC)
            days_left = (expires_at.astimezone(UTC) - datetime.now(UTC)).total_seconds() / 86400
            print(f"현재 토큰 잔여 약 {days_left:.1f}일")
        except ValueError:
            pass

    print("장기 토큰 갱신 중...")
    try:
        resp = requests.get(
            f"{API_BASE}/refresh_access_token",
            params={
                "grant_type": "th_refresh_token",
                "access_token": ACCESS_TOKEN,
            },
            timeout=30,
        )
    except requests.exceptions.RequestException as exc:
        print(f"[ERROR] 네트워크 오류: {exc}")
        sys.exit(1)

    if resp.status_code != 200:
        print(f"[ERROR] 갱신 실패: {resp.status_code} - {resp.text}")
        print("토큰이 만료되었을 수 있습니다. auth.py를 다시 실행해주세요.")
        sys.exit(1)

    data = resp.json()
    new_token = data["access_token"]
    expires_in = data.get("expires_in", "unknown")

    set_key(ENV_PATH, "ACCESS_TOKEN", new_token)
    expires_iso = save_token_expiry(expires_in)

    print(f"갱신 완료! (유효기간: {expires_in}초 = 약 60일)")
    print(f"새 토큰: {new_token[:20]}...")
    if expires_iso:
        print(f"TOKEN_EXPIRES_AT = {expires_iso}")


if __name__ == "__main__":
    refresh_token()
