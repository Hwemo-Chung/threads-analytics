"""장기 토큰 갱신 스크립트 — 만료 전 실행하여 60일 연장"""

import os
import sys

import requests
from dotenv import load_dotenv, set_key

load_dotenv()

ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "")
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
API_BASE = "https://graph.threads.net"


def refresh_token():
    if not ACCESS_TOKEN:
        print("[ERROR] ACCESS_TOKEN이 없습니다.")
        sys.exit(1)

    print("장기 토큰 갱신 중...")
    resp = requests.get(
        f"{API_BASE}/refresh_access_token",
        params={
            "grant_type": "th_refresh_token",
            "access_token": ACCESS_TOKEN,
        },
    )

    if resp.status_code != 200:
        print(f"[ERROR] 갱신 실패: {resp.status_code} - {resp.text}")
        print("토큰이 만료되었을 수 있습니다. auth.py를 다시 실행해주세요.")
        sys.exit(1)

    data = resp.json()
    new_token = data["access_token"]
    expires_in = data.get("expires_in", "unknown")

    set_key(ENV_PATH, "ACCESS_TOKEN", new_token)

    print(f"갱신 완료! (유효기간: {expires_in}초 = 약 60일)")
    print(f"새 토큰: {new_token[:20]}...")


if __name__ == "__main__":
    refresh_token()
