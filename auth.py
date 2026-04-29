"""
Threads OAuth 인증 스크립트
- 자체 서명 SSL 인증서로 로컬 HTTPS 서버 구동
- Authorization Code → 단기 토큰 → 장기 토큰 자동 교환
- .env 파일에 ACCESS_TOKEN, USER_ID 자동 저장
"""

import http.server
import ssl
import urllib.parse
import webbrowser
import os
import sys
import json
import requests
from dotenv import load_dotenv, set_key

load_dotenv()

APP_ID = os.getenv("THREADS_APP_ID", "")
APP_SECRET = os.getenv("THREADS_APP_SECRET", "")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://localhost:8888/callback")
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
CERT_DIR = os.path.dirname(__file__)

SCOPES = "threads_basic,threads_manage_insights,threads_read_replies"

API_BASE = "https://graph.threads.net"


def validate_env():
    if not APP_ID or APP_ID == "your_app_id_here":
        print("[ERROR] .env 파일에 THREADS_APP_ID를 설정해주세요.")
        sys.exit(1)
    if not APP_SECRET or APP_SECRET == "your_app_secret_here":
        print("[ERROR] .env 파일에 THREADS_APP_SECRET를 설정해주세요.")
        sys.exit(1)


def get_auth_url():
    params = urllib.parse.urlencode({
        "client_id": APP_ID,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
        "response_type": "code",
        "state": "threads_auth",
    })
    return f"https://threads.net/oauth/authorize?{params}"


def exchange_code_for_short_token(code: str) -> dict:
    """Authorization Code → 단기 토큰 (1시간)"""
    resp = requests.post(
        f"{API_BASE}/oauth/access_token",
        data={
            "client_id": APP_ID,
            "client_secret": APP_SECRET,
            "grant_type": "authorization_code",
            "redirect_uri": REDIRECT_URI,
            "code": code,
        },
    )
    resp.raise_for_status()
    return resp.json()


def exchange_for_long_token(short_token: str) -> dict:
    """단기 토큰 → 장기 토큰 (60일)"""
    resp = requests.get(
        f"{API_BASE}/access_token",
        params={
            "grant_type": "th_exchange_token",
            "client_secret": APP_SECRET,
            "access_token": short_token,
        },
    )
    resp.raise_for_status()
    return resp.json()


class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    """OAuth redirect를 받는 로컬 HTTPS 핸들러"""

    auth_code = None

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return

        params = urllib.parse.parse_qs(parsed.query)

        if "error" in params:
            error = params.get("error_description", params["error"])[0]
            self.send_response(400)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(f"<h2>인증 실패: {error}</h2><p>터미널로 돌아가주세요.</p>".encode())
            OAuthCallbackHandler.auth_code = None
            return

        code = params.get("code", [None])[0]
        if not code:
            self.send_response(400)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write("<h2>인증 코드를 받지 못했습니다.</h2>".encode())
            return

        OAuthCallbackHandler.auth_code = code

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(
            "<h2>인증 성공! 이 창을 닫고 터미널로 돌아가주세요.</h2>".encode()
        )

    def log_message(self, format, *args):
        pass


def run_oauth_flow():
    validate_env()

    parsed_uri = urllib.parse.urlparse(REDIRECT_URI)
    port = parsed_uri.port or 8888

    cert_file = os.path.join(CERT_DIR, "localhost.crt")
    key_file = os.path.join(CERT_DIR, "localhost.key")

    if not os.path.exists(cert_file) or not os.path.exists(key_file):
        print("[ERROR] SSL 인증서가 없습니다. 아래 명령어로 생성해주세요:")
        print('  openssl req -x509 -newkey rsa:2048 -keyout localhost.key -out localhost.crt -days 365 -nodes -subj "/CN=localhost"')
        sys.exit(1)

    auth_url = get_auth_url()
    print(f"\n{'='*60}")
    print("Threads OAuth 인증 (HTTPS)")
    print(f"{'='*60}")
    print(f"\n1. 브라우저가 자동으로 열립니다.")
    print(f"2. Threads에서 로그인 & 권한 승인")
    print(f"3. 리디렉션 시 '안전하지 않음' 경고가 나오면:")
    print(f"   → '고급' → 'localhost(안전하지 않음)으로 이동' 클릭")
    print(f"\n열리지 않으면 아래 URL을 직접 브라우저에 붙여넣으세요:\n")
    print(auth_url)
    print(f"\n{'='*60}\n")

    webbrowser.open(auth_url)

    server = http.server.HTTPServer(("localhost", port), OAuthCallbackHandler)
    ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ssl_context.load_cert_chain(cert_file, key_file)
    server.socket = ssl_context.wrap_socket(server.socket, server_side=True)
    server.timeout = 5
    print(f"HTTPS 서버 대기 중 (localhost:{port})... Threads에서 인증해주세요.\n")

    import time
    deadline = time.time() + 300
    while OAuthCallbackHandler.auth_code is None and time.time() < deadline:
        try:
            server.handle_request()
        except Exception:
            pass
    server.server_close()

    code = OAuthCallbackHandler.auth_code
    if not code:
        print("[ERROR] 인증 코드를 받지 못했습니다. 다시 시도해주세요.")
        sys.exit(1)

    print(f"인증 코드 수신 완료.")

    print("단기 토큰 교환 중...")
    short_result = exchange_code_for_short_token(code)
    short_token = short_result["access_token"]
    user_id = str(short_result.get("user_id", ""))
    print(f"  User ID: {user_id}")
    print(f"  단기 토큰 발급 완료 (1시간 유효)")

    print("장기 토큰 교환 중...")
    long_result = exchange_for_long_token(short_token)
    long_token = long_result["access_token"]
    expires_in = long_result.get("expires_in", "unknown")
    print(f"  장기 토큰 발급 완료 ({expires_in}초 = 약 60일 유효)")

    set_key(ENV_PATH, "ACCESS_TOKEN", long_token)
    set_key(ENV_PATH, "USER_ID", user_id)

    print(f"\n{'='*60}")
    print(f".env 파일에 저장 완료!")
    print(f"  ACCESS_TOKEN = {long_token[:20]}...")
    print(f"  USER_ID = {user_id}")
    print(f"\n이제 analyze.py를 실행하세요:")
    print(f"  python analyze.py")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    run_oauth_flow()
