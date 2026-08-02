"""
Threads Analyzer 설치 도우미 — `python3 setup.py` 한 번으로 끝난다.

Meta 앱 생성(THREADS_APP_ID / THREADS_APP_SECRET 발급)만 사람이 해야 하고,
나머지(의존성 · SSL 인증서 · .env · 인증)는 전부 여기서 처리한다.
"""

import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(ROOT, ".env")
CRT_PATH = os.path.join(ROOT, "localhost.crt")
KEY_PATH = os.path.join(ROOT, "localhost.key")
REDIRECT_URI = "https://localhost:8888/callback"

# Meta 앱 대시보드에서 복사하는 값들. ID는 숫자, Secret은 32자 hex.
APP_ID_RE = re.compile(r"^\d{10,}$")
APP_SECRET_RE = re.compile(r"^[0-9a-f]{32}$", re.I)

STEP = 0


def say(msg):
    print(msg)


def step(title):
    global STEP
    STEP += 1
    print(f"\n[{STEP}] {title}")


def die(msg):
    print(f"\n[중단] {msg}")
    sys.exit(1)


def check_python():
    step("Python 버전 확인")
    if sys.version_info < (3, 8):
        die(f"Python 3.8 이상이 필요합니다. 현재 {sys.version.split()[0]}\n"
            "      macOS 기본 python은 2.7일 수 있으니 python3 로 실행해주세요.")
    say(f"    OK — Python {sys.version.split()[0]}")


def install_deps():
    step("의존성 설치")
    req = os.path.join(ROOT, "requirements.txt")
    if not os.path.exists(req):
        die("requirements.txt 를 찾을 수 없습니다. 압축을 푼 폴더에서 실행해주세요.")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "-q", "-r", req],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        tail = (result.stderr or result.stdout).strip().splitlines()[-5:]
        die("pip 설치 실패:\n      " + "\n      ".join(tail))
    say("    OK — requests, python-dotenv, tabulate, openpyxl")


def make_cert():
    step("로컬 HTTPS 인증서 생성")
    if os.path.exists(CRT_PATH) and os.path.exists(KEY_PATH):
        say("    건너뜀 — 인증서가 이미 있습니다")
        return
    result = subprocess.run(
        ["openssl", "req", "-x509", "-newkey", "rsa:2048",
         "-keyout", KEY_PATH, "-out", CRT_PATH,
         "-days", "365", "-nodes", "-subj", "/CN=localhost"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        die("openssl 실행 실패. openssl 설치 후 다시 실행해주세요.\n"
            f"      {(result.stderr or '').strip()[:200]}")
    os.chmod(KEY_PATH, 0o600)
    say("    OK — localhost.crt / localhost.key")


def prompt_credentials():
    """이미 유효한 값이 .env에 있으면 묻지 않는다."""
    step("Meta 앱 정보 입력")
    existing = {}
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, encoding="utf-8") as f:
            for line in f:
                if "=" in line and not line.lstrip().startswith("#"):
                    k, _, v = line.partition("=")
                    existing[k.strip()] = v.strip().strip("'\"")

    app_id = existing.get("THREADS_APP_ID", "")
    app_secret = existing.get("THREADS_APP_SECRET", "")
    if APP_ID_RE.match(app_id) and APP_SECRET_RE.match(app_secret):
        say("    건너뜀 — .env 에 이미 앱 정보가 있습니다")
        return existing

    say("    Meta 앱 대시보드 > App settings > Basic 에서 복사해주세요.")
    say("    https://developers.facebook.com/apps/")
    app_id = ask("    Threads App ID", app_id, APP_ID_RE,
                 "숫자만 10자리 이상이어야 합니다")
    app_secret = ask("    Threads App Secret", app_secret, APP_SECRET_RE,
                     "32자리 16진수여야 합니다")
    existing["THREADS_APP_ID"] = app_id
    existing["THREADS_APP_SECRET"] = app_secret
    return existing


def ask(label, current, pattern, hint):
    suffix = " (엔터=기존값 유지)" if current else ""
    while True:
        try:
            value = input(f"{label}{suffix}: ").strip()
        except (EOFError, KeyboardInterrupt):
            die("입력이 취소되었습니다.")
        if not value and current:
            value = current
        if pattern.match(value):
            return value
        say(f"      형식이 올바르지 않습니다 — {hint}")


def write_env(values):
    step(".env 저장")
    defaults = {
        "REDIRECT_URI": REDIRECT_URI,
        "INSIGHTS_CACHE_TTL_DAYS": "7",
        "API_MAX_RETRIES": "3",
        "TOKEN_WARN_DAYS": "7",
    }
    for k, v in defaults.items():
        values.setdefault(k, v)
    for k in ("ACCESS_TOKEN", "USER_ID", "TOKEN_EXPIRES_AT"):
        values.setdefault(k, "")

    order = ["THREADS_APP_ID", "THREADS_APP_SECRET", "REDIRECT_URI",
             "ACCESS_TOKEN", "USER_ID", "TOKEN_EXPIRES_AT",
             "INSIGHTS_CACHE_TTL_DAYS", "API_MAX_RETRIES", "TOKEN_WARN_DAYS"]
    lines = [f"{k}={values[k]}" for k in order if k in values]
    lines += [f"{k}={v}" for k, v in values.items() if k not in order]
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    os.chmod(ENV_PATH, 0o600)
    say("    OK — .env (권한 600)")


def show_callback_checklist():
    step("Meta 앱 대시보드 설정 확인")
    say("    Threads API > Settings 에 아래 3개가 모두 등록되어 있어야 합니다.")
    say(f"      Callback URL       : {REDIRECT_URI}")
    say("      Deauthorize URL    : https://localhost:8888/deauthorize")
    say("      Data Deletion URL  : https://localhost:8888/delete")
    say("    Threads API > Permissions 에서 본인 계정을 테스터로 추가하고,")
    say("    Threads 앱에서 초대를 수락해야 합니다.")
    try:
        input("    완료하셨으면 엔터를 눌러주세요: ")
    except (EOFError, KeyboardInterrupt):
        die("입력이 취소되었습니다.")


def run_auth():
    step("Threads 인증")
    say("    브라우저가 열립니다. 로그인 후 권한을 승인해주세요.")
    say("    브라우저 보안 경고가 뜨면 '고급' > 'localhost(안전하지 않음)으로 이동'을 눌러주세요.")
    result = subprocess.run([sys.executable, os.path.join(ROOT, "auth.py")])
    if result.returncode != 0:
        die("인증에 실패했습니다. 위 오류를 확인한 뒤 python3 auth.py 로 다시 시도해주세요.")


def main():
    say("=" * 60)
    say("Threads Analyzer 설치")
    say("=" * 60)
    check_python()
    install_deps()
    make_cert()
    values = prompt_credentials()
    write_env(values)
    show_callback_checklist()
    run_auth()
    say("\n" + "=" * 60)
    say("설치 완료. 이제 아래 순서로 실행하세요.")
    say("=" * 60)
    say("  python3 analyze.py --export-excel")
    say("\n  전체 게시물 수집에 약 10분이 걸립니다.")
    say("  결과: output/threads_analysis_YYYYMMDD.xlsx (12시트)")


if __name__ == "__main__":
    main()
