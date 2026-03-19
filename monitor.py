import os
import sys
import time
import smtplib
import argparse
import requests
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sys.stdout.reconfigure(line_buffering=True)

SMTP_HOST     = os.getenv("SMTP_HOST", "smtp.office365.com")
SMTP_PORT     = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER     = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
ALERT_TO      = os.getenv("ALERT_EMAIL_TO", "")

BACKEND_HEALTH_URL = os.getenv(
    "BACKEND_HEALTH_URL",
    "http://192.168.10.30:5001/health"
)
CHECK_INTERVAL    = int(os.getenv("CHECK_INTERVAL", "60"))
FAILURE_THRESHOLD = int(os.getenv("FAILURE_THRESHOLD", "3"))


def check_required_env():
    missing = [k for k, v in {
        "SMTP_USER": SMTP_USER,
        "SMTP_PASSWORD": SMTP_PASSWORD,
        "ALERT_EMAIL_TO": ALERT_TO,
    }.items() if not v]
    if missing:
        print(f"[ERROR] 환경변수 미설정: {', '.join(missing)}", flush=True)
        print("  SMTP_USER, SMTP_PASSWORD, ALERT_EMAIL_TO 를 설정하세요.", flush=True)
        return False
    return True


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def to_html(text):
    import html as h
    return h.escape(text).replace("\n", "<br>\n")


def send_email(subject, body):
    try:
        html_body = f"""<html>
<body style="font-family:'Malgun Gothic',Arial,sans-serif;font-size:14px;color:#222;line-height:1.8;">
<pre style="font-family:'Malgun Gothic',Arial,sans-serif;font-size:14px;line-height:1.8;white-space:pre-wrap;">{to_html(body)}</pre>
</body>
</html>"""

        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = SMTP_USER
        msg["To"]      = ALERT_TO
        msg.attach(MIMEText(body,     "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html",  "utf-8"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, ALERT_TO, msg.as_string())

        print(f"[{now()}] 메일 발송 완료 → {ALERT_TO}", flush=True)
        return True
    except Exception as e:
        print(f"[{now()}] 메일 발송 실패: {e}", flush=True)
        return False


def simplify_error(error_msg):
    if not error_msg:
        return "알 수 없는 오류"
    msg = str(error_msg)
    if "timed out" in msg or "Timeout" in msg or "timeout" in msg:
        return "연결 타임아웃 (10초 초과)"
    if "Connection refused" in msg or "ConnectionRefused" in msg:
        return "연결 거부 (서비스 미실행)"
    if "ConnectionError" in msg or "연결 실패" in msg:
        return "연결 실패 (서버 미응답)"
    if msg.startswith("HTTP "):
        return msg
    return msg[:80] + ("..." if len(msg) > 80 else "")


def build_down_body(error_msg, fail_count):
    return (
        f"⚠️  백엔드 서버 다운 감지\n"
        f"{'=' * 50}\n"
        f"시각      : {now()}\n"
        f"서버 주소  : {BACKEND_HEALTH_URL}\n"
        f"서버 URL   : nms.vlan24.co.kr\n"
        f"연속 실패  : {fail_count}회 ({fail_count * CHECK_INTERVAL // 60}분 이상)\n"
        f"오류 내용  : {error_msg}\n"
        f"{'=' * 50}\n"
        f"Zabbix 백엔드 서버(192.168.10.30)를 확인하세요.\n"
        f"서비스 재시작: sudo systemctl restart ppt-backend\n"
        f"로그 확인   : sudo journalctl -u ppt-backend -f\n"
    )


def build_recovery_body(downtime_secs):
    minutes, seconds = divmod(downtime_secs, 60)
    return (
        f"✅  백엔드 서버 복구 감지\n"
        f"{'=' * 50}\n"
        f"시각      : {now()}\n"
        f"서버 주소  : {BACKEND_HEALTH_URL}\n"
        f"서버 URL   : nms.vlan24.co.kr\n"
        f"다운 시간  : {int(minutes)}분 {int(seconds)}초\n"
        f"{'=' * 50}\n"
        f"Zabbix 백엔드 서버가 정상 응답을 시작했습니다.\n"
    )


def health_check():
    try:
        resp = requests.get(BACKEND_HEALTH_URL, timeout=10)
        if resp.status_code == 200:
            return True, None
        return False, f"HTTP {resp.status_code}"
    except requests.exceptions.ConnectionError:
        return False, "연결 실패 (서버 미응답)"
    except requests.exceptions.Timeout:
        return False, "연결 타임아웃 (10초 초과)"
    except Exception as e:
        return False, simplify_error(str(e))


def run_test():
    print("=" * 55, flush=True)
    print("  monitor.py 테스트 모드", flush=True)
    print("=" * 55, flush=True)

    if not check_required_env():
        return

    subject_down     = "[PPT 백엔드] 🚨 서버 다운 감지"
    body_down        = build_down_body("테스트 발송 (실제 오류 아님)", 3)
    subject_recovery = "[PPT 백엔드] ✅ 서버 복구"
    body_recovery    = build_recovery_body(330)

    print("\n[ 발송할 메일 내용 미리보기 ]\n", flush=True)
    print(f"▶ 제목: {subject_down}", flush=True)
    print(body_down, flush=True)
    print(f"▶ 제목: {subject_recovery}", flush=True)
    print(body_recovery, flush=True)

    print("-" * 55, flush=True)
    print(f"수신자  : {ALERT_TO}", flush=True)
    print(f"SMTP    : {SMTP_HOST}:{SMTP_PORT}", flush=True)
    print(f"발신자  : {SMTP_USER}", flush=True)
    print("-" * 55, flush=True)
    print("테스트 메일 발송 중... (다운 알림 + 복구 알림 각 1통)", flush=True)

    ok1 = send_email(subject=subject_down,     body=body_down)
    ok2 = send_email(subject=subject_recovery, body=body_recovery)

    if ok1 and ok2:
        print("✅ 테스트 메일 2통 발송 성공!", flush=True)
        print(f"   → {ALERT_TO} 받은편지함을 확인하세요.", flush=True)
    else:
        print("❌ 일부 메일 발송 실패 — SMTP 설정을 확인하세요.", flush=True)


def run():
    if not check_required_env():
        return

    print(f"[{now()}] 백엔드 헬스 모니터링 시작", flush=True)
    print(f"  대상 URL      : {BACKEND_HEALTH_URL}", flush=True)
    print(f"  체크 주기     : {CHECK_INTERVAL}초", flush=True)
    print(f"  알림 임계값   : {FAILURE_THRESHOLD}회 연속 실패", flush=True)
    print(f"  알림 수신자   : {ALERT_TO}", flush=True)
    print(f"  SMTP          : {SMTP_HOST}:{SMTP_PORT}", flush=True)
    print("-" * 50, flush=True)

    fail_count = 0
    is_down    = False
    down_since = None

    while True:
        ok, err = health_check()

        if ok:
            if is_down:
                downtime = time.time() - down_since
                print(f"[{now()}] ✅ 서버 복구 감지 (다운 시간: {int(downtime)}초)", flush=True)
                send_email(
                    subject="[PPT 백엔드] ✅ 서버 복구",
                    body=build_recovery_body(downtime)
                )
                is_down    = False
                down_since = None
            fail_count = 0
            print(f"[{now()}] ✅ 정상", flush=True)
        else:
            fail_count += 1
            print(f"[{now()}] ❌ 헬스체크 실패 ({fail_count}회) — {err}", flush=True)

            if fail_count >= FAILURE_THRESHOLD and not is_down:
                is_down    = True
                down_since = time.time()
                print(f"[{now()}] 🚨 알림 메일 발송 중...", flush=True)
                send_email(
                    subject="[PPT 백엔드] 🚨 서버 다운 감지",
                    body=build_down_body(err, fail_count)
                )

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PPT 백엔드 헬스체크 모니터")
    parser.add_argument(
        "--test",
        action="store_true",
        help="테스트 메일 발송 후 종료 (실제 모니터링 루프 없음)"
    )
    args = parser.parse_args()

    if args.test:
        run_test()
    else:
        run()
