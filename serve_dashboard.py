#!/usr/bin/env python3
"""대시보드 정적 서버 + Overview overview_state.json 저장(POST) + ETF 개요 메일 발송."""
from __future__ import annotations

import json
import os
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = "overview_state.json"
EMP_EMAIL_TO = "chanhong.kang@shinhanamc.com"
EMP_EMAIL_API = "/api/emp_summary_email"


def _truthy(val: str | None) -> bool:
    return (val or "").strip().lower() in ("1", "true", "yes", "on")


def send_emp_summary_email(*, html: str, subject: str) -> None:
    if not html or not html.strip():
        raise ValueError("Empty email body")
    subject = (subject or "ETF 개요").strip() or "ETF 개요"
    body_html = html if "<html" in html.lower() else (
        "<!DOCTYPE html><html><head><meta charset=\"utf-8\"></head><body>"
        + html
        + "</body></html>"
    )

    if _truthy(os.environ.get("EMP_EMAIL_USE_OUTLOOK")):
        _send_via_outlook(body_html, subject, EMP_EMAIL_TO)
        return

    host = os.environ.get("EMP_SMTP_HOST", "").strip()
    user = os.environ.get("EMP_SMTP_USER", "").strip()
    password = os.environ.get("EMP_SMTP_PASS", "")
    if not host or not user or not password:
        raise RuntimeError(
            "메일 설정 없음: EMP_SMTP_HOST/USER/PASS 환경 변수를 설정하거나 "
            "EMP_EMAIL_USE_OUTLOOK=1 로 Outlook 발송을 사용하세요."
        )

    port = int(os.environ.get("EMP_SMTP_PORT", "587"))
    use_tls = not _truthy(os.environ.get("EMP_SMTP_SSL"))
    from_addr = os.environ.get("EMP_SMTP_FROM", user).strip() or user

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = EMP_EMAIL_TO
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    if use_tls:
        with smtplib.SMTP(host, port, timeout=60) as smtp:
            smtp.ehlo()
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
            smtp.login(user, password)
            smtp.sendmail(from_addr, [EMP_EMAIL_TO], msg.as_string())
    else:
        with smtplib.SMTP_SSL(host, port, timeout=60, context=ssl.create_default_context()) as smtp:
            smtp.login(user, password)
            smtp.sendmail(from_addr, [EMP_EMAIL_TO], msg.as_string())


def _send_via_outlook(html: str, subject: str, to_addr: str) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Outlook 발송에는 pywin32 가 필요합니다: pip install pywin32") from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addr
    mail.Subject = subject
    mail.HTMLBody = html
    mail.Send()


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def do_POST(self) -> None:
        path = self.path.split("?", 1)[0]
        if path in ("/overview_state.json", "/api/overview_state.json"):
            self._handle_overview_state_post()
            return
        if path == EMP_EMAIL_API:
            self._handle_emp_summary_email()
            return
        self.send_error(404)

    def _read_json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        try:
            parsed = json.loads(body.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            raise ValueError("Invalid JSON") from exc
        if not isinstance(parsed, dict):
            raise ValueError("JSON object required")
        return parsed

    def _send_json(self, status: int, payload: dict) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _handle_overview_state_post(self) -> None:
        try:
            parsed = self._read_json_body()
        except ValueError:
            self.send_error(400, "Invalid JSON")
            return
        out_path = os.path.join(ROOT, STATE_FILE)
        with open(out_path, "w", encoding="utf-8", newline="\n") as f:
            json.dump(parsed, f, ensure_ascii=False, indent=2)
            f.write("\n")
        self._send_json(200, {"ok": True})

    def _handle_emp_summary_email(self) -> None:
        try:
            parsed = self._read_json_body()
        except ValueError:
            self.send_error(400, "Invalid JSON")
            return
        html = parsed.get("html")
        if not isinstance(html, str) or not html.strip():
            self.send_error(400, "html required")
            return
        subject = parsed.get("subject")
        if subject is not None and not isinstance(subject, str):
            self.send_error(400, "subject must be string")
            return
        try:
            send_emp_summary_email(html=html, subject=subject or "ETF 개요")
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": str(exc)})
            return
        self._send_json(200, {"ok": True, "to": EMP_EMAIL_TO})

    def log_message(self, fmt: str, *args) -> None:
        if args and (
            args[0].startswith("POST /overview_state")
            or args[0].startswith(f"POST {EMP_EMAIL_API}")
        ):
            return
        super().log_message(fmt, *args)


def main() -> None:
    port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "127.0.0.1")
    server = ThreadingHTTPServer((host, port), DashboardHandler)
    print(f"Dashboard: http://{host}:{port}/dashboard.html")
    print(f"Overview 저장 파일: {os.path.join(ROOT, STATE_FILE)}")
    print(f"ETF 개요 메일 API: http://{host}:{port}{EMP_EMAIL_API}")
    if _truthy(os.environ.get("EMP_EMAIL_USE_OUTLOOK")):
        print("ETF 개요 메일: Outlook(EMP_EMAIL_USE_OUTLOOK=1)")
    elif os.environ.get("EMP_SMTP_HOST"):
        print(f"ETF 개요 메일: SMTP → {EMP_EMAIL_TO}")
    else:
        print(
            "ETF 개요 메일: 미설정 (Outlook: EMP_EMAIL_USE_OUTLOOK=1 또는 SMTP 환경 변수)"
        )
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n종료")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
