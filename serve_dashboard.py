#!/usr/bin/env python3
"""대시보드 정적 서버 + Overview overview_state.json 저장(POST) + ETF 개요 메일 발송."""
from __future__ import annotations

import json
import os
import shutil
import smtplib
import ssl
import tempfile
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = "overview_state.json"
EMP_EMAIL_TO = "chanhong.kang@shinhanamc.com"
EMP_EMAIL_API = "/api/emp_summary_email"
WEEKLY_EMAIL_API = "/api/weekly_report_email"


def _truthy(val: str | None) -> bool:
    return (val or "").strip().lower() in ("1", "true", "yes", "on")


def _outlook_available() -> bool:
    if os.name != "nt":
        return False
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except ImportError:
        return False
    return True


def _use_outlook_email() -> bool:
    if _truthy(os.environ.get("EMP_EMAIL_USE_OUTLOOK")):
        return True
    if os.environ.get("EMP_EMAIL_USE_OUTLOOK", "").strip().lower() in ("0", "false", "no", "off"):
        return False
    return _outlook_available()


def _wrap_html_body(html: str) -> str:
    if "<html" in html.lower():
        return html
    return (
        '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'
        + html
        + "</body></html>"
    )


def _send_via_smtp(
    *,
    html: str,
    subject: str,
    to_addr: str,
    attachments: list[tuple[str, bytes]] | None = None,
) -> None:
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

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = to_addr
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(_wrap_html_body(html), "html", "utf-8"))
    msg.attach(alt)
    for filename, data in attachments or []:
        part = MIMEApplication(data, Name=filename)
        part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", filename))
        msg.attach(part)

    if use_tls:
        with smtplib.SMTP(host, port, timeout=60) as smtp:
            smtp.ehlo()
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
            smtp.login(user, password)
            smtp.sendmail(from_addr, [to_addr], msg.as_string())
    else:
        with smtplib.SMTP_SSL(host, port, timeout=60, context=ssl.create_default_context()) as smtp:
            smtp.login(user, password)
            smtp.sendmail(from_addr, [to_addr], msg.as_string())


def send_emp_summary_email(*, html: str, subject: str) -> None:
    if not html or not html.strip():
        raise ValueError("Empty email body")
    subject = (subject or "ETF 개요").strip() or "ETF 개요"
    body_html = _wrap_html_body(html)

    if _use_outlook_email():
        _send_via_outlook(body_html, subject, EMP_EMAIL_TO)
        return
    _send_via_smtp(html=body_html, subject=subject, to_addr=EMP_EMAIL_TO)


def send_weekly_report_email(*, html: str, subject: str, filename: str) -> None:
    if not html or not html.strip():
        raise ValueError("Empty Word document body")
    subject = (subject or "시장 동향 및 전망").strip() or "시장 동향 및 전망"
    filename = os.path.basename((filename or "weekly_report.doc").strip()) or "weekly_report.doc"
    if not filename.lower().endswith(".doc"):
        filename += ".doc"

    doc_bytes = "\ufeff".encode("utf-8") + html.encode("utf-8")
    body_html = (
        "<p>시장 동향 및 전망 Word 파일(<strong>"
        + filename
        + "</strong>)을 첨부합니다.</p>"
    )

    temp_dir = tempfile.mkdtemp(prefix="weekly_report_")
    temp_path = os.path.join(temp_dir, filename)
    try:
        with open(temp_path, "wb") as tmp:
            tmp.write(doc_bytes)

        if _use_outlook_email():
            _send_via_outlook(body_html, subject, EMP_EMAIL_TO, attachment_paths=[temp_path])
            return
        _send_via_smtp(
            html=body_html,
            subject=subject,
            to_addr=EMP_EMAIL_TO,
            attachments=[(filename, doc_bytes)],
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _send_via_outlook(
    html: str,
    subject: str,
    to_addr: str,
    attachment_paths: list[str] | None = None,
) -> None:
    try:
        import pythoncom
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Outlook 발송에는 pywin32 가 필요합니다: pip install pywin32") from exc

    initialized = False
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
    except pythoncom.com_error:
        pass
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_addr
        mail.Subject = subject
        mail.HTMLBody = html
        for path in attachment_paths or []:
            mail.Attachments.Add(os.path.abspath(path))
        mail.Send()

        # Send()는 보낼편지함에 넣기만 함. Send/Receive로 서버 전송을 즉시 트리거.
        namespace = outlook.GetNamespace("MAPI")
        try:
            namespace.SendAndReceive(False)
        except Exception:
            for i in range(1, namespace.SyncObjects.Count + 1):
                namespace.SyncObjects.Item(i).Start()
    finally:
        if initialized:
            pythoncom.CoUninitialize()


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
        if path == WEEKLY_EMAIL_API:
            self._handle_weekly_report_email()
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

    def _handle_weekly_report_email(self) -> None:
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
        filename = parsed.get("filename")
        if filename is not None and not isinstance(filename, str):
            self.send_error(400, "filename must be string")
            return
        try:
            send_weekly_report_email(
                html=html,
                subject=subject or "시장 동향 및 전망",
                filename=filename or "weekly_report.doc",
            )
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": str(exc)})
            return
        self._send_json(200, {"ok": True, "to": EMP_EMAIL_TO})

    def log_message(self, fmt: str, *args) -> None:
        if args and (
            args[0].startswith("POST /overview_state")
            or args[0].startswith(f"POST {EMP_EMAIL_API}")
            or args[0].startswith(f"POST {WEEKLY_EMAIL_API}")
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
    print(f"Weekly Word 메일 API: http://{host}:{port}{WEEKLY_EMAIL_API}")
    if _use_outlook_email():
        print(f"메일 발송: Outlook → {EMP_EMAIL_TO}")
    elif os.environ.get("EMP_SMTP_HOST"):
        print(f"메일 발송: SMTP → {EMP_EMAIL_TO}")
    else:
        print(
            "메일 발송: 미설정 (Outlook: EMP_EMAIL_USE_OUTLOOK=1 또는 SMTP 환경 변수)"
        )
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n종료")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
