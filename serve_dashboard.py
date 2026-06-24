#!/usr/bin/env python3
"""대시보드 정적 서버 + Overview/EMP 상태 JSON 저장(POST)."""
from __future__ import annotations

import json
import os
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = "overview_state.json"
EMP_STATE_FILE = "emp_state.json"


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def do_POST(self) -> None:
        path = self.path.split("?", 1)[0]
        if path in ("/overview_state.json", "/api/overview_state.json"):
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            try:
                parsed = json.loads(body.decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError):
                self.send_error(400, "Invalid JSON")
                return
            out_path = os.path.join(ROOT, STATE_FILE)
            with open(out_path, "w", encoding="utf-8", newline="\n") as f:
                json.dump(parsed, f, ensure_ascii=False, indent=2)
                f.write("\n")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
            return
        if path in ("/emp_state.json", "/api/emp_state.json"):
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            try:
                parsed = json.loads(body.decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError):
                self.send_error(400, "Invalid JSON")
                return
            out_path = os.path.join(ROOT, EMP_STATE_FILE)
            with open(out_path, "w", encoding="utf-8", newline="\n") as f:
                json.dump(parsed, f, ensure_ascii=False, indent=2)
                f.write("\n")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
            return
        self.send_error(404)

    def log_message(self, fmt: str, *args) -> None:
        if args and (
            args[0].startswith("POST /overview_state") or args[0].startswith("POST /emp_state")
        ):
            return
        super().log_message(fmt, *args)


def main() -> None:
    port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "127.0.0.1")
    server = ThreadingHTTPServer((host, port), DashboardHandler)
    print(f"Dashboard: http://{host}:{port}/dashboard.html")
    print(f"Overview 저장 파일: {os.path.join(ROOT, STATE_FILE)}")
    print(f"EMP 저장 파일: {os.path.join(ROOT, EMP_STATE_FILE)}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n종료")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
