#!/usr/bin/env python3
"""
market_db_dashboard 정적 파일 + Yahoo Finance 차트 JSON 프록시.

Yahoo 웹 차트(https://finance.yahoo.com/chart/...)는 X-Frame-Options 때문에
다른 사이트 iframe에 넣을 수 없습니다. 대신 서버에서
query1.finance.yahoo.com/v8/finance/chart API를 호출해 JSON을 받아
dashboard.html의 Plotly 캔들차트가 /api/yahoo-chart 로 읽도록 합니다.

실행 (기본 포트 8765):
  cd C:\\Users\\infomax\\Documents\\market_db_dashboard
  python serve_market_dashboard.py

브라우저: http://127.0.0.1:8765/dashboard.html#overview
"""
from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("PORT", "8765"))

_YAHOO_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


class DashboardHandler(SimpleHTTPRequestHandler):
    """GET /api/yahoo-chart?symbol=CL=F&interval=1d&range=1y → Yahoo v8 JSON"""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault("directory", ROOT)
        super().__init__(*args, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/yahoo-chart":
            self._yahoo_chart_proxy(parsed.query)
            return
        return super().do_GET()

    def _yahoo_chart_proxy(self, query: str) -> None:
        try:
            qs = urllib.parse.parse_qs(query, keep_blank_values=True)
            symbol = (qs.get("symbol", ["CL=F"])[0] or "CL=F").strip()
            interval = (qs.get("interval", ["1d"])[0] or "1d").strip()
            range_ = (qs.get("range", ["1y"])[0] or "1y").strip()
            if not symbol:
                symbol = "CL=F"
            sym_enc = urllib.parse.quote(symbol, safe="")
            upstream = (
                "https://query1.finance.yahoo.com/v8/finance/chart/"
                + sym_enc
                + "?interval="
                + urllib.parse.quote(interval)
                + "&range="
                + urllib.parse.quote(range_)
            )
            req = urllib.request.Request(upstream, headers={"User-Agent": _YAHOO_UA})
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = resp.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)
        except urllib.error.HTTPError as e:
            err = json.dumps({"chart": {"error": {"description": f"HTTP {e.code}"}}}).encode("utf-8")
            self.send_response(502)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(err)
        except Exception as e:
            err = json.dumps({"chart": {"error": {"description": str(e)}}}).encode("utf-8")
            self.send_response(502)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(err)

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("%s - %s\n" % (self.log_date_time_string(), fmt % args))


def main() -> None:
    os.chdir(ROOT)
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), DashboardHandler)
    print("Serving:", ROOT)
    print("Open:    http://127.0.0.1:%s/dashboard.html" % PORT)
    print("Overview Yahoo proxy: GET /api/yahoo-chart?symbol=CL=F&interval=1d&range=1y")
    print("Stop: Ctrl+C")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
