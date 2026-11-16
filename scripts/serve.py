"""開発用の静的サーバー（キャッシュさせない）。python3 scripts/serve.py [port]"""
import functools
import http.server
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class NoCache(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


port = int(sys.argv[1]) if len(sys.argv) > 1 else 8769
handler = functools.partial(NoCache, directory=ROOT)
print(f"http://127.0.0.1:{port}/web/")
http.server.ThreadingHTTPServer(("127.0.0.1", port), handler).serve_forever()
