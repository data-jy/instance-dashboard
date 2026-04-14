#!/usr/bin/env python3
import http.server, os, json, urllib.request, urllib.error, gzip, io

PORT = 5002
DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = '/home/rocky/license-data.json'
SOFTWARE_EOL_FILE = '/home/rocky/license-dashboard/software_eol.json'
os.chdir(DIR)

# index.html 캐시
_html_cache = None
_html_cache_mtime = 0

def get_html():
    global _html_cache, _html_cache_mtime
    path = os.path.join(DIR, 'index.html')
    mtime = os.path.getmtime(path)
    if _html_cache is None or mtime != _html_cache_mtime:
        with open(path, 'rb') as f:
            raw = f.read()
        _html_cache = gzip.compress(raw, compresslevel=6)
        _html_cache_mtime = mtime
    return _html_cache

class Handler(http.server.BaseHTTPRequestHandler):

    def do_GET(self):
        # favicon - 빠르게 204 반환
        if self.path == '/favicon.ico':
            self.send_response(204)
            self.end_headers()
            return

        # 라이센스 데이터 조회
        if self.path == '/api/sync-licenses-get':
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._resp(200, data)
            except FileNotFoundError:
                self._resp(200, {'licenses': [], 'settings': {}})
            except Exception as e:
                self._resp(500, {'error': str(e)})

        elif self.path == '/api/software-eol':
            try:
                with open(SOFTWARE_EOL_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._resp(200, data)
            except FileNotFoundError:
                self._resp(200, {'software': []})
            except Exception as e:
                self._resp(500, {'error': str(e)})

        elif self.path in ('/', '/index.html'):
            # gzip 압축된 HTML 서빙
            try:
                body = get_html()
                accept = self.headers.get('Accept-Encoding', '')
                if 'gzip' in accept:
                    self.send_response(200)
                    self.send_header('Content-Type', 'text/html; charset=utf-8')
                    self.send_header('Content-Encoding', 'gzip')
                    self.send_header('Content-Length', len(body))
                    self.send_header('Cache-Control', 'no-cache')
                    self.end_headers()
                    self.wfile.write(body)
                else:
                    # gzip 미지원 브라우저
                    with open(os.path.join(DIR, 'index.html'), 'rb') as f:
                        raw = f.read()
                    self.send_response(200)
                    self.send_header('Content-Type', 'text/html; charset=utf-8')
                    self.send_header('Content-Length', len(raw))
                    self.send_header('Cache-Control', 'no-cache')
                    self.end_headers()
                    self.wfile.write(raw)
            except Exception as e:
                self._resp(500, {'error': str(e)})

        else:
            # 기타 정적 파일
            try:
                fpath = os.path.join(DIR, self.path.lstrip('/'))
                if not os.path.isfile(fpath):
                    self.send_response(404)
                    self.end_headers()
                    return
                with open(fpath, 'rb') as f:
                    raw = f.read()
                self.send_response(200)
                self.send_header('Content-Length', len(raw))
                self.end_headers()
                self.wfile.write(raw)
            except Exception:
                self.send_response(404)
                self.end_headers()

    def do_POST(self):
        if self.path == '/webhook-proxy':
            try:
                length  = int(self.headers.get('Content-Length', 0))
                body    = self.rfile.read(length)
                data    = json.loads(body.decode('utf-8'))
                url     = data.get('url', '')
                payload = data.get('payload', {})

                if not url.startswith('https://nhnent.dooray.com/'):
                    self._resp(403, {'ok': False, 'error': 'URL not allowed'})
                    return

                req = urllib.request.Request(
                    url,
                    data=json.dumps(payload).encode('utf-8'),
                    headers={'Content-Type': 'application/json'},
                    method='POST'
                )
                with urllib.request.urlopen(req, timeout=10) as r:
                    self._resp(200, {'ok': True, 'status': r.status})
            except urllib.error.HTTPError as e:
                self._resp(200, {'ok': False, 'error': 'HTTP {}'.format(e.code)})
            except Exception as e:
                self._resp(200, {'ok': False, 'error': str(e)})

        elif self.path == '/api/sync-licenses':
            try:
                length  = int(self.headers.get('Content-Length', 0))
                body    = self.rfile.read(length)
                data    = json.loads(body.decode('utf-8'))
                with open(DATA_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self._resp(200, {'ok': True, 'message': 'Data synced'})
            except Exception as e:
                self._resp(500, {'ok': False, 'error': str(e)})

        elif self.path == '/api/save-software':
            try:
                length  = int(self.headers.get('Content-Length', 0))
                body    = self.rfile.read(length)
                data    = json.loads(body.decode('utf-8'))
                with open(SOFTWARE_EOL_FILE, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                # HTML 캐시 무효화
                global _html_cache
                _html_cache = None
                self._resp(200, {'ok': True, 'message': 'Software data saved'})
            except Exception as e:
                self._resp(500, {'ok': False, 'error': str(e)})

        else:
            self.send_response(404)
            self.end_headers()

    def _resp(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', len(body))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, fmt, *args):
        pass

print('License Dashboard running on port {}'.format(PORT))
httpd = http.server.HTTPServer(('0.0.0.0', PORT), Handler)
httpd.serve_forever()
