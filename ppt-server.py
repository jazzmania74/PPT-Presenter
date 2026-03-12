#!/usr/bin/env python3
"""
PPT Presenter Server
노트북에서 실행하면 스마트폰으로 PowerPoint를 원격 제어할 수 있습니다.
사용법: python3 ppt-server.py
"""
import http.server
import json
import os
import subprocess
import socket
import glob
import sys
import platform

PORT = 8080
SLIDE_DIR = '/tmp/ppt-presenter-slides'


class PPTController:
    """Controls Microsoft PowerPoint via AppleScript (macOS)"""

    def run_applescript(self, script):
        try:
            result = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True, text=True, timeout=10
            )
            return result.stdout.strip(), result.returncode
        except Exception as e:
            return str(e), 1

    def is_powerpoint_running(self):
        script = ('tell application "System Events" to return '
                  '(name of processes) contains "Microsoft PowerPoint"')
        out, _ = self.run_applescript(script)
        return out == 'true'

    def get_status(self):
        try:
            script = '''
tell application "Microsoft PowerPoint"
    try
        set ssView to slide show view of slide show window 1
        set currentSlide to slide number of ssView
        set totalSlides to count of slides of active presentation
        set presName to name of active presentation
        return "" & currentSlide & "|" & totalSlides & "|" & presName
    on error
        try
            set totalSlides to count of slides of active presentation
            set presName to name of active presentation
            return "0|" & totalSlides & "|" & presName
        on error
            return "0|0|"
        end try
    end try
end tell'''
            out, rc = self.run_applescript(script)
            if rc != 0 or not out:
                return {'presenting': False, 'currentSlide': 0,
                        'totalSlides': 0, 'filename': ''}
            parts = out.split('|', 2)
            current = int(parts[0]) if parts[0] else 0
            return {
                'presenting': current > 0,
                'currentSlide': current,
                'totalSlides': int(parts[1]) if len(parts) > 1 else 0,
                'filename': parts[2] if len(parts) > 2 else ''
            }
        except Exception:
            return {'presenting': False, 'currentSlide': 0,
                    'totalSlides': 0, 'filename': ''}

    def start_slideshow(self):
        script = '''
tell application "Microsoft PowerPoint"
    activate
    set thePresentation to active presentation
    run slide show slide show settings of thePresentation
end tell'''
        return self.run_applescript(script)

    def start_from_current(self):
        script = '''
tell application "Microsoft PowerPoint"
    activate
    set thePresentation to active presentation
    set slideNum to slide index of slide of view of active window
    set starting slide of slide show settings of thePresentation to slideNum
    run slide show slide show settings of thePresentation
end tell'''
        return self.run_applescript(script)

    def end_slideshow(self):
        script = '''
tell application "Microsoft PowerPoint"
    try
        exit slide show slide show view of slide show window 1
    end try
end tell'''
        return self.run_applescript(script)

    def next_slide(self):
        script = '''
tell application "Microsoft PowerPoint"
    try
        go to next slide slide show view of slide show window 1
    end try
end tell'''
        return self.run_applescript(script)

    def prev_slide(self):
        script = '''
tell application "Microsoft PowerPoint"
    try
        go to previous slide slide show view of slide show window 1
    end try
end tell'''
        return self.run_applescript(script)

    def goto_slide(self, n):
        script = f'''
tell application "Microsoft PowerPoint"
    try
        go to slide slide show view of slide show window 1 number {int(n)}
    end try
end tell'''
        return self.run_applescript(script)

    def black_screen(self):
        script = '''
tell application "Microsoft PowerPoint"
    try
        set blackScreen to black screen of slide show view of slide show window 1
        if blackScreen then
            set black screen of slide show view of slide show window 1 to false
        else
            set black screen of slide show view of slide show window 1 to true
        end if
    end try
end tell'''
        return self.run_applescript(script)

    def export_slides(self):
        """Export all slides as PNG images"""
        os.makedirs(SLIDE_DIR, exist_ok=True)
        # Clear existing
        for f in glob.glob(os.path.join(SLIDE_DIR, '**/*'), recursive=True):
            if os.path.isfile(f):
                os.remove(f)

        script = f'''
tell application "Microsoft PowerPoint"
    set thePresentation to active presentation
    set presName to name of thePresentation
    save thePresentation in "{SLIDE_DIR}/slides" as save as PNG
    return presName
end tell'''
        out, rc = self.run_applescript(script)
        if rc != 0:
            return [], out

        # PowerPoint creates a subfolder with the presentation name
        files = sorted(glob.glob(os.path.join(SLIDE_DIR, '**/*.png'),
                                 recursive=True))
        if not files:
            files = sorted(glob.glob(os.path.join(SLIDE_DIR, '**/*.PNG'),
                                     recursive=True))
        return files, ''


class RequestHandler(http.server.BaseHTTPRequestHandler):
    ppt = PPTController()
    slide_files = []

    def do_GET(self):
        path = self.path.split('?')[0]

        if path == '/' or path == '/index.html':
            self.serve_file('index.html', 'text/html')
        elif path == '/api/status':
            status = self.ppt.get_status()
            status['slideCount'] = len(RequestHandler.slide_files)
            self.send_json(status)
        elif path == '/api/slides':
            urls = [f'/api/slide-image/{i}' for i in
                    range(len(RequestHandler.slide_files))]
            self.send_json({'slides': urls, 'count': len(urls)})
        elif path.startswith('/api/slide-image/'):
            try:
                idx = int(path.split('/')[-1])
                self.serve_slide_image(idx)
            except (ValueError, IndexError):
                self.send_error(404)
        elif path == '/api/check':
            running = self.ppt.is_powerpoint_running()
            self.send_json({'running': running, 'platform': platform.system()})
        elif path == '/api/ip':
            self.send_json({'ip': get_local_ip(), 'port': PORT})
        else:
            self.send_error(404)

    def do_POST(self):
        path = self.path.split('?')[0]

        if path == '/api/slideshow/start':
            self.ppt.start_slideshow()
            self.send_json({'ok': True})
        elif path == '/api/slideshow/start-current':
            self.ppt.start_from_current()
            self.send_json({'ok': True})
        elif path == '/api/slideshow/end':
            self.ppt.end_slideshow()
            self.send_json({'ok': True})
        elif path == '/api/slide/next':
            self.ppt.next_slide()
            self.send_json({'ok': True})
        elif path == '/api/slide/prev':
            self.ppt.prev_slide()
            self.send_json({'ok': True})
        elif path == '/api/slide/goto':
            body = self.read_body()
            n = body.get('slide', 1)
            self.ppt.goto_slide(n)
            self.send_json({'ok': True})
        elif path == '/api/slide/black':
            self.ppt.black_screen()
            self.send_json({'ok': True})
        elif path == '/api/export':
            files, err = self.ppt.export_slides()
            RequestHandler.slide_files = files
            self.send_json({'ok': True, 'count': len(files),
                           'error': err if err else None})
        else:
            self.send_error(404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def read_body(self):
        length = int(self.headers.get('Content-Length', 0))
        if length:
            return json.loads(self.rfile.read(length))
        return {}

    def serve_slide_image(self, idx):
        files = RequestHandler.slide_files
        if 0 <= idx < len(files) and os.path.exists(files[idx]):
            self.send_response(200)
            self.send_header('Content-Type', 'image/png')
            self.send_header('Cache-Control', 'public, max-age=3600')
            self.end_headers()
            with open(files[idx], 'rb') as f:
                self.wfile.write(f.read())
        else:
            self.send_error(404)

    def serve_file(self, filename, content_type):
        filepath = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), filename)
        if os.path.exists(filepath):
            self.send_response(200)
            self.send_header('Content-Type', content_type + '; charset=utf-8')
            self.end_headers()
            with open(filepath, 'rb') as f:
                self.wfile.write(f.read())
        else:
            self.send_error(404)

    def send_json(self, data):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode())

    def log_message(self, format, *args):
        pass


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


if __name__ == '__main__':
    ip = get_local_ip()
    print()
    print('  ╔══════════════════════════════════════╗')
    print('  ║       📽️  PPT Presenter Server       ║')
    print('  ╠══════════════════════════════════════╣')
    print(f'  ║  Local:   http://localhost:{PORT}      ║')
    print(f'  ║  Network: http://{ip}:{PORT}  ║')
    print('  ╠══════════════════════════════════════╣')
    print('  ║  스마트폰에서 Network URL로 접속하세요  ║')
    print('  ║  Ctrl+C로 서버 종료                   ║')
    print('  ╚══════════════════════════════════════╝')
    print()

    server = http.server.HTTPServer(('0.0.0.0', PORT), RequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n  서버를 종료합니다...')
        server.shutdown()
