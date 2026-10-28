import argparse
import hmac
import json
import os
import secrets
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from .common import atomic_json, canonical, inside, now
from .store import Store
from .folders import FolderWorker
from .excel import ExcelListener

DEFAULTS = dict(roots=[], scan_seconds=5, settle_seconds=2, snapshot_limit_mb=512,
                comparison_cells=200000, excel_cache_cells=200000, excel_event_cells=10000)
PACKAGE_ROOT = Path(__file__).resolve().parent.parent

def default_data():
    return Path(os.environ.get('LOCALAPPDATA', str(Path.home()/'.local/share'))) / 'FileChangeLog'

def validate_config(data, output):
    config = dict(DEFAULTS)
    if not isinstance(data, dict) or not isinstance(data.get('roots', []), list):
        raise ValueError('Provide a list of folder paths.')
    roots = []
    for raw in data.get('roots', []):
        if not isinstance(raw, str) or not raw.strip():
            raise ValueError('Folder paths must be nonempty text.')
        if not os.path.isabs(raw):
            raise ValueError('Use an absolute folder path, such as C:\\Royalty or \\\\server\\share.')
        path = canonical(raw)
        if inside(path, output) or inside(output, path):
            raise ValueError('A monitored folder cannot contain recorder history or be inside it.')
        if any(inside(path, other) or inside(other, path) for other in roots):
            raise ValueError('Remove duplicate or nested monitored folders; subfolders are already included.')
        roots.append(path)
    if len(roots) > 20:
        raise ValueError('Pilot supports up to 20 folder roots.')
    config['roots'] = roots
    for key, low, high in [('scan_seconds', 2, 3600), ('settle_seconds', 1, 300),
                           ('snapshot_limit_mb', 1, 16384), ('comparison_cells', 100, 1000000),
                           ('excel_cache_cells', 100, 1000000), ('excel_event_cells', 1, 50000)]:
        value = int(data.get(key, DEFAULTS[key]))
        if not low <= value <= high:
            raise ValueError(f'{key} must be between {low} and {high}.')
        config[key] = value
    return config

class Recorder:
    def __init__(self, directory):
        self.store = Store(directory)
        self.config_path = self.store.directory/'settings.json'
        self.config = validate_config(json.loads(self.config_path.read_text(encoding='utf-8')) if self.config_path.exists() else {}, self.store.directory)
        self.paused = False
        self.workers = []
        self.excel = None
        self.started = now()
        self.lock = threading.RLock()

    def start(self):
        self.store.emit('recorder_started', version='0.1.0', platform=sys.platform,
                        recorded_account=os.environ.get('USERNAME', os.environ.get('USER', 'unknown')),
                        capture_scope='selected_folders_only', missing_intermediate_history_not_recoverable=True)
        self.reconfigure_workers()
        self.excel = ExcelListener(self.store, self.config, lambda: self.config['roots'], lambda: not self.paused)
        self.excel.thread.start()

    def reconfigure_workers(self):
        for w in self.workers:
            w.stop.set()
        self.workers = []
        for root in self.config['roots']:
            w = FolderWorker(root, self.store, self.config, enabled=lambda r=root: not self.paused and r in self.config['roots'])
            self.workers.append(w)
            w.begin()

    def configure(self, value):
        new = validate_config(value, self.store.directory)
        with self.lock:
            self.config = new
            atomic_json(self.config_path, new)
            self.store.emit('configuration_changed', roots=new['roots'], settings=new)
            if self.excel:
                self.excel.settings = new
                self.excel.reset.set()
            self.reconfigure_workers()

    def status(self):
        return dict(started=self.started, paused=self.paused, settings=self.config, history=str(self.store.directory),
                    storage_error=self.store.error, folders=[dict(w.health) for w in self.workers],
                    excel=dict(self.excel.health) if self.excel else {}, startup=startup_enabled())

    def close(self):
        self.paused = True
        for w in self.workers:
            w.stop.set()
        if self.excel:
            self.excel.stop.set()
        self.store.emit('recorder_stopping', remaining_unsaved_edits_not_implied_committed=True)
        # Blocked network operations can outlive shutdown. They are daemon threads;
        # the process exit ends them, without making false deletion claims.

def _scheduled_task():
    import win32com.client
    service = win32com.client.Dispatch('Schedule.Service')
    service.Connect()
    return service.GetFolder('\\\\'), service


def startup_enabled():
    if os.name != 'nt':
        return False
    try:
        folder, _ = _scheduled_task()
        folder.GetTask('FileChangeLogRecorder')
        return True
    except Exception:
        # Legacy packages registered startup in the current-user Run key.
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\\Microsoft\\Windows\\CurrentVersion\\Run') as k:
                return bool(winreg.QueryValueEx(k, 'FileChangeLog')[0])
        except OSError:
            return False


def set_startup(enabled, directory):
    if os.name != 'nt':
        raise ValueError('Sign-in startup is available on Windows only.')
    import winreg
    folder, service = _scheduled_task()
    if enabled:
        task = service.NewTask(0)
        task.RegistrationInfo.Description = 'File Change Log: restart after unexpected exit in this interactive user session'
        task.Principal.LogonType = 3  # TASK_LOGON_INTERACTIVE_TOKEN, no stored password.
        trigger = task.Triggers.Create(9)  # TASK_TRIGGER_LOGON.
        trigger.Enabled = True
        settings = task.Settings
        settings.Enabled = True
        settings.StartWhenAvailable = True
        settings.ExecutionTimeLimit = 'PT0S'
        settings.RestartInterval = 'PT1M'
        settings.RestartCount = 999
        executable = Path(sys.executable).with_name('pythonw.exe')
        if not executable.exists():
            executable = Path(sys.executable)
        action = task.Actions.Create(0)
        action.Path = str(executable)
        action.Arguments = subprocess.list2cmdline([str(PACKAGE_ROOT/'run.py'), '--background', '--data-dir', str(directory)])
        action.WorkingDirectory = str(PACKAGE_ROOT)
        # Create/update the current user's task; do not elevate or store credentials.
        folder.RegisterTaskDefinition('FileChangeLogRecorder', task, 6, '', '', 3)
    else:
        try:
            folder.DeleteTask('FileChangeLogRecorder', 0)
        except Exception as exc:
            if 'cannot find' not in str(exc).lower() and 'not found' not in str(exc).lower():
                raise
    # Only remove the legacy registration after the scheduled task was set.
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r'Software\\Microsoft\\Windows\\CurrentVersion\\Run') as key:
        try:
            winreg.DeleteValue(key, 'FileChangeLog')
        except FileNotFoundError:
            pass

def make_handler(recorder, token, shutdown):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass  # Never expose session tokens in HTTP access logs.

        def authorized(self):
            host = self.headers.get('Host', '').split(':')[0]
            if host not in {'127.0.0.1', 'localhost'}:
                return False
            supplied = self.headers.get('X-Recorder-Token', '') or parse_qs(urlparse(self.path).query).get('token', [''])[0]
            return hmac.compare_digest(supplied, token)

        def respond(self, status, value, content_type='application/json; charset=utf-8', download=None):
            body = json.dumps(value, ensure_ascii=False).encode() if not isinstance(value, bytes) else value
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Referrer-Policy', 'no-referrer')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; frame-ancestors 'none'; form-action 'none'")
            if download:
                self.send_header('Content-Disposition', f'attachment; filename="{download}"')
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            if not self.authorized():
                self.respond(403, {'error': 'Open the recorder using Open Recorder.bat.'})
                return
            url = urlparse(self.path)
            try:
                if url.path == '/':
                    self.respond(200, (PACKAGE_ROOT/'web/index.html').read_bytes(), 'text/html; charset=utf-8')
                elif url.path == '/api/status':
                    self.respond(200, recorder.status())
                elif url.path == '/api/events':
                    self.respond(200, recorder.store.recent(search=parse_qs(url.query).get('q', [''])[0]))
                elif url.path == '/api/export':
                    q = parse_qs(url.query).get('q', [''])[0]
                    # Bounded UI export; complete JSONL history remains in logs/.
                    events = list(reversed(recorder.store.recent(limit=1000, search=q)))
                    body = ''.join(json.dumps(e, ensure_ascii=False)+'\n' for e in events).encode()
                    self.respond(200, body, 'application/x-ndjson', 'recorder-events.jsonl')
                else:
                    self.respond(404, {'error': 'Not found'})
            except Exception as e:
                self.respond(500, {'error': str(e)})

        def do_POST(self):
            if not self.authorized():
                self.respond(403, {'error': 'Invalid session token'})
                return
            try:
                length = int(self.headers.get('Content-Length', 0))
                if not 0 < length <= 200000:
                    raise ValueError('Invalid request size')
                data = json.loads(self.rfile.read(length))
                route = urlparse(self.path).path
                if route == '/api/settings':
                    recorder.configure(data)
                elif route == '/api/note':
                    text = data.get('text', '')
                    if not isinstance(text, str) or not text.strip() or len(text) > 20000:
                        raise ValueError('Enter an explanation of at most 20,000 characters.')
                    scope = data.get('scope', 'case_note')
                    if scope not in {'case_note', 'proposed_rule'}:
                        raise ValueError('Unknown note scope')
                    recorder.store.emit('user_explanation', text=text.strip(), linked_event_id=str(data.get('event_id', '')),
                                        scope=scope, reusable_rule_not_approved=True)
                elif route == '/api/pause':
                    paused = bool(data.get('paused'))
                    recorder.paused = paused
                    if recorder.excel:
                        recorder.excel.reset.set()
                    recorder.store.emit('monitoring_paused' if paused else 'monitoring_resumed', capture_gap=True)
                elif route == '/api/startup':
                    set_startup(bool(data.get('enabled')), recorder.store.directory)
                    recorder.store.emit('signin_startup_changed', enabled=bool(data.get('enabled')))
                elif route == '/api/stop':
                    self.respond(200, {'ok': True})
                    threading.Thread(target=shutdown, daemon=True).start()
                    return
                else:
                    self.respond(404, {'error': 'Not found'})
                    return
                self.respond(200, {'ok': True})
            except (ValueError, TypeError, KeyError) as e:
                self.respond(400, {'error': str(e)})
            except Exception as e:
                self.respond(500, {'error': str(e)})
    return Handler

def main():
    parser = argparse.ArgumentParser(description='Local folder and desktop Excel event recorder')
    parser.add_argument('--data-dir', type=Path, default=default_data())
    parser.add_argument('--background', action='store_true')
    parser.add_argument('--install-startup', action='store_true')
    parser.add_argument('--remove-startup', action='store_true')
    args = parser.parse_args()
    args.data_dir.mkdir(parents=True, exist_ok=True)
    if args.install_startup or args.remove_startup:
        set_startup(args.install_startup, args.data_dir)
        return
    # Bind a single-instance port before any capture threads start. Use a local
    # socket, not a stale PID file, to determine whether this recorder is alive.
    gate = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        gate.bind(('127.0.0.1', 47653))
        gate.listen(1)
    except OSError:
        gate.close()
        session_file = args.data_dir/'session.json'
        if not args.background and session_file.exists():
            webbrowser.open(json.loads(session_file.read_text())['url'])
        return
    recorder = Recorder(args.data_dir)
    server = None
    token = secrets.token_urlsafe(32)
    try:
        server = ThreadingHTTPServer(('127.0.0.1', 0), make_handler(recorder, token, lambda: server.shutdown()))
        server.daemon_threads = True
        url = f'http://127.0.0.1:{server.server_port}/?token={token}'
        atomic_json(args.data_dir/'session.json', dict(url=url, pid=os.getpid(), started=now()))
        recorder.start()
        if not args.background:
            webbrowser.open(url)
        server.serve_forever(poll_interval=.3)
    except KeyboardInterrupt:
        pass
    finally:
        recorder.close()
        if server:
            server.server_close()
        gate.close()
        (args.data_dir/'session.json').unlink(missing_ok=True)

if __name__ == '__main__':
    main()
