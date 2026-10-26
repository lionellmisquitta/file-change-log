import hashlib
import os
import shutil
import threading
import time
import uuid
from pathlib import Path
from .common import SNAPSHOT_EXTENSIONS, canonical, ignored, now
from .compare import compare_files

def signature(path):
    s = os.stat(path)
    return [s.st_size, s.st_mtime_ns, s.st_ino, s.st_dev]

def walk_files(root, data_dir):
    """Incomplete enumeration is an outage, never evidence of mass deletion."""
    if not os.path.isdir(root):
        raise OSError('Folder unavailable or not a directory')
    errors = []
    found = {}
    def error(e):
        errors.append(e)
    for folder, dirs, files in os.walk(root, followlinks=False, onerror=error):
        dirs[:] = [d for d in dirs if not ignored(Path(folder)/d, data_dir)
                   and not os.path.islink(Path(folder)/d)
                   and not getattr(os.path, 'isjunction', lambda p: False)(Path(folder)/d)]
        for name in files:
            path = str(Path(folder)/name)
            if ignored(path, data_dir) or os.path.islink(path):
                continue
            try:
                found[path] = signature(path)
            except OSError as e:
                errors.append(e)
    if errors:
        raise OSError(f'Incomplete scan: {errors[0]}')
    return found

class FolderWorker:
    def __init__(self, root, store, settings, enabled=lambda: True):
        self.root = canonical(root)
        self.store = store
        self.settings = settings
        self.enabled = enabled
        self.stop = threading.Event()
        self.pending = {}
        self.health = dict(path=self.root, state='starting', last_scan=None, error=None)
        self.offline = None
        self.thread = threading.Thread(target=self.run, daemon=True, name='Folder scan')
        self.observer = None

    def begin(self):
        self.thread.start()

    def native_listener(self):
        try:
            from watchdog.events import FileSystemEventHandler
            from watchdog.observers import Observer
            worker = self
            class Handler(FileSystemEventHandler):
                def on_any_event(self, event):
                    if event.event_type not in {'created', 'modified', 'moved', 'deleted'}:
                        return
                    if worker.stop.is_set() or not worker.enabled() or ignored(event.src_path, worker.store.directory):
                        return
                    try:
                        worker.store.emit('folder_notification', event.src_path, action=event.event_type,
                                          destination=getattr(event, 'dest_path', None), is_directory=event.is_directory,
                                          source='filesystem_notification', contents_not_implied=True)
                    except Exception:
                        pass  # Store exposes its write failure prominently in status.
            observer = Observer()
            observer.schedule(Handler(), self.root, recursive=True)
            observer.start()
            self.observer = observer
        except Exception as e:
            self.store.emit('folder_notification_unavailable', self.root, reason=str(e), fallback='periodic_scan')

    def capture(self, path, expected):
        ext = Path(path).suffix.lower().lstrip('.')
        if ext not in SNAPSHOT_EXTENSIONS:
            self.store.emit('snapshot_not_selected', path, reason='File type outside selected document/data snapshot scope')
            return None
        maximum = int(self.settings.get('snapshot_limit_mb', 512)) * 1024 * 1024
        if expected[0] > maximum:
            self.store.emit('snapshot_skipped', path, reason=f'File exceeds configured {maximum//(1024*1024)} MiB per-file limit', size=expected[0])
            return None
        if shutil.disk_usage(self.store.directory).free < max(256*1024*1024, expected[0]*2):
            raise OSError('Insufficient free space for snapshot; original left untouched')
        parent = self.store.snapshots / hashlib.sha256(canonical(path).encode()).hexdigest()[:24]
        parent.mkdir(exist_ok=True)
        temporary = parent / (uuid.uuid4().hex + '.partial')
        digest = hashlib.sha256()
        try:
            with open(path, 'rb') as src, temporary.open('xb') as dst:
                for block in iter(lambda: src.read(1024*1024), b''):
                    if self.stop.is_set():
                        raise OSError('Capture cancelled')
                    digest.update(block)
                    dst.write(block)
                dst.flush()
                os.fsync(dst.fileno())
            # Never label a changing file as a stable capture.
            if signature(path) != expected:
                raise OSError('File changed while copying; will retry')
            destination = parent / (digest.hexdigest() + Path(path).suffix.lower())
            if destination.exists():
                temporary.unlink()
            else:
                os.replace(temporary, destination)
            return str(destination.relative_to(self.store.directory))
        finally:
            temporary.unlink(missing_ok=True)

    def scan_once(self):
        if not self.enabled():
            return
        try:
            current = walk_files(self.root, self.store.directory)
        except OSError as e:
            if self.offline is not True:
                self.store.emit('folder_unavailable', self.root, reason=str(e), deletion_inference_disabled=True)
            self.offline = True
            self.health.update(state='unavailable', error=str(e))
            return
        if self.offline:
            self.store.emit('folder_reconnected', self.root, intermediate_activity_unknown=True)
        self.offline = False
        old = self.store.inventory(self.root)
        for path in set(old) - set(current):
            if not self.enabled():
                return
            self.store.emit('file_absent_on_scan', path, root=self.root, last_snapshot=old[path][1],
                            exact_deletion_time_unknown=True)
            self.store.drop_file(self.root, path)
            self.pending.pop(path, None)
        for path, sig in current.items():
            if self.stop.is_set() or not self.enabled():
                return
            previous = old.get(path)
            if previous and sig == previous[0]:
                self.pending.pop(path, None)
                continue
            pending = self.pending.get(path)
            if pending is None or pending[0] != sig:
                self.pending[path] = (sig, time.monotonic())
                continue
            if time.monotonic() - pending[1] < float(self.settings.get('settle_seconds', 2)):
                continue
            try:
                snapshot = self.capture(path, sig)
                event = self.store.emit('file_version_captured', path, root=self.root,
                    capture='saved_file_bytes', classification='changed' if previous else 'first_observed',
                    size=sig[0], modified_time_ns=sig[1], snapshot=snapshot,
                    previous_snapshot=previous[1] if previous else None)
                self.store.put_file(self.root, path, sig, snapshot)
                self.pending.pop(path, None)
                if snapshot and previous and previous[1] and snapshot != previous[1]:
                    try:
                        compare_files(self.store.directory/previous[1], self.store.directory/snapshot,
                                      self.store.emit, path, max_cells=int(self.settings.get('comparison_cells', 200000)))
                    except Exception as e:
                        self.store.emit('comparison_failed', path, reason=str(e), snapshots_retained=True)
            except Exception as e:
                # Back off transient locks without losing the old baseline.
                self.pending[path] = (sig, time.monotonic())
                self.store.emit('snapshot_retry', path, reason=str(e))
        self.health.update(state='watching', last_scan=now(), error=None, files=len(current), pending=len(self.pending))

    def run(self):
        self.native_listener()
        while not self.stop.is_set():
            try:
                self.scan_once()
            except Exception as e:
                self.health.update(state='error', error=str(e))
                try:
                    self.store.emit('folder_scan_error', self.root, reason=str(e))
                except Exception:
                    pass
            if self.observer and not self.observer.is_alive():
                self.health.update(error='Native notifications stopped; periodic scans continue.')
            self.stop.wait(float(self.settings.get('scan_seconds', 5)))
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)
