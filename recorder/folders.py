import hashlib
import os
import shutil
import threading
import time
import uuid
from pathlib import Path
from .common import SNAPSHOT_EXTENSIONS, canonical, ignored, inside as inside_path, now
from .compare import compare_files, workbook_cells

def signature(path):
    s = os.stat(path)
    return [s.st_size, s.st_mtime_ns, s.st_ino, s.st_dev]

class ScanInventory(dict):
    def __init__(self, values, unreadable=()):
        super().__init__(values)
        self.unreadable = tuple(unreadable)

def walk_files(root, data_dir):
    """Incomplete enumeration is an outage, never evidence of mass deletion."""
    if not os.path.isdir(root):
        raise OSError('Folder unavailable or not a directory')
    unreadable = []
    found = {}
    def error(e):
        unreadable.append(getattr(e, 'filename', None) or root)
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
            except FileNotFoundError:
                continue  # Files can vanish between directory walk and stat.
            except (PermissionError, OSError):
                unreadable.append(path)
    return ScanInventory(found, unreadable)

class FolderWorker:
    def __init__(self, root, store, settings, enabled=lambda: True):
        self.root = canonical(root)
        self.store = store
        self.settings = settings
        self.enabled = enabled
        self.stop = threading.Event()
        self.pending = {}
        self.failures = {}
        self.baseline_paths = None
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

    def capture(self, path, expected, initial=False):
        ext = Path(path).suffix.lower().lstrip('.')
        if ext not in SNAPSHOT_EXTENSIONS:
            self.store.emit('snapshot_not_selected', path, reason='Unsupported document/data file type')
            return None
        size = expected[0]
        maximum = int(self.settings.get('snapshot_limit_mb', 50)) * 1024 * 1024
        budget = int(self.settings.get('snapshot_budget_mb', 500)) * 1024 * 1024
        baseline_budget = int(self.settings.get('baseline_budget_mb', 250)) * 1024 * 1024
        if size > maximum:
            self.store.emit('snapshot_skipped_size', path, size=size, limit_bytes=maximum)
            return None
        with self.store.snapshot_lock:
            if self.store.snapshot_bytes + size > budget or (initial and self.store.baseline_bytes + size > baseline_budget):
                if not self.store.snapshot_suspended_issued:
                    self.store.snapshot_suspended_issued = True
                    self.store.emit('snapshot_budget_reached', path,
                                    used_bytes=self.store.snapshot_bytes,
                                    total_budget_bytes=budget,
                                    no_automatic_deletion=True)
                self.store.emit('snapshot_skipped_budget', path, size=size, used_bytes=self.store.snapshot_bytes,
                                total_budget_bytes=budget, baseline=initial)
                return None
            if shutil.disk_usage(self.store.directory).free < size + 1024*1024*1024:
                self.store.emit('snapshot_skipped_disk_space', path, size=size, minimum_free_bytes=1024*1024*1024)
                return None
            temporary = self.store.snapshots / (uuid.uuid4().hex + '.partial')
            digest = hashlib.sha256()
            try:
                # Windows: avoid blocking Excel's atomic replace/rename during Save.
                if os.name == 'nt':
                    import win32con
                    import win32file
                    handle = win32file.CreateFile(
                        path, win32con.GENERIC_READ,
                        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                        None, win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None)
                    import msvcrt
                    fd = msvcrt.open_osfhandle(int(handle.Detach()), os.O_RDONLY)
                    src = os.fdopen(fd, 'rb')
                else:
                    src = open(path, 'rb')
                with src, temporary.open('xb') as dst:
                    for block in iter(lambda: src.read(1024*1024), b''):
                        if self.stop.is_set():
                            raise OSError('Capture cancelled')
                        digest.update(block)
                        dst.write(block)
                    dst.flush()
                    os.fsync(dst.fileno())
                if signature(path) != expected:
                    raise OSError('File changed while copying; will retry')
                destination = self.store.snapshots / (digest.hexdigest() + Path(path).suffix.lower())
                if destination.exists():
                    temporary.unlink()
                else:
                    os.replace(temporary, destination)
                    self.store.snapshot_bytes += size
                    if self.store.snapshot_bytes >= budget * .7 and not self.store.snapshot_warning_issued:
                        self.store.snapshot_warning_issued = True
                        self.store.emit('storage_warning', used_bytes=self.store.snapshot_bytes,
                                        total_budget_bytes=budget, threshold_percent=70)
                    if initial:
                        self.store.baseline_bytes += size
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
        if self.baseline_paths is None:
            self.baseline_paths = set(current) if not old else set()
        blocked = getattr(current, 'unreadable', ())
        if blocked:
            self.health['unreadable_paths'] = len(blocked)
        else:
            self.health.pop('unreadable_paths', None)
        # A same-volume rename has a stable device/inode pair. Preserve the
        # old snapshot without fabricating content changes or a deletion.
        missing = {path: old[path] for path in set(old) - set(current)}
        arrived = {path: sig for path, sig in current.items() if path not in old}
        moved_from = {}
        used_old = set()
        for destination, sig in arrived.items():
            if not sig[2] or not sig[3]:
                continue
            matches = [source for source, prior in missing.items()
                       if source not in used_old and prior[0][2:] == sig[2:] ]
            if len(matches) == 1:
                source = matches[0]
                moved_from[destination] = source
                used_old.add(source)
        for path in set(old) - set(current):
            if any(inside_path(path, subtree) for subtree in blocked):
                continue  # An unreadable subtree is not evidence of deletion.
            if not self.enabled():
                return
            if path in used_old:
                destination = next(dest for dest, source in moved_from.items() if source == path)
                self.store.emit('file_moved', path, destination=destination, root=self.root,
                                continuity='device_and_inode_match')
            else:
                self.store.emit('file_absent_on_scan', path, root=self.root, last_snapshot=old[path][1],
                                exact_deletion_time_unknown=True)
            self.store.drop_file(self.root, path)
            self.pending.pop(path, None)
        # Capture recently modified working files first when baseline allowance is finite.
        for path, sig in sorted(current.items(), key=lambda item: item[1][1], reverse=True):
            if self.stop.is_set() or not self.enabled():
                return
            previous = old.get(path)
            if path in moved_from:
                carried = old[moved_from[path]]
                if sig == carried[0]:
                    self.store.put_file(self.root, path, sig, carried[1])
                    self.pending.pop(path, None)
                    continue
                previous = carried
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
                snapshot = self.capture(path, sig, initial=(previous is None and path in self.baseline_paths))
                event = self.store.emit('file_version_captured', path, root=self.root,
                    capture='saved_file_bytes', classification='moved_and_changed' if path in moved_from else ('changed' if previous else ('baseline' if path in self.baseline_paths else 'arrived')),
                    size=sig[0], modified_time_ns=sig[1], snapshot=snapshot,
                    previous_snapshot=previous[1] if previous else None)
                self.store.put_file(self.root, path, sig, snapshot or (previous[1] if previous else None))
                self.pending.pop(path, None)
                if path in self.failures:
                    self.store.emit('snapshot_recovered', path, attempts=self.failures[path]['attempts'])
                    del self.failures[path]
                if snapshot and not previous and Path(path).suffix.lower() in {'.xlsx', '.xlsm'}:
                    # A first snapshot is an observation, not proof of when each
                    # cell was entered. Show populated cells without inventing edits.
                    try:
                        baseline = workbook_cells(self.store.directory/snapshot,
                                                  int(self.settings.get('comparison_cells', 200000)))
                        count = sum(len(cells) for cells in baseline.values())
                        maximum = int(self.settings.get('comparison_details', 5000))
                        details = [
                            dict(sheet=sheet, address=address, value=value)
                            for sheet, cells in baseline.items()
                            for address, value in cells.items()
                        ][:maximum]
                        self.store.emit('snapshot_initial_cells', path, cells=details,
                                        total_populated_cells=count, details_truncated=count > maximum,
                                        observation='first_saved_version_not_proof_of_when_cells_were_entered',
                                        snapshot=snapshot)
                    except Exception as e:
                        self.store.emit('snapshot_initial_cells_unavailable', path,
                                        reason=str(e), snapshot_retained=True)
                if snapshot and previous and previous[1] and snapshot != previous[1]:
                    try:
                        compare_files(self.store.directory/previous[1], self.store.directory/snapshot,
                                      self.store.emit, path, max_cells=int(self.settings.get('comparison_cells', 200000)))
                    except Exception as e:
                        self.store.emit('comparison_failed', path, reason=str(e), snapshots_retained=True)
            except Exception as e:
                # Back off transient locks without losing the old baseline.
                self.pending[path] = (sig, time.monotonic())
                failure = self.failures.get(path)
                current_time = time.monotonic()
                if failure is None:
                    failure = {'first': current_time, 'last_notice': 0, 'attempts': 0}
                    self.failures[path] = failure
                failure['attempts'] += 1
                if failure['attempts'] == 1 or current_time - failure['last_notice'] > 300:
                    self.store.emit('snapshot_delayed' if current_time - failure['first'] > 120 else 'snapshot_retry',
                                    path, reason=str(e), attempts=failure['attempts'])
                    failure['last_notice'] = current_time

        self.health.update(state='watching', last_scan=now(), error=None, files=len(current), pending=len(self.pending), delayed=len(self.failures))

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
                self.health['notifications'] = 'stopped; periodic scans remain active'
            self.stop.wait(float(self.settings.get('scan_seconds', 5)))
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)
