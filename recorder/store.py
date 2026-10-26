import json
import os
import sqlite3
import threading
import uuid
from pathlib import Path
from .common import now, clean

class Store:
    """SQLite is a rebuildable index; JSONL files are the durable event stream."""
    def __init__(self, directory):
        self.directory = Path(directory).resolve()
        self.logs = self.directory / 'logs'
        self.snapshots = self.directory / 'snapshots'
        for path in (self.logs, self.snapshots):
            path.mkdir(parents=True, exist_ok=True)
        self.lock = threading.RLock()
        self.error = None
        self.db = sqlite3.connect(self.directory / 'index.sqlite', check_same_thread=False)
        self.db.execute('PRAGMA journal_mode=WAL')
        self.db.execute('CREATE TABLE IF NOT EXISTS events (seq INTEGER PRIMARY KEY AUTOINCREMENT, id TEXT UNIQUE, kind TEXT, path TEXT, body TEXT)')
        self.db.execute('CREATE TABLE IF NOT EXISTS inventory (root TEXT, path TEXT, signature TEXT, snapshot TEXT, PRIMARY KEY(root,path))')
        self.db.commit()
        # Recover entries appended just before an interrupted SQLite insert.
        for log in sorted(self.logs.glob('events-*.jsonl')):
            with log.open(encoding='utf-8') as f:
                for line in f:
                    try:
                        event = json.loads(line)
                        self._index(event)
                    except (ValueError, KeyError):
                        self.error = f'Malformed/incomplete line found in {log.name}; original retained.'
        self.db.commit()

    def _index(self, event):
        self.db.execute('INSERT OR IGNORE INTO events(id,kind,path,body) VALUES(?,?,?,?)',
                        (event['id'], event['event'], event.get('path', ''), json.dumps(event, ensure_ascii=False)))

    def emit(self, kind, path='', **fields):
        event = clean(dict(id=uuid.uuid4().hex, timestamp_utc=now(), event=kind, path=str(path), **fields))
        with self.lock:
            try:
                line = json.dumps(event, ensure_ascii=False, allow_nan=False)
                with (self.logs / f'events-{event["timestamp_utc"][:10]}.jsonl').open('a', encoding='utf-8') as f:
                    f.write(line + '\n')
                    f.flush()
                    os.fsync(f.fileno())
                self._index(event)
                self.db.commit()
                return event
            except Exception as e:
                self.error = f'Unable to write event history: {type(e).__name__}: {e}'
                raise

    def recent(self, limit=200, search='', after=0):
        with self.lock:
            # LIKE wildcards are deliberately treated as ordinary search characters.
            term = search.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_')
            rows = self.db.execute("SELECT seq,body FROM events WHERE seq>? AND body LIKE ? ESCAPE '\\' ORDER BY seq DESC LIMIT ?",
                                   (after, '%' + term + '%', min(limit, 1000))).fetchall()
            return [dict(json.loads(body), sequence=seq) for seq, body in rows]

    def inventory(self, root):
        with self.lock:
            return {p: (json.loads(sig), snap) for p, sig, snap in self.db.execute('SELECT path,signature,snapshot FROM inventory WHERE root=?', (root,))}

    def put_file(self, root, path, signature, snapshot=None):
        with self.lock:
            self.db.execute('INSERT OR REPLACE INTO inventory VALUES(?,?,?,?)', (root, path, json.dumps(signature), snapshot))
            self.db.commit()

    def drop_file(self, root, path):
        with self.lock:
            self.db.execute('DELETE FROM inventory WHERE root=? AND path=?', (root, path))
            self.db.commit()

    def close(self):
        with self.lock:
            self.db.close()
