import datetime as dt
import json
import math
import os
from pathlib import Path

SNAPSHOT_EXTENSIONS = set('xlsx xlsm xlsb xls csv tsv ods doc docx docm odt pdf rtf txt md json jsonl xml yaml yml sql log ini cfg html htm eml msg ppt pptx pptm odp'.split())

def now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec='milliseconds')

def clean(value):
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else str(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(k): clean(v) for k, v in value.items()}
    if isinstance(value, (tuple, list)):
        return [clean(v) for v in value]
    return str(value)

def canonical(path):
    # Resolve Windows 8.3 aliases and reparse-path spelling before containment
    # checks; the history directory is resolved by pathlib as well.
    path = os.path.realpath(os.path.abspath(os.path.expanduser(os.fspath(path))))
    if os.name == 'nt':
        try:
            import win32wnet
            path = win32wnet.WNetGetUniversalName(path)
        except Exception:
            pass
    return os.path.normcase(os.path.normpath(path))

def inside(path, root):
    try:
        return os.path.commonpath([canonical(path), canonical(root)]) == canonical(root)
    except (ValueError, OSError):
        return False

def atomic_json(path, value):
    path = Path(path)
    temporary = path.with_suffix(path.suffix + '.tmp')
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temporary, path)

def ignored(path, data_dir):
    p = Path(path)
    return inside(p, data_dir) or p.name.startswith('~$') or p.suffix.lower() in {'.tmp', '.part', '.crdownload'}

def cell_address(row, col):
    letters = ''
    while col:
        col, r = divmod(col - 1, 26)
        letters = chr(65 + r) + letters
    return f'{letters}{row}'
