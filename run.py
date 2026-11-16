from pathlib import Path
import sys

# Embedded Python uses an explicit path list; add this package deliberately.
sys.path.insert(0, str(Path(__file__).resolve().parent))
if __name__ == '__main__':
    try:
        from recorder.app import main
        main()
    except Exception:
        import os, traceback
        log = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'FileChangeLog' / 'startup-error.txt'
        log.parent.mkdir(parents=True, exist_ok=True)
        log.write_text(traceback.format_exc(), encoding='utf-8')
        if sys.stderr:
            print(traceback.format_exc(), file=sys.stderr)
        raise
