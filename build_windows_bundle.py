"""Build a self-contained Windows x64 ZIP using official CPython and PyPI wheels.

Run on any OS: python build_windows_bundle.py
Downloads are pinned and checked against provider SHA256 metadata. No pip or
admin installation is required on the target Windows machine.
"""
import hashlib
import io
import json
import shutil
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PYTHON_VERSION = '3.13.17'
PACKAGES = {'pywin32': '312', 'watchdog': '6.0.0', 'openpyxl': '3.1.5', 'et-xmlfile': '2.0.0', 'defusedxml': '0.7.1'}

def fetch(url):
    with urllib.request.urlopen(url, timeout=90) as r:
        data = r.read()
        if r.headers.get('Content-Encoding') == 'gzip':
            import gzip
            data = gzip.decompress(data)
        return data

def unpack(data, destination):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for member in z.infolist():
            target = (destination / member.filename).resolve()
            if not target.is_relative_to(destination.resolve()):
                raise ValueError('Unsafe archive member')
        z.extractall(destination)

def build():
    out = ROOT/'dist'
    out.mkdir(exist_ok=True)
    package = out/'FileChangeLog-Windows-x64'
    if package.exists():
        shutil.rmtree(package)
    package.mkdir()
    for name in ['recorder', 'web', 'docs', 'tests']:
        shutil.copytree(ROOT/name, package/name, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    for name in ['run.py', 'README.md', 'START-HERE.txt', 'Start Recorder.bat', 'Open Recorder.bat',
                 'Start Quietly.vbs', 'requirements.txt', 'build_windows_bundle.py']:
        shutil.copy2(ROOT/name, package/name)
    runtime = package/'runtime'
    runtime.mkdir()
    url = f'https://www.python.org/ftp/python/{PYTHON_VERSION}/python-{PYTHON_VERSION}-embed-amd64.zip'
    data = fetch(url)
    # SHA256 is extracted from the official release page rather than guessed.
    release = fetch('https://www.python.org/downloads/release/python-'+PYTHON_VERSION.replace('.', '')+'/').decode()
    from html.parser import HTMLParser
    class ReleaseParser(HTMLParser):
        def __init__(self):
            super().__init__(); self.row=''; self.inrow=False; self.rows=[]
        def handle_starttag(self, tag, attrs):
            if tag=='tr': self.inrow=True; self.row=''
            if self.inrow:
                self.row+=' '.join(v or '' for _,v in attrs)+' '
        def handle_endtag(self, tag):
            if tag=='tr' and self.inrow: self.rows.append(self.row); self.inrow=False
        def handle_data(self, value):
            if self.inrow: self.row+=value+' '
    import re
    parser=ReleaseParser();parser.feed(release)
    row=next((r for r in parser.rows if f'python-{PYTHON_VERSION}-embed-amd64.zip' in r), '')
    candidates=re.findall(r'\b[0-9a-f]{64}\b', row)
    digest=hashlib.sha256(data).hexdigest()
    if not candidates or digest not in candidates:
        raise RuntimeError('Official CPython SHA256 could not be verified; refusing to bundle runtime.')
    unpack(data, runtime)
    manifest=[dict(name='CPython', version=PYTHON_VERSION, source=url, sha256=digest)]
    site=runtime/'Lib/site-packages'
    site.mkdir(parents=True)
    for name, version in PACKAGES.items():
        meta=json.loads(fetch(f'https://pypi.org/pypi/{name}/{version}/json'))
        candidates=[f for f in meta['urls'] if f['filename'].endswith('.whl') and
                    (f['filename'].endswith('py3-none-any.whl') or
                     ('cp313-cp313-win_amd64.whl' in f['filename']) or
                     (name=='watchdog' and 'win_amd64.whl' in f['filename']))]
        if not candidates:
            raise RuntimeError(f'No pinned Windows/Python 3.13 wheel for {name}')
        item=candidates[0]; wheel=fetch(item['url'])
        if hashlib.sha256(wheel).hexdigest()!=item['digests']['sha256']:
            raise RuntimeError('Dependency checksum mismatch')
        unpack(wheel, site)
        manifest.append(dict(name=name, version=version, source=item['url'], sha256=item['digests']['sha256']))
        print(f'Bundled {name} {version}', flush=True)
    (runtime/'python313._pth').write_text('python313.zip\n.\nLib\\site-packages\n..\nimport site\n', encoding='utf-8')
    (package/'DEPENDENCIES.json').write_text(json.dumps(manifest, indent=2), encoding='utf-8')
    archive=out/'FileChangeLog-Windows-x64.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for file in sorted(package.rglob('*')):
            if file.is_file(): z.write(file, str(Path(package.name)/file.relative_to(package)))
    print(f'Created {archive} ({archive.stat().st_size:,} bytes)', flush=True)
    return archive

if __name__=='__main__':
    build()
