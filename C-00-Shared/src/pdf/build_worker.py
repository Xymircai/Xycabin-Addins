"""Build an isolated, offline PDF worker; restore only hash-pinned NuGet inputs."""
from pathlib import Path
import hashlib
import io
import json
import os
from process_support import run_process
import urllib.request
import zipfile

ROOT = Path(PROJECT_ROOT).resolve()


def build():
    out = ROOT / 'build/pdf-worker'
    out.mkdir(parents=True, exist_ok=True)
    lock = json.loads((source_path('src/pdf/packages.lock.json')).read_text(encoding='utf-8'))
    notices = {}
    for item in lock:
        name, version = item['id'].lower(), item['version']
        cached = ROOT / 'build/pdf-packages' / (name + '.' + version + '.nupkg')
        cached.parent.mkdir(parents=True, exist_ok=True)
        if not cached.exists():
            url = 'https://api.nuget.org/v3-flatcontainer/{0}/{1}/{0}.{1}.nupkg'.format(name, version)
            data = urllib.request.urlopen(url, timeout=60).read()
        else:
            data = cached.read_bytes()
        if hashlib.sha256(data).hexdigest() != item['sha256']:
            raise ValueError('PDF dependency checksum mismatch: ' + name)
        cached.write_bytes(data)
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            for member in item['files']:
                (out / Path(member).name).write_bytes(package.read(member))
            for member in package.namelist():
                if Path(member).name.upper() in ('LICENSE.TXT', 'THIRD-PARTY-NOTICES.TXT'):
                    notices[item['id'] + '-' + Path(member).name] = package.read(member)
    csc = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if not csc.exists():
        csc = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
    run_process([str(csc), '/nologo', '/target:exe', '/platform:anycpu', '/codepage:65001',
                    '/out:' + str(out / 'PdfWorker.exe'), '/r:System.Drawing.dll',
                    '/r:System.Xml.dll', '/r:Microsoft.CSharp.dll',
                    '/r:' + str(out / 'PdfSharp-gdi.dll'),
                    str(source_path('src/pdf/PdfWorker.cs'))], check=True)
    (out / 'PdfWorker.exe.config').write_bytes((source_path('src/pdf/PdfWorker.exe.config')).read_bytes())
    notice = source_path('src/pdf/THIRD-PARTY-NOTICES.txt')
    bundle = ROOT / 'build/PdfWorker.zip'
    with zipfile.ZipFile(bundle, 'w', zipfile.ZIP_DEFLATED) as archive:
        # Select from the lock, never include stale DLLs from a previous build.
        names = {Path(member).name for item in lock for member in item['files']}
        names.update(('PdfWorker.exe', 'PdfWorker.exe.config'))
        for name in sorted(names):
            archive.write(out / name, name)
        for name, data in sorted(notices.items()):
            archive.writestr(name, data)
        archive.write(notice, notice.name)
    return bundle


if __name__ == '__main__':
    print(build())
