"""Content-addressed local build/test receipts; delivery gates are never cached."""
from pathlib import Path
from contextlib import contextmanager
import ctypes
import hashlib
import json
import os
import platform
import re
import subprocess
import sys
import time
import uuid
from document_assets import referenced_paths, strip_delivery_audit
from configurators import CONFIGURE

sys.dont_write_bytecode = True
SITE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SITE.parent / "C-00-Shared"))
from process_support import run_process
_ENVIRONMENT = None


def digest(path):
    value = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            value.update(block)
    return value.hexdigest()


def key(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()


def environment():
    global _ENVIRONMENT
    if _ENVIRONMENT is None:
        import winreg
        fonts = []
        for hive, label in ((winreg.HKEY_LOCAL_MACHINE, 'machine'), (winreg.HKEY_CURRENT_USER, 'user')):
            try:
                registry = winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
            except FileNotFoundError:
                continue
            with registry:
                for i in range(winreg.QueryInfoKey(registry)[1]):
                    name, filename, _ = winreg.EnumValue(registry, i)
                    if not isinstance(filename, str):
                        fonts.append((label, name, filename, None))
                        continue
                    file = Path(filename)
                    if not file.is_absolute():
                        font_root = Path(os.environ['SystemRoot']) / 'Fonts' if label == 'machine' else Path(os.environ['LOCALAPPDATA']) / 'Microsoft/Windows/Fonts'
                        file = font_root / file
                    info = file.stat() if file.is_file() else None
                    fonts.append((label, name, filename, (info.st_size, info.st_mtime_ns) if info else None))
        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
        except AttributeError:
            dpi = 96
        framework = Path(os.environ['SystemRoot']) / 'Microsoft.NET/Framework64/v4.0.30319'
        if not framework.is_dir():
            framework = framework.parent.parent / 'Framework/v4.0.30319'
        _ENVIRONMENT = key({'python': sys.version, 'executable': sys.executable,
                            'windows': platform.platform(), 'dpi': dpi, 'fonts': sorted(fonts),
                            'language': ctypes.windll.kernel32.GetUserDefaultUILanguage(),
                            'framework': [(p.name, digest(p)) for p in
                                          (framework / n for n in ('csc.exe', 'mscorlib.dll', 'System.Windows.Forms.dll', 'System.Drawing.dll')) if p.is_file()]})
    return _ENVIRONMENT


def inputs(base, snapshot):
    base = Path(base).resolve()
    paths = set(base.glob('*.py')) | {base / 'project.json'}
    for directory in ('src', 'AC-Original Code', 'lib', 'AB-Graphic'):
        paths.update(p for p in (base / directory).rglob('*') if p.is_file() and '__pycache__' not in p.parts)
    paths.update(p for p in (base.parent / 'C-04-Devlog').rglob('*.md') if p.is_file())
    paths.add(base / 'AA-Program' / CONFIGURE[base.name])
    paths.update(referenced_paths(base))
    paths.add(base.parent / 'verify_shared.py')
    paths.update(p for p in (base.parent / 'C-00-Shared').rglob('*') if p.is_file() and '__pycache__' not in p.parts)
    paths.update(p for p in (SITE / 'src').rglob('*') if p.is_file() and '__pycache__' not in p.parts)
    paths.add(SITE / 'AA-Program/C3-XycabinSite_Configure.py')
    values = []
    for path in sorted(paths):
        if path.suffix.lower() == '.md':
            text = path.read_text(encoding='utf-8-sig')
            text = strip_delivery_audit(text)
            fingerprint = hashlib.sha256(text.encode('utf-8')).hexdigest()
        else:
            fingerprint = digest(path)
        values.append((str(path.relative_to(base.parent)), fingerprint))
    from site_core import embedded_snapshot
    page = embedded_snapshot(Path(snapshot).read_text(encoding='utf-8'))
    return key({'schema': 1, 'inputs': values, 'site': hashlib.sha256(page.encode('utf-8')).hexdigest(), 'environment': environment()})


def input_fingerprint(root, addin_roots):
    """Raw-byte fingerprint of the embedded document surface; detects mid-build edits."""
    root = Path(root).resolve()
    paths = set((root.parent / 'C-04-Devlog').glob('*.md'))
    paths.add(root.parent / 'verify_shared.py')
    paths.update(p for p in (root.parent / 'C-00-Shared').rglob('*') if p.is_file() and '__pycache__' not in p.parts)
    for base in addin_roots:
        base = Path(base).resolve()
        for item in referenced_paths(base):
            path = Path(item)
            paths.update([path if path.is_absolute() else base / path])
    values = [(str(path.relative_to(root.parent)), digest(path)) for path in sorted(paths)]
    return key({'schema': 1, 'fingerprint': values})


def atomic_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def receipt(path, expected, outputs=()):
    try:
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        if data['schema'] != 1 or data['key'] != expected:
            return None
        if any(not Path(p).is_file() or digest(p) != value for p, value in data['outputs'].items()):
            return None
        if set(map(str, outputs)) != set(data['outputs']):
            return None
        if hashlib.sha256(data.get('output', '').encode('utf-8')).hexdigest() != data['evidence_sha256']:
            return None
        return data
    except (OSError, ValueError, KeyError, TypeError):
        return None


def save(path, expected, outputs=(), output=''):
    data = {'schema': 1, 'key': expected, 'time': time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()),
            'outputs': {str(p): digest(p) for p in outputs}, 'output': output,
            'evidence_sha256': hashlib.sha256(output.encode('utf-8')).hexdigest()}
    pending = os.environ.get('XY_CACHE_PENDING')
    if pending:
        atomic_json(Path(pending) / (key(str(Path(path).resolve())) + '.json'), {'target': str(Path(path).resolve()), 'receipt': data})
    else:
        atomic_json(path, data)


def commit(pending):
    allowed = [SITE / 'build/incremental-cache'] + [SITE.parent / name / 'build/cache' for name in ('C-01-ExcelIterate', 'C-02-WordUpdater')]
    for file in Path(pending).glob('*.json'):
        item = json.loads(file.read_text(encoding='utf-8'))
        target = Path(item['target']).resolve()
        if not any(target.is_relative_to(base.resolve()) for base in allowed):
            raise ValueError('缓存目标越界：' + str(target))
        atomic_json(target, item['receipt'])


def compile_cached(command, cwd, mode='full'):
    """Use every actual source/reference/resource byte, not a modification-time guess."""
    command = list(map(str, command))
    cwd = Path(cwd)
    target = Path(next(a[5:] for a in command if a.startswith('/out:')))
    normalized = []
    framework = Path(command[0]).parent
    for argument in command:
        if argument.startswith('/out:'):
            normalized.append('/out:' + target.name)
            continue
        prefix, value = ('', argument)
        suffix = ''
        if argument.startswith('/r:'):
            prefix, value = '/r:', argument[3:]
        elif argument.startswith('/resource:'):
            prefix, value = '/resource:', argument[10:]
            value, separator, resource = value.partition(',')
            suffix = separator + resource
        path = Path(value)
        if not path.is_absolute():
            path = cwd / path
        if prefix == '/r:' and not path.is_file():
            path = framework / value
        normalized.append((prefix, digest(path), suffix) if path.is_file() else argument)
    expected = key({'command': normalized, 'environment': environment()})
    cache = cwd / 'build/cache/compiler' / (target.name + '.json')
    if mode == 'incremental' and receipt(cache, expected, [target]):
        print('CACHE: 编译 ' + target.name, flush=True)
        return
    cache.unlink(missing_ok=True)
    # The feature directory layout can exceed Windows CreateProcess's command
    # length. Keep the full input list in the cache key; only transport it via rsp.
    invocation = command
    if len(subprocess.list2cmdline(command)) > 24000:
        response = cache.with_suffix('.rsp')
        response.parent.mkdir(parents=True, exist_ok=True)
        response.write_text('\n'.join(subprocess.list2cmdline([arg]) for arg in command[1:]) + '\n',
                            encoding='utf-8-sig')
        invocation = [command[0], '@' + str(response)]
    run_process(invocation, cwd=cwd, check=True)
    save(cache, expected, [target])


@contextmanager
def build_lock():
    import msvcrt
    path = SITE / 'build/unified-build.lock'
    path.parent.mkdir(exist_ok=True)
    stream = path.open('a+b')
    if not path.stat().st_size:
        stream.write(b'0')
        stream.flush()
    stream.seek(0)
    try:
        try:
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise ValueError('已有统一构建正在运行，请等待其完成。') from exc
        try:
            yield
        finally:
            stream.seek(0)
            msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
    finally:
        stream.close()
