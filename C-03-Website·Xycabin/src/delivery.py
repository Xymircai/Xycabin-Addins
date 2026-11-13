"""One website source, staged installers, and rollback on delivery failure."""
from pathlib import Path
import hashlib
import importlib.util
import json
import os
import sys
import uuid
from configurators import CONFIGURE
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "C-00-Shared"))

RUN_STAGES = []

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]
PROJECTS = (('C-01-ExcelIterate', 'Excel_Iterate', 'Excel'),
            ('C-02-WordUpdater', 'WordUpdater', 'Word'))


def projects(root=ROOT):
    result = []
    for folder, assembly, host in PROJECTS:
        base = Path(root).parent / folder
        for item in ('project.json', 'rebuild.py', 'validate.py', 'AA-Program/' + CONFIGURE[folder]):
            if not (base / item).is_file():
                raise FileNotFoundError('统一重建需要两个加载项项目，缺少：' + str(base / item))
        documented = base.parent / 'C-04-Devlog/B1-Project-Files.md'
        if not documented.is_file():
            raise FileNotFoundError('统一重建需要仓库根文档，缺少：' + str(documented))
        identity = json.loads((base / 'project.json').read_text(encoding='utf-8-sig'))
        if identity['assembly'] != assembly or identity['host'] != host:
            raise ValueError('加载项身份不匹配：' + str(base))
        result.append((base, assembly))
    return result


def version(base):
    spec = importlib.util.spec_from_file_location('_xy_program', base / 'AA-Program' / CONFIGURE[base.name])
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module.get_version()


def run(command, cwd, log):
    from build_report import run_command
    result = run_command(command, cwd, log)
    RUN_STAGES.append(result)


def payload_probe(root=ROOT):
    root = Path(root)
    source = root / 'src/tools/PayloadProbe.cs'
    target = root / 'build/PayloadProbe.exe'
    compiler = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if not compiler.is_file():
        compiler = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
    if not target.is_file() or source.stat().st_mtime_ns > target.stat().st_mtime_ns:
        target.parent.mkdir(parents=True, exist_ok=True)
        run([compiler, '/nologo', '/target:exe', '/codepage:65001', '/out:' + str(target), source],
            root, root / 'build/payload-probe-build.log')
    return target


def verify_payload(base, assembly, installer, site, root=ROOT):
    from document_assets import collect, manifest_bytes
    expected = Path(root) / 'build' / (assembly + '-assets-' + uuid.uuid4().hex[:12] + '.xml')
    expected.parent.mkdir(parents=True, exist_ok=True)
    try:
        expected.write_bytes(manifest_bytes(collect(base)))
        run([payload_probe(root), base, assembly, installer, site, version(base), expected], root,
            Path(root) / ('build/' + assembly + '-payload.log'))
    finally:
        expected.unlink(missing_ok=True)


def publish(updates, replace=os.replace):
    """Prepare sibling files first; roll back every committed target on failure.

    Backups are retained if rollback itself fails, with their paths in the error.
    All mutations use literal Path objects and remain beside the supplied targets.
    """
    prepared, committed = [], []
    token = uuid.uuid4().hex[:12]
    try:
        targets = [Path(target).resolve() for source, target in updates]
        if len(set(targets)) != len(targets):
            raise ValueError('重复的交付目标')
        for (source, target), absolute in zip(updates, targets):
            absolute.parent.mkdir(parents=True, exist_ok=True)
            temporary = absolute.with_name('.xycabin-' + token + '-' + absolute.name + '.new')
            backup = absolute.with_name('.xycabin-' + token + '-' + absolute.name + '.old')
            prepared.append((absolute, temporary, backup))
            with temporary.open('wb') as output:
                output.write(Path(source).read_bytes())
                output.flush()
                os.fsync(output.fileno())
        for target, temporary, backup in prepared:
            existed = target.exists()
            if existed:
                replace(target, backup)
            committed.append((target, backup, existed))
            replace(temporary, target)
    except Exception as failure:
        recovery_errors = []
        for target, backup, existed in reversed(committed):
            try:
                if existed:
                    os.replace(backup, target)
                elif target.exists():
                    target.unlink()
            except OSError as error:
                recovery_errors.append(str(backup) + ': ' + str(error))
        if recovery_errors:
            raise RuntimeError('交付失败且回滚未完成；保留恢复文件：' + '; '.join(recovery_errors)) from failure
        raise
    else:
        for target, temporary, backup in prepared:
            if backup.exists():
                try:
                    backup.unlink()
                except OSError:
                    print('交付已完成，旧备份暂未清理：' + str(backup), flush=True)
    finally:
        for target, temporary, backup in prepared:
            if temporary.exists():
                try:
                    temporary.unlink()
                except OSError:
                    print('暂存文件暂未清理：' + str(temporary), flush=True)


def fingerprints(paths):
    return {str(p): {'bytes': p.stat().st_size, 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()}
            for p in paths}
