"""Verify published artifacts; developer mechanism regressions are an explicit profile."""
from pathlib import Path
import argparse
import json
import sys
sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'src'))
from delivery import projects, run, verify_payload, version, fingerprints, RUN_STAGES
from site_core import verify_installers
from build_cache import atomic_json
from build_report import BuildReport, DEVELOPER_GATES


def verify_delivery():
    bases = projects(ROOT)
    site = ROOT / 'AA-Program/XyCabin.html'
    record_path = ROOT / 'build/current-delivery.json'
    record = json.loads(record_path.read_text(encoding='utf-8'))
    targets = [base / 'AA-Program' / (assembly + '_Setup.exe') for base, assembly in bases] + [site]
    actual = fingerprints(targets)
    for target, details in actual.items():
        expected = record.get(target)
        if expected is None or any(expected.get(key) != details[key] for key in ('bytes', 'sha256')):
            raise ValueError('正式产物与交付记录不一致：' + target)
    packages = {}
    for base, assembly in bases:
        installer = base / 'AA-Program' / (assembly + '_Setup.exe')
        expected_version = version(base)
        if record[str(installer.resolve())].get('version') != expected_version:
            raise ValueError('交付记录版本不一致：' + assembly)
        packages['word' if assembly == 'WordUpdater' else 'excel'] = (installer, expected_version)
        # Verify real CLR identities, pinned publisher and both manifest signatures directly.
        run([sys.executable, base / 'verify_package.py'], base, ROOT / ('build/' + assembly + '-signatures.log'))
        verify_payload(base, assembly, installer, site, ROOT)
    verify_installers(site.read_text(encoding='utf-8'), packages)
    print('PASS: 正式产物摘要、固定发布者、清单双签名、五载荷、内嵌资源及网页两包字节一致')
    return {'artifacts': actual, 'record': str(record_path)}


def developer_checks(report):
    with report.gate('G-03'):
        run([sys.executable, ROOT / 'src/tools/workflow_probe.py'], ROOT, ROOT / 'build/workflow-validation.log')
    with report.gate('G-05'):
        run([sys.executable, ROOT / 'src/tools/cleanup_probe.py'], ROOT, ROOT / 'build/cleanup-validation.log')
    with report.gate('G-06'):
        run([sys.executable, ROOT / 'src/tools/cache_probe.py'], ROOT, ROOT / 'build/cache-validation.log')
    with report.gate('G-07'):
        run([sys.executable, ROOT / 'src/tools/cleanup_review_probe.py'], ROOT, ROOT / 'build/cleanup-review-validation.log')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    scope = parser.add_mutually_exclusive_group()
    scope.add_argument('--site-only', action='store_true', help='仅网页检查，摘要为部分验证')
    scope.add_argument('--developer-checks', action='store_true', help='单独运行 G-03/05/06/07 机制回归')
    parser.add_argument('--skip-ui', action='store_true', help='保留兼容；前台交互始终单列验收')
    parser.add_argument('--background', action='store_true', help='不主动激活窗口，子进程无控制台，焦点只读记录')
    parser.add_argument('--verbose', action='store_true')
    args = parser.parse_args()
    developer = args.developer_checks
    report = BuildReport(ROOT.parent, ROOT / 'build', 'development' if developer else 'closing',
                         DEVELOPER_GATES if developer else ['G-23'], args.mode, args.verbose, args.background)
    parent_run_id = None
    if not developer:
        try:
            previous = json.loads((ROOT / 'build/last-run-summary.json').read_text(encoding='utf-8'))
            if previous.get('status') == 'success':
                parent_run_id = previous.get('run_id')
        except (OSError, ValueError):
            pass
    status, evidence = 'failed', {}
    try:
        if developer:
            developer_checks(report)
        else:
            with report.gate('G-23') as gate:
                if args.site_only:
                    run([sys.executable, ROOT / 'src/tools/site_probe.py', '--html', ROOT / 'AA-Program/XyCabin.html'],
                        ROOT, ROOT / 'build/site-validation.log')
                    gate['status'] = 'partial'
                else:
                    evidence = verify_delivery()
        status = 'partial' if args.site_only else 'success'
    finally:
        target = 'development-run-summary.json' if developer else 'validate-run-summary.json'
        data = report.summary(status, stages=RUN_STAGES, parent_run_id=parent_run_id,
                              classification_gate=dict(rule='R-30', cached=False, status='not_run',
                                                       reason='本入口仅核正式产物或开发机制；源码分类在重建与产品验证执行'),
                              **evidence)
        atomic_json(ROOT / 'build' / target, data)
    if developer:
        print('开发自检完成；日志：' + str(report.directory))
    else:
        report.manual_items()


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, RuntimeError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
