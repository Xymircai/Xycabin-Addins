"""Generate one immutable page, build both add-ins, then publish together."""
from pathlib import Path
import argparse
import json
import sys
import uuid
import shutil
import os
import time
sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'src'))
from site_core import build, attach_installers, embedded_snapshot, verify_installers
from delivery import projects, run, verify_payload, publish, fingerprints, RUN_STAGES
from build_cache import inputs, input_fingerprint, receipt, save, commit, build_lock, atomic_json
from build_report import BuildReport, BUILD_GATES


def write_summary(report, mode, status, started, reused, skipped):
    source_stage = next((stage for stage in RUN_STAGES if stage['stage'] == 'shared-source'), None)
    classification = {'rule': 'R-30', 'cached': False, 'status':
                      'not_run' if source_stage is None else 'fail' if source_stage['returncode'] else 'pass'}
    inprocess = [dict(gate=item['gate'], status=item['status']) for item in report.gates
                 if item['gate'] in ('G-04', 'G-08', 'G-09', 'G-17', 'G-19', 'G-20', 'G-21', 'G-22')]
    atomic_json(ROOT / 'build/last-run-summary.json', report.summary(
        status, reused_builds=reused, stages=RUN_STAGES, inprocess_gates=inprocess,
        classification_gate=classification, skipped=skipped))


def rebuild():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--skip-ui', action='store_true', help='只跳过窗口布局预览；业务、签名和载荷门禁仍执行')
    parser.add_argument('--background', action='store_true', help='后台验证模式，不运行会显示或激活窗口的探针；前台项目单独列为待验收')
    parser.add_argument('--verbose', action='store_true', help='在控制台显示完整构建日志')
    args = parser.parse_args()
    report = BuildReport(ROOT.parent, ROOT / 'build', 'rebuild',
                         BUILD_GATES, args.mode, args.verbose, args.background)
    started = time.perf_counter()
    reused = []
    cache_updates = []
    skipped = {'G-12': '--skip-ui', 'G-15': '--skip-ui'} if args.skip_ui else {}
    try:
        # 文档与共享输入先核对；机制自检由 --developer-checks 单独运行。
        with report.gate('G-01'):
            run([sys.executable, ROOT / 'src/tools/devlog_probe.py'], ROOT, ROOT / 'build/devlog-validation.log')
        with report.gate('G-02'):
            checker = ROOT.parent / 'verify_shared.py'
            if not checker.is_file():
                raise FileNotFoundError('缺少共享检查入口：' + str(checker))
            run([sys.executable, checker], ROOT.parent, ROOT / 'build/shared-source.log')
        with report.gate('G-04'):
            bases = projects(ROOT)  # Fail before generating or changing any output.
        with report.gate('G-08'):
            before = input_fingerprint(ROOT, [base for base, _ in bases])
        stage = ROOT / ('build/stage-' + uuid.uuid4().hex[:10])
        stage.mkdir(parents=True)
        pending = stage / 'cache-pending'
        previous_pending = os.environ.get('XY_CACHE_PENDING')
        os.environ['XY_CACHE_PENDING'] = str(pending)
        try:
            with report.gate('G-09'):
                snapshot = build(ROOT, target=stage / 'embedded.snapshot.html')
            with report.gate('G-10'):
                run([sys.executable, ROOT / 'src/tools/site_probe.py', '--html', snapshot], ROOT,
                    ROOT / 'build/site-validation.log')
            updates, delivered, packages = [], [], {}
            for index, (base, assembly) in enumerate(bases, 1):
                with report.gate('G-11' if assembly == 'Excel_Iterate' else 'G-14'):
                    setup = stage / (assembly + '_Setup.exe')
                    identity = json.loads((base / 'project.json').read_text(encoding='utf-8-sig'))
                    common = ('MakeCertCapi', 'RibbonProbe', 'AddInProbe', 'GetPkt', 'ValidateManifest', 'ValidateXmlSig',
                              'SignCmi', 'TrustAdd', 'UiPreviewProbe', 'DialogSizesProbe', 'InstallFlowProbe', 'AgentExtensionProbe', 'AgentDocumentCoreProbe', 'SetupDoctor', 'GripProbe', 'LiveProbeSafetyFixture')
                    outputs = [base / 'build' / (name + '.exe') for name in common] + [base / 'build' / (t['name'] + '.exe') for t in identity['tools']]
                    outputs += [base / 'build' / name for name in (assembly+'.dll', 'Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
                                'Interop.Microsoft.Office.Core.dll', assembly+'.dll.manifest', assembly+'.vsto')]
                    target = base / 'AA-Program' / setup.name
                    outputs.append(target)
                    cache = ROOT / 'build/incremental-cache' / (assembly + '.json')
                    expected = inputs(base, snapshot)
                    hit = receipt(cache, expected, outputs) if args.mode == 'incremental' else None
                    if hit:
                        report.reuse(hit['time'], whole_gate=True)
                        print('CACHE: ' + assembly + ' 构建；原构建 ' + hit['time'], flush=True)
                        # Current key availability/expiry is checked even when reusing signed bytes.
                        run([sys.executable, '-c', 'from resign import ensure_certificate; ensure_certificate()'], base,
                            ROOT / ('build/' + assembly + '-certificate.log'))
                        shutil.copyfile(target, setup)
                        (report.directory / (assembly + '-rebuild.log')).write_text('CACHED '+hit['time']+'\n', encoding='utf-8')
                        reused.append(assembly)
                    else:
                        print('MISS: ' + assembly + '：完整模式、依赖变化或缓存/产物未通过核对', flush=True)
                        cache.unlink(missing_ok=True)
                        run([sys.executable, base / 'rebuild.py', '--verbose', *(['--background'] if args.background else []), '--mode', args.mode, '--site-html', snapshot, '--installer-output', setup],
                            base, ROOT / ('build/' + assembly + '-rebuild.log'))
                with report.gate('G-12' if assembly == 'Excel_Iterate' else 'G-15'):
                    validate = [sys.executable, base / 'validate.py', '--verbose', '--mode', args.mode, '--site-html', snapshot, '--installer', setup]
                    cache_updates.append((cache, inputs(base, snapshot), outputs))
                    if args.skip_ui:
                        validate.append('--skip-ui')
                    if args.background:
                        validate.append('--background')
                    run(validate, base, ROOT / ('build/' + assembly + '-validation.log'))
                with report.gate('G-13' if assembly == 'Excel_Iterate' else 'G-16'):
                    verify_payload(base, assembly, setup, snapshot, ROOT)
                target = base / 'AA-Program' / setup.name
                packages["word" if assembly=="WordUpdater" else "excel"]=setup
                updates.append((setup, target))
                delivered.append(target)
            with report.gate('G-17'):
                complete=stage/'site-downloads.html'
                canonical_text=snapshot.read_bytes().decode('utf-8')
                complete_text=attach_installers(canonical_text,packages)
                complete.write_bytes(complete_text.encode('utf-8'))
                if embedded_snapshot(complete.read_bytes().decode('utf-8'))!=canonical_text: raise ValueError('下载网页改变了内嵌正文')
                from delivery import version
                expected_packages={('word' if assembly=='WordUpdater' else 'excel'):
                    (packages['word' if assembly=='WordUpdater' else 'excel'],version(base)) for base,assembly in bases}
                verify_installers(complete_text,expected_packages)
            with report.gate('G-18'):
                run([sys.executable, ROOT / 'src/tools/site_probe.py', '--html', complete],ROOT,ROOT/'build/site-downloads-validation.log')
            updates.append((complete, ROOT / 'AA-Program/XyCabin.html'))
            evidence = {str(target): list(fingerprints([Path(source)]).values())[0] for source, target in updates}
            for base,assembly in bases: evidence[str(base/'AA-Program'/(assembly+'_Setup.exe'))]['version']=version(base)
            record = stage / 'delivery-record.json'
            record.write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding='utf-8')
            updates.append((record, ROOT / 'build/current-delivery.json'))
            with report.gate('G-19'):
                if input_fingerprint(ROOT, [base for base, _ in bases]) != before:
                    raise ValueError('构建期间文档源被修改，已中止发布；请等待构建结束后再修改并重跑')
            with report.gate('G-20'):
                publish(updates)
            with report.gate('G-21'):
                from tools.layout_probe import verify as verify_layout
                verify_layout(ROOT.parent)
            with report.gate('G-22'):
                for values in cache_updates:
                    save(*values)
                commit(pending)
            write_summary(report, args.mode, 'success', started, reused, skipped)
        finally:
            if previous_pending is None:
                os.environ.pop('XY_CACHE_PENDING', None)
            else:
                os.environ['XY_CACHE_PENDING'] = previous_pending
            # Stage contains only compiler snapshots, packages and a copied delivery record.
            if stage.parent.resolve() != (ROOT / 'build').resolve() or not stage.name.startswith('stage-'):
                raise ValueError('Unsafe build stage cleanup')
            shutil.rmtree(stage)
    except BaseException:
        try:
            write_summary(report, args.mode, 'failed', started, reused, skipped)
        except OSError:
            pass
        raise
    print('两个项目已同步完成；详细日志：' + str(report.directory), flush=True)
    for target in delivered:
        print(target, flush=True)
    print('网站预览：' + str(ROOT / 'AA-Program/XyCabin.html'), flush=True)


def main():
    with build_lock():
        rebuild()


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, RuntimeError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
