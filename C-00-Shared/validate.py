"""Run offline regressions and render previews; never connect to WPS or install."""
from pathlib import Path
import re
import argparse
import shutil
import subprocess
import sys
import time
import json
sys.dont_write_bytecode = True
from project_support import ROOT, BUILD, OUT, PROJECT, SITE_ROOT, load_site, DESIGN_OUT, verify_features
from verify_package import verify
sys.path.insert(0, str(SITE_ROOT / "src"))
from build_cache import inputs, key, digest, receipt, save


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--skip-ui', action='store_true')
    parser.add_argument('--background', action='store_true', help='只运行已核对不会显示或激活窗口的探针，其余明确列为待验收')
    parser.add_argument('--site-html', type=Path, default=SITE_ROOT / 'AA-Program/XyCabin.html')
    parser.add_argument('--installer', type=Path, default=OUT / (PROJECT['assembly'] + '_Setup.exe'))
    args = parser.parse_args()
    mode, background = args.mode, args.background
    classification = verify_features()
    load_site().validate_snapshot(args.site_html, ROOT)
    verify()
    dll = BUILD / (PROJECT['assembly'] + '.dll')
    tests = [(Path(sys.executable), [ROOT / 'src/tools/devlog_probe.py']), (Path(sys.executable), [ROOT / 'src/tools/configure_probe.py']),
             (Path(sys.executable), [source_path('@shared/src/tools/source_layout_probe.py'), '--project-root', ROOT]),
             (Path(sys.executable), [ROOT / 'src/tools/dialog_catalog_probe.py']),
             (Path(sys.executable), [ROOT / 'src/tools/window_thread_probe.py']),
             (Path(sys.executable), [SITE_ROOT / 'src/tools/document_assets_probe.py', '--root', ROOT]),
             (Path(sys.executable), [SITE_ROOT / 'src/tools/site_probe.py', '--html', args.site_html])]
    # --skip-ui alone does not suppress forms opened by business/Agent probes.
    background_tools = {'PdfSelectionProbe', 'RuntimeQualityProbe', 'HostDialogProbe', 'FormulaProbe', 'FormulaInteractionProbe', 'Feature099Probe', 'TrackerProbe', 'ErrorSourceProbe', 'WorkbookToolsProbe', 'FileOperationsProbe', 'MergeHeadersProbe', 'EliminationProbe', 'DisciplineProbe',
                        'LinkViewProbe', 'LinkMapProbe', 'UpdateProbe', 'LinkViewFormProbe', 'EtCleanupProbe', 'EtSelectionReaderProbe', 'ClipboardBoundaryFixture'}
    pending = []
    for tool in PROJECT['tools']:
        if not tool.get('offline'):
            continue
        if args.background and tool['name'] not in background_tools:
            pending.append(tool['name'])
        else:
            tests.append((BUILD / (tool['name'] + '.exe'), ['--core'] if args.background and tool['name'] == 'FormulaProbe' else []))
            if args.background and tool['name'] == 'FormulaProbe':
                pending.append('FormulaProbe 前台窗口路径')
    tests += [(BUILD / 'PdfAsyncUiProbe.exe', [dll, BUILD / 'pdf-ui-regression']),
              (BUILD / 'PdfWorkerHostProbe.exe', [BUILD / 'pdf-worker/PdfWorker.exe']),
              (BUILD / 'AgentDocumentCoreProbe.exe', [dll]),
              (BUILD / 'GripProbe.exe', [dll]), (BUILD / 'LiveProbeSafetyFixture.exe', []),
              (BUILD / 'InstallFlowProbe.exe', []),
              (BUILD / 'RibbonProbe.exe', [dll]),
              (BUILD / 'ValidateManifest.exe', [BUILD / (PROJECT['assembly'] + '.dll.manifest'), BUILD / (PROJECT['assembly'] + '.vsto')])]
    if args.background:
        tests += [(BUILD / 'DialogSizesProbe.exe', [dll, '--core'])]
        pending += ['DialogSizesProbe 前台预览/拖动', 'AgentExtensionProbe', 'UiPreviewProbe']
    else:
        tests += [(BUILD / 'DialogSizesProbe.exe', [dll]),
                  (BUILD / 'AgentExtensionProbe.exe', [dll, args.site_html])]
    render_ui = not (args.skip_ui or args.background)
    if render_ui:
        installer = args.installer
        if not installer.is_file():
            raise FileNotFoundError('Installer preview requires a packaged installer: ' + str(installer))
        rendered = BUILD / 'ui-preview-current'
        rendered.mkdir(exist_ok=True)
        tests += [(BUILD / 'UiPreviewProbe.exe', [dll, installer, rendered])]
    logs = ['待前台验收（本轮未运行）：' + ', '.join(pending)] if pending else []
    if pending:
        print(logs[0], flush=True)
    log_path = BUILD / ('background-validation.log' if args.background else 'offline-validation.log')
    dependency_key = inputs(ROOT, args.site_html)
    pending_receipts, timings = [], []
    for exe, args in tests:
        started = time.perf_counter()
        cacheable = background and exe.stem in background_tools
        cache = BUILD / 'cache/validation' / (exe.stem + '.json')
        expected = key([dependency_key, digest(exe), digest(dll), list(map(str, args))]) if cacheable else None
        hit = receipt(cache, expected, [exe, dll]) if cacheable and mode == 'incremental' else None
        print('TEST: ' + exe.name + ('（复用）' if hit else '（执行）'), flush=True)
        if hit:
            output = hit['output']
            logs.append(exe.name + ' CACHED ' + hit['time'] + '\n' + output)
            log_path.write_text('\n'.join(logs), encoding='utf-8')
            print('CACHE: ' + exe.name + '；原检查 ' + hit['time'], flush=True)
            timings.append({'test': exe.name, 'seconds': time.perf_counter()-started, 'cached': True, 'evidence_time': hit['time']})
            continue
        if cacheable:
            cache.unlink(missing_ok=True)
        result = subprocess.run([str(exe), *map(str, args)], cwd=ROOT, capture_output=True, timeout=120)
        codec = 'gbk' if exe.stem == 'ValidateManifest' else 'utf-8'
        output = result.stdout.decode(codec, 'replace') + result.stderr.decode(codec, 'replace')
        logs.append(exe.name + '\n' + output)
        log_path.write_text('\n'.join(logs), encoding='utf-8')
        if result.returncode or 'FAILED:' in output:
            print(output)
            raise ValueError(exe.name + ' failed with exit ' + str(result.returncode))
        summary = next((s for s in reversed(output.splitlines()) if s.strip()), '')
        elapsed = time.perf_counter() - started
        timings.append({'test': exe.name, 'seconds': elapsed, 'cached': False})
        print('PASS:', exe.name, summary, '（%.2f秒）' % elapsed, flush=True)
        if cacheable:
            pending_receipts.append((cache, expected, [exe, dll], output))
    log_path.write_text('\n'.join(logs), encoding='utf-8')
    for values in pending_receipts:
        save(*values)
    (BUILD / 'validation-timings.json').write_text(json.dumps(
        {'mode': mode, 'classification_gate': classification, 'tests': timings},
        ensure_ascii=False, indent=2), encoding='utf-8')
    if render_ui:
        preview = DESIGN_OUT
        preview.mkdir(exist_ok=True)
        for pattern in ('Agent-README-*.png', 'Agent-使用说明-*.png'):
            for retired in preview.glob(pattern):
                retired.unlink()
        # A native frame may be skipped when another window occludes it. Never
        # retain an older exported frame after the probe removed its build copy.
        for previous in preview.glob('*-原生标题栏-100.png'):
            if not (BUILD / 'ui-preview-current' / previous.name).is_file():
                previous.unlink()
        for f in (BUILD / 'ui-preview-current').glob('*.png'):
            if f.name.endswith('-100.png') or f.name.startswith('交互弹窗设置-') or f.name in ('功能区图标.png', '功能区图标-前后对照.png'):
                (preview / f.name).unlink(missing_ok=True)
                shutil.move(str(f), str(preview / f.name))
        # Keep all four settings-table scales; other alternate scales remain represented in the log.
        shutil.rmtree(BUILD / 'ui-preview-current')


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, subprocess.SubprocessError) as exc:
        print('FAIL:', exc)
        sys.exit(1)
