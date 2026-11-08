"""Compile the WPS add-in and diagnostics with the shared C# 5 pipeline."""
import os
from pathlib import Path
import argparse
import subprocess
from process_support import run_process
import sys
sys.dont_write_bytecode = True
from project_support import ROOT, SRC, BUILD, LIB, CSC, COMMON_REF, PROJECT, DEFINES, load_configure, write_project_info, load_site, SITE_ROOT, DOC_OUT, OUT, verify_layout, verify_features

sys.path.insert(0, str(SITE_ROOT / "src"))
from build_cache import compile_cached
from document_assets import collect as collect_document_assets, document_overview, manifest_bytes, MANIFEST_RESOURCE
from tool_catalog import tools
MODE = "full"


def run(args, cwd=ROOT):
    args = [str(a) for a in args]
    print('[cmd] ' + subprocess.list2cmdline(args), flush=True)
    if Path(args[0]) == CSC:
        compile_cached(args, cwd, MODE)
    else:
        run_process(args, cwd=cwd, check=True)


def compile_tool(name, source, references=(), output=None, identity=False, additional_sources=()):
    target = output or BUILD / (name + '.exe')
    args = [CSC, '/nologo', '/target:exe', '/platform:anycpu', '/codepage:65001', '/out:' + str(target)] + DEFINES + COMMON_REF + [source_path('src/tools/SourcePaths.cs')]
    for ref in references:
        args.append('/r:' + (str(BUILD / (PROJECT['assembly'] + '.dll')) if ref == '@addin' else ref))
    args.append(source_path(source))
    args += [source_path(path) for path in additional_sources]
    if identity:
        args.append(SRC / 'shared/ProjectInfo.g.cs')
    run(args)
    print('OK ->', target, flush=True)


def document_resources(site_html=None):
    resources = []
    bundle = collect_document_assets(ROOT)
    for document in bundle['documents']:
        filename, resource = document['path'], document['name']
        text = document['content'].decode('utf-8')
        try:
            document_overview(text)
        except ValueError as error:
            raise ValueError(filename + ": " + str(error)) from error
        # Embed normalized UTF-8, including files copied from a BOM/CRLF editor.
        target = BUILD / ("Embedded" + resource + ".md")
        target.write_bytes(document['content'])
        resources.append('/resource:' + str(target) + ',WpsShared.Agent.' + resource + '.md')
    manifest = BUILD / 'AgentAssets.xml'
    manifest.write_bytes(manifest_bytes(bundle))
    resources.append('/resource:' + str(manifest) + ',' + MANIFEST_RESOURCE)
    resources += ['/resource:' + str(asset['source']) + ',' + asset['resource'] for asset in bundle['assets']]
    website = load_site()
    site = website.validate_snapshot(site_html, ROOT, require_embedded=True) if site_html else website.build(SITE_ROOT, target=BUILD / "site-embedded.snapshot.html")
    resources.append('/resource:' + str(site) + ',WpsShared.Agent.Home.html')

    return resources


def main():
    global MODE
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--site-html')
    parser.add_argument('--with-testharness', action='store_true')
    args = parser.parse_args()
    MODE = args.mode
    verify_features()
    website = load_site()
    if args.site_html:
        website.validate_snapshot(args.site_html, ROOT, require_embedded=True)
    verify_layout(source_only=True)
    if "--with-testharness" in sys.argv:
        raise ValueError("旧 COM 注册测试器已归档；使用 project.json 中的现行工具清单")
    run([sys.executable, ROOT / "src/tools/devlog_probe.py"])
    run([sys.executable, ROOT / "src/tools/dialog_catalog_probe.py"])
    run([sys.executable, source_path('@shared/src/tools/source_layout_probe.py'), '--project-root', ROOT])
    checker = ROOT.parent / 'verify_shared.py'
    if all((ROOT.parent / n).is_dir() for n in ('C-01-ExcelIterate', 'C-02-WordUpdater')):
        if not checker.is_file():
            raise FileNotFoundError('缺少共享检查入口：' + str(checker))
        run([sys.executable, checker])
    if not CSC.is_file():
        raise FileNotFoundError('csc.exe not found: ' + str(CSC))
    BUILD.mkdir(exist_ok=True)
    OUT.mkdir(exist_ok=True)
    cfg = load_configure()
    cfg.main()  # Always regenerate, so an edited setting never leaves stale callbacks/text.
    write_project_info()
    vsto = ['Microsoft.Office.Tools.Common.v4.0.Utilities.dll', 'Microsoft.Office.Tools.' + PROJECT['host'] + '.dll',
            'Microsoft.VisualStudio.Tools.Applications.Runtime.dll', 'Microsoft.Office.Tools.dll',
            'Interop.Microsoft.Office.Core.dll', 'Microsoft.Office.Tools.Common.dll']
    target = BUILD / (PROJECT['assembly'] + '.dll')
    pdf_resources = []
    if PROJECT.get('pdf_worker'):
        import importlib.util
        spec = importlib.util.spec_from_file_location('_pdf_worker_build', SRC / 'pdf/build_worker.py')
        pdf_builder = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(pdf_builder)
        pdf_resources = ['/resource:' + str(pdf_builder.build()) + ',WpsShared.PdfWorker.zip']
    try:
        run([CSC, '/nologo', '/target:library', '/platform:anycpu', '/optimize+', '/codepage:65001', '/out:' + str(target)]
            + DEFINES + COMMON_REF + pdf_resources + document_resources(args.site_html) + ['/r:System.Windows.Forms.dll']
            + ['/r:' + str(LIB / n) for n in vsto]
            + [source_path(f, addin=True) for f in PROJECT['sources']])
    finally:
        # Embedded documents stay in build/ as the exact inputs of this compilation; they are
        # whole normalized copies of C-04-Devlog, regenerated on every build, never edited by hand.
        # Cleanup runs on failure too, so an aborted build never leaves them behind for the layout gate.
        for name in ('AgentAssets.xml',):
            (BUILD / name).unlink(missing_ok=True)
    # Only these two dependencies are distributed; host libraries remain GAC references.
    for name in ('Microsoft.Office.Tools.Common.v4.0.Utilities.dll', 'Interop.Microsoft.Office.Core.dll'):
        (BUILD / name).write_bytes((LIB / name).read_bytes())
    for tool in tools(PROJECT):
        compile_tool(tool['name'], tool['source'], tool['references'], identity=tool.get('identity', False),
                     additional_sources=tool.get('additional_sources', ()))
    print('Build complete. Re-sign with resign.py before packaging.', flush=True)


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        print('FAIL:', exc)
        sys.exit(1)
