"""Build a diagnostic PDF-only launcher. This is deliberately NOT a VSTO installer.

The Office Core compile fixture is used only in this diagnostic build; the normal
build.py still requires the real Microsoft assemblies and original signing key.
"""
from pathlib import Path
import argparse
import importlib.util
import json
import subprocess
import sys


def build(output):
    root = Path(PROJECT_ROOT).resolve()
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    spec = importlib.util.spec_from_file_location('_pdf_trial_identity', root / 'project_support.py')
    project = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(project)
    project.write_project_info()
    subprocess.run([sys.executable, str(root / 'src/pdf/build_worker.py')], check=True)
    scratch = root / 'build/pdf-trial-compile'
    scratch.mkdir(parents=True, exist_ok=True)
    stub = scratch / 'OfficeCoreCompileFixture.cs'
    stub.write_text('namespace Microsoft.Office.Core { public interface IRibbonExtensibility { string GetCustomUI(string id); } '
                    'public interface IRibbonUI { void Invalidate(); void InvalidateControl(string id); } }', encoding='utf-8')
    sources = [project.source_path(source, addin=True) for source in project.PROJECT['sources'] if Path(source).name != 'ThisAddIn.cs']
    sources += [root / 'src/shared/ProjectInfo.g.cs', stub]
    library = output / (project.PROJECT['assembly'] + '.dll')
    subprocess.run([str(project.CSC), '/nologo', '/target:library', '/codepage:65001', '/out:' + str(library)]
                   + project.DEFINES + project.COMMON_REF + ['/r:System.Windows.Forms.dll',
                   '/resource:' + str(root / 'build/PdfWorker.zip') + ',WpsShared.PdfWorker.zip']
                   + list(map(str, sources)), check=True)
    executable = output / ('Word-PDF验证版.exe' if project.PROJECT['host'] == 'Word' else 'Excel-PDF验证版.exe')
    subprocess.run([str(project.CSC), '/nologo', '/target:winexe', '/codepage:65001', '/out:' + str(executable)]
                   + project.DEFINES + ['/r:System.Windows.Forms.dll', '/r:System.Drawing.dll',
                   str(project.source_path('src/tools/PdfTrialLauncher.cs'))], check=True)
    Path(str(executable) + '.config').write_text('<?xml version="1.0"?><configuration><startup>'
        '<supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.8"/></startup></configuration>', encoding='utf-8')
    subprocess.run([str(executable), '--check'], check=True, timeout=60)
    return executable


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    print(build(parser.parse_args().output))
