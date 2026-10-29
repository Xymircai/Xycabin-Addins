"""Verify a diagnostic build without installing or connecting to WPS.

Usage: python src/pdf/verify_workflow.py [--render]
Failures retain self-owned fixtures under build/pdf-workflow-evidence.
"""
from pathlib import Path
import argparse
import struct
import subprocess
import sys
import xml.etree.ElementTree as ET
import zlib
import uuid


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--render', action='store_true')
    args = parser.parse_args()
    root = Path(PROJECT_ROOT).resolve()
    sys.path.insert(0, str(root))
    import project_support as project
    evidence = root / 'build/pdf-workflow-evidence' / uuid.uuid4().hex[:12]
    evidence.mkdir(parents=True, exist_ok=True)
    trial = root / 'build/pdf-review'
    subprocess.run([sys.executable, str(root / 'src/pdf/build_trial.py'), '--output', str(trial)], check=True)
    for name, sources, refs in [
        ('PdfSelectionProbe', ['src/testharness/PdfSelectionProbe.cs', 'src/addin/PdfFileSelection.cs'], []),
        ('PdfToolsProbe', ['src/testharness/PdfToolsProbe.cs', 'src/tools/SourcePaths.cs'], ['System.Windows.Forms.dll', 'System.Drawing.dll', 'System.Runtime.Serialization.dll']),
        ('PdfAsyncUiProbe', ['src/testharness/PdfAsyncUiProbe.cs'], ['System.Windows.Forms.dll', 'System.Drawing.dll']),
        ('PdfWorkerHostProbe', ['src/testharness/PdfWorkerHostProbe.cs'], []),
        ('PdfCurrentDocumentProbe', ['src/testharness/PdfCurrentDocumentProbe.cs', 'src/tools/SourcePaths.cs'], ['Microsoft.CSharp.dll', 'System.Runtime.Serialization.dll']),
    ]:
        subprocess.run([str(project.CSC), '/nologo', '/codepage:65001', '/out:' + str(root / 'build' / (name + '.exe'))]
                       + project.DEFINES + ['/r:System.Core.dll'] + ['/r:' + ref for ref in refs]
                       + [str(project.source_path(source)) for source in sources], check=True)
    subprocess.run([str(root / 'build/PdfSelectionProbe.exe')], check=True)
    subprocess.run([str(root / 'build/PdfWorkerHostProbe.exe'), str(root / 'build/pdf-worker/PdfWorker.exe')], check=True)
    subprocess.run([str(root / 'build/PdfAsyncUiProbe.exe'), str(trial / (project.PROJECT['assembly'] + '.dll')), str(evidence)], check=True, timeout=30)
    subprocess.run([str(root / 'build/PdfCurrentDocumentProbe.exe'), str(trial / (project.PROJECT['assembly'] + '.dll')), str(evidence)], check=True, timeout=60)
    if args.render:
        subprocess.run([str(root / 'build/PdfToolsProbe.exe'), str(trial / (project.PROJECT['assembly'] + '.dll')), str(evidence), '--render-only'], check=True, timeout=120)

    # Generate a PNG using only the standard library, then exercise the real worker.
    def chunk(kind, data):
        return struct.pack('!I', len(data)) + kind + data + struct.pack('!I', zlib.crc32(kind + data) & 0xffffffff)
    image = evidence / 'source.png'
    image.write_bytes(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('!2I5B', 2, 2, 8, 2, 0, 0, 0))
                      + chunk(b'IDAT', zlib.compress(b'\x00\xff\x00\x00\xff\x00\x00' * 2)) + chunk(b'IEND', b''))
    count = 0

    def job(mode, sources, folders=None, expected=0, **attributes):
        nonlocal count
        count += 1
        output = evidence / ('case-' + str(count))
        output.mkdir()
        xml = ET.Element('PdfJob', mode=mode, output=str(output), **attributes)
        for i, source in enumerate(sources):
            entry = ET.SubElement(xml, 'File')
            entry.text = str(source)
            if folders:
                entry.set('folder', folders[i])
        request = evidence / ('request-' + str(count) + '.xml')
        ET.ElementTree(xml).write(request, encoding='utf-8', xml_declaration=True)
        process = subprocess.run([str(root / 'build/pdf-worker/PdfWorker.exe'), str(request)], timeout=60)
        report = ET.parse(str(request) + '.result').getroot()
        assert process.returncode == expected, ET.tostring(report, encoding='unicode')
        return output, report

    output, report = job('convert', [image, image], ['Project/Dept', 'Project (2)/Dept'])
    assert len(report.findall('Success')) == 2 and len(list(output.rglob('*.pdf'))) == 2
    pdf = Path(report.find('Success').text)
    output, report = job('convert', [image, image])
    assert (output / 'source.pdf').exists() and (output / 'source (1).pdf').exists()
    output, report = job('convert', [image], ['../escape'], expected=1)
    assert not list(output.rglob('*.pdf')) and report.find('Failure') is not None
    output, report = job('split', [pdf], ['Project/Dept'], ranges='')
    assert (output / 'Project/Dept/source_001.pdf').exists()
    output, report = job('merge', [pdf, pdf])
    assert len(report.findall('Success')) == 1 and len(list(output.rglob('*.pdf'))) == 1
    bad = evidence / 'invalid.pdf'
    bad.write_text('invalid pdf')
    output, report = job('merge', [pdf, bad], expected=1)
    assert not list(output.rglob('*.pdf')), 'Partial merge was published'
    output, report = job('convertmerge', [image, bad], expected=1)
    assert not list(output.rglob('*.pdf')), 'Partial conversion merge was published'
    output, report = job('publish', [pdf], name='Published', ranges='1')
    assert (output / 'Published.pdf').exists()
    output, report = job('watermark', [pdf], ['Project/Dept'], text='Test', opacity='20', size='48', angle='-35', color='808080', layout='center')
    assert len(report.findall('Success')) == 1
    assert image.exists() and pdf.exists(), 'Source changed'
    print('PASS: 9 worker workflows; hierarchy, collisions, traversal rejection, split, merge, failed atomic merge, current-document publish and watermark.')
    print('Evidence:', evidence)


if __name__ == '__main__':
    main()
