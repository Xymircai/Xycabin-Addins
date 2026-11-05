"""Shared project identity, configuration loading and compiler setup.

Each sibling has its own project.json and builds with the sibling C-03 website directory.
Versions and user-facing button settings still belong to the AA-Program configurator.
"""
from pathlib import Path
import importlib.util
import json
import subprocess
import os
import re
import sys

ROOT = Path(PROJECT_ROOT).resolve()
BUILD = ROOT / 'build'
PROGRAM_OUT = ROOT / 'AA-Program'
DESIGN_OUT = ROOT / 'AB-Graphic'
OUT = PROGRAM_OUT
DOCS = ROOT.parent / 'C-04-Devlog'
DOC_FILES = ('B0-Entry.md', 'A-UserGuide.md', 'B1-Project-Files.md',
             'B2-DevExperience.md', 'B3-Spec Coding.md', 'B4-Code Index.md')
CONFIGURE_NAME = {'C-01-ExcelIterate': 'C1-Exceliterate_Configure.py',
                  'C-02-WordUpdater': 'C2-Wordupdater_Configure.py'}[ROOT.name]
DOC_OUT = DOCS
SRC = ROOT / 'src'
LIB = ROOT / 'lib'
PROJECT = json.loads((ROOT / 'project.json').read_text(encoding='utf-8-sig'))
if PROJECT['host'] not in ('Word', 'Excel'):
    raise ValueError('project.json host must be Word or Excel')
for key in ('assembly', 'namespace', 'addin_key'):
    if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', PROJECT[key]):
        raise ValueError('Invalid project identity: ' + key)
if PROJECT['whitelist_host'] != ('WPS' if PROJECT['host'] == 'Word' else 'ET'):
    raise ValueError('Host and WPS whitelist must agree')
for rel in PROJECT['sources'] + [t['source'] for t in PROJECT['tools']]:
    source_path(rel, addin=rel in PROJECT['sources'])

CSC = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if not CSC.is_file():
    CSC = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
DEFINES = ['/define:WPS_WORD' if PROJECT['host'] == 'Word' else '/define:WPS_EXCEL']
PAYLOAD = [PROJECT['assembly'] + '.dll', 'Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
           'Interop.Microsoft.Office.Core.dll', PROJECT['assembly'] + '.dll.manifest',
           PROJECT['assembly'] + '.vsto']
COMMON_REF = ['/r:System.dll', '/r:System.Core.dll', '/r:Microsoft.CSharp.dll', '/r:System.Drawing.dll', '/r:System.IO.Compression.dll', '/r:System.Xml.dll', '/r:System.Security.dll', '/r:System.Runtime.Serialization.dll', '/r:System.Net.Http.dll']
INFO_CS = SRC / 'shared/ProjectInfo.g.cs'


def assembly_metadata(paths, reader=None):
    """Read actual CLR identities/references without executing the payload DLLs."""
    reader = Path(reader) if reader is not None else BUILD / 'GetPkt.exe'
    if not reader.is_file():
        raise ValueError('缺少程序集标识读取工具，请先运行 build.py：' + str(reader))
    paths = [Path(path).resolve() for path in paths]
    result = subprocess.run([str(reader), '--json', *map(str, paths)], capture_output=True, timeout=60)
    output = result.stdout.decode('utf-8-sig', 'replace')
    if result.returncode:
        raise ValueError('无法读取实际程序集标识：' + output + result.stderr.decode('utf-8', 'replace'))
    records = json.loads(output)
    if not isinstance(records, list) or len(records) != len(paths):
        raise ValueError('程序集标识读取结果不完整')
    for path, record in zip(paths, records):
        if Path(record['path']).resolve() != path or not all(key in record['identity'] for key in ('name', 'version', 'language', 'processorArchitecture')):
            raise ValueError('程序集标识读取结果与请求不一致：' + str(path))
    return records


def load_configure():
    previous = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec = importlib.util.spec_from_file_location('_wps_configuration', OUT / CONFIGURE_NAME)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        sys.dont_write_bytecode = previous


def write_project_info():
    """Generate only compile-time constants; JSON is never a runtime dependency."""
    p = PROJECT
    legacy = p['legacy_name']
    values = dict(AssemblyName=p['assembly'], AddinNamespace=p['namespace'], Host=p['host'],
                  DisplayName=p['display_name'], HostDisplayName=p['host_display_name'],
                  AddinName=p['addin_key'], FriendlyName=p['friendly_name'], Description=p['description'],
                  AddinKeyPath='Software\\Microsoft\\Office\\' + p['host'] + '\\Addins\\' + p['addin_key'],
                  SetupKeyPath='Software\\' + p['assembly'] + 'Setup',
                  AddinsWlKeyPath='Software\\Kingsoft\\Office\\' + p['whitelist_host'] + '\\AddinsWL',
                  PayloadVstoName=p['assembly'] + '.vsto', CertificateSubject='CN=' + p['assembly'] + ' Dev',
                  LegacyAddinName=legacy, LegacyVstoName=legacy + '.vsto' if legacy else '',
                  LegacyAddinKeyPath='Software\\Microsoft\\Office\\' + p['host'] + '\\Addins\\' + legacy if legacy else '',
                  LegacySetupKeyPath='Software\\' + legacy + 'Setup' if legacy else '')
    quote = lambda s: '@"' + s.replace('"', '""') + '"'
    lines = ['// Generated from project.json by project_support.py; do not edit.',
             'namespace WpsShared', '{', '    internal static class ProjectInfo', '    {']
    for name, value in values.items():
        lines.append('        internal const string ' + name + ' = ' + quote(value) + ';')
    for name, values in [('Payload', PAYLOAD), ('ProgIds', p['progids']),
                         ('LegacyPayload', [legacy + '.dll', PAYLOAD[1], PAYLOAD[2], legacy + '.dll.manifest', legacy + '.vsto'] if legacy else [])]:
        lines.append('        internal static readonly string[] ' + name + ' = new string[] { ' + ', '.join(quote(s) for s in values) + ' };')
    text = '\r\n'.join(lines + ['    }', '}', '']).encode('utf-8')
    INFO_CS.parent.mkdir(parents=True, exist_ok=True)
    if not INFO_CS.exists() or INFO_CS.read_bytes() != text:
        INFO_CS.write_bytes(text)
    return INFO_CS


SITE_ROOT = ROOT.parent / 'C-03-Website·Xycabin'


def load_site():
    """The website source must exist; generated previews are never a fallback."""
    for rel in ('AA-Program/C3-XycabinSite_Configure.py', 'src/site_core.py', 'src/site_template.html',
                'src/tools/site_probe.py', 'src/tools/layout_probe.py'):
        if not (SITE_ROOT / rel).is_file():
            raise FileNotFoundError('缺少集中网站目录文件，请与加载项一起携带 C-03-Website·Xycabin：' + str(SITE_ROOT / rel))
    for name in DOC_FILES:
        if not (DOCS / name).is_file():
            raise FileNotFoundError('缺少集中文档，请携带仓库根 C-04-Devlog：' + str(DOCS / name))
    spec = importlib.util.spec_from_file_location('_xy_site_core', SITE_ROOT / 'src/site_core.py')
    module = importlib.util.module_from_spec(spec)
    site_src = str(SITE_ROOT / 'src')
    if site_src not in sys.path:
        sys.path.insert(0, site_src)
    spec.loader.exec_module(module)
    return module


def verify_features():
    """R-30 is deliberately executed outside compiler and validation caches."""
    from feature_index import check
    return check(ROOT.parent, project_root=ROOT)


def verify_layout(source_only=True):
    """Source checks allow absent generated artifacts; completed deliveries require them."""
    load_site()
    spec = importlib.util.spec_from_file_location('_xy_layout_probe', SITE_ROOT / 'src/tools/layout_probe.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.verify(ROOT.parent, source_only=source_only, folders=(ROOT.name, SITE_ROOT.name))
