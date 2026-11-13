"""Read-only gate for the delivery folders and the single document home."""
from pathlib import Path
import argparse
import re
import sys
import tempfile
from urllib.parse import unquote

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src'))
from document_assets import collect as collect_document_assets
PROGRAM, DESIGN = 'AA-Program', 'AB-Graphic'
DOCS = {'B0-Entry.md', 'A-UserGuide.md', 'B1-Project-Files.md',
        'B2-DevExperience.md', 'B3-Spec Coding.md', 'B4-Code Index.md'}
DEVLOG = 'C-04-Devlog'
CONFIGURE = {'C-01-ExcelIterate': 'C1-Exceliterate_Configure.py',
             'C-02-WordUpdater': 'C2-Wordupdater_Configure.py',
             'C-03-Website·Xycabin': 'C3-XycabinSite_Configure.py'}
PROJECTS = {'C-01-ExcelIterate': 'Excel_Iterate', 'C-02-WordUpdater': 'WordUpdater', 'C-03-Website·Xycabin': None}
AUDIT_HEAD = re.compile(r'^###\s+\(\d+\)\.\s*本轮交付证据\s*$')


def audit_spans(text):
    """Offsets of every delivery-evidence subsection, found by heading only."""
    spans, offset, start = [], 0, None
    for line in text.split('\n'):
        if AUDIT_HEAD.match(line):
            start = offset
        elif start is not None and re.match(r'^#{1,2}\s', line):
            spans.append((start, offset))
            start = None
        offset += len(line) + 1
    if start is not None:
        spans.append((start, offset))
    return spans


# rule=R-12
# rule=R-02
def verify(parent=ROOT.parent, source_only=False, folders=None, strict_evidence=False):
    count = 0
    pending_outputs = {(Path(parent) / folder / PROGRAM /
                        (assembly + '_Setup.exe' if assembly else 'XyCabin.html')).resolve()
                       for folder, assembly in PROJECTS.items()}
    for folder, assembly in PROJECTS.items():
        if folders is not None and folder not in folders:
            continue
        base = Path(parent) / folder
        if not base.is_dir():
            continue  # An add-in + Website pair is a supported checkout.
        program = {CONFIGURE[folder], '一键重建并打包.bat',
                   assembly + '_Setup.exe' if assembly else 'XyCabin.html'}
        path = base / PROGRAM
        actual = {p.name for p in path.iterdir()} if path.is_dir() else set()
        required = program - ({assembly + '_Setup.exe'} if assembly else {'XyCabin.html'}) if source_only else program
        if not path.is_dir() or not required.issubset(actual) or not actual.issubset(program) \
                or any(not p.is_file() or p.is_symlink() for p in path.iterdir()):
            raise ValueError('输出目录必须只包含规定文件：' + str(path) + '；应为 ' + ', '.join(sorted(program)))
        count += len(program)
        art = base / DESIGN
        if not art.is_dir():
            raise ValueError('图形目录必须存在：' + str(art))
        for item in art.iterdir():
            allowed = item.is_file() and not item.is_symlink() and (
                item.suffix.lower() in ('.png', '.svg') or item.name in ('README.txt', '.gitkeep'))
            if not allowed:
                raise ValueError('图形目录只允许 PNG、SVG、README.txt 与 .gitkeep：' + str(item))
        for obsolete in ('AP1-程序', 'AP2-美工', 'AP3-说明', 'icons', 'src/documentation',
                         '输出文件—程序整合', '输出文件—美工设计', '输出文件—使用说明', '输出',
                         'README.md', '开发日志.md', 'AGENTS.md', 'Configure_program.py',
                         'Configure_Site.py', '一键重建并打包.bat'):
            if (base / obsolete).exists():
                raise ValueError('旧维护位置必须消失：' + str(base / obsolete))
        unique = {name: path / name for name in program}
        for file in base.rglob('*'):
            if not file.is_file():
                continue
            if source_only and any(part.startswith('stage-') for part in file.relative_to(base).parts):
                continue
            if file.name in unique and file.resolve() != unique[file.name].resolve():
                raise ValueError('重复维护或交付文件：' + str(file))
            if file.name in ('README.md', 'readme_probe.py', 'Configure.py', 'AgentReadme.md',
                             'AgentUserGuide.md', 'AgentHome.html'):
                raise ValueError('遗留维护副本：' + str(file))
            if assembly and file.name in ('XyCabin.html', 'C3-XycabinSite_Configure.py'):
                raise ValueError('加载项不得保留网站副本：' + str(file))
        if assembly:
            collect_document_assets(base)  # Includes inline PNGs, SVGs, and path/content checks.
        launcher = (base / PROGRAM / '一键重建并打包.bat').read_bytes()
        if launcher.startswith(b'\xef\xbb\xbf') or b'\n' in launcher.replace(b'\r\n', b'') or b'chcp 65001' not in launcher or b'PYTHONUTF8=1' not in launcher or b'PYTHONIOENCODING=utf-8' not in launcher:
            raise ValueError('批处理必须 UTF-8 无 BOM、CRLF、65001 和 Python UTF-8：' + str(base))
        launcher.decode('utf-8', 'strict')

    docs = Path(parent) / DEVLOG
    actual = {p.name for p in docs.iterdir()} if docs.is_dir() else set()
    if actual != DOCS or any(not p.is_file() or p.is_symlink() for p in docs.iterdir()):
        raise ValueError('文档目录必须只包含六份规定文件：' + str(docs) + '；应为 ' + ', '.join(sorted(DOCS)))
    count += len(DOCS)
    for doc in sorted(docs.glob('*.md')):
        text = doc.read_text(encoding='utf-8-sig')
        audits = audit_spans(text)
        for link in re.finditer(r'(?<!!)\[[^\]]*\]\(([^)]+)\)', text):
            target = link.group(1)
            target = unquote(target.split('#', 1)[0].strip('<>'))
            if not target or re.match(r'^[a-zA-Z]+:', target):
                continue
            destination = (doc.parent / target).resolve()
            optional = next((Path(parent) / name for name in PROJECTS if name in destination.parts), None)
            if optional is not None and not optional.exists():
                continue
            if not destination.exists():
                if source_only and destination in pending_outputs:
                    continue
                in_audit = any(start <= link.start() < end for start, end in audits)
                generated = (destination.is_relative_to(Path(parent).resolve() / 'build')
                             or (optional is not None and destination.is_relative_to(optional / 'build')))
                if in_audit and generated and not strict_evidence:
                    print('INFO: 可选审计证据未携带，不作为重建输入：' + str(destination), flush=True)
                    continue
                raise ValueError('文档链接失效：' + str(doc) + ' -> ' + target)
    print('PASS: 三类输出、唯一文档位置、旧路径、文档链接及 UTF-8 批处理；%d 项维护/交付文件' % count, flush=True)


def self_test():
    with tempfile.TemporaryDirectory(prefix='xycabin-layout-中文·') as temporary:
        parent = Path(temporary)
        base = parent / 'C-01-ExcelIterate'
        (base / PROGRAM).mkdir(parents=True)
        (base / DESIGN).mkdir()
        (base / PROGRAM / 'C1-Exceliterate_Configure.py').write_text('', encoding='utf-8')
        (base / PROGRAM / 'Excel_Iterate_Setup.exe').write_text('', encoding='utf-8')
        (base / PROGRAM / '一键重建并打包.bat').write_bytes(
            b'@echo off\r\nchcp 65001\r\nset PYTHONUTF8=1\r\nset PYTHONIOENCODING=utf-8\r\n')
        (base / DESIGN / 'shot.png').write_bytes(b'')
        (base / DESIGN / 'ui.svg').write_text('<svg xmlns="http://www.w3.org/2000/svg"/>', encoding='utf-8')
        (base / DESIGN / 'README.txt').write_text('', encoding='utf-8')
        docs = parent / DEVLOG
        docs.mkdir()
        (docs / 'B0-Entry.md').write_text('# 一. 入口\n', encoding='utf-8')
        (docs / 'B1-Project-Files.md').write_text('**项目档案**\n\n最后更新：2026-10-18\n\n# 一. Agent及开发者接手概览\n概览。\n', encoding='utf-8')
        (docs / 'B2-DevExperience.md').write_text('**经验**\n\n# 一. Word与Excel共享开发经验\n正文。\n', encoding='utf-8')
        (docs / 'B3-Spec Coding.md').write_text('**门禁**\n\n# 一. Word与Excel共享门禁\n正文。\n', encoding='utf-8')
        (docs / 'B4-Code Index.md').write_text('**代码索引**\n\n# 一. Excel-Iterate\n索引。\n# 二. WordUpdater\n索引。\n', encoding='utf-8')
        (docs / 'A-UserGuide.md').write_text('**使用说明**\n\n概览。\n\n# 一. Excel-Iterate\n[入口](../C-01-ExcelIterate/AA-Program/C1-Exceliterate_Configure.py)\n', encoding='utf-8')
        verify(parent, folders={'C-01-ExcelIterate'})
        (base / PROGRAM / 'Excel_Iterate_Setup.exe').unlink()
        verify(parent, source_only=True, folders={'C-01-ExcelIterate'})
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('Final gate accepted an absent artifact')
        (base / PROGRAM / 'Excel_Iterate_Setup.exe').write_text('', encoding='utf-8')
        (base / DESIGN / 'bad.txt').write_text('', encoding='utf-8')
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('图形目录接受了非图片文件')
        (base / DESIGN / 'bad.txt').unlink()
        (base / PROGRAM / 'extra.py').write_text('', encoding='utf-8')
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('输出目录接受了多余文件')
        (base / PROGRAM / 'extra.py').unlink()
        # New joint-build evidence lives under repository/build as well as
        # product/build. R-12 still requires it in strict mode and never
        # exempts links outside delivery-audit sections.
        archive = docs / 'B1-Project-Files.md'
        original_archive = archive.read_text(encoding='utf-8')
        archive.write_text(original_archive + '\n### (1). 本轮交付证据\n[历史报告](../build/history.json)\n', encoding='utf-8')
        verify(parent, folders={'C-01-ExcelIterate'})
        try:
            verify(parent, folders={'C-01-ExcelIterate'}, strict_evidence=True)
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('Strict evidence accepted a missing root-build report')
        guide = docs / 'A-UserGuide.md'
        original_guide = guide.read_text(encoding='utf-8')
        guide.write_text(original_guide + '\n[正文报告](../build/history.json)\n', encoding='utf-8')
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('Non-audit missing root-build link was accepted')
        guide.write_text(original_guide, encoding='utf-8')
        archive.write_text(original_archive, encoding='utf-8')
        (docs / 'B3-Spec Coding.md').unlink()
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('文档目录接受了缺失文件')
        (docs / 'B3-Spec Coding.md').write_text('**门禁**\n\n# 一. Word与Excel共享门禁\n正文。\n', encoding='utf-8')
        canonical = docs / 'B3-Spec Coding.md'
        for obsolete in ('Spec Coding.md', 'B3-Gate Control.md'):
            legacy = docs / obsolete
            canonical.rename(legacy)
            try:
                verify(parent, folders={'C-01-ExcelIterate'})
            except (OSError, ValueError):
                pass
            else:
                raise AssertionError('文档目录接受了旧规范名称：' + obsolete)
            legacy.rename(canonical)
            legacy.write_bytes(canonical.read_bytes())
            try:
                verify(parent, folders={'C-01-ExcelIterate'})
            except (OSError, ValueError):
                pass
            else:
                raise AssertionError('文档目录接受了新旧规范并存：' + obsolete)
            legacy.unlink()
        verify(parent, folders={'C-01-ExcelIterate'})
        (docs / 'A-UserGuide.md').write_text('**使用说明**\n\n概览。\n\n# 一. Excel-Iterate\n[失效](missing.md)\n', encoding='utf-8')
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('Accepted missing document link')
        (docs / 'A-UserGuide.md').write_text('**使用说明**\n\n概览。\n\n# 一. Excel-Iterate\n[入口](../C-01-ExcelIterate/AA-Program/C1-Exceliterate_Configure.py)\n', encoding='utf-8')
        (base / 'AP1-程序').mkdir()
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('Accepted the obsolete AP1 directory')
        (base / 'AP1-程序').rmdir()
        (base / PROGRAM / '一键重建并打包.bat').write_bytes(b'@echo off\nchcp 65001\n')
        try:
            verify(parent, folders={'C-01-ExcelIterate'})
        except (OSError, ValueError):
            pass
        else:
            raise AssertionError('Accepted a launcher without CRLF/Python UTF-8')
    print('LayoutProbe self-test: 16 PASS, 0 FAIL', flush=True)


if __name__ == '__main__':
    try:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument('--self-test', action='store_true')
        parser.add_argument('--strict-evidence', action='store_true', help='正式审计同时要求所有证据链接存在')
        args = parser.parse_args()
        if args.self_test:
            self_test()
        verify(strict_evidence=args.strict_evidence)
    except (OSError, ValueError, AssertionError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
