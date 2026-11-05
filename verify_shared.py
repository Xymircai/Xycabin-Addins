"""Check shared source and documentation; never compare old build artifacts."""
from pathlib import Path
import ast
import sys

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'C-00-Shared'))
from source_layout import manifest, resolve_source

DOCS = {'B0-Entry.md', 'A-UserGuide.md', 'B1-Project-Files.md',
        'B2-DevExperience.md', 'B3-Spec Coding.md', 'B4-Code Index.md'}
SHARED_TITLES = ('# 一. Word与Excel共享开发经验', '# 一. Word与Excel共享门禁')


def check_entry(project, relative):
    """Product Python entries may only dispatch with an explicit product root."""
    if not relative.endswith('.py'):
        return
    if not (ROOT / 'C-00-Shared' / relative).is_file():
        raise ValueError('缺少公共脚本实现：' + relative)
    depth = len(Path(relative).parts) - 1
    expected = '''"""Product entry point; implementation is compiled/executed from C-00-Shared."""
from pathlib import Path as _Path
import sys as _sys
_root = _Path(__file__).resolve().parents[%d]
_sys.dont_write_bytecode = True
_sys.path.insert(0, str(_root))
_sys.path.insert(1, str(_root.parent / "C-00-Shared"))
from source_layout import run_entry as _run_entry
_run_entry(globals(), %r, _root)
''' % (depth, relative)
    path = project / relative
    if not path.is_file() or ast.dump(ast.parse(path.read_text(encoding='utf-8-sig'))) != ast.dump(ast.parse(expected)):
        raise ValueError('公共脚本产品入口只能保留显式产品根目录的薄封装：' + str(path))


# rule=R-01
def check_documents():
    """Documents live once in C-04-Devlog; projects keep no copies."""
    agents = ROOT / 'AGENTS.md'
    if not agents.is_file():
        raise ValueError('缺少仓库根 AGENTS.md 指针')
    pointer = agents.read_text(encoding='utf-8-sig')
    if (len(pointer.encode('utf-8')) > 300 or pointer.count('](') != 1
            or '](C-04-Devlog/B0-Entry.md)' not in pointer):
        raise ValueError('AGENTS.md 必须只放一个指向 C-04-Devlog/B0-Entry.md 的链接')
    if '只作指针' not in pointer or '<!--' in pointer:
        raise ValueError('AGENTS.md 不得维护规则或包含注释')
    docs = ROOT / 'C-04-Devlog'
    actual = {p.name for p in docs.iterdir()} if docs.is_dir() else set()
    if actual != DOCS:
        raise ValueError('C-04-Devlog 必须恰好包含六份文档：' + ', '.join(sorted(actual)))
    texts = {name: (docs / name).read_text(encoding='utf-8-sig') for name in DOCS}
    for title in SHARED_TITLES:
        count = sum(1 for text in texts.values() for line in text.splitlines() if line.strip() == title)
        if count != 1:
            raise ValueError('共享章节必须只保留一份：' + title)
    for name, text in texts.items():
        if 'shared-wps-experience' in text:
            raise ValueError(name + ': 旧的共享标记必须删除')
    for project in (ROOT / 'C-01-ExcelIterate', ROOT / 'C-02-WordUpdater', ROOT / 'C-03-Website·Xycabin'):
        for name in ('开发日志.md', '使用说明.md', 'AGENTS.md'):
            leftovers = [p for p in project.rglob(name) if 'build' not in p.parts]
            if leftovers:
                raise ValueError('项目内不得保留文档副本：' + str(leftovers[0]))


# rule=R-15
def main():
    left, right = ROOT / 'C-01-ExcelIterate', ROOT / 'C-02-WordUpdater'
    for project in (left, right):
        if not project.is_dir():
            raise FileNotFoundError('共享检查需要两个加载项目录：' + str(project))
    from feature_index import check as check_feature_index
    check_feature_index(ROOT)
    errors = []
    layout = manifest()
    for relative in layout['linked']:
        if not (ROOT / 'C-00-Shared' / relative).is_file():
            errors.append('缺少公共实现：' + relative)
        for project in (left, right):
            if (project / relative).exists():
                errors.append('公共实现不得保留产品副本：' + str(project / relative))
    for project in (left, right):
        import json
        for relative in layout['entries']:
            try: check_entry(project, relative)
            except (OSError, ValueError, SyntaxError) as error: errors.append(str(error))
        data = json.loads((project / 'project.json').read_text(encoding='utf-8-sig'))
        for value in data['sources']:
            try: resolve_source(project, value, addin=True)
            except ValueError as error: errors.append(str(error))
        for tool in data['tools']:
            for value in [tool['source']] + tool.get('additional_sources', []):
                try: resolve_source(project, value)
                except ValueError as error: errors.append(str(error))
    for relative in layout['binary_checks']:
        if (left / relative).read_bytes() != (right / relative).read_bytes():
            errors.append('公共依赖不同：' + relative)
    try:
        check_documents()
    except (OSError, ValueError) as error:
        errors.append(str(error))
    for relative in ('src/document_assets.py', 'src/site_core.py', 'src/site_template.html',
                     'AA-Program/C3-XycabinSite_Configure.py'):
        path = ROOT / 'C-03-Website·Xycabin' / relative
        if not path.is_file():
            errors.append('缺少集中网站接入文件：' + str(path))
    if errors:
        raise ValueError('\n'.join(errors))
    print('PASS: C-04 documents, %d unique linked implementations and both product source manifests' % len(layout['linked']))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
