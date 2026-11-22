"""Feature source ownership and portable Markdown navigation, derived from configuration."""
from pathlib import Path
import argparse
from collections import defaultdict
import copy
import importlib.util
import os
import re
import sys
from unittest.mock import patch

from source_layout import addin_sources, manifest, project_data, resolve_source

sys.dont_write_bytecode = True
PRODUCTS = (('C-01-ExcelIterate', 'C1-Exceliterate_Configure.py', 'Excel-Iterate'),
            ('C-02-WordUpdater', 'C2-Wordupdater_Configure.py', 'WordUpdater'))
INDEX_NAME = 'B4-Code Index.md'
FEATURE_ROOT = 'AC-Original Code'
INFRASTRUCTURE = {
    'C-01-ExcelIterate': {'AssemblyInfo.cs', 'Ribbon.g.cs', 'Connect.cs', 'ThisAddIn.cs',
                        'WorkbookFeatures.cs', 'CustomizationCallbacks.g.cs', 'ApiPlatforms.g.cs'},
    'C-02-WordUpdater': {'AssemblyInfo.cs', 'Ribbon.g.cs', 'Connect.cs', 'ThisAddIn.cs',
                       'CustomizationCallbacks.g.cs'},
}


def configuration(root):
    name = next(name for directory, name, _ in PRODUCTS if directory == root.name)
    spec = importlib.util.spec_from_file_location('_feature_index_' + root.name, root / 'AA-Program' / name)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def menu_nodes(root):
    """Read stable node identities, including hidden tabs and disabled leaves."""
    config = configuration(root)
    tabs = [(config.TAB_ID, config.TAB_LABEL, config.BUTTONS)]
    tabs += [(tab['id'], tab['label'], tab['buttons']) for tab in config.EXTRA_TABS]
    result = []
    def visit(items, labels, ids):
        for item in items:
            if item.get('type') == 'menu' and not item.get('children'):
                continue
            branch, identifiers = labels + [item['label']], ids + [item['id']]
            result.append(dict(item, labels=branch, identifiers=identifiers))
            if item.get('children'):
                visit(item['children'], branch, identifiers)
    for tab_id, tab, buttons in tabs:
        for button in buttons:
            visit([button], [tab, button['group'][1]], [tab_id, button['group'][0]])
    return [(label, buttons) for _, label, buttons in tabs], result


def commands(root):
    tabs, nodes = menu_nodes(root)
    return tabs, [node for node in nodes if not node.get('children')]


def expected_directories(root, project):
    _, nodes = menu_nodes(root)
    identities = {identity for node in nodes for identity in node['identifiers']}
    layout = project.get('code_layout', {})
    if set(layout) != identities:
        raise ValueError('菜单节点英文目录映射缺失或多余：' + str(sorted(set(layout) ^ identities)))
    if any(not isinstance(slug, str) or not re.fullmatch(r'[A-Z][A-Za-z0-9]*', slug)
           for slug in layout.values()):
        raise ValueError('菜单节点目录名必须采用英文 PascalCase')
    seen, result, parents = {}, {}, {FEATURE_ROOT}
    for node in nodes:
        parts = [FEATURE_ROOT] + [layout[identity] for identity in node['identifiers']]
        directory = '/'.join(parts)
        for depth in range(2, len(parts) + 1):
            path = '/'.join(parts[:depth])
            identity = tuple(node['identifiers'][:depth - 1])
            previous = seen.setdefault(path.casefold(), identity)
            if previous != identity:
                raise ValueError('不同菜单节点映射到同一目录：' + path)
            parents.add(path)
        result[node['id']] = directory
    return nodes, result, parents


def code(text):
    pattern = r'//[^\n]*|/\*[\s\S]*?\*/|@"(?:""|[^"])*"|"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\''
    return re.sub(pattern, lambda m: ''.join('\n' if c == '\n' else ' ' for c in m[0]), text)


def callback(command):
    return command.get('callback') or 'OnOpenUrl_' + command['id']


def _verify_product(root, project=None):
    root = Path(root).resolve()
    project = project if project is not None else project_data(root)
    _, leaves = commands(root)
    expected = {leaf['id'] for leaf in leaves}
    features = project.get('features', {})
    if set(features) != expected:
        raise ValueError(root.name + ': 功能映射与菜单叶子不一致：' + str(sorted(set(features) ^ expected)))
    paths = addin_sources(root, project)
    if len(paths) != len(set(paths)):
        raise ValueError(root.name + ': 编译源码重复')
    compiled = set(project['sources'])
    outside = {ref for ref in compiled if not ref.startswith(('@shared/', FEATURE_ROOT + '/'))}
    allowed = {'src/addin/' + name for name in INFRASTRUCTURE[root.name]}
    if outside - allowed:
        raise ValueError('业务源码必须位于 AC，基础设施白名单外存在编译输入：' + str(sorted(outside - allowed)))
    common = {'@shared/' + ref for ref in manifest()['linked']}
    if {ref for ref in compiled if ref.startswith('@shared/')} - common:
        raise ValueError('公共编译输入必须登记在 C-00-Shared/source-manifest.json')
    nodes, directories, shared_parents = expected_directories(root, project)
    for old, references in project.get('source_aliases', {}).items():
        if not references or len(references) != len(set(references)) or any(ref not in compiled for ref in references):
            raise ValueError(root.name + ': 拆分源码读取别名无效：' + old)
    root_directory = root / FEATURE_ROOT
    if not root_directory.is_dir() or root_directory.is_symlink():
        raise ValueError('缺少功能源码目录：' + str(root_directory))
    actual = set()
    for file in root_directory.rglob('*'):
        if file.is_symlink() or not file.resolve().is_relative_to(root_directory.resolve()):
            raise ValueError('功能目录不得包含越界路径或链接：' + str(file))
        if file.is_file():
            if file.suffix != '.cs':
                raise ValueError('功能目录只放功能 C# 源码：' + str(file))
            actual.add(file.relative_to(root).as_posix())
        if any(not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.-]*', part)
               for part in file.relative_to(root_directory).parts):
            raise ValueError('功能路径必须采用稳定英文名称：' + str(file))
        if len(str(file)) >= 260:
            raise ValueError('功能路径超过当前构建链兼容长度：' + str(file))
    declared = {ref for ref in compiled if ref.startswith(FEATURE_ROOT + '/')}
    if actual != declared:
        raise ValueError(root.name + ': 功能目录与实际编译清单不一致：' + str(sorted(actual ^ declared)))
    texts = {ref: code(resolve_source(root, ref).read_text(encoding='utf-8-sig')) for ref in project['sources']}
    combined = '\n'.join(texts.values())
    seen_directories, seen_entries, hierarchy = set(), set(), {}
    for leaf in leaves:
        bid = leaf['id']
        feature = features[bid]
        directory, entry = feature['directory'], feature['entry']
        parts = Path(directory).parts
        if directory != directories[bid]:
            raise ValueError('功能目录与菜单 ID 路径不一致：' + bid + '；应为 ' + directories[bid])
        if directory in seen_directories or entry in seen_entries:
            raise ValueError('多个末级功能共用目录或入口：' + bid)
        seen_directories.add(directory)
        seen_entries.add(entry)
        for depth in range(1, len(leaf['identifiers'])):
            key = tuple(leaf['identifiers'][:depth])
            prefix = '/'.join(parts[:depth + 1])
            if key in hierarchy and hierarchy[key] != prefix:
                raise ValueError('同一菜单节点映射到多个目录：' + bid)
            hierarchy[key] = prefix
        if entry not in compiled or Path(entry).parent.as_posix() != directory:
            raise ValueError('功能入口必须位于其末级目录并参与编译：' + bid)
        name = callback(leaf)
        signature = r'\bpublic\s+void\s+' + re.escape(name) + r'\s*\('
        if len(re.findall(signature, texts[entry])) != 1 or len(re.findall(signature, combined)) != 1:
            raise ValueError('按钮回调必须唯一且存在于实际入口：' + bid + '/' + name)
        image = r'\bpublic\s+IPictureDisp\s+GetImage_' + re.escape(bid) + r'\s*\('
        if len(re.findall(image, texts[entry])) != 1 or len(re.findall(image, combined)) != 1:
            raise ValueError('按钮图标回调必须唯一且存在于实际入口：' + bid)
        if len(feature.get('implementation', [])) != len(set(feature.get('implementation', []))):
            raise ValueError('功能实现依赖重复：' + bid)
        for reference in feature.get('implementation', []):
            if reference not in compiled:
                raise ValueError('功能实现依赖必须参与当前产品编译：' + bid + '/' + reference)
            resolve_source(root, reference)
        tool_sources = {tool['source'] for tool in project['tools']}
        for reference in feature.get('tests', []):
            if reference not in tool_sources:
                raise ValueError('功能验证入口必须来自工具清单：' + bid + '/' + reference)
            resolve_source(root, reference)
    owners = defaultdict(set)
    for bid, feature in features.items():
        for reference in [feature['entry']] + feature.get('implementation', []):
            owners[reference].add(bid)
    menu_entries = set()
    for node in nodes:
        if not node.get('children'):
            continue
        entry = directories[node['id']] + '/Connect.' + re.sub(r'^menu', '', node['id']) + '.Entry.cs'
        expected_image = 'GetImage_' + node['id']
        if entry not in compiled:
            raise ValueError('父菜单缺少已编译图标入口：' + node['id'])
        methods = re.findall(r'\b(?:public|private|internal)\s+(?:static\s+)?[\w<>,.\[\]]+\s+(\w+)\s*\(', texts[entry])
        if methods != [expected_image]:
            raise ValueError('父菜单目录只允许本菜单图标接线，业务实现应归叶子或 _Shared：' + entry)
        if len(re.findall(r'\bpublic\s+IPictureDisp\s+' + re.escape(expected_image) + r'\s*\(', combined)) != 1:
            raise ValueError('父菜单图标回调必须唯一：' + node['id'])
        menu_entries.add(entry)
    for reference in declared - menu_entries:
        consumers = owners.get(reference)
        if not consumers:
            raise ValueError('功能源码未归入任何功能映射：' + reference)
        parent = Path(reference).parent.as_posix()
        if Path(parent).name == '_Shared':
            ancestor = Path(parent).parent.as_posix()
            closest = os.path.commonpath([features[bid]['directory'] for bid in consumers]).replace('\\', '/')
            if len(consumers) < 2 or ancestor not in shared_parents or ancestor != closest:
                raise ValueError('共用源码必须位于使用功能的最近公共父级 _Shared：' + reference + '；应为 ' + closest + '/_Shared')
        elif not any(features[bid]['directory'] == parent for bid in consumers):
            raise ValueError('专属实现必须直接位于所属末级功能目录：' + reference)
    return leaves


# rule=R-30
def verify_product(root, project=None):
    try:
        return _verify_product(root, project)
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ValueError('R-30 功能源码分类门禁：' + str(error)) from error


def literal(value):
    return chr(96) + value + chr(96)


def link(root, reference, label=None):
    source = resolve_source(root, reference)
    relative = source.relative_to(root.parent).as_posix()
    return '[' + (label or source.name) + '](<../' + relative + '>)'


def tree(root, leaves, features, tab):
    labels, branches = {}, set()
    for leaf in leaves:
        if leaf['labels'][0] != tab:
            continue
        parts = Path(features[leaf['id']]['directory']).parts[1:]
        for depth, label in enumerate(leaf['labels'], 1):
            key = parts[:depth]
            branches.add(key)
            labels[key] = label
    tab_name = next(key[0] for key in branches)
    for source in project_data(root)['sources']:
        if source.startswith(FEATURE_ROOT + '/' + tab_name + '/') and '/_Shared/' in source:
            parts = Path(source).parent.parts[1:]
            for depth in range(1, len(parts) + 1):
                branches.add(parts[:depth])
    children = {}
    for key in sorted(branches):
        children.setdefault(key[:-1], []).append(key)
    output = [FEATURE_ROOT + '/']
    def emit(parent, prefix):
        entries = children.get(parent, [])
        for position, key in enumerate(entries):
            last = position == len(entries) - 1
            label = labels.get(key, '同项目共用业务代码' if key[-1] == '_Shared' else '')
            output.append(prefix + ('└─ ' if last else '├─ ') + key[-1] + '/' + ('（' + label + '）' if label else ''))
            emit(key, prefix + ('   ' if last else '│  '))
    emit((), '')
    return '\n'.join(output)


def render(repository, project_root=None):
    output = ['**功能代码索引**', '',
              '本文件按两个加载项的默认菜单结构定位功能源码。先在树中找到中文功能，再打开表中的英文目录、按钮入口和核心实现；文件内的方法名可直接搜索。', '',
              '路径以仓库根目录为基准，链接相对于本文件，可随整个仓库移动。界面自定名称、排序或隐藏状态不改变源码目录，稳定控件 ID 用于核对身份。', '',
              '按钮入口负责接线；专属算法和弹窗位于对应末级目录，产品内部共用业务代码位于就近的 _Shared，跨产品实现只在 C-00-Shared 维护。'
              '加载框架、生成文件、测试、安装器、依赖及构建产物分别保留在 src、lib、build 等工程位置。复现需同时携带表列依赖及工程底座，不将末级目录视为独立项目。', '',
              '内容由配置器菜单、各产品 project.json 的 code_layout 英文节点名及 features 功能映射生成。更新命令：' + literal('python C-00-Shared/feature_index.py --write') + '；'
              '只读检查：' + literal('python C-00-Shared/feature_index.py') + '；验证入口另记录于各功能的 tests。', '',
              '构建与签名入口不在本功能索引内：项目统一发布者与共享 PFX 见 ' + literal('C-00-Shared/signing/signing.json') + '，'
              '配置校验见 ' + literal('C-00-Shared/signing_config.py') + '、自动导入见 ' + literal('C-00-Shared/resign.py') + '；'
              '两产品配置器的 ' + literal('SIGNING_CERT_THUMBPRINT') + ' 固定同一指纹。', '']
    for number, (directory, _, title) in enumerate(PRODUCTS, 1):
        root = Path(repository) / directory
        if project_root is not None and root.resolve() != Path(project_root).resolve():
            continue
        if not root.is_dir():
            continue
        leaves = verify_product(root)
        project = project_data(root)
        features = project['features']
        output += ['# ' + ('一' if number == 1 else '二') + '. ' + title + ' 功能代码索引', '',
                   '项目根目录：' + literal(directory + '/') + '。当前收录 ' + str(len(leaves)) + ' 个末级入口，含禁用占位项。', '',
                   '工程底座：' + '、'.join(link(root, ref) for ref in
                     ('src/addin/ThisAddIn.cs', 'src/addin/Connect.cs', '@shared/src/addin/UiKit.cs', '@shared/src/addin/HostDialogService.cs')) + '。', '']
        tabs = list(dict.fromkeys(leaf['labels'][0] for leaf in leaves))
        for index, tab in enumerate(tabs, 1):
            output += ['## ' + str(index) + '. ' + tab + (' 主选项卡' if index == 1 else ' 选项卡'), '',
                       '### (1). 功能目录树', '', '~~~text', tree(root, leaves, features, tab), '~~~', '',
                       '### (2). 末级功能源码定位表', '',
                       '| 中文功能路径 | 稳定控件 ID | 功能目录（相对仓库根） | 入口文件与回调 | 核心实现与共用依赖（阅读顺序） |',
                       '| --- | --- | --- | --- | --- |']
            for leaf in leaves:
                if leaf['labels'][0] != tab:
                    continue
                bid = leaf['id']
                feature = features[bid]
                label = ' → '.join(leaf['labels'][1:])
                if leaf.get('enabled') is False:
                    label += '（禁用占位，尚未开放）'
                directory_path = directory + '/' + feature['directory']
                cell = '[' + directory_path + '/](<../' + directory_path + '/>)'
                entry = link(root, feature['entry']) + '；' + literal(callback(leaf) + '()')
                dependencies = ' → '.join(link(root, ref, Path(ref).name + ('（共享）' if ref.startswith('@shared/') else ''))
                                          for ref in feature.get('implementation', [])) or '逻辑直接位于入口文件。'
                output.append('| ' + ' | '.join((label, literal(bid), cell, entry, dependencies)) + ' |')
            output.append('')
    return '\n'.join(output)


def section(text, project_root):
    title = next(title for directory, _, title in PRODUCTS if directory == Path(project_root).name)
    match = re.search(r'(?m)^# [一二]\. ' + re.escape(title) + r' 功能代码索引\s*$', text)
    if match is None:
        raise ValueError('R-30 B4 缺少产品章节：' + title)
    end = re.search(r'(?m)^# ', text[match.end():])
    return text[match.start():match.end() + end.start() if end else len(text)].rstrip()


def check(repository, project_root=None):
    repository = Path(repository).resolve()
    if project_root is not None:
        project_root = Path(project_root).resolve()
        if project_root.parent != repository or project_root.name not in INFRASTRUCTURE or not project_root.is_dir():
            raise ValueError('R-30 --project-root 必须指向本仓库的实际加载项目录')
    expected = render(repository, project_root)
    target = repository / 'C-04-Devlog' / INDEX_NAME
    actual = target.read_text(encoding='utf-8-sig') if target.is_file() else ''
    if project_root is not None:
        actual, expected = section(actual, project_root), section(expected, project_root)
    if actual != expected:
        raise ValueError('R-30 B4 代码索引与实际功能映射不一致，请运行 feature_index.py --write')
    print('PASS R-30: menu IDs, source ownership, nearest shared folders and B4' +
          ('; ' + Path(project_root).name if project_root else '; both products'), flush=True)
    selected = [project_root] if project_root is not None else [
        repository / directory for directory, _, _ in PRODUCTS if (repository / directory).is_dir()]
    return {'rule': 'R-30', 'status': 'pass', 'cached': False,
            'products': [{'project': root.name, 'leaf_count': len(project_data(root)['features'])}
                         for root in selected]}


def self_test(repository):
    count = 0
    for directory, _, _ in PRODUCTS:
        root = Path(repository) / directory
        if not root.is_dir():
            continue
        project = project_data(root)
        verify_product(root, project)
        count += 1
        bid = next(iter(project['features']))
        invalid = []
        missing = copy.deepcopy(project); del missing['features'][bid]; invalid.append(missing)
        wrong = copy.deepcopy(project); wrong['features'][bid]['entry'] = 'src/addin/Connect.cs'; invalid.append(wrong)
        missing_dependency = copy.deepcopy(project); missing_dependency['features'][bid]['implementation'].append('missing.cs'); invalid.append(missing_dependency)
        duplicate = copy.deepcopy(project); duplicate['sources'].append(duplicate['sources'][0]); invalid.append(duplicate)
        alias = copy.deepcopy(project); alias['source_aliases']['bad.cs'] = ['missing.cs']; invalid.append(alias)
        depth = copy.deepcopy(project); depth['features'][bid]['directory'] += '/Unexpected'; invalid.append(depth)
        for candidate in invalid:
            try:
                verify_product(root, candidate)
            except (ValueError, OSError):
                count += 1
            else:
                raise AssertionError('Invalid feature source mapping accepted')
    check(repository)
    target = Path(repository) / 'C-04-Devlog' / INDEX_NAME
    original_read = Path.read_text
    def stale_index(path, *args, **kwargs):
        text = original_read(path, *args, **kwargs)
        return text + '\n索引已过期。\n' if path == target else text
    with patch.object(Path, 'read_text', stale_index):
        try:
            check(repository)
        except ValueError:
            count += 1
        else:
            raise AssertionError('Stale B4 content accepted')
    from feature_gate_tests import run as run_fixtures
    count += run_fixtures(repository)
    print('FeatureIndex self-test: %d PASS, 0 FAIL' % count)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write', action='store_true', help='Regenerate B4 explicitly; default is read-only')
    parser.add_argument('--self-test', action='store_true')
    parser.add_argument('--project-root', type=Path, help='Read-only gate for one product and its B4 chapter')
    args = parser.parse_args()
    root = Path(__file__).resolve().parent.parent
    if args.write:
        if args.project_root or not all((root / directory).is_dir() for directory, _, _ in PRODUCTS):
            raise ValueError('--write 需要两个产品完整源码；--project-root 仅用于只读核验')
        (root / 'C-04-Devlog' / INDEX_NAME).write_text(render(root), encoding='utf-8', newline='\n')
    elif args.self_test:
        self_test(root)
    else:
        check(root, args.project_root)


if __name__ == '__main__':
    main()
