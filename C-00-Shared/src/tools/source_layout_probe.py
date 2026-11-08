"""Exercise shared source confinement and both products' cache dependencies."""
from pathlib import Path
import argparse
import json
import re
import sys
from tempfile import TemporaryDirectory
from unittest.mock import patch

sys.dont_write_bytecode = True
COMMON = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(COMMON))
from source_layout import resolve_source, addin_sources
from tool_catalog import tools, validation


def verify(project_root):
    root = Path(project_root).resolve()
    repository = root.parent
    checks = 0

    def check(condition, label):
        nonlocal checks
        if not condition:
            raise AssertionError(label)
        checks += 1

    sources = addin_sources(root)
    check(len(sources) == len(set(sources)), 'a product cannot compile the same implementation twice')
    check(all(path.is_file() for path in sources), 'all compiled source references exist')
    verify_architecture(root, sources, check)
    verify_tool_catalog(root, check)
    for reference in ('@shared/../verify_shared.py', '../AGENTS.md', '@shared/missing-source.cs'):
        try:
            resolve_source(root, reference)
        except ValueError:
            checks += 1
        else:
            raise AssertionError('invalid reference accepted: ' + reference)
    site = repository / 'C-03-Website·Xycabin'
    sys.path.insert(0, str(site / 'src'))
    import build_cache
    from site_core import build
    # A source-only checkout has no published HTML yet. Probe fresh inputs
    # without reading or replacing the previous official delivery.
    with TemporaryDirectory(prefix='xy-source-layout-') as temporary:
        snapshot = build(site, target=Path(temporary) / 'snapshot.html')
        verify_cache_dependencies(root, build_cache, snapshot, check)
    print('SourceLayoutProbe: %d PASS, 0 FAIL' % checks)


def verify_architecture(root, sources, check):
    by_name = {path.name: path for path in sources}
    names = ('MigrationPlanner.cs', 'MigrationConfiguration.cs', 'MigrationNameIndex.cs') if root.name == 'C-01-ExcelIterate' else ('LinkFieldCode.cs',)
    forbidden = r'Windows\.Forms|\bdynamic\b|Marshal\.|\b(?:File|Directory)\.|DataMigrationForm|DataMigrationHost|LinkPasteEngine|UpdateEngine'
    for name in names:
        text = by_name[name].read_text(encoding='utf-8-sig')
        check(not re.search(forbidden, text), name + ' has no UI, COM, file access or business-engine dependency')
    if root.name == 'C-01-ExcelIterate':
        form = by_name['DataMigrationForm.cs'].read_text(encoding='utf-8-sig')
        for method in ('TriggerSummary', 'TriggerTag', 'MatchedFiles', 'ShowPairDetails'):
            body = re.search(r'^        private [^\n]+\b' + method + r'\([^\n]+\n.*?^        }', form, re.M | re.S)
            check(body is not None and not re.search(r'RefreshTrigger\(|TargetIndex\(|MigrationPlanner.UnusedTargets\(', body[0]),
                  method + ' consumes prepared matches instead of querying the target pool')
    if root.name == 'C-02-WordUpdater':
        update = by_name['UpdateEngine.cs'].read_text(encoding='utf-8-sig')
        check('static object GetItem(' not in update, 'COM collection access belongs to the Word adapter')
        for path in sources:
            check('UpdateEngine.GetItem(' not in path.read_text(encoding='utf-8-sig'), 'no business borrows update-engine adapters')


def verify_tool_catalog(root, check):
    project = json.loads((root / 'project.json').read_text(encoding='utf-8-sig'))
    catalog = tools(project)
    check(all(resolve_source(root, tool['source']).is_file() for tool in catalog), 'all tool sources resolve')
    context = dict(dll='addin.dll', manifest='addin.manifest', vsto='addin.vsto', site_html='site.html', installer='setup.exe', rendered='preview')
    background, pending, cacheable = validation(project, True, False, context)
    common = ['AgentDocumentCoreProbe', 'GripProbe', 'LiveProbeSafetyFixture', 'InstallFlowProbe', 'RibbonProbe', 'ValidateManifest', 'DialogSizesProbe']
    expected = [tool['name'] for tool in project['tools'] if tool.get('offline')]
    check([name for name, _ in background] == expected + common, 'all previous mandatory background tools retain their order')
    check(len(background) + 8 == (32 if project['host'] == 'Excel' else 25), 'background regression count is unchanged')
    check(cacheable == set(expected), 'only previously cacheable product probes use validation receipts')
    check(all(not tool.get('background') and 'background' not in tool for tool in project['tools'] if not tool.get('offline')), 'live tools cannot enter background validation')
    arguments = dict(background)
    check(arguments['DialogSizesProbe'] == ['addin.dll', '--core'], 'background dialog check never creates forms')
    if project['host'] == 'Excel':
        check(arguments['FormulaProbe'] == ['--core'] and 'FormulaProbe 前台窗口路径' in pending, 'formula core and foreground paths stay distinct')
    foreground, _, foreground_cache = validation(project, False, True, context)
    check([name for name, _ in foreground] == expected + common + ['AgentExtensionProbe', 'UiPreviewProbe'], 'foreground sequence retains UI and Agent checks')
    check(dict(foreground)['UiPreviewProbe'] == ['addin.dll', 'setup.exe', 'preview'], 'UI preview consumes catalog arguments')
    check(cacheable == foreground_cache, 'execution mode does not redefine the cache boundary')
    check(pending[-3:] == ['DialogSizesProbe 前台预览/拖动', 'AgentExtensionProbe', 'UiPreviewProbe'], 'background omissions remain explicit')


def verify_cache_dependencies(root, build_cache, snapshot, check):
    repository = root.parent
    products = [repository / name for name in ('C-01-ExcelIterate', 'C-02-WordUpdater')]
    existing = [product for product in products if product.is_dir()]
    before = [build_cache.inputs(product, snapshot) for product in existing]
    digest = build_cache.digest
    shared_source = resolve_source(root, '@shared/src/addin/UiKit.cs')

    def changed_digest(path):
        return 'changed-shared-source' if Path(path).resolve() == shared_source else digest(path)

    # Inject a hash change instead of editing source during a build or test.
    with patch.object(build_cache, 'digest', side_effect=changed_digest):
        after = [build_cache.inputs(product, snapshot) for product in existing]
    for old, new in zip(before, after):
        check(old != new, 'shared source change invalidates each product cache')
    check(before == [build_cache.inputs(product, snapshot) for product in existing],
          'unchanged inputs retain stable cache fingerprints')
    for selected in existing:
        feature = next(path for path in addin_sources(selected)
                       if path.is_relative_to(selected / 'AC-Original Code'))
        def changed_feature(path):
            return 'changed-feature-source' if Path(path).resolve() == feature else digest(path)
        with patch.object(build_cache, 'digest', side_effect=changed_feature):
            modified = [build_cache.inputs(product, snapshot) for product in existing]
        check(all((old != new) == (product == selected)
                  for product, old, new in zip(existing, before, modified)),
              'a product feature edit invalidates only that product cache')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project-root', type=Path, required=True)
    verify(parser.parse_args().project_root)
