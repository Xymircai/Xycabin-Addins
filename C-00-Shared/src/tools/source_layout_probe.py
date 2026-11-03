"""Exercise shared source confinement and both products' cache dependencies."""
from pathlib import Path
import argparse
import sys
from tempfile import TemporaryDirectory
from unittest.mock import patch

sys.dont_write_bytecode = True
COMMON = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(COMMON))
from source_layout import resolve_source, addin_sources


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
