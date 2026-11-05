"""R-30 regression fixtures; all source mutations stay in a disposable checkout."""
import copy
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from unittest.mock import patch

import feature_index as index
import source_layout


def run(repository):
    repository = Path(repository)
    checks = 0
    original_sys_path = list(sys.path)

    def passed(label):
        nonlocal checks
        checks += 1
        print('PASS R-30 fixture: ' + label, flush=True)

    def reject(root, project, message):
        try:
            index.verify_product(root, project)
        except ValueError as error:
            if 'R-30' not in str(error) or message not in str(error):
                raise AssertionError('Wrong rejection: ' + str(error)) from error
            passed(message)
        else:
            raise AssertionError('Invalid classification accepted: ' + message)

    try:
        with tempfile.TemporaryDirectory(prefix='r30-中文 空格-') as temporary:
            fixture = Path(temporary)
            excluded = {'build', 'lib', 'AB-Graphic', '__pycache__', '.workbuddy', '.ai-memory'}
            suffixes = {'.py', '.cs', '.json', '.md', '.bat', '.html', '.js', '.txt'}
            for directory in ['C-00-Shared', 'C-01-ExcelIterate', 'C-02-WordUpdater',
                              'C-03-Website·Xycabin', 'C-04-Devlog']:
                for current, directories, filenames in os.walk(repository / directory):
                    directories[:] = [name for name in directories if name not in excluded]
                    for filename in filenames:
                        source = Path(current) / filename
                        if source.suffix not in suffixes or filename in ('XyCabin.html', 'index.html'):
                            continue
                        target = fixture / source.relative_to(repository)
                        target.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copyfile(source, target)
            for name in ('verify_shared.py', 'verify_devlog.py', 'AGENTS.md'):
                shutil.copyfile(repository / name, fixture / name)
            with patch.object(source_layout, 'SHARED_ROOT', fixture / 'C-00-Shared'):
                for directory, _, _ in index.PRODUCTS:
                    root = fixture / directory
                    project = source_layout.project_data(root)
                    index.verify_product(root)
                    passed('valid product and linked common sources: ' + directory)
                    _, leaves = index.commands(root)
                    first, second = leaves[:2]
                    bid = first['id']
                    other = second['id']
                    wrong = copy.deepcopy(project)
                    wrong['features'][bid]['directory'] = wrong['features'][other]['directory']
                    reject(root, wrong, '目录与菜单 ID 路径不一致')
                    missing = copy.deepcopy(project)
                    del missing['code_layout'][first['identifiers'][1]]
                    reject(root, missing, '节点英文目录映射缺失')
                    misplaced = copy.deepcopy(project)
                    misplaced['sources'].append('src/addin/R30Misplaced.cs')
                    bad_source = root / 'src/addin/R30Misplaced.cs'
                    bad_source.write_text('class R30Misplaced {}', encoding='utf-8')
                    try:
                        reject(root, misplaced, '基础设施白名单外')
                    finally:
                        bad_source.unlink()
                    orphan = root / project['features'][bid]['directory'] / 'R30Orphan.cs'
                    orphan.write_text('class R30Orphan {}', encoding='utf-8')
                    try:
                        reject(root, copy.deepcopy(project), '实际编译清单不一致')
                        unowned = copy.deepcopy(project)
                        unowned['sources'].append(orphan.relative_to(root).as_posix())
                        reject(root, unowned, '未归入任何功能映射')
                    finally:
                        orphan.unlink()
                    duplicate = root / project['features'][bid]['directory'] / 'R30Duplicate.cs'
                    duplicate.write_text('public partial class Connect { public void ' +
                                         index.callback(first) + '(object control) {} }', encoding='utf-8')
                    duplicated = copy.deepcopy(project)
                    duplicated['sources'].append(duplicate.relative_to(root).as_posix())
                    duplicated['features'][bid]['implementation'].append(duplicate.relative_to(root).as_posix())
                    try:
                        reject(root, duplicated, '按钮回调必须唯一')
                    finally:
                        duplicate.unlink()
                    common = os.path.commonpath([project['features'][bid]['directory'],
                                                 project['features'][other]['directory']]).replace('\\', '/')
                    shared = root / common / '_Shared/R30SharedFixture.cs'
                    shared.parent.mkdir(parents=True, exist_ok=True)
                    shared.write_text('internal static class R30SharedFixture { internal static int Value() { return 1; } }', encoding='utf-8')
                    combined = copy.deepcopy(project)
                    ref = shared.relative_to(root).as_posix()
                    combined['sources'].append(ref)
                    for owner in (bid, other):
                        combined['features'][owner]['implementation'].append(ref)
                    try:
                        index.verify_product(root, combined)
                        passed('nearest common folder for two consumers: ' + directory)
                        one_consumer = copy.deepcopy(combined)
                        one_consumer['features'][other]['implementation'].remove(ref)
                        reject(root, one_consumer, '最近公共父级')
                    finally:
                        shared.unlink()
                    configuration = root / 'AA-Program' / next(name for folder, name, _ in index.PRODUCTS if folder == directory)
                    original_config = configuration.read_text(encoding='utf-8-sig')
                    configuration.write_text(original_config + '\nBUTTONS.reverse()\nBUTTONS[0]["label"] += "（更名样例）"\n', encoding='utf-8')
                    try:
                        index.verify_product(root)
                        passed('labels and sibling order retain directory identities: ' + directory)
                    finally:
                        configuration.write_text(original_config, encoding='utf-8')
                    # A genuinely new UI leaf needs a node mapping, compiled entry and owner.
                    new_id = 'btnR30Fixture'
                    # Choose a root button so the new leaf has tab/group/leaf depth.
                    new_directory = '/'.join([index.FEATURE_ROOT,
                        project['code_layout'][first['identifiers'][0]],
                        project['code_layout'][first['identifiers'][1]], 'R30Fixture'])
                    entry_ref = new_directory + '/Connect.R30Fixture.Entry.cs'
                    entry = root / entry_ref
                    entry.parent.mkdir(parents=True, exist_ok=True)
                    entry.write_text('public partial class Connect { public void OnR30Fixture(object control) {} '
                                     'public IPictureDisp GetImage_btnR30Fixture(object control) { return null; } }', encoding='utf-8')
                    addition = {'id': new_id, 'label': '门禁样例', 'group': (first['identifiers'][1], first['labels'][1]),
                                'action': 'feature', 'callback': 'OnR30Fixture', 'builtin_icon': 'table'}
                    configuration.write_text(original_config + '\nBUTTONS.append(' + repr(addition) + ')\n', encoding='utf-8')
                    fresh = copy.deepcopy(project)
                    fresh['code_layout'][new_id] = 'R30Fixture'
                    fresh['sources'].append(entry_ref)
                    fresh['features'][new_id] = {'directory': new_directory, 'entry': entry_ref, 'implementation': [], 'tests': []}
                    try:
                        index.verify_product(root, fresh)
                        passed('new menu leaf with complete source registration: ' + directory)
                    finally:
                        configuration.write_text(original_config, encoding='utf-8')
                        entry.unlink()
                        entry.parent.rmdir()
                # Test independent checkout validation against the full portable B4 document.
                word = fixture / 'C-02-WordUpdater'
                hidden = fixture / 'unused-word-checkout'
                assert word.resolve().is_relative_to(fixture) and hidden.resolve().is_relative_to(fixture)
                word.rename(hidden)
                try:
                    index.check(fixture, fixture / 'C-01-ExcelIterate')
                    passed('one-product checkout validates only its B4 chapter')
                finally:
                    hidden.rename(word)
            # Real process boundaries prove all entry points reject before compiling.
            for directory, _, _ in index.PRODUCTS:
                config = fixture / directory / 'project.json'
                data = json.loads(config.read_text(encoding='utf-8-sig'))
                first = next(iter(data['features'].values()))
                first['directory'] += '/InvalidDepth'
                config.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
            environment = dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8', PYTHONDONTWRITEBYTECODE='1')
            for directory in ('C-01-ExcelIterate', 'C-02-WordUpdater', 'C-03-Website·Xycabin'):
                entries = ('rebuild.py', 'build.py', 'validate.py') if directory != 'C-03-Website·Xycabin' else ('rebuild.py', 'validate.py')
                for entry in entries:
                    for mode in ('full', 'incremental'):
                        args = [sys.executable, '-B', str(fixture / directory / entry), '--mode', mode]
                        if entry != 'build.py':
                            args.append('--background')
                        result = subprocess.run(args, cwd=fixture, env=environment, capture_output=True,
                                                text=True, encoding='utf-8', timeout=45)
                        output = result.stdout + result.stderr
                        if result.returncode == 0 or 'R-30' not in output:
                            raise AssertionError(directory + '/' + entry + ' did not fail at R-30:\n' + output[-2500:])
                        passed(directory + '/' + entry + ' rejects ' + mode + ' before build/publish')
            if any(path.suffix.lower() in ('.dll', '.exe') for path in fixture.rglob('*') if path.is_file()):
                raise AssertionError('Invalid fixture produced binaries')
            passed('invalid entry points produced no DLL/EXE')
    finally:
        sys.path[:] = original_sys_path
    return checks
