"""Read-only configuration tests; never generate or modify user configuration."""
from pathlib import Path
import contextlib
import copy
import io
import itertools
import sys
import xml.etree.ElementTree as ET

sys.dont_write_bytecode = True
sys.path.insert(0, str(PROJECT_ROOT))
from project_support import load_configure


def main():
    core = load_configure()._core
    original, original_connect, original_templates = core.BUTTONS, core.CONNECT_CS, core.RESULT_TEMPLATES
    original_extras, original_onload = getattr(core, "EXTRA_TABS", []), getattr(core, "RIBBON_ON_LOAD", None)
    original_defaults = getattr(core, "ARITHMETIC_DEFAULTS", None)
    passed = 0

    def check(condition, message):
        nonlocal passed
        if not condition:
            raise AssertionError(message)
        passed += 1

    def xml_shape(node):
        return node.tag, list(node.attrib.items()), [xml_shape(child) for child in node]

    def serialize_xml(groups, tab):
        # Use the production C# literal and text-writer settings, without files.
        memory = io.BytesIO()
        writer = io.TextIOWrapper(memory, encoding='utf-8', newline='\r\n')
        try:
            writer.write(core.cs_verbatim(core.build_xml(groups, tab)))
            writer.flush()
            source = memory.getvalue()
        finally:
            writer.close()
        check(b'\r\r\n' not in source, 'serialized Ribbon has exactly one CR before LF')
        literal = source.decode('utf-8')
        return ET.fromstring(literal[2:-1].replace('""', '"'))

    groups, model, tab, _ = core.build_model()
    document = ET.fromstring(core.build_xml(groups, tab))
    check(xml_shape(serialize_xml(groups, tab)) == xml_shape(document),
          'actual serialization preserves all controls, ordered attributes and children')
    check(len(model) == len(document.findall('.//{*}button')) + len(document.findall('.//{*}menu')) + len(document.findall('.//{*}toggleButton')), 'actual control count')
    check(len(document.findall('.//{*}menu')) == sum(1 for item in model if item['type']=='menu' and not item['hidden']), 'actual native menus')
    boxes = document.findall('.//{*}box')
    check(not boxes, 'current small controls are direct group children')
    check([group[1] for group in groups] == (['数据','操作','反馈'] if any(item['id']=='menuOneClickFormula' for item in model) else ['链接','操作','反馈']) + ['扩展','详情','操作'], 'primary and Agent-Evo detail/operation groups')
    # Product-level PDF navigation contract, independent of display names/customization.
    file_menu = document.find('.//{*}menu[@id="menuFileOperations"]')
    check(file_menu is not None, 'file operations entrance exists')
    pdf_menu = file_menu.find('{*}menu[@id="menuFileConversion"]')
    check(pdf_menu is not None and [node.get('id') for node in pdf_menu] ==
          ['btnCurrentPdf', 'btnWatermarkPdf'], 'two PDF entries in conversion subtree')
    batch_menu = file_menu.find('{*}menu[@id="menuBatchProcessing"]')
    if any(item['id'] == 'menuOneClickFormula' for item in model):
        check(batch_menu is not None and [node.get('id') for node in batch_menu] ==
              ['btnOrganizeFiles', 'btnRenameFiles', 'btnFileNames'], 'original three batch command identities preserved')
        check([node.get('id') for node in file_menu] == ['menuBatchProcessing', 'menuFileConversion'], 'Excel two-subtree navigation')
    else:
        check(batch_menu is None and [node.get('id') for node in file_menu] == ['menuFileConversion'], 'Word contains PDF tools without Excel batch features')
    # 首个分组前三个是小图标，之后新增的根层项用大图标；数量随功能增减，不锁死条目个数。
    sizes = [item['size'] for item in groups[0][2]]
    check(sizes[:3] == ['normal']*3 and all(size == 'large' for size in sizes[3:]), 'native small rows then large roots')
    check(all(item['size']=='large' for group in groups[1:] for item in group[2]), 'action and feedback use large roots')
    for group_id, _, items in groups:
        group = next(node for node in document.findall('.//{*}group') if node.get('id') == group_id)
        entrances = [row for node in group for row in (list(node) if node.tag.endswith('}box') else [node])]
        check([node.get('id') for node in entrances] == [item['id'] for item in items if not item['hidden']],
              'configured root order: ' + group_id)
    # Every submenu follows the editable children list, independently of labels or function names.
    for entry in model:
        if entry['type'] == 'menu' and not entry['hidden']:
            menu = next(node for node in document.findall('.//{*}menu') if node.get('id') == entry['id'])
            check([node.get('id') for node in menu] == [child['id'] for child in entry['children'] if not child['hidden']],
                  'configured submenu order: ' + entry['id'])
    try:
        check(core.arithmetic_defaults()['divide'] != '0', 'nonzero default divisor')
        for values in ({'add':1},dict(add=True,subtract=100,multiply=100,divide=10000),dict(add=float('nan'),subtract=100,multiply=100,divide=10000),dict(add=100,subtract=100,multiply=100,divide=0)):
            core.ARITHMETIC_DEFAULTS=values
            rejected=False
            with contextlib.redirect_stderr(io.StringIO()):
                try: core.arithmetic_defaults()
                except SystemExit: rejected=True
            check(rejected,'invalid arithmetic defaults')
        core.ARITHMETIC_DEFAULTS=dict(add=1.5,subtract=-2,multiply=10,divide=0.5)
        check(core.arithmetic_defaults()==dict(add='1.5',subtract='-2',multiply='10',divide='0.5'),'custom default numbers normalized')
        if original_defaults is None: del core.ARITHMETIC_DEFAULTS
        else: core.ARITHMETIC_DEFAULTS=original_defaults
        # Synthetic ids need no real public wrappers: production wrappers were checked above.
        core.EXTRA_TABS = []; core.RIBBON_ON_LOAD = None
        core.RESULT_TEMPLATES = {}
        core.CONNECT_CS = str(Path(__file__).with_name('absent-probe-connect.cs'))
        leaf = dict(id='leaf', label='item & <text>', action='feature', callback='Fake', builtin_icon='formula',
                    tip_width=1000, screentip='title & <text>\r\nsecond "line"',
                    supertip='first & <one>\r\nsecond "two"\r\n\r\nfourth')
        tree = [dict(type='menu', id='rootMenu', group=('testGroup', 'test'), label='menu', size='large',
                     children=[dict(type='menu', id='subMenu', label='sub', children=[leaf]),
                               dict(id='urlLeaf', label='web', action='url', url='https://example.com')])]

        def generate(items):
            core.BUTTONS = items
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                groups, model, tab, _ = core.build_model()
                return ET.fromstring(core.build_xml(groups, tab)), model

        doc, model = generate(copy.deepcopy(tree))
        groups, _, tab, _ = core.build_model()
        serialized = serialize_xml(groups, tab)
        check(xml_shape(serialized) == xml_shape(doc), 'multiline tooltip serialization preserves XML semantics')
        serialized_leaf = next(node for node in serialized.findall('.//{*}button') if node.get('id') == 'leaf')
        check(serialized_leaf.get('screentip') == leaf['screentip'], 'serialized screentip preserves CRLF and escapes')
        check(serialized_leaf.get('supertip') == leaf['supertip'], 'serialized supertip preserves blank lines and escapes')
        check([m['px'] for m in model] == [32, 16, 16, 16], 'nested logical sizes')
        check([m['type'] for m in model] == ['menu', 'menu', 'button', 'button'], 'backward-compatible default button')
        root = doc.find('.//{*}menu')
        check(root.get('size') == 'large' and root.get('onAction') is None, 'root menu no action')
        check(all(c.get('size') is None for c in root.iter() if c is not root), 'unsized nested controls')
        check(doc.find('.//{*}button').get('label') == 'item & <text>', 'XML escapes label')
        check(doc.findall('.//{*}button')[-1].get('onAction') == 'OnOpenUrl_urlLeaf', 'nested URL wrapper')
        ordered = [dict(type='menu', id='orderedRoot', group=('testGroup', 'test'), label='root', size='large',
                        children=[dict(type='menu', id='firstFunction', label='first', children=[dict(leaf, id='firstLeaf')]),
                                  dict(type='menu', id='secondFunction', label='second', children=[dict(leaf, id='secondLeaf')]),
                                  dict(leaf, id='thirdFunction')])]
        for permutation in itertools.permutations(ordered[0]['children']):
            configured = copy.deepcopy(ordered); configured[0]['children'] = list(permutation)
            doc, _ = generate(configured)
            menu = doc.find('.//{*}menu')
            check([node.get('id') for node in menu] == [child['id'] for child in permutation], 'custom sibling order preserved')
            check(all(node.get('onAction') == 'Fake' for node in doc.findall('.//{*}button')),
                  'reordering preserves leaf callbacks')
        # Hide all children of one submenu; the empty submenu disappears rather than emitting invalid XML.
        hidden = copy.deepcopy(tree); hidden[0]['children'][0]['children'][0]['hidden'] = True
        doc, _ = generate(hidden)
        check(len(doc.findall('.//{*}menu')) == 1, 'hidden empty submenu pruned')
        column = [dict(id='quickA', group=('testGroup', 'test'), label='A', column='quickColumn', size='normal', action='feature', callback='Fake'),
                  dict(id='quickB', group=('testGroup', 'test'), label='B', column='quickColumn', size='normal', action='feature', callback='Fake'),
                  dict(id='big', group=('testGroup', 'test'), label='large', size='large', action='feature', callback='Fake')]
        doc, model = generate(copy.deepcopy(column))
        box = doc.find('.//{*}box')
        check(box.get('boxStyle') == 'vertical' and [b.get('id') for b in box] == ['quickA', 'quickB'], 'ordered vertical column')
        check([m['px'] for m in model] == [16, 16, 32], 'mixed icon sizes')
        hidden = copy.deepcopy(column); hidden[0]['hidden'] = True
        doc, _ = generate(hidden)
        check([b.get('id') for b in doc.find('.//{*}box')] == ['quickB'], 'hidden row removed')
        hidden[1]['hidden'] = True
        doc, _ = generate(hidden)
        check(not doc.findall('.//{*}box'), 'empty column removed')
        mixed = copy.deepcopy(column[:2])
        mixed.append(dict(type='menu', id='quickMenu', group=('testGroup', 'test'), label='menu', column='quickColumn',
                          size='normal', children=[copy.deepcopy(leaf)]))
        doc, model = generate(mixed)
        check([node.get('id') for node in doc.find('.//{*}box')] == ['quickA', 'quickB', 'quickMenu'],
              'normal root menu fills third row')
        check(all(item['px'] == 16 for item in model) and doc.find('.//{*}menu').get('onAction') is None,
              'column menu size and no action')
        mixed[2]['children'][0]['hidden'] = True
        doc, _ = generate(mixed)
        check([node.get('id') for node in doc.find('.//{*}box')] == ['quickA', 'quickB'], 'hidden column menu pruned')
        for case in ('invalid', 'large', 'cross_group', 'noncontiguous', 'control_id', 'group_id', 'tab_id', 'large_menu', 'nested'):
            bad = copy.deepcopy(column)
            if case == 'invalid': bad[0]['column'] = 'bad-id'
            if case == 'large': bad[0]['size'] = 'large'
            if case == 'cross_group': bad[1]['group'] = ('secondGroup', 'second')
            if case == 'noncontiguous': bad = [bad[0], bad[2], bad[1]]
            if case == 'control_id': bad[0]['column'] = 'quickB'
            if case == 'group_id': bad[0]['column'] = 'testGroup'
            if case == 'tab_id': bad[0]['column'] = core.TAB_ID
            if case == 'large_menu': bad[0].update(type='menu', size='large', action=None, callback=None, children=[copy.deepcopy(leaf)])
            if case == 'nested':
                bad = copy.deepcopy(tree); bad[0]['children'][0]['children'][0]['column'] = 'nestedColumn'
            rejected = False
            try: generate(bad)
            except SystemExit as error: rejected = error.code == 2
            check(rejected, 'column configuration rejected: ' + case)
        for case in ('size', 'group', 'action', 'duplicate', 'empty', 'bad_type', 'leaf_children', 'global_id'):
            bad = copy.deepcopy(tree)
            child = bad[0]['children'][0]
            if case == 'size': child['size'] = 'normal'
            if case == 'group': child['group'] = ('other', 'other')
            if case == 'action': child['callback'] = 'Fake'
            if case == 'duplicate': child['id'] = 'rootMenu'
            if case == 'empty': child['children'] = []
            if case == 'bad_type': child['type'] = 'popup'
            if case == 'leaf_children': child['children'][0]['children'] = [copy.deepcopy(leaf)]
            if case == 'global_id': bad[0]['id'] = 'testGroup'
            rejected = False
            try:
                generate(bad)
            except SystemExit as error:
                rejected = error.code == 2
            check(rejected, 'configuration rejected: ' + case)
        # Optional extensions preserve the legacy shape when absent.
        doc, _ = generate(copy.deepcopy(column))
        check(len(doc.findall('.//{*}tab')) == 1 and not doc.get('onLoad'), 'legacy configuration unchanged')
        core.EXTRA_TABS = [dict(id='extraTab', label='Agent', get_visible='GetVisible', buttons=[
            dict(id='extraButton', group=('extraGroup','resources'), label='extra', size='large', action='feature', callback='Fake')])]
        core.RIBBON_ON_LOAD = 'Loaded'
        doc, _ = generate(copy.deepcopy(column))
        check(len(doc.findall('.//{*}tab')) == 2 and doc.get('onLoad') == 'Loaded', 'optional multiple tabs')
        check(doc.findall('.//{*}tab')[1].get('getVisible') == 'GetVisible', 'tab visibility callback')
        toggle = dict(id='switch', group=('testGroup','test'), type='toggleButton', label='toggle', size='large', action='feature', callback='Toggle', get_pressed='GetPressed')
        doc, _ = generate([toggle])
        check(doc.find('.//{*}toggleButton').get('getPressed') == 'GetPressed', 'toggle native XML')
        for case in ('tab', 'group', 'control', 'container', 'duplicate_tab', 'bad_toggle', 'url_toggle', 'nested_toggle'):
            core.EXTRA_TABS = copy.deepcopy([dict(id='extraTab', label='Agent', buttons=[dict(id='extraButton', group=('extraGroup','resources'), label='extra', action='feature', callback='Fake')])])
            items = copy.deepcopy(column)
            if case == 'tab': core.EXTRA_TABS[0]['id'] = core.TAB_ID
            if case == 'group': core.EXTRA_TABS[0]['buttons'][0]['group'] = ('testGroup','test')
            if case == 'control': core.EXTRA_TABS[0]['buttons'][0]['id'] = 'quickA'
            if case == 'container': core.EXTRA_TABS[0]['buttons'][0]['id'] = 'quickColumn'
            if case == 'duplicate_tab': core.EXTRA_TABS.append(copy.deepcopy(core.EXTRA_TABS[0]))
            if case == 'bad_toggle': items = [dict(toggle, get_pressed=None)]
            if case == 'url_toggle': items = [dict(toggle, action='url', url='https://example.com')]
            if case == 'nested_toggle': items = copy.deepcopy(tree); items[0]['children'] = [dict(toggle)]; items[0]['children'][0].pop('group'); items[0]['children'][0].pop('size')
            rejected = False
            try: generate(items)
            except SystemExit: rejected = True
            check(rejected, 'cross-tab or toggle conflict: ' + case)
        source = 'public void Load(object ui) {} public bool Visible(object c) {} public void Toggle(object c, bool pressed) {}'
        for name, result, pressed in [('Load','void',False), ('Visible','bool',False), ('Toggle','void',True)]:
            core.require_callback(source,name,result,pressed); check(True, 'callback signature ' + name)
        for name, result, pressed in [('Load','bool',False), ('Visible','void',False), ('Toggle','void',False), ('missing','void',True)]:
            rejected = False
            with contextlib.redirect_stderr(io.StringIO()):
                try: core.require_callback(source,name,result,pressed)
                except SystemExit: rejected = True
            check(rejected, 'invalid callback signature ' + name)
    finally:
        if original_defaults is None: core.__dict__.pop("ARITHMETIC_DEFAULTS",None)
        else: core.ARITHMETIC_DEFAULTS=original_defaults
        core.EXTRA_TABS, core.RIBBON_ON_LOAD = original_extras, original_onload
        core.BUTTONS, core.CONNECT_CS, core.RESULT_TEMPLATES = original, original_connect, original_templates
    print('ConfigureProbe: %d PASS, 0 FAIL' % passed)


if __name__ == '__main__':
    main()
