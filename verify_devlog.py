"""Read-only gate for the six documents in C-04-Devlog.

Structure is derived from level-one chapter titles; no invisible HTML markers
are used. Scopes, overview and delivery-evidence sections are computed by the
same helpers that compose the embedded documents, so gate and build agree.
"""
from pathlib import Path
import ast
import json
import re
import sys

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
DOCS = ROOT / 'C-04-Devlog'
sys.path.insert(0, str(ROOT / 'C-00-Shared'))
from build_report import ordered_errors, profile_errors, BUILD_GATES, DEVELOPER_GATES, GATE_SCOPES, GATE_ROW
FILES = ('B0-Entry.md', 'A-UserGuide.md', 'B1-Project-Files.md',
         'B2-DevExperience.md', 'B3-Spec Coding.md', 'B4-Code Index.md')
LIMIT = 200
AUDIT_HEAD = re.compile(r'^###\s+\(\d+\)\.\s*本轮交付证据\s*$')
EXPECTED_SCOPES = {
    'A-UserGuide.md': ['excel', 'word', 'site'],
    'B1-Project-Files.md': ['shared', 'excel', 'word', 'site'],
    'B2-DevExperience.md': ['shared', 'excel', 'word', 'site'],
    'B3-Spec Coding.md': ['shared', 'site'],
    'B4-Code Index.md': ['excel', 'word'],
}
RULE_LINE = re.compile(r'(?m)^\*\*(R-\d{2})\*\*')
RULE_MARK = re.compile(r'rule=(R-\d{2})')
RULES = {
    'R-01': ('L1', 'verify_shared.py', 'check_documents'),
    'R-02': ('L1', 'C-03-Website·Xycabin/src/tools/layout_probe.py', 'verify/layout'),
    'R-03': ('L1', 'C-00-Shared/src/tools/markdown_gate.py', 'hierarchy'),
    'R-04': ('L1', 'verify_devlog.py', 'check_dates'),
    'R-05': ('L1', 'verify_devlog.py', 'version_rows'),
    'R-06': ('L1', 'verify_devlog.py', 'verify_file/guide'),
    'R-07': ('L1', 'verify_devlog.py', 'verify_file/chapters'),
    'R-08': ('L1', 'C-03-Website·Xycabin/src/document_assets.py', 'chapter_scope'),
    'R-09': ('L1', 'C-03-Website·Xycabin/src/document_assets.py', 'document_overview'),
    'R-10': ('L1', 'C-03-Website·Xycabin/src/document_assets.py', 'strip_delivery_audit'),
    'R-11': ('L1', 'verify_devlog.py', 'check_no_comments'),
    'R-12': ('L1', 'C-03-Website·Xycabin/src/tools/layout_probe.py', 'verify/links'),
    'R-13': ('L1', 'C-03-Website·Xycabin/src/document_assets.py', '_regular_asset'),
    'R-14': ('L1', 'C-03-Website·Xycabin/src/document_assets.py', 'collect'),
    'R-15': ('L2', 'verify_shared.py', 'main/shared'),
    'R-16': ('L2', 'C-00-Shared/src/configuration/configure_core.py', 'get_version'),
    'R-17': ('L2', 'C-00-Shared/rebuild.py', 'main'),
    'R-18': ('L2', 'C-00-Shared/build_installer.py', 'main'),
    'R-19': ('L2', 'C-00-Shared/resign.py', 'main'),
    'R-20': ('L2', 'verify_devlog.py', 'check_gate_inventory'),
    'R-21': ('L2', 'C-03-Website·Xycabin/src/tools/cleanup.py', 'main'),
    'R-22': ('L2', 'C-03-Website·Xycabin/src/tools/site_probe.py', 'main'),
    'R-23': ('L2', 'C-03-Website·Xycabin/src/tools/workflow_probe.py', 'main'),
    'R-24': ('L2', 'C-03-Website·Xycabin/src/tools/PayloadProbe.cs', 'Main'),
    'R-25': ('L2', 'C-01-ExcelIterate/src/testharness/DisciplineProbe.cs', 'Main'),
    'R-26': ('L2', 'C-00-Shared/src/tools/dialog_catalog_probe.py', 'main'),
    'R-27': ('L2', 'C-00-Shared/src/testharness/GridHeadersProbe.cs', 'GridHeaderChecks'),
    'R-28': ('L2', 'C-00-Shared/src/testharness/Feature099Probe.cs', 'OutputFormatChecks'),
    'R-29': ('L2', 'C-00-Shared/src/testharness/LiveProbeSafetyFixture.cs', 'Main'),
    'R-30': ('L1', 'C-00-Shared/feature_index.py', 'verify_product'),
}

H_LINE = re.compile(r'(?m)^\*\*(H-\d{2})\*\*')
H_ITEMS = {
    'H-01': '维护顺序', 'H-02': '构建期禁编辑', 'H-03': '安装与卸载报告', 'H-04': '排障顺序',
    'H-05': 'COM 方法指引', 'H-06': '弹窗真实交互验收', 'H-07': '清理触发', 'H-08': '冗余候选条件',
    'H-09': '五类盘点', 'H-10': 'fixture 与删除边界', 'H-11': '网站经验归属', 'H-12': '固定图样不读桌面',
}


def check_h_items(text):
    """Manual gates are numbered consecutively, in document order, with acceptance evidence."""
    errors = []
    found = H_LINE.findall(text)
    if len(found) != len(set(found)):
        errors.append('B3: 人工门禁条目 H 编号重复')
    missing = sorted(set(H_ITEMS) - set(found))
    extra = sorted(set(found) - set(H_ITEMS))
    if missing:
        errors.append('B3: 缺少人工门禁条目 ' + ', '.join(missing))
    if extra:
        errors.append('B3: 未登记人工门禁条目 ' + ', '.join(extra))
    expected = ['H-%02d' % index for index in range(1, len(H_ITEMS) + 1)]
    if found and found != expected:
        errors.append('B3: 人工门禁条目顺序必须为 H-01…H-%02d' % len(H_ITEMS))
    for line in text.split('\n'):
        if H_LINE.match(line) and '验收：' not in line:
            errors.append('B3: ' + line[:6] + ' 缺少“验收/证据”记录')
    return errors


def markdown_gate():
    for folder in ('C-01-ExcelIterate', 'C-02-WordUpdater', 'C-03-Website·Xycabin'):
        tools = ROOT / folder / 'src/tools'
        if (tools / 'markdown_gate.py').is_file():
            sys.path.insert(0, str(tools))
            import markdown_gate as module
            return module
    raise FileNotFoundError('找不到 markdown_gate.py')


def document_rules():
    source = ROOT / 'C-03-Website·Xycabin/src'
    if (source / 'document_assets.py').is_file():
        if str(source) not in sys.path:
            sys.path.insert(0, str(source))
        import document_assets
        return document_assets
    raise FileNotFoundError('找不到 document_assets.py')


# rule=R-11
def check_no_comments(text, label):
    return [label + ': 文档不得包含 HTML 注释'] if '<!--' in text else []


def check_rules(text):
    """Every rule ID appears once in B3, has an implementation, and no orphan marks."""
    errors = []
    defined = RULE_LINE.findall(text)
    if len(defined) != len(set(defined)):
        errors.append('B3: 规则编号重复')
    if set(defined) != set(RULES):
        missing = sorted(set(RULES) - set(defined))
        extra = sorted(set(defined) - set(RULES))
        if missing:
            errors.append('B3: 缺少规则编号 ' + ', '.join(missing))
        if extra:
            errors.append('B3: 未登记规则编号 ' + ', '.join(extra))
    for rule, (layer, relative, check) in sorted(RULES.items()):
        path = ROOT / relative
        if not path.is_file():
            errors.append(rule + ': 实现文件不存在 ' + relative)
            continue
        if ('rule=' + rule) not in path.read_text(encoding='utf-8-sig'):
            errors.append(rule + ': ' + relative + ' 缺少 rule=' + rule + ' 标记')
        marker = check.split('/')[0]
        if not function_exists(path, marker):
            errors.append(rule + ': ' + relative + ' 缺少检查点 ' + marker)
    for folder in ('C-00-Shared', 'C-01-ExcelIterate', 'C-02-WordUpdater', 'C-03-Website·Xycabin'):
        for path in (ROOT / folder).rglob('*'):
            if not path.is_file() or path.suffix.lower() not in ('.py', '.cs'):
                continue
            if 'build' in path.parts or '__pycache__' in path.parts:
                continue
            for rule in RULE_MARK.findall(path.read_text(encoding='utf-8-sig')):
                if rule not in RULES:
                    errors.append(str(path) + ': 未登记的实现标记 rule=' + rule)
    return errors


GATE_TABLE = re.compile(r'^###\s+\(1\)\.\s*统一重建门禁总览\s*$', re.M)
SEMANTICS = {
    'G-01': 'FAIL', 'G-02': 'FAIL', 'G-03': 'FAIL', 'G-04': 'FAIL',
    'G-05': 'FAIL', 'G-06': 'FAIL', 'G-07': 'FAIL', 'G-08': 'FAIL',
    'G-09': 'FAIL', 'G-10': 'FAIL+SKIP(node)',
    'G-11': 'FAIL(HIT 保留原检查时间)', 'G-12': 'FAIL', 'G-13': 'FAIL',
    'G-14': 'FAIL(HIT 保留原检查时间)', 'G-15': 'FAIL', 'G-16': 'FAIL',
    'G-17': 'FAIL', 'G-18': 'FAIL+SKIP(node)', 'G-19': 'FAIL',
    'G-20': 'FAIL(回滚)', 'G-21': 'FAIL', 'G-22': 'FAIL', 'G-23': 'FAIL',
}
GATE_REGISTRY = {
    'G-01': ('六份文档与规则注册表', 'R-01–R-16,R-20'),
    'G-02': ('公共实现与功能源码分类', 'R-15,R-30'),
    'G-03': ('三个 CMD 入口契约', 'R-23'),
    'G-04': ('加载项身份预检', 'R-17'),
    'G-05': ('清理机制回归（隔离夹具）', 'R-21'),
    'G-06': ('缓存收据核对', ''),
    'G-07': ('清理审阅机制回归（隔离夹具）', 'R-21'),
    'G-08': ('构建前文档指纹', 'R-20'),
    'G-09': ('网站快照构建', 'R-14'),
    'G-10': ('快照网页门禁', 'R-22'),
    'G-11': ('Excel 重建链', 'R-16–R-19,R-30'),
    'G-12': ('Excel 后台验证', 'R-17,R-25,R-26,R-27,R-28,R-29,R-30'),
    'G-13': ('Excel 五载荷与内嵌资产', 'R-18,R-24'),
    'G-14': ('Word 重建链', 'R-16–R-19,R-30'),
    'G-15': ('Word 后台验证', 'R-17,R-25,R-26,R-27,R-29,R-30'),
    'G-16': ('Word 五载荷与内嵌资产', 'R-18,R-24'),
    'G-17': ('下载页组装与两包核对', 'R-24'),
    'G-18': ('下载页网页门禁', 'R-22'),
    'G-19': ('发布前文档指纹', 'R-20'),
    'G-20': ('事务发布', ''),
    'G-21': ('发布后布局门禁', 'R-02,R-23'),
    'G-22': ('缓存提交与运行摘要', 'R-20'),
    'G-23': ('正式产物收尾核对', 'R-14,R-17–R-20,R-24'),
}


def clause_rules(cell):
    """Expand every R-nn and R-nn–R-mm reference inside a clause cell."""
    rules = set()
    for match in re.finditer(r'R-(\d{2})(?:\s*[–-]\s*R-(\d{2}))?', cell):
        low = int(match.group(1))
        high = int(match.group(2) or match.group(1))
        rules.update('R-%02d' % index for index in range(low, high + 1))
    return rules


# rule=R-20
def check_gate_inventory(text):
    """The B3 gate inventory and its registry describe the same gates and clauses."""
    errors = []
    if GATE_TABLE.search(text) is None:
        return ['B3: 缺少 ### (1). 统一重建门禁总览 门禁清单']
    rows = GATE_ROW.findall(text)
    if len(rows) != len(GATE_REGISTRY):
        errors.append('B3: 门禁清单 %d 行，注册表 %d 项' % (len(rows), len(GATE_REGISTRY)))
    for index, row in enumerate(rows, 1):
        gate, title, _, clause, semantic, scope = row
        if gate != 'G-%02d' % index:
            errors.append('B3: 门禁清单第 %d 行应为 G-%02d，实际 %s' % (index, index, gate))
        registered = GATE_REGISTRY.get(gate)
        if registered is None:
            errors.append('B3: 未登记门禁 ' + gate)
            continue
        if title.strip() != registered[0]:
            errors.append(gate + ': 门禁名称与注册表不一致：' + title.strip())
        if scope.strip() != GATE_SCOPES[gate]:
            errors.append(gate + ': 执行范围与入口不一致')
        rules = clause_rules(clause)
        unknown = sorted(rules - set(RULES))
        if unknown:
            errors.append(gate + ': 对应条款引用未登记规则 ' + ', '.join(unknown))
        if rules != clause_rules(registered[1]):
            errors.append(gate + ': 对应条款与注册表不一致：' + clause.strip())
        if semantic.strip() != SEMANTICS.get(gate):
            errors.append(gate + ': 失败语义应为 ' + str(SEMANTICS.get(gate)) + '，实际 ' + semantic.strip())
    for gate in sorted(set(GATE_REGISTRY) - {row[0] for row in rows}):
        errors.append('B3: 门禁清单缺少 ' + gate)
    errors += gate_consistency()
    errors += check_execution_points()
    return errors


STEM_GATES = {
    'devlog-validation': ('G-01',),
    'shared-source': ('G-02',),
    'workflow-validation': ('G-03',),
    'cleanup-validation': ('G-05',),
    'cache-validation': ('G-06',),
    'cleanup-review-validation': ('G-07',),
    'site-validation': ('G-10',),
    'site-downloads-validation': ('G-18',),
    'Excel_Iterate-rebuild': ('G-11',),
    'Excel_Iterate-validation': ('G-12',),
    'Excel_Iterate-payload': ('G-13',),
    'WordUpdater-rebuild': ('G-14',),
    'WordUpdater-validation': ('G-15',),
    'WordUpdater-payload': ('G-16',),
    'Excel_Iterate-certificate': ('G-11',),
    'WordUpdater-certificate': ('G-14',),
    'payload-probe-build': ('G-13', 'G-16'),
}


def read_summary(path):
    try:
        data = json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) and 'status' in data else None


def gate_consistency():
    """The previous run summaries must account for every registered gate."""
    root = ROOT / 'C-03-Website·Xycabin/build'
    errors, notes, accounted, unknown = [], [], set(), []
    last = read_summary(root / 'last-run-summary.json')
    closing = read_summary(root / 'validate-run-summary.json')
    development = read_summary(root / 'development-run-summary.json')
    if development is not None:
        errors += profile_errors(development)
    if last is not None and last.get('schema') in (2, 3):
        expected = BUILD_GATES if last.get('schema') == 3 else ['G-%02d' % i for i in range(1, 23)]
        if last.get('schema') == 3 and last.get('profile') != 'build':
            errors.append('统一重建摘要范围必须为 build')
        errors += ordered_errors(last, expected) + profile_errors(last)
        if closing is not None and closing.get('parent_run_id') == last.get('run_id'):
            if closing.get('schema') == 3 and closing.get('profile') != 'delivery':
                errors.append('收尾摘要范围必须为 delivery')
            errors += ordered_errors(closing, ['G-23']) + profile_errors(closing)
            if closing.get('status') != 'success':
                # A failed previous attempt is diagnostic; it must not prevent fixing and rerunning G-23.
                notes.append('本轮 G-23 上次未成功，当前重试仍须完整执行')
        else:
            notes.append('G-23 尚无属于本轮的独立验证记录')
        for note in notes:
            print('INFO: ' + note, flush=True)
        return errors
    if last is None:
        notes.append('无上一轮统一重建摘要，跳过门禁执行核对')
    else:
        for data, label in ((last, 'last-run-summary.json'), (closing, 'validate-run-summary.json')):
            if data is None:
                notes.append('缺少运行摘要：' + label)
                continue
            for stage in data.get('stages', []):
                gates = STEM_GATES.get(stage.get('stage'))
                if gates is None:
                    unknown.append(str(stage.get('stage')))
                elif stage.get('returncode'):
                    notes.append('阶段失败：' + str(stage.get('stage')))
                else:
                    accounted.update(gates)
            for item in data.get('inprocess_gates', []):
                if item.get('status') == 'pass':
                    accounted.add(item.get('gate'))
        if closing is None:
            notes.append('收尾独立验证尚未执行（G-23）')
        elif closing.get('status') == 'success':
            accounted.add('G-23')
        else:
            errors.append('收尾独立验证未成功：G-23')
        for assembly in last.get('reused_builds', []):
            notes.append('缓存复用，保留原检查时间：' + str(assembly))
        if last.get('status') == 'success':
            missing = sorted(set(GATE_REGISTRY) - accounted - ({'G-23'} if closing is None else set()))
            if missing:
                errors.append('上一轮重建未执行门禁：' + ', '.join(missing))
        else:
            notes.append('上一轮重建未成功，仅核对已执行门禁无异常')
    for stage in unknown:
        errors.append('运行摘要出现未登记门禁阶段：' + stage)
    for note in notes:
        print('INFO: ' + note, flush=True)
    return errors


# rule=R-20
def check_execution_points():
    """Bind every top-level gate to an actual report context in its executable entry."""
    errors = []
    def bindings(path):
        if not path.is_file():
            errors.append('缺少统一重建执行入口：' + str(path))
            return []
        tree = ast.parse(path.read_text(encoding='utf-8-sig'))
        result = []
        for node in ast.walk(tree):
            if isinstance(node, ast.With):
                for item in node.items:
                    call = item.context_expr
                    if isinstance(call, ast.Call) and isinstance(call.func, ast.Attribute) and call.func.attr == 'gate':
                        for arg in call.args:
                            result.extend(child.value for child in ast.walk(arg)
                                          if isinstance(child, ast.Constant) and isinstance(child.value, str)
                                          and re.fullmatch(r'G-\d{2}', child.value))
        return result
    root = ROOT / 'C-03-Website·Xycabin'
    found = bindings(root / 'rebuild.py')
    expected = set(BUILD_GATES)
    if set(found) != expected or len(found) != len(expected):
        errors.append('rebuild.py 的实际报告关口应绑定常规 build 范围：' + repr(found))
    validation = bindings(root / 'validate.py')
    expected_validation = set(DEVELOPER_GATES) | {'G-23'}
    if set(validation) != expected_validation or len(validation) != len(expected_validation):
        errors.append('validate.py 必须绑定 G-23 与四个显式开发自检')
    return errors


def function_exists(path, name):
    """A registered check point must exist as a real definition in its file."""
    text = path.read_text(encoding='utf-8-sig')
    if path.suffix.lower() == '.py':
        try:
            tree = ast.parse(text)
        except SyntaxError:
            return False
        return any(isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == name
                   for node in ast.walk(tree))
    return bool(re.search(r'\b' + re.escape(name) + r'\s*\(', text))


# rule=R-04
def check_dates(text, label, required):
    cleaned = re.sub(r'(?ms)^```.*?^```[ \t]*$', '', text)
    cleaned = re.sub(r'https?://[^\s)]+|`[^`\n]+`', '', cleaned)
    dates = re.findall(r'\d{4}-\d{2}-\d{2}', cleaned)
    top = re.findall(r'^最后更新：(\d{4}-\d{2}-\d{2})', text, re.M)
    if len(top) > 1:
        return [label + ': 只允许一处“最后更新”日期']
    if dates and (len(top) != 1 or dates != top):
        return [label + ': 正文不得出现“最后更新”以外的日期']
    if required and len(top) != 1:
        return [label + ': 缺少顶部“最后更新”日期']
    return []


# rule=R-05
def version_rows(block, label):
    errors, seen, previous = [], set(), None
    section = re.search(r'^## \d+\. 版本摘要\s*\n(.*?)(?=^## |\Z)', block, re.M | re.S)
    if section is None:
        return [label + ': 缺少版本摘要章节']
    for line in section.group(1).splitlines():
        row = re.match(r'^\s*\|\s*(\d+\.\d+\.\d+\.\d+)\s*\|\s*(.*?)\s*\|\s*$', line)
        if not row:
            continue
        version, summary = row.groups()
        if version in seen:
            errors.append(label + ': 重复版本 ' + version)
        seen.add(version)
        numeric = tuple(map(int, version.split('.')))
        if previous is not None and numeric < previous:
            errors.append(label + ': 版本摘要须升序：' + version)
        previous = numeric
        count = len(summary.strip())
        if count > LIMIT:
            errors.append(label + ': ' + version + ' 摘要 ' + str(count) + ' 字，超过 200 字')
    if not seen:
        errors.append(label + ': 版本摘要没有可识别的版本行')
    return errors


def chapter_overview(lines):
    index = 0
    while index < len(lines) and (not lines[index].strip() or lines[index].startswith('#')):
        index += 1
    return index < len(lines) and bool(lines[index].strip())


# rule=R-06
# rule=R-07
def verify_file(path):
    gate = markdown_gate()
    rules = document_rules()
    text = path.read_text(encoding='utf-8-sig')
    name = path.name
    errors, titles = gate.hierarchy(text, name)
    try:
        gate.verify_local_anchors(path)
    except ValueError as error:
        errors += str(error).splitlines()
    errors += check_no_comments(text, name)
    errors += check_dates(text, name, required=(name == 'B1-Project-Files.md'))
    try:
        rules.strip_delivery_audit(text)
    except ValueError as error:
        errors.append(name + ': ' + str(error))

    if name == 'B0-Entry.md':
        required = ('入口与阅读顺序', '三个输出的关系', '共享与同步', '交付边界')
        missing = [item for item in required if not re.search(r'(?m)^## \d+\. ' + re.escape(item) + r'$', text)]
        if missing:
            errors.append(name + ': 缺少章节 ' + ', '.join(missing))
        return errors

    try:
        _, chapters = rules.split_chapters(text)
        scopes = [rules.chapter_scope(title, name) for title, _ in chapters]
    except ValueError as error:
        return errors + [str(error)]
    if scopes != EXPECTED_SCOPES[name]:
        errors.append(name + ': 一级章节范围应为 ' + ', '.join(EXPECTED_SCOPES[name])
                      + '，实际为 ' + ', '.join(scopes))

    if name == 'A-UserGuide.md':
        products = [re.sub(r'^[一二三四五六七八九十]+\.\s*', '', title) for title, _ in chapters]
        if products != ['Excel-Iterate', 'WordUpdater', 'XyCabin']:
            errors.append(name + ': 一级章节必须依次为三个产品')
        for title, lines in chapters:
            if not chapter_overview(lines):
                errors.append(name + ': 章节缺少导言：' + title)
        try:
            rules.document_overview(text)
        except ValueError as error:
            errors.append(name + ': ' + str(error))
    elif name == 'B1-Project-Files.md':
        try:
            rules.document_overview(text)
        except ValueError as error:
            errors.append(name + ': ' + str(error))
        if len(chapters) != 4:
            errors.append(name + ': 必须包含四个一级章节')
        audits = sum(1 for line in text.split('\n') if AUDIT_HEAD.match(line))
        if audits != 3:
            errors.append(name + ': 「本轮交付证据」必须恰好三处，实际 ' + str(audits))
        for title, lines in chapters[1:]:
            body = '\n'.join(lines)
            found = sum(1 for line in body.split('\n') if AUDIT_HEAD.match(line))
            if found != 1:
                errors.append(name + ': 每个项目档案须有一处交付证据：' + title)
        for title, lines in chapters:
            if rules.chapter_scope(title, name) in ('excel', 'word'):
                errors += version_rows('\n'.join(lines), name + '/' + title)
    elif name == 'B2-DevExperience.md':
        if len(chapters) != 4:
            errors.append(name + ': 必须包含四个一级章节')
    elif name == 'B3-Spec Coding.md':
        if len(chapters) != 2:
            errors.append(name + ': 必须包含两个一级章节')
    elif name == 'B4-Code Index.md':
        if len(chapters) != 2:
            errors.append(name + ': 必须依次包含 Excel-Iterate 与 WordUpdater 两章')
    return errors


def verify(files=None):
    paths = [DOCS / name for name in FILES] if files is None else [Path(item) for item in files]
    errors = []
    for path in paths:
        if not path.is_file():
            errors.append('缺少文档：' + str(path))
            continue
        errors += verify_file(path)
    b3 = DOCS / 'B3-Spec Coding.md'
    if b3.is_file():
        text = b3.read_text(encoding='utf-8-sig')
        errors += check_rules(text)
        errors += check_h_items(text)
        errors += check_gate_inventory(text)
    if errors:
        raise ValueError('\n'.join(errors))
    print('PASS: C-04 六份文档的章节范围、概览、交付证据、版本表、日期、R/H 门禁分类与规则注册表；' + str(len(paths)) + ' 份')


def self_test():
    import tempfile
    gate = markdown_gate()
    rules = document_rules()
    assert gate.hierarchy('# 一. 甲\n## 1. 乙\n', 'ok')[0] == []
    assert version_rows('## 2. 版本摘要\n| 0.9.1.0 | 甲 |\n| 0.9.2.0 | 乙 |\n', 'ok') == []
    assert version_rows('## 2. 版本摘要\n| 0.9.2.0 | 甲 |\n| 0.9.1.0 | 乙 |\n', 'bad')
    assert version_rows('## 2. 版本摘要\n| 0.9.1.0 | ' + '中' * 201 + ' |\n', 'long')
    assert check_no_comments('<!-- x -->', 'bad')
    assert RULE_LINE.findall('**R-01** 甲\n**R-02** 乙\n') == ['R-01', 'R-02']
    assert check_gate_inventory('无清单正文') == ['B3: 缺少 ### (1). 统一重建门禁总览 门禁清单']
    sample = '### (1). 统一重建门禁总览\n\n| 编号 | 门禁 | 执行入口 | 对应条款 | 失败语义 | 执行范围 |\n| --- | --- | --- | --- | --- | --- |\n'
    for gate, (title, clause) in sorted(GATE_REGISTRY.items()):
        sample += '| %s | %s | 入口 | %s | %s | %s |\n' % (gate, title, clause or '二.2 正文', SEMANTICS[gate], GATE_SCOPES[gate])
    assert check_gate_inventory(sample) == []
    assert any('G-23' in item for item in check_gate_inventory(sample.replace('| G-23 |', '| 无关 |', 1)))
    assert any('失败语义' in item for item in check_gate_inventory(sample.replace('| FAIL |', '| 硬错 |', 1)))
    assert set(RULES) == {'R-%02d' % index for index in range(1, 31)}
    assert H_LINE.findall('**H-01** 甲。验收：x。证据：y。\n**H-02** 乙。验收：x。证据：y。\n') == ['H-01', 'H-02']
    hsample = '\n'.join('**H-%02d** 甲。验收：人工。证据：记录。' % index for index in range(1, 13))
    assert check_h_items(hsample) == []
    assert any('缺少人工门禁条目 H-12' in item for item in check_h_items('\n'.join(hsample.split('\n')[:-1])))
    assert any('H 编号重复' in item for item in check_h_items(hsample + '\n**H-01** 乙。验收：人工。证据：记录。'))
    swapped = hsample.replace('**H-01**', '**H-X1**').replace('**H-02**', '**H-01**').replace('**H-X1**', '**H-02**')
    assert any('顺序' in item for item in check_h_items(swapped))
    assert check_execution_points() == []
    assert rules.chapter_scope('一. Agent及开发者接手概览', 'B1-Project-Files.md') == 'shared'
    assert rules.chapter_scope('二. Excel-Iterate 项目档案', 'B1-Project-Files.md') == 'excel'
    try:
        rules.chapter_scope('一. 未知章节', 'B1-Project-Files.md')
    except ValueError:
        pass
    else:
        raise AssertionError('unknown chapter accepted')
    try:
        rules.chapter_scope('一. Excel-Iterate 与 WordUpdater', 'B2-DevExperience.md')
    except ValueError:
        pass
    else:
        raise AssertionError('ambiguous chapter accepted')
    try:
        rules.document_overview('# 一. 甲\n\n## 1. 乙\n')
    except ValueError:
        pass
    else:
        raise AssertionError('empty overview accepted')
    with tempfile.TemporaryDirectory() as temporary:
        path = Path(temporary) / 'B2-DevExperience.md'
        path.write_text('# 一. Word与Excel共享开发经验\n## 1. 乙\n# 二. Excel-Iterate 专属\n# 三. WordUpdater 专属\n# 四. XyCabin 专属\n', encoding='utf-8')
        assert verify_file(path) == []
        path = Path(temporary) / 'B3-Spec Coding.md'
        text = (DOCS / path.name).read_text(encoding='utf-8-sig')
        path.write_text(text, encoding='utf-8')
        assert verify_file(path) == []
        assert check_rules(text) == [] and check_h_items(text) == [] and check_gate_inventory(text) == []
        path.write_text(text + '\n# 三. Word与Excel共享门禁\n', encoding='utf-8')
        assert any('两个一级章节' in item for item in verify_file(path))
        assert check_rules(text.replace('**R-30**', '**R-99**'))
    print('VerifyDevlog self-test: 25 PASS, 0 FAIL', flush=True)


def main():
    arguments = sys.argv[1:]
    if arguments and arguments[0] == '--self-test':
        self_test()
        return
    verify(arguments or None)


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, AssertionError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
