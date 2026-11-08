"""Spec-ordered build results and durable per-run logs shared by all launchers."""
from contextlib import contextmanager, redirect_stdout, redirect_stderr
from datetime import datetime, timezone
from pathlib import Path
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import uuid
from process_support import configure_background, FocusTrace, popen_process

_ACTIVE = None
GATE_ROW = re.compile(r'^\|\s*(G-\d{2})\s*\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|\s*$', re.M)
H_ROW = re.compile(r'^\*\*(H-\d{2})\*\*\s*(.+)$', re.M)
H_SUMMARY_ROW = re.compile(r'^\|\s*(H-\d{2})\s*\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|\s*$', re.M)
H_ITEMS = {
    'H-01': '维护顺序', 'H-02': '构建期禁编辑', 'H-03': '安装与卸载报告', 'H-04': '排障顺序',
    'H-05': '实机 COM 验收', 'H-06': '弹窗真实交互验收', 'H-07': '清理触发', 'H-08': '冗余候选条件',
    'H-09': '五类盘点', 'H-10': 'fixture 与删除边界', 'H-11': '网站经验归属', 'H-12': '固定图样不读桌面',
    'H-13': '特定示例代码优化',
}
H_DEFAULT_INACTIVE = frozenset(('H-03', 'H-04', 'H-07', 'H-08', 'H-09', 'H-10', 'H-13'))
DEVELOPER_GATES = ('G-03', 'G-05', 'G-06', 'G-07')
BUILD_GATES = tuple('G-%02d' % i for i in range(1, 23) if 'G-%02d' % i not in DEVELOPER_GATES)
GATE_SCOPES = {'G-%02d' % i: 'developer' if 'G-%02d' % i in DEVELOPER_GATES else 'delivery' if i == 23 else 'build'
               for i in range(1, 24)}


def manual_spec_items(text):
    """Read manual guidance metadata only; this never invokes its maintenance actions."""
    definitions = H_ROW.findall(text)
    found = [key for key, _ in definitions]
    expected = ['H-%02d' % i for i in range(1, len(H_ITEMS) + 1)]
    if len(found) != len(set(found)):
        raise ValueError('人工门禁条目 H 编号重复')
    missing, extra = sorted(set(H_ITEMS) - set(found)), sorted(set(found) - set(H_ITEMS))
    if missing:
        raise ValueError('缺少人工门禁条目 ' + ', '.join(missing))
    if extra:
        raise ValueError('未登记人工门禁条目 ' + ', '.join(extra))
    if found != expected:
        raise ValueError('人工门禁条目顺序必须为 H-01…H-%02d' % len(H_ITEMS))
    humans = [(key, body.split('：', 1)[0].strip()) for key, body in definitions]
    if any(title != H_ITEMS[key] for key, title in humans):
        raise ValueError('人工条款主题与注册表不一致')
    if any('验收：' not in body or '证据：' not in body for _, body in definitions):
        raise ValueError('人工条款缺少“验收/证据”记录')
    rows = H_SUMMARY_ROW.findall(text)
    if [row[0] for row in rows] != expected:
        raise ValueError('人工条款汇总缺项、重复或顺序不一致')
    for key, title, applicable, trigger, method, evidence, display in rows:
        if title.strip() != H_ITEMS[key] or not all(cell.strip() for cell in (applicable, trigger, method, evidence)):
            raise ValueError(key + ' 人工条款汇总主题或核对信息不一致')
        default = '本轮不适用' if key in H_DEFAULT_INACTIVE else '人工核对'
        if display.strip() != default:
            raise ValueError(key + ' 人工条款构建显示必须为“' + default + '”')
    return humans


def spec_items(repository):
    text = (Path(repository) / 'C-04-Devlog/B3-Spec Coding.md').read_text(encoding='utf-8-sig')
    rows = GATE_ROW.findall(text)
    gates = {key: {'name': name.strip(), 'rules': rules.strip(), 'semantics': semantics.strip(), 'scope': scope.strip()}
             for key, name, _, rules, semantics, scope in rows}
    humans = manual_spec_items(text)
    if len(rows) != 23 or list(gates) != ['G-%02d' % i for i in range(1, 24)]:
        raise ValueError('B3 门禁总览必须按 G-01～G-23 排列')
    if any(info['scope'] != GATE_SCOPES[key] for key, info in gates.items()):
        raise ValueError('B3 门禁执行范围与实际入口不一致')
    return gates, humans


def ordered_errors(data, expected):
    """Check actual chronology, including incomplete failed runs, without sorting it."""
    records = data.get('gates', [])
    actual = [item.get('gate') for item in records]
    errors = []
    if len(actual) != len(set(actual)):
        errors.append('门禁编号重复')
    if actual != list(expected)[:len(actual)]:
        errors.append('门禁缺失、未知编号或执行顺序错误：' + repr(actual))
    if data.get('status') == 'success' and actual != list(expected):
        errors.append('成功摘要缺少门禁：' + repr(list(expected)[len(actual):]))
    for index, item in enumerate(records):
        state = item.get('status')
        if state not in ('pass', 'cached', 'fail', 'partial'):
            errors.append(str(item.get('gate')) + ' 状态无效')
        if state == 'cached' and not item.get('evidence_time'):
            errors.append(str(item.get('gate')) + ' 缺少原检查时间')
        if data.get('status') == 'success' and state not in ('pass', 'cached'):
            errors.append(str(item.get('gate')) + ' 未通过却标记整轮成功')
        if state == 'fail' and index != len(records) - 1:
            errors.append('失败后仍执行后续门禁')
    return errors


def profile_errors(data):
    """Schema 2 keeps its historical schedule; schema 3 declares its scope."""
    if data.get('schema') != 3:
        return []
    profile = data.get('profile')
    schedules = {'build': BUILD_GATES, 'delivery': ('G-23',), 'developer': DEVELOPER_GATES}
    if profile in schedules:
        return ordered_errors(data, schedules[profile])
    if profile in ('product-build', 'product-validation'):
        allowed = ('G-11', 'G-14') if profile == 'product-build' else ('G-12', 'G-15')
        records = data.get('gates', [])
        first = records[0].get('gate') if records else allowed[0]
        return ordered_errors(data, [first if first in allowed else allowed[0]])
    return [] if profile == 'fixture' else ['未知摘要执行范围：' + str(profile)]


class _Capture:
    def __init__(self, stream, console, verbose):
        self.stream, self.console, self.verbose = stream, console, verbose

    def write(self, value):
        self.stream.write(value)
        self.stream.flush()
        if self.verbose:
            self.console.write(value)
            self.console.flush()
        return len(value)

    def flush(self):
        self.stream.flush()
        if self.verbose:
            self.console.flush()

    def isatty(self):
        return False


class BuildReport:
    def __init__(self, repository, build, phase, expected, mode='full', verbose=False,
                 background=False, console=None, profile=None):
        self.metadata, self.humans = spec_items(repository)
        self.expected = tuple(expected)
        self.mode, self.verbose, self.background = mode, verbose, background
        configure_background(background)
        self.profile = profile or ('build' if phase == 'rebuild' and self.expected == BUILD_GATES else
                                   'delivery' if phase == 'closing' else 'developer' if phase == 'development' else
                                   'product-validation' if phase == 'validation' else 'product-build' if phase == 'rebuild' else 'fixture')
        self.run_id = os.environ.get('XY_REPORT_RUN_ID') or uuid.uuid4().hex
        if not re.fullmatch(r'[a-f0-9]{32}', self.run_id):
            raise ValueError('无效的报告运行标识')
        self.phase = phase
        self.directory = Path(build) / 'runs' / self.run_id / phase
        self.directory.mkdir(parents=True, exist_ok=True)
        self.console = sys.stdout if console is None else console
        self.gates, self.current = [], None
        self.started = time.perf_counter()
        self.focus = FocusTrace(self.directory, lambda: self.current['gate'] if self.current else None)
        self.focus.start()

    @contextmanager
    def gate(self, identifier):
        global _ACTIVE
        index = len(self.gates)
        if index >= len(self.expected) or identifier != self.expected[index]:
            raise ValueError('门禁次序错误：实际 %s，期望 %s' %
                             (identifier, self.expected[index:index + 1]))
        info = self.metadata[identifier]
        record = dict(gate=identifier, name=info['name'], rules=info['rules'], status='running',
                      started_at=datetime.now(timezone.utc).isoformat(), cached=False, skipped=[], logs=[])
        self.gates.append(record)
        self.current = record
        path = self.directory / (identifier + '.log')
        record['log'] = str(path.resolve())
        stop = threading.Event()
        started = time.perf_counter()
        tty = bool(getattr(self.console, 'isatty', lambda: False)()) and not self.verbose

        def progress():
            while not stop.wait(15):
                self.console.write('\r[%s] 运行中｜%s｜%.0f秒' %
                                   (identifier, info['name'], time.perf_counter() - started))
                self.console.flush()

        thread = threading.Thread(target=progress, daemon=True) if tty else None
        previous = _ACTIVE
        previous_gate = os.environ.get('XY_BUILD_GATE')
        os.environ['XY_BUILD_GATE'] = identifier
        _ACTIVE = self
        try:
            with path.open('w', encoding='utf-8') as log:
                sink = _Capture(log, self.console, self.verbose)
                if thread:
                    thread.start()
                try:
                    with redirect_stdout(sink), redirect_stderr(sink):
                        yield record
                    if record['status'] == 'running':
                        record['status'] = 'cached' if record['cached'] else 'pass'
                except BaseException as error:
                    record['status'], record['error'] = 'fail', str(error)
                    traceback.print_exception(type(error), error, error.__traceback__, file=log)
                    raise
        except BaseException as error:
            if record['status'] == 'running':
                record['status'], record['error'] = 'fail', str(error)
            raise
        finally:
            stop.set()
            if thread and thread.is_alive():
                thread.join(timeout=2)
            _ACTIVE, self.current = previous, None
            if previous_gate is None:
                os.environ.pop('XY_BUILD_GATE', None)
            else:
                os.environ['XY_BUILD_GATE'] = previous_gate
            record['seconds'] = round(time.perf_counter() - started, 3)
            output = path.read_text(encoding='utf-8') if path.is_file() else ''
            for line in output.splitlines():
                if line.startswith(('SKIP:', '待前台验收（本轮未运行）：')) and line not in record['skipped']:
                    record['skipped'].append(line)
            state = {'pass': '通过', 'cached': '复用', 'fail': '失败', 'partial': '部分验证'}[record['status']]
            if record['skipped']:
                state += '（含未执行项）'
            evidence = '｜原检查 ' + record['evidence_time'] if record.get('evidence_time') else ''
            self.console.write(('\r' + ' ' * 120 + '\r' if tty else '') + '[%s] %s｜%s｜%s｜%.1f秒%s\n' %
                               (identifier, state, info['name'], info['rules'], record['seconds'], evidence))
            if record['status'] == 'fail':
                self.console.write('错误：%s\n日志：%s\n' % (record['error'], path.resolve()))
                self.console.write('\n'.join(output.splitlines()[-10:]) + '\n')
            self.console.flush()

    def reuse(self, evidence_time, label='', whole_gate=False):
        if self.current is None:
            raise ValueError('缓存证据必须属于正在执行的门禁')
        if whole_gate:
            self.current.update(cached=True, evidence_time=evidence_time)
        else:
            self.current.setdefault('reused_checks', []).append(dict(check=label, evidence_time=evidence_time))

    def summary(self, status, **extra):
        data = dict(schema=3, profile=self.profile, run_id=self.run_id, phase=self.phase, mode=self.mode, status=status,
                    seconds=round(time.perf_counter() - self.started, 3), gates=self.gates, focus_trace=self.focus.stop())
        data.update(extra)
        errors = ordered_errors(data, self.expected) + profile_errors(data)
        if errors:
            raise ValueError('\n'.join(errors))
        return data

    def manual_items(self, site=True):
        if os.environ.get('XY_REPORT_NESTED') == '1':
            return
        for key, title in self.humans:
            inactive = key in H_DEFAULT_INACTIVE or (not site and key == 'H-11')
            self.console.write('[%s] %s｜%s\n' % (key, '本轮不适用' if inactive else '人工核对', title))
        self.console.write('详细日志：' + str(self.directory.resolve()) + '\n')
        self.console.flush()


def active_report():
    return _ACTIVE


def relay_omissions(output):
    """Keep omitted subchecks visible to the gate even when a probe is captured or cached."""
    for line in output.splitlines():
        if line.startswith(('SKIP:', '待前台验收（本轮未运行）：')):
            print(line, flush=True)


def run_command(command, cwd, log):
    """Stream complete child output into the current gate and its own durable log."""
    report = _ACTIVE
    log = (report.directory / Path(log).name) if report else Path(log)
    log.parent.mkdir(parents=True, exist_ok=True)
    if report and report.current is not None:
        report.current['logs'].append(str(log.resolve()))
    environment = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUTF8='1',
                       PYTHONDONTWRITEBYTECODE='1', XY_REPORT_NESTED='1')
    if report:
        environment['XY_REPORT_RUN_ID'] = report.run_id
    started = time.perf_counter()
    child = popen_process(list(map(str, command)), cwd=cwd, env=environment, stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace')
    try:
        with log.open('w', encoding='utf-8') as stream:
            for line in child.stdout:
                stream.write(line)
                stream.flush()
                print(line, end='', flush=True)
        code = child.wait()
    except BaseException:
        child.kill()
        child.wait()
        raise
    finally:
        child.stdout.close()
    result = dict(stage=log.stem, seconds=round(time.perf_counter() - started, 3), returncode=code,
                  log=str(log.resolve()))
    if code:
        raise RuntimeError('阶段 %s 失败，退出码 %s；日志：%s' % (log.stem, code, log.resolve()))
    return result


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.tmp-' + uuid.uuid4().hex)
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temporary, path)
