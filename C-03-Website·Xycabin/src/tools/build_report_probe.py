"""Exercise real reporting contexts and child processes in disposable directories."""
from contextlib import redirect_stdout
from pathlib import Path
import copy
import io
import os
import sys
import tempfile
import subprocess
from unittest.mock import patch

sys.dont_write_bytecode = True
REPOSITORY = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPOSITORY / 'C-00-Shared'))
from process_support import run_process
from build_report import BuildReport, ordered_errors, relay_omissions, run_command, spec_items


def verify_report():
    checks = 0

    def check(value, label):
        nonlocal checks
        if not value:
            raise AssertionError(label)
        checks += 1

    with tempfile.TemporaryDirectory(prefix='xy-report-中文·') as temporary:
        root = Path(temporary)
        console = io.StringIO()
        report = BuildReport(REPOSITORY, root, 'rebuild', ['G-11', 'G-12'], console=console, profile='fixture')
        with report.gate('G-11'):
            run_command([sys.executable, '-X', 'utf8', '-c', "print('内部细节：中文')"], root, root / 'child.log')
            report.reuse('original-evidence-time', whole_gate=True)
        with report.gate('G-12'):
            captured = run_process([sys.executable, '-X', 'utf8', '-c',
                                       "print('SKIP: 缺少 Node.js；下载逻辑未执行'); print('PASS: final probe line')"],
                                      capture_output=True, text=True, encoding='utf-8', check=True).stdout
            relay_omissions(captured)
            relay_omissions(captured)  # A cached replay must keep the original reason without duplicate summary entries.
            print('待前台验收（本轮未运行）：窗口拖动')
            report.reuse('test-evidence-time', 'fixture.exe')
        data = report.summary('success')
        output = console.getvalue()
        check(output.count('[G-11]') == 1 and output.count('[G-12]') == 1, 'one default result per actual gate')
        check('内部细节' not in output and '内部细节' in Path(data['gates'][0]['log']).read_text(encoding='utf-8'),
              'quiet console retains complete child output in its gate log')
        check('原检查 original-evidence-time' in output and data['gates'][0]['status'] == 'cached', 'cache time is preserved')
        check(len(data['gates'][1]['skipped']) == 2, 'node and foreground omissions are recorded')
        check(data['gates'][1]['skipped'].count('SKIP: 缺少 Node.js；下载逻辑未执行') == 1,
              'captured and cached probe output retains its Node omission exactly once')
        check(data['gates'][1]['reused_checks'][0]['evidence_time'] == 'test-evidence-time', 'partial cache evidence is retained')
        check(ordered_errors(data, ['G-11', 'G-12']) == [], 'valid execution chronology')
        for records, label in [([], 'missing'), (list(reversed(data['gates'])), 'out of order'),
                               ([data['gates'][0]] * 2, 'duplicate'),
                               ([dict(data['gates'][0], gate='G-99')], 'unknown')]:
            changed = dict(data, gates=records)
            check(bool(ordered_errors(changed, ['G-11', 'G-12'])), 'reject ' + label + ' gates')
        changed = copy.deepcopy(data)
        changed['gates'][0].pop('evidence_time')
        check(bool(ordered_errors(changed, ['G-11', 'G-12'])), 'cached success requires the original evidence time')
        failure_console = io.StringIO()
        failure = BuildReport(REPOSITORY, root, 'closing', ['G-23'], console=failure_console)
        try:
            with failure.gate('G-23'):
                run_command([sys.executable, '-c', "print('failure detail'); raise SystemExit(7)"], root, root / 'child.log')
        except RuntimeError as error:
            check('退出码 7' in str(error), 'the real child exit code propagates')
        else:
            raise AssertionError('failing child accepted')
        check(failure.gates[0]['status'] == 'fail' and 'failure detail' in failure_console.getvalue(),
              'failure identifies its gate and shows useful diagnostics')
        check(report.gates[0]['logs'][0] != failure.gates[0]['logs'][0], 'closing logs never overwrite rebuild logs')
        check('内部细节' in Path(report.gates[0]['logs'][0]).read_text(encoding='utf-8'), 'original log survives later validation')
        verbose_console = io.StringIO()
        verbose = BuildReport(REPOSITORY, root, 'validation', ['G-15'], verbose=True, console=verbose_console)
        with verbose.gate('G-15'):
            print('verbose internal output')
        check('verbose internal output' in verbose_console.getvalue(), 'verbose mode shows internal output')
        try:
            with verbose.gate('G-16'):
                raise AssertionError('unexpected gate body ran')
        except ValueError:
            check(True, 'runtime guard blocks an undeclared next gate before its body')
        human_console = io.StringIO()
        human = BuildReport(REPOSITORY, root, 'human', [], console=human_console)
        nested = os.environ.pop('XY_REPORT_NESTED', None)
        try:
            human.manual_items(site=False)
        finally:
            if nested is not None:
                os.environ['XY_REPORT_NESTED'] = nested
        lines = [line for line in human_console.getvalue().splitlines() if line.startswith('[H-')]
        check([line[1:5] for line in lines] == ['H-%02d' % i for i in range(1, 13)], 'manual items follow Spec order')
        check(all('通过' not in line for line in lines), 'manual items never claim automatic success')
        fake = root / 'repository'
        (fake / 'C-04-Devlog').mkdir(parents=True)
        spec = (REPOSITORY / 'C-04-Devlog/B3-Spec Coding.md').read_text(encoding='utf-8-sig')
        (fake / 'C-04-Devlog/B3-Spec Coding.md').write_text(spec + '\n| G-01 | duplicate | entry | R-01 | FAIL | build |\n', encoding='utf-8')
        try:
            spec_items(fake)
        except ValueError:
            check(True, 'a duplicate Spec table row cannot disappear in a dictionary')
        else:
            raise AssertionError('duplicate Spec row accepted')
        partial = dict(data, status='success', gates=[dict(data['gates'][0], status='partial')])
        check(bool(ordered_errors(partial, ['G-11'])), 'partial verification cannot claim a complete successful gate')
        class Terminal(io.StringIO):
            def isatty(self):
                return True
        denied = BuildReport(REPOSITORY, root, 'denied', ['G-14'], console=Terminal())
        try:
            with patch.object(Path, 'open', side_effect=PermissionError('fixture log unavailable')):
                with denied.gate('G-14'):
                    raise AssertionError('unwritable gate body ran')
        except PermissionError as error:
            check(str(error) == 'fixture log unavailable', 'an unopened log preserves the original I/O error before the progress thread starts')
        else:
            raise AssertionError('an unwritable gate log was accepted')
    print('BuildReportProbe: %d PASS, 0 FAIL' % checks, flush=True)
    return checks


if __name__ == '__main__':
    verify_report()
