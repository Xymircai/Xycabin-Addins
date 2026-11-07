"""Regress quiet process launches, scoped schedules and nonrecursive delivery checks."""
from contextlib import ExitStack, redirect_stdout
from pathlib import Path
import importlib.util
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
from unittest.mock import patch

sys.dont_write_bytecode = True
SITE = Path(__file__).resolve().parents[2]
REPOSITORY = SITE.parent
sys.path.insert(0, str(REPOSITORY / 'C-00-Shared'))
sys.path.insert(0, str(SITE / 'src'))
from process_support import run_process, process_options, FocusTrace, BACKGROUND_ENV, TRACE_ENV
from build_report import BuildReport, BUILD_GATES, DEVELOPER_GATES, profile_errors, ordered_errors
from delivery import fingerprints


def module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def verify_quiet_build():
    checks = 0

    def check(value, label):
        nonlocal checks
        if not value:
            raise AssertionError(label)
        checks += 1

    def rejected(action, label):
        try:
            action()
        except (OSError, ValueError, RuntimeError, subprocess.SubprocessError):
            check(True, label)
        else:
            raise AssertionError('Accepted: ' + label)

    with tempfile.TemporaryDirectory(prefix='xy-quiet-中文·') as temporary, patch.dict(os.environ):
        root = Path(temporary)
        os.environ[BACKGROUND_ENV] = '1'
        os.environ['XY_REPORT_NESTED'] = '1'
        os.environ.pop(TRACE_ENV, None)
        if os.name == 'nt':
            child_code = "import ctypes,sys,json;print(json.dumps([int(ctypes.windll.kernel32.GetConsoleWindow() or 0),sys.stdin.read()]))"
            result = run_process([sys.executable, '-c', child_code], capture_output=True, text=True, check=True)
            check(json.loads(result.stdout) == [0, ''], 'background child has no console and receives EOF')
            nested_code = ("import sys;sys.path.insert(0,sys.argv[1]);from process_support import run_process;"
                           "p=run_process([sys.executable,'-c',sys.argv[2]],capture_output=True,text=True,check=True);print(p.stdout)")
            result = run_process([sys.executable, '-c', nested_code, str(REPOSITORY / 'C-00-Shared'), child_code],
                                 capture_output=True, text=True, check=True)
            check(json.loads(result.stdout) == [0, ''], 'nested child also has no console and receives EOF')
            rejected(lambda: process_options({'creationflags': subprocess.CREATE_NEW_CONSOLE}), 'new-console request rejected')
            rejected(lambda: process_options({'creationflags': subprocess.DETACHED_PROCESS}), 'detached-console request rejected')
        result = run_process([sys.executable, '-X', 'utf8', '-c', "import sys;print(sys.stdin.read());print('错误细节',file=sys.stderr)"],
                             input='合法输入2', capture_output=True, text=True, encoding='utf-8', check=True)
        check(result.stdout.strip() == '合法输入2' and result.stderr.strip() == '错误细节', 'explicit stdin and both UTF-8 streams preserved')
        try:
            run_process([sys.executable, '-c', "print('failure detail');raise SystemExit(9)"], capture_output=True, check=True)
        except subprocess.CalledProcessError as error:
            check(error.returncode == 9 and b'failure detail' in error.output, 'child exit code and diagnostics preserved')
        else:
            raise AssertionError('child failure swallowed')
        rejected(lambda: run_process([sys.executable, '-c', 'import time;time.sleep(5)'], timeout=.1), 'timeout terminates child')

        os.environ.pop('XY_REPORT_NESTED')
        trace = FocusTrace(root / 'focus', lambda: 'G-fixture')
        trace.start()
        # Wait only for a read-only initial sample, without sending input or activating any window.
        trace.stop_event.wait(.15)
        observation = trace.stop()
        if os.name == 'nt':
            events = [json.loads(line) for line in trace.path.read_text(encoding='utf-8').splitlines()]
            check(observation['status'] == 'observed' and events[0]['gate'] == 'G-fixture', 'foreground sample is associated with its gate')
            check(all('input_age_ms' in item and 'user_input_recent' in item for item in events), 'input timing is observed without attributing user intent')
            check(TRACE_ENV not in os.environ, 'trace environment restored on completion')
        os.environ['XY_REPORT_NESTED'] = '1'

        schedules = {'build': BUILD_GATES, 'developer': DEVELOPER_GATES, 'delivery': ('G-23',)}
        for profile, gates in schedules.items():
            data = dict(schema=3, profile=profile, status='success', gates=[dict(gate=gate, status='pass') for gate in gates])
            check(not profile_errors(data), 'valid scope ' + profile)
            for changed in (data['gates'][:-1], list(reversed(data['gates'])) if len(gates) > 1 else data['gates'] * 2,
                            data['gates'] + [dict(gate='G-99', status='pass')]):
                check(bool(profile_errors(dict(data, gates=changed))), 'reject missing, duplicate, unknown or reordered scope ' + profile)
        historical = dict(schema=2, status='success', gates=[dict(gate='G-%02d' % i, status='pass') for i in range(1, 23)])
        check(not profile_errors(historical) and not ordered_errors(historical, ['G-%02d' % i for i in range(1, 23)]), 'historical schema 2 retains its old schedule')
        check(len(BUILD_GATES) + 1 == 19 and not set(BUILD_GATES) & set(DEVELOPER_GATES), 'default and developer gates do not overlap')

        closing = module(SITE / 'validate.py', '_quiet_closing')
        fixture = root / 'delivery'
        page = fixture / 'AA-Program/XyCabin.html'
        page.parent.mkdir(parents=True)
        page.write_text('fixture page', encoding='utf-8')
        bases = []
        for name in ('Excel_Iterate', 'WordUpdater'):
            base = root / name
            (base / 'AA-Program').mkdir(parents=True)
            (base / 'AA-Program' / (name + '_Setup.exe')).write_bytes(b'MZ-fixture-' + name.encode())
            bases.append((base, name))
        targets = [base / 'AA-Program' / (name + '_Setup.exe') for base, name in bases] + [page]
        record = fingerprints(targets)
        for target in targets[:-1]:
            record[str(target)]['version'] = '0.9.9.3'
        (fixture / 'build').mkdir()
        record_path = fixture / 'build/current-delivery.json'
        record_path.write_text(json.dumps(record), encoding='utf-8')
        commands = []
        with patch.object(closing, 'ROOT', fixture), patch.object(closing, 'projects', return_value=bases), \
             patch.object(closing, 'version', return_value='0.9.9.3'), \
             patch.object(closing, 'run', side_effect=lambda command, *args: commands.append(command)), \
             patch.object(closing, 'verify_payload') as payload, patch.object(closing, 'verify_installers') as installers, redirect_stdout(io.StringIO()):
            closing.verify_delivery()
            check([Path(command[1]).name for command in commands] == ['verify_package.py', 'verify_package.py'], 'closing runs cryptographic gates, never product validate.py')
            check(payload.call_count == 2 and installers.call_count == 1, 'closing binds both five-payload probes and HTML byte verification')
            for target in (targets[0], page):
                original = target.read_bytes()
                target.write_bytes(original + b'tampered')
                rejected(closing.verify_delivery, 'tampered EXE or HTML rejected by delivery record')
                target.write_bytes(original)
            stale = json.loads(json.dumps(record))
            stale[str(targets[0])]['version'] = '0.0.0.0'
            record_path.write_text(json.dumps(stale), encoding='utf-8')
            rejected(closing.verify_delivery, 'stale version rejected')
            record_path.write_text(json.dumps(record), encoding='utf-8')
            with patch.object(closing, 'run', side_effect=RuntimeError('signature verification failed')):
                rejected(closing.verify_delivery, 'signature failure propagates out of G-23')
            with patch.object(closing, 'verify_installers', side_effect=ValueError('embedded byte mismatch')):
                rejected(closing.verify_delivery, 'embedded installer mismatch propagates out of G-23')

        # Execute the actual orchestrator with fake tools. A failed business probe must never publish.
        rebuild = module(SITE / 'rebuild.py', '_quiet_rebuild')
        from tools import layout_probe
        import delivery
        for mode, fail in (('full', False), ('incremental', False), ('full', True)):
            work = root / ('pipeline-' + mode + str(fail))
            work.mkdir()
            (work / 'verify_shared.py').write_text('', encoding='utf-8')
            site = work / 'site'
            (site / 'AA-Program').mkdir(parents=True)
            formal = site / 'AA-Program/XyCabin.html'
            formal.write_text('old delivery', encoding='utf-8')
            for base, name in bases:
                (base / 'project.json').write_text('{"tools":[]}', encoding='utf-8')
            calls = []
            def run_tool(command, *args):
                calls.append(command)
                if '--installer-output' in command:
                    Path(command[command.index('--installer-output') + 1]).write_bytes(b'MZ-candidate')
                if fail and Path(command[1]).name == 'validate.py':
                    raise RuntimeError('business probe failed')
            def make_snapshot(unused, target):
                target.write_text('new delivery', encoding='utf-8')
                return target
            def report_factory(unused, *args, **kwargs):
                return BuildReport(REPOSITORY, *args, **dict(kwargs, console=io.StringIO()))
            with ExitStack() as stack:
                for name, value in dict(ROOT=site, projects=lambda unused: bases, run=run_tool,
                                        BuildReport=report_factory, input_fingerprint=lambda *a: 'fixed',
                                        inputs=lambda *a: {}, build=make_snapshot,
                                        attach_installers=lambda text, packages: text, embedded_snapshot=lambda text: text,
                                        verify_installers=lambda *a: None, verify_payload=lambda *a: None,
                                        save=lambda *a: None, commit=lambda *a: None,
                                        receipt=lambda *a: {'time': 'original-evidence'} if mode == 'incremental' else None,
                                        RUN_STAGES=[]).items():
                    stack.enter_context(patch.object(rebuild, name, value))
                stack.enter_context(patch.object(delivery, 'version', return_value='0.9.9.3'))
                stack.enter_context(patch.object(layout_probe, 'verify', return_value=None))
                stack.enter_context(patch.object(sys, 'argv', ['rebuild.py', '--mode', mode, '--background']))
                stack.enter_context(redirect_stdout(io.StringIO()))
                if fail:
                    rejected(rebuild.rebuild, 'failed business probe aborts build')
                    check(formal.read_text(encoding='utf-8') == 'old delivery', 'failure preserves previous published HTML')
                    check(not any(Path(command[1]).name == 'rebuild.py' and command[1].parent == bases[1][0] for command in calls), 'failure prevents later product stages')
                else:
                    rebuild.rebuild()
                    check(sum(Path(command[1]).name == 'validate.py' for command in calls) == 2, 'exactly one business validation per product in ' + mode)
                    check(not any(Path(command[1]).stem in ('workflow_probe', 'cleanup_probe', 'cache_probe', 'cleanup_review_probe') for command in calls), 'developer mechanisms excluded in ' + mode)
                    check(formal.read_text(encoding='utf-8') == 'new delivery', 'successful candidate published in ' + mode)
                    summary = json.loads((site / 'build/last-run-summary.json').read_text(encoding='utf-8'))
                    check([item['gate'] for item in summary['gates']] == list(BUILD_GATES), 'real orchestrator follows scoped order in ' + mode)
                    if mode == 'incremental':
                        check(all(item['evidence_time'] == 'original-evidence' for item in summary['gates'] if item['status'] == 'cached'), 'incremental build retains original evidence time')
                check(not list((site / 'build').glob('stage-*')), 'disposable candidates removed after success or failure')
    print('QuietBuildProbe: %d PASS, 0 FAIL' % checks, flush=True)
    return checks


if __name__ == '__main__':
    verify_quiet_build()
