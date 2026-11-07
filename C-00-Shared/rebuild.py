"""Run the shared generate/build/sign/package sequence; never install or stop WPS."""
import os
import argparse
import subprocess
import sys
sys.dont_write_bytecode = True
from project_support import ROOT, OUT, PROJECT, CONFIGURE_NAME, load_site, verify_layout, verify_features
from build_report import BuildReport, run_command, write_json


# rule=R-17
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--background', action='store_true', help='后台构建，不打开安装器或宿主窗口')
    parser.add_argument('--open-installer-only', action='store_true')
    parser.add_argument('--site-html')
    parser.add_argument('--installer-output')
    parser.add_argument('--verbose', action='store_true', help='在控制台显示完整构建日志')
    args = parser.parse_args()
    if args.background and args.open_installer_only:
        raise ValueError('--background 不能与 --open-installer-only 同时使用')
    if args.open_installer_only:
        os.startfile(str(OUT / (PROJECT['assembly'] + '_Setup.exe')))
        return
    gate = 'G-11' if PROJECT['host'] == 'Excel' else 'G-14'
    report = BuildReport(ROOT.parent, ROOT / 'build', 'rebuild', [gate], args.mode, args.verbose, args.background)
    status = 'failed'
    try:
        with report.gate(gate):
            rebuild(args, report)
        status = 'success'
    finally:
        write_json(ROOT / 'build/rebuild-run-summary.json', report.summary(status))


def rebuild(args, report):
    verify_features()
    website = load_site()
    verify_layout(source_only=True)
    if args.site_html:
        website.validate_snapshot(args.site_html, ROOT, require_embedded=True)
    for index, script in enumerate(('AA-Program/' + CONFIGURE_NAME, 'build.py', 'resign.py', 'build_installer.py'), 1):
        print('[%d/4] %s' % (index, script), flush=True)
        command = [sys.executable, str(ROOT / script)]
        if script in ('build.py', 'build_installer.py'):
            command += ['--mode', args.mode]
        if script == 'build.py' and args.site_html:
            command += ['--site-html', args.site_html]
        if script == 'build_installer.py' and args.installer_output:
            command += ['--output', args.installer_output]
        run_command(command, ROOT, report.directory / (str(index) + '-' + script.split('/')[-1] + '.log'))
    verify_layout(source_only=bool(args.installer_output))
    print('完成：' + str(args.installer_output or (OUT / (PROJECT['assembly'] + '_Setup.exe'))))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, RuntimeError, subprocess.CalledProcessError) as exc:
        print('FAIL:', exc)
        sys.exit(1)
