"""Delegate the document gate to the repository-level verifier."""
from pathlib import Path
import subprocess
import sys

sys.dont_write_bytecode = True
ROOT = Path(PROJECT_ROOT).resolve()
sys.path.insert(0, str(ROOT.parent / 'C-00-Shared'))
from process_support import run_process
TARGET = ROOT.parent / 'verify_devlog.py'


def main():
    if '--self-test' in sys.argv:
        command = [sys.executable, str(TARGET), '--self-test']
    else:
        command = [sys.executable, str(TARGET)] + sys.argv[1:]
    raise SystemExit(run_process(command).returncode)


if __name__ == '__main__':
    main()
