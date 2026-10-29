"""Product entry point; implementation is compiled/executed from C-00-Shared."""
from pathlib import Path as _Path
import sys as _sys
_root = _Path(__file__).resolve().parents[2]
_sys.dont_write_bytecode = True
_sys.path.insert(0, str(_root))
_sys.path.insert(1, str(_root.parent / "C-00-Shared"))
from source_layout import run_entry as _run_entry
_run_entry(globals(), 'src/pdf/verify_workflow.py', _root)
