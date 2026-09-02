"""sitecustomize — make bundled native rtlsdr DLLs findable at interpreter start.

pyrtlsdr's ctypes loader needs 'rtlsdr.dll' (+ msvcr100/pthreadVC2) on its
search path. We prepend python/rtlsdr_libs/ to both os.environ['PATH'] and
sys.path so CDLL() finds it regardless of how the interpreter was started
(uv run, python -c, pytest, etc.). Self-contained per-repo; no admin needed.

This file lives at python/sitecustomize.py and is picked up because the
python/ dir (the interpreter's cwd) is on sys.path at startup.
"""
import os
import sys
from pathlib import Path

_LIBS = Path(__file__).resolve().parent / "rtlsdr_libs"
if _LIBS.is_dir():
    _libs_str = str(_LIBS)
    if _libs_str not in sys.path:
        sys.path.insert(0, _libs_str)
    os.environ["PATH"] = _libs_str + os.pathsep + os.environ.get("PATH", "")
    if hasattr(os, "add_dll_directory"):
        try:
            os.add_dll_directory(_libs_str)
        except Exception:
            pass