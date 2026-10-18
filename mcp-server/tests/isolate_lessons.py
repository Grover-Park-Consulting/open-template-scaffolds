"""Keep the general tests away from the developer's real lesson store.

Importing this module points the store at a temporary folder and stops the sync from starting
PowerShell, so a test that fetches a template's `lessons` part reads the library's own lessons
from the text files and writes nothing outside the temporary folder. `REAL_POWERSHELLS` is the
original lookup, for the tests of the sync itself.
"""
import atexit
import os
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import lessons_sync  # noqa: E402

REAL_POWERSHELLS = getattr(lessons_sync, "REAL_POWERSHELLS", lessons_sync._powershells)

if not getattr(lessons_sync, "REAL_POWERSHELLS", None):
    _store = tempfile.mkdtemp(prefix="ots-lessons-test-")
    atexit.register(shutil.rmtree, _store, True)
    os.environ["OTS_LESSONS_STORE"] = _store
    lessons_sync.REAL_POWERSHELLS = lessons_sync._powershells
    lessons_sync._powershells = lambda: []
