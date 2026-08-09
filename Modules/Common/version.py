from __future__ import annotations

import re
import sys
from pathlib import Path


GITHUB_OWNER = "Mrbinggrae"
GITHUB_REPO = "UnHelper"
_FALLBACK_VERSION = "1.0.4"


def _history_candidates() -> list[Path]:
    if getattr(sys, "frozen", False):
        executable_dir = Path(sys.executable).resolve().parent
        meipass = Path(getattr(sys, "_MEIPASS", executable_dir))
        return [
            meipass / "UPDATE_HISTORY.txt",
            executable_dir / "_internal" / "UPDATE_HISTORY.txt",
            executable_dir / "UPDATE_HISTORY.txt",
        ]
    return [Path(__file__).resolve().parents[2] / "UPDATE_HISTORY.txt"]


def read_current_version() -> str:
    for history_path in _history_candidates():
        if not history_path.is_file():
            continue
        try:
            with history_path.open("r", encoding="utf-8") as handle:
                for line in handle:
                    match = re.search(r"\[v(\d+(?:\.\d+)+)\]", line)
                    if match:
                        return match.group(1)
        except OSError:
            continue
    return _FALLBACK_VERSION


CURRENT_VERSION = read_current_version()
