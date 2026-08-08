from __future__ import annotations

import sys
from pathlib import Path


APP_NAME = "UnHelper"


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def bundled_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return project_root()


def chromedriver_path() -> Path:
    return bundled_root() / "chromedriver.exe"


def default_download_dir() -> Path:
    path = Path.home() / "Downloads" / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


def release_manifest_path() -> Path:
    prefix = APP_NAME.replace(" ", "_")
    if getattr(sys, "frozen", False):
        return project_root() / f"{prefix}_manifest.json"
    return project_root() / "release" / f"{prefix}_manifest.json"
