from __future__ import annotations

import argparse
import ctypes
import sys

from PySide6.QtCore import QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from Modules.Common.paths import app_icon_path
from Modules.GUI.MainWindow import MainWindow
from Modules.GUI.Theme import apply_dark_theme


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="UnHelper")
    parser.add_argument("--smoke-test", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args(argv)


def set_windows_app_identity() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Mrbinggrae.UnHelper"
        )
    except (AttributeError, OSError):
        pass


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    set_windows_app_identity()
    app = QApplication(sys.argv[:1])
    app.setApplicationName("UnHelper")
    app.setOrganizationName("Mrbinggrae")
    icon_path = app_icon_path()
    if icon_path.is_file():
        app.setWindowIcon(QIcon(str(icon_path)))
    apply_dark_theme(app)
    window = MainWindow(smoke_test=args.smoke_test)
    if args.smoke_test:
        window.show()
        QTimer.singleShot(250, app.quit)
    else:
        window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
