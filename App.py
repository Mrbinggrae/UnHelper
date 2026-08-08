from __future__ import annotations

import argparse
import sys

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from Modules.GUI.MainWindow import MainWindow
from Modules.GUI.Theme import apply_dark_theme


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="UnHelper")
    parser.add_argument("--smoke-test", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    app = QApplication(sys.argv[:1])
    app.setApplicationName("UnHelper")
    app.setOrganizationName("Mrbinggrae")
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
