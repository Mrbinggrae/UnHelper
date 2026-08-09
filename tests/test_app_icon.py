from __future__ import annotations

import os
import unittest
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from Modules.Common.paths import app_icon_path


class AppIconTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        cls.app = QApplication.instance() or QApplication([])

    def test_project_icon_exists_and_qt_can_load_it(self) -> None:
        icon_path = app_icon_path()

        self.assertEqual(icon_path.name, "app-icon.ico")
        self.assertTrue(icon_path.is_file())

        icon = QIcon(str(icon_path))
        self.assertFalse(icon.isNull())
        sizes = {(size.width(), size.height()) for size in icon.availableSizes()}
        self.assertIn((16, 16), sizes)
        self.assertIn((32, 32), sizes)
        self.assertIn((256, 256), sizes)

    def test_packaging_surfaces_reference_the_same_icon(self) -> None:
        root = Path(__file__).resolve().parents[1]
        spec = (root / "UnHelper.spec").read_text(encoding="utf-8")
        installer = (root / "UnHelper.nsi").read_text(encoding="utf-8")

        self.assertIn('asset_dir / "app-icon.ico"', spec)
        self.assertIn('!define APP_ICON     "assets\\app-icon.ico"', installer)
        self.assertIn('"DisplayIcon"', installer)


if __name__ == "__main__":
    unittest.main()
