from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from PySide6.QtCore import QSettings
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QApplication, QFrame, QPushButton, QScrollArea

from Modules.Common.ErrorReport import FailureDetails
from Modules.GUI.Dialogs import ErrorReportDialog, UpdateHistoryDialog
from Modules.GUI.MainWindow import MainWindow, SettingsDialog
from Modules.GUI.ProductMemoryDialog import ProductMemoryDialog
from Modules.GUI.Theme import APP_STYLESHEET, COLORS, apply_dark_theme
from Modules.WMS.ProductMemory import ProductMemory


class DarkThemeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])
        apply_dark_theme(cls.app)

    def test_palette_and_stylesheet_use_dark_layered_surfaces(self) -> None:
        self.assertEqual(
            self.app.palette().color(QPalette.ColorRole.Window).name().upper(),
            COLORS["canvas"].upper(),
        )
        self.assertEqual(self.app.styleSheet(), APP_STYLESHEET)
        self.assertIn("QTabBar#MainTabBar::tab", APP_STYLESHEET)
        self.assertIn("QTabBar#SubTabBar::tab", APP_STYLESHEET)
        self.assertIn("QScrollBar:vertical", APP_STYLESHEET)
        self.assertIn("QComboBox QAbstractItemView", APP_STYLESHEET)
        self.assertIn("QDateEdit", APP_STYLESHEET)
        self.assertNotIn("#DDE1E5", APP_STYLESHEET.upper())
        self.assertGreaterEqual(self._contrast(COLORS["primary"], "#FFFFFF"), 4.5)
        self.assertGreaterEqual(self._contrast(COLORS["primary_hover"], "#FFFFFF"), 4.5)

    def test_main_window_uses_semantic_roles_and_renders(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            window.show()
            self.app.processEvents()

            self.assertEqual(window.main_tabs.tabBar().objectName(), "MainTabBar")
            self.assertEqual(window.raw_tabs.tabBar().objectName(), "SubTabBar")
            self.assertEqual(window.raw_table.objectName(), "RawTable")
            self.assertEqual(window.log_view.objectName(), "LogView")
            self.assertEqual(window.settings_button.objectName(), "SettingsButton")
            self.assertEqual(window.settings_button.height(), 40)
            self.assertIsNotNone(window.findChild(QFrame, "BaseDatePanel"))
            self.assertEqual(window.base_date_mode.objectName(), "BaseDateMode")
            self.assertEqual(window.manual_base_date.objectName(), "ManualBaseDate")
            self.assertIsNotNone(window.findChild(QFrame, "DataCard"))
            self.assertFalse(window.grab().isNull())
        finally:
            window.close()
            self.app.processEvents()

    def test_dialogs_share_cards_and_semantic_roles(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(str(root / "settings.ini"), QSettings.Format.IniFormat)
            settings_dialog = SettingsDialog(settings, memory_path=root / "memory.json")
            settings_dialog.resize(700, 600)
            memory_dialog = ProductMemoryDialog(ProductMemory(root / "memory.json"))
            error_dialog = ErrorReportDialog(
                "테마 확인",
                FailureDetails("요약", "상세"),
            )
            history_path = root / "history.txt"
            history_path.write_text("## [v0.0.0] 테마 확인", encoding="utf-8")
            history_dialog = UpdateHistoryDialog(history_path=history_path)
            dialogs = (settings_dialog, memory_dialog, error_dialog, history_dialog)
            try:
                for dialog in dialogs:
                    dialog.show()
                    self.app.processEvents()
                    self.assertFalse(dialog.grab().isNull())

                cards = [
                    frame
                    for frame in settings_dialog.findChildren(QFrame)
                    if frame.property("card") is True
                ]
                self.assertEqual(len(cards), 3)
                self.assertEqual(error_dialog.summary_label.objectName(), "ErrorSummary")
                self.assertEqual(error_dialog.report_button.objectName(), "PrimaryButton")
                self.assertEqual(memory_dialog.table.objectName(), "StoredProductTable")
                self.assertEqual(history_dialog.history_view.objectName(), "DocumentView")
                settings_scroll = settings_dialog.findChild(QScrollArea, "SettingsScroll")
                self.assertIsNotNone(settings_scroll)
                self.assertGreater(settings_scroll.verticalScrollBar().maximum(), 0)
                save_button = next(
                    button
                    for button in settings_dialog.findChildren(QPushButton)
                    if button.text() == "저장하고 닫기"
                )
                save_button_bottom = save_button.mapTo(
                    settings_dialog,
                    save_button.rect().bottomRight(),
                )
                self.assertTrue(settings_dialog.rect().contains(save_button_bottom))
            finally:
                for dialog in dialogs:
                    dialog.close()
                self.app.processEvents()

    @staticmethod
    def _contrast(first: str, second: str) -> float:
        def luminance(value: str) -> float:
            channels = [int(value[index : index + 2], 16) / 255 for index in (1, 3, 5)]
            linear = [
                channel / 12.92
                if channel <= 0.04045
                else ((channel + 0.055) / 1.055) ** 2.4
                for channel in channels
            ]
            return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

        brighter, darker = sorted(
            (luminance(first), luminance(second)),
            reverse=True,
        )
        return (brighter + 0.05) / (darker + 0.05)


if __name__ == "__main__":
    unittest.main()
