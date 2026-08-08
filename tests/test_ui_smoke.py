from __future__ import annotations

import os
import tempfile
import time
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings, QThread
from PySide6.QtWidgets import QApplication

from Modules.GUI.MainWindow import MainWindow, SettingsDialog
from Modules.Shipments.DailyInbound import MilkrunProductRow


class BriefWorker(QThread):
    def run(self) -> None:
        time.sleep(0.08)


class MainWindowSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_raw_milkrun_is_default_and_button_exists(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            self.assertEqual(window.main_tabs.tabText(window.main_tabs.currentIndex()), "RAW")
            self.assertEqual(window.raw_tabs.tabText(window.raw_tabs.currentIndex()), "Milkrun")
            self.assertEqual(window.get_data_button.text(), "데이터 얻기")
            self.assertEqual(window.raw_table.columnCount(), 6)
            self.assertEqual(window.raw_table.horizontalHeaderItem(0).text(), "거래처 이름")
            self.assertEqual(window.raw_table.horizontalHeaderItem(5).text(), "SKU 명")
        finally:
            window.close()

    def test_close_waits_for_update_worker(self) -> None:
        window = MainWindow(smoke_test=True)
        worker = BriefWorker()
        window.update_check_worker = worker
        worker.finished.connect(lambda: window._on_worker_finished("update_check_worker", worker))
        window.show()
        worker.start()
        self.app.processEvents()

        window.close()
        self.assertTrue(window._closing_after_workers)

        self.assertTrue(worker.wait(2000))
        for _ in range(5):
            self.app.processEvents()
            time.sleep(0.01)
        self.assertFalse(window._update_worker_running())
        window.close()

    def test_settings_persists_linked_milkrun_excel_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            target = root / "입고스케줄관리.xlsx"
            target.write_bytes(b"placeholder")
            settings = QSettings(str(root / "settings.ini"), QSettings.Format.IniFormat)
            dialog = SettingsDialog(settings)
            try:
                self.assertTrue(dialog.excel_path.isReadOnly())
                dialog.excel_path.setText(str(target))
                self.assertIsNotNone(dialog._persist())
                self.assertEqual(
                    Path(str(settings.value("milkrun_excel_path"))),
                    target.resolve(),
                )
            finally:
                dialog.close()

    def test_milkrun_products_populate_six_column_table(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            products = (
                MilkrunProductRow("거래처", "10813478", "10", "720", "72246083", "상품 A"),
                MilkrunProductRow("거래처", "10813478", "10", "720", "72246115", "상품 B"),
            )
            window._populate_milkrun_products(products)

            self.assertEqual(window.raw_table.rowCount(), 2)
            self.assertEqual(window.raw_table.item(0, 0).text(), "거래처")
            self.assertEqual(window.raw_table.item(1, 4).text(), "72246115")
            self.assertEqual(window.raw_table.item(1, 5).text(), "상품 B")
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
