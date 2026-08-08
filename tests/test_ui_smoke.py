from __future__ import annotations

import os
import time
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QThread
from PySide6.QtWidgets import QApplication

from Modules.GUI.MainWindow import MainWindow


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
            self.assertEqual(window.raw_table.horizontalHeaderItem(0).text(), "거래처")
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


if __name__ == "__main__":
    unittest.main()
