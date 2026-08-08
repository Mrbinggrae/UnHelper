from __future__ import annotations

import os
import tempfile
import time
import unittest
from pathlib import Path
from decimal import Decimal
from unittest.mock import patch
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings, QThread
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

from Modules.Common.Credentials import WMSCredentialStore
from Modules.GUI.MainWindow import MainWindow, SettingsDialog, _open_product_memory_with_recovery
from Modules.GUI.ProductMemoryDialog import ProductMemoryDialog
from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.WMS.ProductMemory import ProductMemory


class BriefWorker(QThread):
    def run(self) -> None:
        time.sleep(0.08)


class RunningWorker:
    def isRunning(self) -> bool:
        return True


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
            self.assertEqual(window.raw_table.columnCount(), 10)
            self.assertEqual(window.raw_table.horizontalHeaderItem(0).text(), "거래처 이름")
            self.assertEqual(window.raw_table.horizontalHeaderItem(6).text(), "SKU 명")
            self.assertEqual(window.raw_table.horizontalHeaderItem(9).text(), "분류")
            self.assertFalse(window.raw_table.wordWrap())
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

    def test_settings_persists_dpapi_protected_wms_credentials(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(str(root / "settings.ini"), QSettings.Format.IniFormat)
            dialog = SettingsDialog(settings, memory_path=root / "memory.json")
            try:
                dialog.wms_id_input.setText("worker-id")
                dialog.wms_password_input.setText("secret-password")
                self.assertIsNotNone(dialog._persist())
                credentials = WMSCredentialStore(settings).load()
                self.assertEqual(credentials.wms_id, "worker-id")
                self.assertEqual(credentials.password, "secret-password")
                self.assertFalse(settings.contains("wms_password"))
                self.assertFalse(settings.contains("wms_pw"))
            finally:
                dialog.close()

    def test_milkrun_products_populate_weight_table(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            products = (
                MilkrunProductRow("거래처", "10813478", "10", "720", "72246083", "상품 A"),
                MilkrunProductRow("거래처", "10813478", "10", "720", "72246115", "상품 B"),
            )
            window._populate_milkrun_products(products)

            self.assertEqual(window.raw_table.rowCount(), 2)
            self.assertEqual(window.raw_table.item(0, 0).text(), "거래처")
            self.assertEqual(window.raw_table.item(0, 4).text(), "72")
            self.assertEqual(window.raw_table.item(1, 5).text(), "72246115")
            self.assertEqual(window.raw_table.item(1, 6).text(), "상품 B")
            self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "?")
        finally:
            window.close()

    def test_one_pallet_two_boxes_displays_two_before_wms_and_keeps_manual_category(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            window = MainWindow(smoke_test=True)
            window.product_memory_file = root / "memory.json"
            try:
                product = MilkrunProductRow("거래처", "M1", "1", "2", "123", "상품")
                window._populate_milkrun_products((product,))

                self.assertEqual(window.raw_table.item(0, 4).text(), "2")

                manual = ProductMemory(window.product_memory_file).set_manual_category(
                    "123",
                    "고단",
                    "상품",
                )
                window._on_weight_record_ready(manual, False)

                self.assertEqual(window.raw_table.item(0, 4).text(), "2")
                self.assertEqual(window.raw_table.item(0, 7).text(), "-")
                self.assertEqual(window.raw_table.item(0, 8).text(), "-")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "고단")
            finally:
                window.close()

    def test_empty_today_result_skips_memory_recovery_and_wms_worker(self) -> None:
        window = MainWindow(smoke_test=True)
        window._credential_load_failure = object()
        try:
            with (
                patch("Modules.GUI.MainWindow._open_product_memory_with_recovery") as open_memory,
                patch.object(window, "_finalize_weight_if_ready") as finalize,
            ):
                window._start_weight_lookup(())

            open_memory.assert_not_called()
            finalize.assert_called_once_with()
            self.assertIsNone(window.weight_worker)
            self.assertIsNone(window._credential_load_failure)
            self.assertEqual(window._pending_weight_summary.total_skus, 0)
            self.assertEqual(window.status_label.text(), "오늘 표시할 상품 없음 · WMS 조회 생략")
        finally:
            window.close()

    def test_queued_milkrun_completion_after_close_does_not_start_wms(self) -> None:
        window = MainWindow(smoke_test=True)
        window._closing_after_cancel = True
        try:
            with patch.object(window, "_start_weight_lookup") as start_weight:
                window._on_milkrun_completed(object())
            start_weight.assert_not_called()
            self.assertIsNone(window.weight_worker)
        finally:
            window._closing_after_cancel = False
            window.close()

    def test_cancelled_corrupt_memory_recovery_finalizes_without_signal_order_dependency(self) -> None:
        window = MainWindow(smoke_test=True)
        product = MilkrunProductRow("거래처", "M1", "1", "1", "123", "상품")
        try:
            with (
                patch(
                    "Modules.GUI.MainWindow._open_product_memory_with_recovery",
                    return_value=None,
                ),
                patch.object(window, "_finalize_weight_if_ready") as finalize,
            ):
                window._start_weight_lookup((product,))

            finalize.assert_called_once_with()
            self.assertTrue(window._weight_finalize_pending)
            self.assertIsNotNone(window._pending_weight_failure)
            self.assertIsNone(window.weight_worker)
        finally:
            window.close()

    def test_weight_render_and_single_button_manual_category_are_persisted(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            window = MainWindow(smoke_test=True)
            window.product_memory_file = root / "memory.json"
            try:
                products = (
                    MilkrunProductRow("거래처", "10813478", "2", "100", "123", "상품 A/\n상품 B"),
                    MilkrunProductRow("거래처", "10813478", "2", "100", "123", "상품 A/상품 B"),
                )
                window._populate_milkrun_products(products)
                record = ProductMemory(window.product_memory_file).upsert_measurement(
                    "123", "상품 A/ 상품 B", Decimal("6000"), "100", "2"
                )
                window._on_weight_record_ready(record, False)

                self.assertEqual(window.raw_table.item(0, 6).text(), "상품 A/ 상품 B")
                self.assertEqual(window.raw_table.item(0, 4).text(), "50")
                self.assertEqual(window.raw_table.item(0, 7).text(), "6000")
                self.assertEqual(window.raw_table.item(0, 8).text(), "300")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "중량")

                window.raw_table.cellWidget(0, 9).click()
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "경량")
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "경량")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "경량",
                )

                window.raw_table.cellWidget(0, 9).click()
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "중량")

                window.raw_table.cellWidget(0, 9).click()
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "고단")
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "고단")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "고단",
                )

                window.raw_table.cellWidget(0, 9).click()
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "중량")
                self.assertIsNone(
                    ProductMemory(window.product_memory_file).get("123").category_override
                )
            finally:
                window.close()

    def test_invalid_sku_row_does_not_block_valid_sku_category_edit(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow("거래처", "M1", "1", "1", "BAD/SKU", "잘못된 SKU"),
                    MilkrunProductRow("거래처", "M2", "2", "100", "123", "정상/상품"),
                )
                window._populate_milkrun_products(products)
                record = ProductMemory(window.product_memory_file).upsert_measurement(
                    "123", "정상/상품", Decimal("6000"), "100", "2"
                )
                window._on_weight_record_ready(record, False)

                with patch("Modules.GUI.MainWindow.ErrorReportDialog") as error_dialog:
                    window.raw_table.cellWidget(1, 9).click()

                error_dialog.assert_not_called()
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "경량")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "경량",
                )
            finally:
                window.close()

    def test_weight_worker_blocks_update_dialog_download_and_apply(self) -> None:
        window = MainWindow(smoke_test=True)
        window.weight_worker = RunningWorker()
        try:
            with (
                patch("Modules.GUI.MainWindow.QMessageBox.warning") as warning,
                patch("Modules.Common.AutoUpdater.AutoUpdater.apply_update") as apply_update,
            ):
                window._show_update_dialog(object())
                window._download_update(object())
                window._apply_downloaded_update("update.zip", object())

            self.assertIsNone(window.update_dialog)
            self.assertIsNone(window.update_download_worker)
            self.assertEqual(warning.call_count, 2)
            apply_update.assert_not_called()
        finally:
            window.weight_worker = None
            window.close()

    def test_settings_memory_change_refreshes_current_raw_table(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            window = MainWindow(smoke_test=True)
            window.product_memory_file = root / "memory.json"
            try:
                product = MilkrunProductRow("거래처", "M1", "1", "280", "123", "상품")
                window._populate_milkrun_products((product, product))
                record = ProductMemory(window.product_memory_file).upsert_measurement(
                    "123", "상품", "1000", "280", "1"
                )
                window._on_weight_record_ready(record, False)
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "중량")

                def remove_and_emit(dialog: SettingsDialog) -> int:
                    ProductMemory(window.product_memory_file).delete("123")
                    dialog.product_memory_changed.emit()
                    return 0

                with mock.patch.object(SettingsDialog, "exec", remove_and_emit):
                    window.show_settings()

                self.assertEqual(window.raw_table.item(0, 4).text(), "280")
                self.assertEqual(window.raw_table.item(0, 7).text(), "-")
                self.assertEqual(window.raw_table.item(0, 8).text(), "-")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "?")
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "?")
                self.assertNotIn("123", window._weight_records)
            finally:
                window.close()

    def test_settings_memory_refresh_skips_invalid_sku_and_updates_valid_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow("거래처", "M1", "1", "1", "잘못된 SKU", "오류 행"),
                    MilkrunProductRow("거래처", "M1", "1", "280", "123", "정상 행"),
                )
                window._populate_milkrun_products(products)
                ProductMemory(window.product_memory_file).upsert_measurement(
                    "123", "정상 행", "1000", "280", "1"
                )

                window._refresh_current_product_memory()

                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "중량")
                self.assertIn("123", window._weight_records)
            finally:
                window.close()

    def test_product_memory_dialog_emits_changes_for_import_and_delete(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            destination = ProductMemory(root / "destination.json")
            destination.upsert_measurement("100", "기존", "1000", "100", "1")
            source = ProductMemory(root / "source.json")
            source.upsert_measurement("200", "신규", "1000", "100", "1")
            exported = source.export_to(root / "import.json")
            dialog = ProductMemoryDialog(destination)
            changes = []
            dialog.memory_changed.connect(lambda: changes.append(True))
            try:
                with (
                    mock.patch.object(QFileDialog, "getOpenFileName", return_value=(str(exported), "")),
                    mock.patch.object(QMessageBox, "information"),
                ):
                    dialog._import_records()
                self.assertEqual(len(changes), 1)
                self.assertIsNotNone(destination.get("200"))

                dialog.table.selectRow(0)
                with mock.patch.object(
                    QMessageBox,
                    "question",
                    return_value=QMessageBox.StandardButton.Yes,
                ):
                    dialog._delete_selected()
                self.assertEqual(len(changes), 2)
            finally:
                dialog.close()

    def test_corrupt_memory_recovery_requires_confirmation_and_preserves_backup(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "memory.json"
            original = b"{broken-json"
            path.write_bytes(original)

            with mock.patch.object(
                QMessageBox,
                "question",
                return_value=QMessageBox.StandardButton.No,
            ) as question:
                self.assertIsNone(_open_product_memory_with_recovery(path, None))
            self.assertEqual(path.read_bytes(), original)
            self.assertIn("모두 초기화", question.call_args.args[2])
            self.assertIn("백업", question.call_args.args[2])

            with (
                mock.patch.object(
                    QMessageBox,
                    "question",
                    return_value=QMessageBox.StandardButton.Yes,
                ),
                mock.patch.object(QMessageBox, "information"),
            ):
                recovered = _open_product_memory_with_recovery(path, None)

            self.assertIsNotNone(recovered)
            self.assertEqual(recovered.entries(), ())
            backups = list(path.parent.glob("memory.corrupt_*.json"))
            self.assertEqual(len(backups), 1)
            self.assertEqual(backups[0].read_bytes(), original)

    def test_settings_recovery_close_refreshes_current_raw_table(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            window = MainWindow(smoke_test=True)
            window.product_memory_file = root / "memory.json"
            try:
                product = MilkrunProductRow("거래처", "M1", "1", "280", "123", "상품")
                window._populate_milkrun_products((product,))
                record = ProductMemory(window.product_memory_file).upsert_measurement(
                    "123", "상품", "1000", "280", "1"
                )
                window._on_weight_record_ready(record, False)
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "중량")

                window.product_memory_file.write_bytes(b"{broken-json")

                def open_memory_list(dialog: SettingsDialog) -> int:
                    dialog._show_product_memory()
                    return 0

                memory_dialog = mock.Mock()
                with (
                    mock.patch.object(SettingsDialog, "exec", open_memory_list),
                    mock.patch(
                        "Modules.GUI.MainWindow.ProductMemoryDialog",
                        return_value=memory_dialog,
                    ),
                    mock.patch.object(
                        QMessageBox,
                        "question",
                        return_value=QMessageBox.StandardButton.Yes,
                    ),
                    mock.patch.object(QMessageBox, "information"),
                ):
                    window.show_settings()

                self.assertEqual(window.raw_table.item(0, 4).text(), "280")
                self.assertEqual(window.raw_table.item(0, 7).text(), "-")
                self.assertEqual(window.raw_table.item(0, 8).text(), "-")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "?")
                self.assertNotIn("123", window._weight_records)
                backups = list(root.glob("memory.corrupt_*.json"))
                self.assertEqual(len(backups), 1)
                self.assertEqual(backups[0].read_bytes(), b"{broken-json")
            finally:
                window.close()


if __name__ == "__main__":
    unittest.main()
