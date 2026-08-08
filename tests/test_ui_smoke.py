from __future__ import annotations

import os
import tempfile
import time
import unittest
from datetime import date
from pathlib import Path
from decimal import Decimal
from unittest.mock import patch
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QDate, QSettings, QThread
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

from Modules.Common.Credentials import WMSCredentialStore
from Modules.Common.ErrorReport import FailureDetails
from Modules.GUI.MainWindow import MainWindow, SettingsDialog, _open_product_memory_with_recovery
from Modules.GUI.ProductMemoryDialog import ProductMemoryDialog
from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.Shipments.MilkrunDownloader import MilkrunDownloadRequest
from Modules.Shipments.TruckDownloader import TruckDownloadRequest
from Modules.WMS.ProductMemory import ProductMemory
from Modules.WMS.ProductWeightWorker import ProductWeightSummary, SkuWeightFailure


class BriefWorker(QThread):
    def run(self) -> None:
        time.sleep(0.08)


class RunningWorker:
    def isRunning(self) -> bool:
        return True


class FinishedWorker:
    def __init__(self) -> None:
        self.cancel_requests = 0

    def isRunning(self) -> bool:
        return False

    def request_cancel(self) -> None:
        self.cancel_requests += 1

    def deleteLater(self) -> None:
        pass


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

    def test_raw_truck_has_functional_table_and_independent_state(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            self.assertEqual(window.raw_tabs.tabText(0), "트럭")
            self.assertEqual(window.truck_get_data_button.text(), "데이터 얻기")
            self.assertEqual(window.truck_table.horizontalHeaderItem(1).text(), "예약번호")
            self.assertEqual(window.truck_table.horizontalHeaderItem(3).text(), "유닛 수")
            self.assertEqual(window.truck_table.horizontalHeaderItem(4).text(), "팔렛트당 유닛")

            milkrun = MilkrunProductRow(
                "밀크런 거래처", "10813478", "1", "2", "100", "밀크런 상품", "M1"
            )
            truck = MilkrunProductRow(
                "트럭 거래처", "상세 번호", Decimal("1"), Decimal("2"),
                "200", "트럭 상품", "T3372829"
            )
            window._populate_milkrun_products((milkrun,))
            window._populate_truck_products((truck,))

            self.assertEqual(window.raw_table.rowCount(), 1)
            self.assertEqual(window.raw_table.item(0, 0).text(), "밀크런 거래처")
            self.assertEqual(window.truck_table.rowCount(), 1)
            self.assertEqual(window.truck_table.item(0, 1).text(), "T3372829")
            self.assertEqual(window.truck_table.item(0, 3).text(), "2")
            self.assertEqual(window.truck_table.item(0, 4).text(), "2")
            self.assertGreaterEqual(window.truck_table.columnWidth(9), 84)
        finally:
            window.close()

    def test_multi_sku_truck_merges_vehicle_identity_and_keeps_sku_calculations(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow(
                        "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                        "123", "상품 A", "T3372829",
                    ),
                    MilkrunProductRow(
                        "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                        "456", "상품 B", "T3372829",
                    ),
                )
                window._populate_truck_products(products)
                memory = ProductMemory(window.product_memory_file)
                first = memory.upsert_measurement("123", "상품 A", "1000", "20", "2")
                first = memory.set_manual_category("123", "고단")
                second = memory.upsert_measurement("456", "상품 B", "2000", "20", "2")
                second = memory.set_manual_category("456", "중량")

                window._render_weight_record(first, "truck")
                window._render_weight_record(second, "truck")

                self.assertEqual(window.truck_table.rowSpan(0, 0), 2)
                self.assertEqual(window.truck_table.rowSpan(0, 1), 2)
                self.assertEqual(window.truck_table.rowSpan(0, 2), 1)
                self.assertEqual(window.truck_table.rowSpan(0, 9), 1)
                self.assertEqual(window.truck_table.item(0, 1).text(), "T3372829")
                self.assertEqual(window.truck_table.item(1, 1).text(), "T3372829")
                self.assertEqual(window.truck_table.item(0, 7).text(), "1000")
                self.assertEqual(window.truck_table.item(1, 7).text(), "2000")
                self.assertEqual(window.truck_table.item(0, 8).text(), "10")
                self.assertEqual(window.truck_table.item(1, 8).text(), "20")
                first_button = window.truck_table.cellWidget(0, 9)
                second_button = window.truck_table.cellWidget(1, 9)
                self.assertEqual(first_button.text(), "고단")
                self.assertEqual(second_button.text(), "중량")
                self.assertEqual(
                    window.truck_table.item(0, 3).background().color(),
                    window.truck_table.item(1, 3).background().color(),
                )

                first_button.click()

                self.assertEqual(first_button.text(), "경량")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    None,
                )
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("456").category_override,
                    "중량",
                )

                window._populate_truck_products((products[0],))
                self.assertEqual(window._truck_group_categories, {})
                self.assertEqual(window.truck_table.rowSpan(0, 0), 1)
                self.assertEqual(window.truck_table.rowSpan(0, 1), 1)
                self.assertEqual(window.truck_table.rowSpan(0, 9), 1)
            finally:
                window.close()

    def test_multi_sku_milkrun_merges_identity_and_persists_each_sku_category(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow(
                        "거래처 A", "10813478", "2", "20", "123", "상품 A", "M3370492"
                    ),
                    MilkrunProductRow(
                        "거래처 B", "10799314", "1", "5", "789", "다른 밀크런", "M3370492"
                    ),
                    MilkrunProductRow(
                        "거래처 A", "10813478", "2", "20", "456", "상품 B", "M3370492"
                    ),
                )
                window._populate_milkrun_products(products)
                memory = ProductMemory(window.product_memory_file)
                first = memory.upsert_measurement("123", "상품 A", "1000", "20", "2")
                first = memory.set_manual_category("123", "고단")
                second = memory.upsert_measurement("456", "상품 B", "2000", "20", "2")
                second = memory.set_manual_category("456", "중량")

                window._render_weight_record(first, "milkrun")
                window._render_weight_record(second, "milkrun")

                self.assertEqual(
                    tuple(window.raw_table.item(row, 1).text() for row in range(3)),
                    ("10813478", "10813478", "10799314"),
                )
                self.assertEqual(window.raw_table.rowSpan(0, 0), 2)
                self.assertEqual(window.raw_table.rowSpan(0, 1), 2)
                self.assertEqual(window.raw_table.rowSpan(0, 2), 1)
                self.assertEqual(window.raw_table.rowSpan(0, 9), 1)
                self.assertEqual(window.raw_table.item(0, 7).text(), "1000")
                self.assertEqual(window.raw_table.item(1, 7).text(), "2000")
                self.assertEqual(window.raw_table.item(0, 8).text(), "10")
                self.assertEqual(window.raw_table.item(1, 8).text(), "20")
                first_button = window.raw_table.cellWidget(0, 9)
                second_button = window.raw_table.cellWidget(1, 9)
                self.assertEqual(first_button.text(), "고단")
                self.assertEqual(second_button.text(), "중량")
                self.assertIsNotNone(window.raw_table.cellWidget(2, 9))
                self.assertEqual(
                    window.raw_table.item(0, 3).background().color(),
                    window.raw_table.item(1, 3).background().color(),
                )
                self.assertNotEqual(
                    window.raw_table.item(0, 3).background().color(),
                    window.raw_table.item(2, 3).background().color(),
                )

                first_button.click()

                self.assertEqual(first_button.text(), "경량")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    None,
                )
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("456").category_override,
                    "중량",
                )

                window._refresh_current_product_memory()
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "경량")

                window._populate_milkrun_products((products[0],))
                self.assertEqual(window._milkrun_group_categories, {})
                self.assertEqual(window.raw_table.rowSpan(0, 0), 1)
                self.assertEqual(window.raw_table.rowSpan(0, 1), 1)
                self.assertEqual(window.raw_table.rowSpan(0, 9), 1)
            finally:
                window.close()

    def test_multi_sku_milkrun_weight_refresh_replaces_stale_button_error(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow("거래처", "10813478", "2", "20", "123", "상품 A"),
                    MilkrunProductRow("거래처", "10813478", "2", "20", "456", "상품 B"),
                )
                window._populate_milkrun_products(products)
                failure = SkuWeightFailure(
                    sku_id="123",
                    product_name="상품 A",
                    details=FailureDetails(summary="이전 WMS 실패", detail="이전 WMS 실패"),
                )

                window._on_weight_sku_failed(failure)
                self.assertEqual(
                    window.raw_table.cellWidget(0, 9).toolTip(),
                    "이전 WMS 실패",
                )

                record = ProductMemory(window.product_memory_file).upsert_weight_only(
                    "123", "상품 A", "1000"
                )
                window._render_weight_record(record, "milkrun")
                self.assertEqual(window.raw_table.item(0, 7).text(), "1000")
                self.assertNotIn(
                    "이전 WMS 실패",
                    window.raw_table.cellWidget(0, 9).toolTip(),
                )

                window._on_weight_sku_failed(failure)
                window._render_unknown_sku("123", "milkrun")
                self.assertEqual(window.raw_table.item(0, 7).text(), "-")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "?")
            finally:
                window.close()

    def test_sku_shared_by_truck_groups_uses_saved_category_and_row_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow(
                        "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                        "123", "공유 상품", "T11111",
                    ),
                    MilkrunProductRow(
                        "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                        "456", "다중 상품", "T11111",
                    ),
                    MilkrunProductRow(
                        "트럭 거래처", "상세", Decimal("1"), Decimal("2"),
                        "123", "공유 상품", "T22222",
                    ),
                )
                window._populate_truck_products(products)
                memory = ProductMemory(window.product_memory_file)
                shared = memory.upsert_measurement("123", "공유 상품", "1000", "300", "1")
                shared = memory.set_manual_category("123", "중량")

                window._render_weight_record(shared, "truck")

                self.assertEqual(window.truck_table.item(0, 8).text(), "10")
                self.assertEqual(window.truck_table.cellWidget(0, 9).text(), "중량")
                self.assertEqual(window.truck_table.item(2, 8).text(), "2")
                self.assertEqual(window.truck_table.cellWidget(2, 9).text(), "중량")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "중량",
                )
            finally:
                window.close()

    def test_sku_shared_by_milkrun_groups_uses_saved_category_and_row_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow("거래처", "10813478", "2", "20", "123", "공유 상품", "M1"),
                    MilkrunProductRow("거래처", "10813478", "2", "20", "456", "다중 상품", "M1"),
                    MilkrunProductRow("거래처", "10799314", "1", "2", "123", "공유 상품", "M1"),
                )
                window._populate_milkrun_products(products)
                memory = ProductMemory(window.product_memory_file)
                shared = memory.upsert_measurement("123", "공유 상품", "1000", "300", "1")
                shared = memory.set_manual_category("123", "중량")

                window._render_weight_record(shared, "milkrun")

                self.assertEqual(window.raw_table.item(0, 8).text(), "10")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "중량")
                self.assertEqual(window.raw_table.item(2, 8).text(), "2")
                self.assertEqual(window.raw_table.cellWidget(2, 9).text(), "중량")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "중량",
                )
            finally:
                window.close()

    def test_multi_sku_truck_failure_is_shown_on_each_sku_button(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            products = (
                MilkrunProductRow(
                    "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                    "123", "상품 A", "T3372829",
                ),
                MilkrunProductRow(
                    "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                    "456", "상품 B", "T3372829",
                ),
            )
            window._populate_truck_products(products)

            for row_index, sku_id in enumerate(("123", "456")):
                summary = f"SKU {sku_id} 조회 실패"
                window._on_weight_sku_failed(
                    SkuWeightFailure(
                        sku_id=sku_id,
                        product_name=f"상품 {sku_id}",
                        details=FailureDetails(summary=summary, detail=summary),
                    )
                )
                self.assertEqual(
                    window.truck_table.cellWidget(row_index, 9).toolTip(),
                    summary,
                )

            self.assertEqual(window.truck_table.cellWidget(0, 9).text(), "?")
            self.assertEqual(window.truck_table.cellWidget(1, 9).text(), "?")
        finally:
            window.close()

    def test_multi_sku_truck_completion_uses_individual_classification(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            products = (
                MilkrunProductRow(
                    "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                    "123", "상품 A", "T3372829",
                ),
                MilkrunProductRow(
                    "트럭 거래처", "상세", Decimal("2"), Decimal("20"),
                    "456", "상품 B", "T3372829",
                ),
            )
            window._populate_truck_products(products)
            window._pending_weight_summary = ProductWeightSummary(
                total_skus=2,
                cache_hits=2,
                wms_successes=0,
                failures=(),
            )

            with patch.object(QMessageBox, "information") as information:
                window._finalize_weight_lookup()

            message = information.call_args.args[2]
            self.assertIn("WMS 무게 분류를 완료했습니다.", message)
            self.assertNotIn("수동 분류", message)
        finally:
            window.close()

    def test_multi_sku_milkrun_completion_uses_individual_classification(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            products = (
                MilkrunProductRow("거래처", "10813478", "2", "20", "123", "상품 A"),
                MilkrunProductRow("거래처", "10813478", "2", "20", "456", "상품 B"),
            )
            window._populate_milkrun_products(products)
            window._pending_weight_summary = ProductWeightSummary(
                total_skus=2,
                cache_hits=2,
                wms_successes=0,
                failures=(),
            )

            with patch.object(QMessageBox, "information") as information:
                window._finalize_weight_lookup()

            message = information.call_args.args[2]
            self.assertIn("WMS 무게 분류를 완료했습니다.", message)
            self.assertNotIn("수동 분류", message)
        finally:
            window.close()

    def test_truck_weight_uses_units_per_pallet_and_280kg_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            try:
                product = MilkrunProductRow(
                    "트럭 거래처",
                    "상세 번호",
                    Decimal("1"),
                    Decimal("2"),
                    "56913939",
                    "상품",
                    "T3372829",
                )
                window._populate_truck_products((product,))
                record = ProductMemory(Path(temp) / "memory.json").upsert_measurement(
                    "56913939",
                    "상품",
                    Decimal("140000"),
                    Decimal("2"),
                    Decimal("1"),
                )

                window._render_weight_record(record, "truck")

                self.assertEqual(window.truck_table.item(0, 4).text(), "2")
                self.assertEqual(window.truck_table.item(0, 8).text(), "280")
                self.assertEqual(window.truck_table.cellWidget(0, 9).text(), "중량")
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

    def test_finished_booking_reference_blocks_close_until_queued_completion_is_drained(self) -> None:
        window = MainWindow(smoke_test=True)
        window.milkrun_worker = FinishedWorker()
        try:
            self.assertTrue(window._automation_worker_running())
            with patch.object(
                QMessageBox,
                "question",
                return_value=QMessageBox.StandardButton.Yes,
            ):
                window.close()
            self.assertTrue(window._closing_after_cancel)

            with patch.object(window, "_start_weight_lookup") as start_weight:
                window._on_milkrun_completed(object())
            start_weight.assert_not_called()
        finally:
            window.milkrun_worker = None
            window._closing_after_cancel = False
            window.close()

    def test_stop_intent_blocks_queued_booking_completion_from_starting_wms(self) -> None:
        window = MainWindow(smoke_test=True)
        worker = FinishedWorker()
        window.milkrun_worker = worker
        try:
            window.cancel_milkrun_download()

            self.assertEqual(worker.cancel_requests, 1)
            self.assertTrue(window._automation_cancel_requested)
            self.assertEqual(window.status_label.text(), "작업 중지 중...")
            with patch.object(window, "_start_weight_lookup") as start_weight:
                window._on_milkrun_completed(object())
            start_weight.assert_not_called()
            self.assertEqual(window.status_label.text(), "작업 취소됨")
        finally:
            window.milkrun_worker = None
            window._automation_cancel_requested = False
            window.close()

    def test_stop_intent_converts_queued_weight_completion_to_cancelled(self) -> None:
        window = MainWindow(smoke_test=True)
        worker = FinishedWorker()
        window.weight_worker = worker
        try:
            window.cancel_milkrun_download()
            window._on_weight_completed(
                ProductWeightSummary(
                    total_skus=1,
                    cache_hits=1,
                    wms_successes=0,
                    failures=(),
                )
            )

            self.assertEqual(worker.cancel_requests, 1)
            self.assertTrue(window._automation_cancel_requested)
            self.assertIsNone(window._pending_weight_summary)
            self.assertIn("중지", window._pending_weight_cancel)
            with patch.object(QMessageBox, "information") as information:
                window._on_weight_finished()
            self.assertIsNone(window.weight_worker)
            self.assertEqual(information.call_args.args[1], "WMS 무게 조회 취소")
        finally:
            window.weight_worker = None
            window._automation_cancel_requested = False
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
                self.assertFalse(hasattr(dialog, "base_date_mode"))
                self.assertFalse(hasattr(dialog, "manual_base_date"))
                dialog.excel_path.setText(str(target))
                self.assertIsNotNone(dialog._persist())
                self.assertEqual(
                    Path(str(settings.value("milkrun_excel_path"))),
                    target.resolve(),
                )
            finally:
                dialog.close()

    def test_main_window_persists_auto_or_manual_base_date(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(str(root / "settings.ini"), QSettings.Format.IniFormat)
            window = MainWindow(smoke_test=True, settings=settings)
            try:
                self.assertEqual(window.base_date_mode.currentData(), "auto")
                self.assertFalse(window.manual_base_date.isEnabled())
                self.assertIsNone(window._configured_base_date())

                window.base_date_mode.setCurrentIndex(1)
                window.manual_base_date.setDate(QDate(2026, 8, 8))
                self.assertTrue(window.manual_base_date.isEnabled())

                self.assertEqual(settings.value("base_date_mode"), "manual")
                self.assertEqual(settings.value("manual_base_date"), "2026-08-08")
                self.assertEqual(window._configured_base_date(), date(2026, 8, 8))

                window._set_automation_working(True)
                self.assertFalse(window.base_date_mode.isEnabled())
                self.assertFalse(window.manual_base_date.isEnabled())
                window._set_automation_working(False)
                self.assertTrue(window.base_date_mode.isEnabled())
                self.assertTrue(window.manual_base_date.isEnabled())

                window.base_date_mode.setCurrentIndex(0)
                self.assertEqual(settings.value("base_date_mode"), "auto")
                self.assertIsNone(window._configured_base_date())
            finally:
                window.close()

            reopened = MainWindow(smoke_test=True, settings=settings)
            try:
                self.assertEqual(reopened.base_date_mode.currentData(), "auto")
                self.assertEqual(reopened.manual_base_date.date(), QDate(2026, 8, 8))
            finally:
                reopened.close()

    def test_manual_base_date_reaches_milkrun_and_truck_start_button_requests(self) -> None:
        cases = (
            (
                "milkrun",
                "get_data_button",
                MilkrunDownloadRequest,
                "Modules.GUI.MainWindow.MilkrunExcelImporter.validate_target_path",
            ),
            (
                "truck",
                "truck_get_data_button",
                TruckDownloadRequest,
                "Modules.GUI.MainWindow.TruckExcelImporter.validate_target_path",
            ),
        )
        for booking_type, button_name, request_type, validator_path in cases:
            with self.subTest(booking_type=booking_type), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                driver = root / "chromedriver.exe"
                driver.write_bytes(b"driver")
                workbook = root / "SAN2_입고스케줄관리.xlsx"
                workbook.write_bytes(b"workbook")
                settings = QSettings(
                    str(root / "settings.ini"),
                    QSettings.Format.IniFormat,
                )
                settings.setValue("download_dir", str(root / "downloads"))
                settings.setValue("milkrun_excel_path", str(workbook))
                settings.setValue("base_date_mode", "manual")
                settings.setValue("manual_base_date", "2026-08-08")
                settings.sync()

                window = MainWindow(smoke_test=True, settings=settings)
                fake_worker = mock.Mock()
                try:
                    with (
                        patch(
                            "Modules.GUI.MainWindow.chromedriver_path",
                            return_value=driver,
                        ),
                        patch(validator_path, return_value=workbook),
                        patch(
                            "Modules.GUI.MainWindow.MilkrunWorker",
                            return_value=fake_worker,
                        ) as worker_class,
                    ):
                        getattr(window, button_name).click()

                    request = worker_class.call_args.args[0]
                    self.assertIsInstance(request, request_type)
                    self.assertEqual(request.base_date, date(2026, 8, 8))
                    self.assertEqual(
                        worker_class.call_args.kwargs["booking_type"],
                        booking_type,
                    )
                    fake_worker.start.assert_called_once_with()
                finally:
                    window.milkrun_worker = None
                    window._closing_after_cancel = False
                    window.close()

    def test_corrupt_base_date_mode_or_value_warns_without_starting_worker(self) -> None:
        cases = (
            (
                "invalid mode",
                "get_data_button",
                "broken-mode",
                "2026-08-08",
                "Modules.GUI.MainWindow.MilkrunExcelImporter.validate_target_path",
            ),
            (
                "invalid date",
                "truck_get_data_button",
                "manual",
                "not-a-date",
                "Modules.GUI.MainWindow.TruckExcelImporter.validate_target_path",
            ),
        )
        for case_name, button_name, mode, raw_date, validator_path in cases:
            with self.subTest(case=case_name), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                driver = root / "chromedriver.exe"
                driver.write_bytes(b"driver")
                workbook = root / "SAN2_입고스케줄관리.xlsx"
                workbook.write_bytes(b"workbook")
                settings = QSettings(
                    str(root / "settings.ini"),
                    QSettings.Format.IniFormat,
                )
                settings.setValue("milkrun_excel_path", str(workbook))
                settings.setValue("base_date_mode", mode)
                settings.setValue("manual_base_date", raw_date)
                settings.sync()

                window = MainWindow(smoke_test=True, settings=settings)
                try:
                    with (
                        patch(
                            "Modules.GUI.MainWindow.chromedriver_path",
                            return_value=driver,
                        ),
                        patch(validator_path, return_value=workbook),
                        patch("Modules.GUI.MainWindow.MilkrunWorker") as worker_class,
                        patch.object(QMessageBox, "warning") as warning,
                    ):
                        getattr(window, button_name).click()

                    worker_class.assert_not_called()
                    warning.assert_called_once()
                    self.assertEqual(warning.call_args.args[1], "기준일 확인")
                    self.assertEqual(window.base_date_mode.currentData(), "invalid")
                    self.assertIn("메인 화면", warning.call_args.args[2])
                    self.assertIsNone(window.milkrun_worker)

                    window.base_date_mode.setCurrentIndex(
                        window.base_date_mode.findData("auto")
                    )
                    self.assertEqual(window.base_date_mode.findData("invalid"), -1)
                    self.assertEqual(settings.value("base_date_mode"), "auto")
                finally:
                    window.milkrun_worker = None
                    window._closing_after_cancel = False
                    window.close()

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
            self.assertEqual(window.raw_table.rowSpan(0, 0), 2)
            self.assertEqual(window.raw_table.rowSpan(0, 1), 2)
            self.assertEqual(window.raw_table.rowSpan(0, 2), 1)
            self.assertEqual(window.raw_table.rowSpan(0, 9), 1)
            self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "?")
            self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "?")
        finally:
            window.close()

    def test_console_log_can_collapse_expand_and_persists_state(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            settings = QSettings(
                str(Path(temp) / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            window = MainWindow(smoke_test=True, settings=settings)
            try:
                self.assertFalse(window.log_view.isHidden())
                self.assertEqual(window.log_toggle_button.text(), "로그 접기 ▲")

                window.log_toggle_button.click()

                self.assertTrue(window.log_view.isHidden())
                self.assertEqual(window.log_toggle_button.text(), "로그 펼치기 ▼")
                self.assertFalse(settings.value("log_expanded", True, type=bool))

                window.log_toggle_button.click()

                self.assertFalse(window.log_view.isHidden())
                self.assertTrue(settings.value("log_expanded", False, type=bool))
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

    def test_category_column_and_button_have_readable_width(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            product = MilkrunProductRow("거래처", "M1", "1", "2", "123", "상품")
            window._populate_milkrun_products((product,))
            button = window.raw_table.cellWidget(0, 9)
            window._configure_category_button(button, "중량", manual=False, enabled=True)
            window.resize(1024, 650)
            window.show()
            self.app.processEvents()

            self.assertGreaterEqual(window.raw_table.columnWidth(9), 84)
            self.assertGreaterEqual(button.width(), button.minimumSizeHint().width())
            cell_rect = window.raw_table.visualRect(window.raw_table.model().index(0, 9))
            self.assertTrue(cell_rect.contains(button.geometry()))
            self.assertEqual(window.raw_table.horizontalScrollBar().maximum(), 0)
        finally:
            window.close()

    def test_same_cached_weight_is_classified_per_row_pallet_ratio_and_high_stays_manual(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            window = MainWindow(smoke_test=True)
            window.product_memory_file = Path(temp) / "memory.json"
            try:
                products = (
                    MilkrunProductRow("거래처", "M1", "1", "100", "123", "상품"),
                    MilkrunProductRow("거래처", "M2", "1", "200", "123", "상품"),
                )
                window._populate_milkrun_products(products)
                memory = ProductMemory(window.product_memory_file)
                record = memory.upsert_measurement("123", "상품", "2000", "100", "1")

                window._on_weight_record_ready(record, True)

                self.assertEqual(window.raw_table.item(0, 4).text(), "100")
                self.assertEqual(window.raw_table.item(1, 4).text(), "200")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "경량")
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "중량")

                high = memory.set_manual_category("123", "고단")
                window._on_weight_record_ready(high, True)

                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "고단")
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "고단")
                self.assertIn("이후 데이터 조회에서도 유지", window.raw_table.cellWidget(0, 9).toolTip())
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
            self.assertEqual(window.status_label.text(), "기준일 표시 상품 없음 · WMS 조회 생략")
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

                self.assertEqual(window.raw_table.item(1, 7).text(), "1000")
                self.assertEqual(window.raw_table.item(1, 8).text(), "280")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "?")
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

    def test_product_memory_dialog_labels_weight_only_record_as_weight_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            memory = ProductMemory(Path(temp) / "memory.json")
            memory.upsert_weight_only("123", "무게 전용 상품", "1000")
            dialog = ProductMemoryDialog(memory)
            try:
                self.assertEqual(dialog.table.item(0, 2).text(), "1000")
                self.assertEqual(dialog.table.item(0, 3).text(), "미분류")
                self.assertEqual(dialog.table.item(0, 4).text(), "무게만")
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
