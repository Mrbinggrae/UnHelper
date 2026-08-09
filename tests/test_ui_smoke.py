from __future__ import annotations

import os
import tempfile
import time
import unittest
from datetime import date, datetime
from pathlib import Path
from decimal import Decimal
from unittest.mock import patch
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QDate, QPoint, QSettings, QThread, Qt
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox, QScrollArea

from Modules.Common.BookingSnapshotStore import BookingSnapshotStore
from Modules.Common.Credentials import WMSCredentialStore
from Modules.Common.ErrorReport import FailureDetails
from Modules.Excel.ArrivalSequenceReader import (
    ArrivalSequenceEntry,
    ArrivalSequenceSnapshot,
    ArrivalSummary,
    BookingFloorAssignment,
)
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
            self.assertEqual(window.raw_table.horizontalHeaderItem(1).text(), "배차번호")
            self.assertEqual(window.raw_table.horizontalHeaderItem(3).text(), "유닛 수")
            self.assertEqual(window.raw_table.horizontalHeaderItem(4).text(), "팔렛트당 유닛")
            self.assertEqual(window.raw_table.horizontalHeaderItem(6).text(), "SKU 명")
            self.assertEqual(window.raw_table.horizontalHeaderItem(9).text(), "분류")
            page = window.raw_tabs.widget(1)
            self.assertIs(window.milkrun_search_input.parentWidget().parentWidget(), page)
            self.assertLess(
                window.milkrun_search_input.mapTo(
                    page,
                    QPoint(0, 0),
                ).y(),
                window.raw_table.mapTo(page, QPoint(0, 0)).y(),
            )
            self.assertFalse(window.raw_table.wordWrap())
            self.assertTrue(window.operation_progress.isHidden())
        finally:
            window.close()

    def test_raw_excel_apply_checkbox_is_synchronized_and_persisted(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            settings = QSettings(
                str(Path(temp) / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            window = MainWindow(smoke_test=True, settings=settings)
            try:
                self.assertTrue(window.milkrun_apply_excel_checkbox.isChecked())
                self.assertTrue(window.truck_apply_excel_checkbox.isChecked())

                window.milkrun_apply_excel_checkbox.setChecked(False)

                self.assertFalse(window.truck_apply_excel_checkbox.isChecked())
                self.assertFalse(settings.value("apply_raw_to_excel", type=bool))
            finally:
                window.close()

            reopened = MainWindow(smoke_test=True, settings=settings)
            try:
                self.assertFalse(reopened.milkrun_apply_excel_checkbox.isChecked())
                self.assertFalse(reopened.truck_apply_excel_checkbox.isChecked())
            finally:
                reopened.close()

    def test_arrival_tab_has_dashboard_and_requires_raw_before_refresh(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            self.assertEqual(window.main_tabs.tabText(0), "입차순번")
            self.assertEqual(window.arrival_refresh_button.text(), "새로고침")
            self.assertIsNotNone(window.findChild(QScrollArea, "ArrivalScroll"))
            summary_table = window.arrival_summary_tables["outside_waiting"]
            self.assertGreaterEqual(summary_table.minimumHeight(), 122)
            self.assertGreaterEqual(summary_table.verticalHeader().minimumSectionSize(), 30)
            self.assertGreaterEqual(summary_table.horizontalHeader().height(), 30)
            self.assertEqual(
                set(window.arrival_detail_tables["outside_waiting"]),
                {"first", "second", "previous"},
            )
            self.assertEqual(
                set(window.arrival_detail_tables["floor_targets"]),
                {"first", "second"},
            )
            self.assertGreaterEqual(
                window.arrival_detail_tables["outside_waiting"]["second"].minimumHeight(),
                188,
            )
            self.assertEqual(
                window.arrival_detail_tables["outside_waiting"]["second"].rowCount(),
                6,
            )
            self.assertEqual(
                window.arrival_detail_tables["outside_waiting"]["second"].columnCount(),
                2,
            )
            self.assertEqual(
                window.arrival_summary_tables["floor_targets"].verticalHeaderItem(2).text(),
                "합계",
            )
            with patch.object(QMessageBox, "information") as information:
                window.refresh_arrival_sequence()
            information.assert_called_once()
            self.assertEqual(information.call_args.args[1], "RAW 데이터 필요")
            self.assertIsNone(window.arrival_worker)
        finally:
            window.close()

    def test_arrival_cards_use_full_width_without_vehicle_detail_table(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            window._arrival_auto_refreshed = True
            window.resize(1024, 650)
            window.main_tabs.setCurrentIndex(0)
            window.show()
            self.app.processEvents()

            self.assertFalse(hasattr(window, "arrival_detail_table"))
            self.assertGreaterEqual(
                window.arrival_summary_tables["outside_waiting"].width(),
                240,
            )
            self.assertGreaterEqual(
                window.arrival_detail_tables["outside_waiting"]["second"].width(),
                240,
            )
            outside = window.arrival_summary_tables["outside_waiting"]
            departure = window.arrival_summary_tables["departure"]
            floor_targets = window.arrival_summary_tables["floor_targets"]
            arrival_page = window.findChild(QScrollArea, "ArrivalScroll").widget()
            outside_pos = outside.mapTo(arrival_page, QPoint(0, 0))
            departure_pos = departure.mapTo(arrival_page, QPoint(0, 0))
            floor_pos = floor_targets.mapTo(arrival_page, QPoint(0, 0))
            self.assertEqual(outside_pos.y(), departure_pos.y())
            self.assertEqual(departure_pos.y(), floor_pos.y())
            self.assertLess(outside_pos.x(), departure_pos.x())
            self.assertLess(departure_pos.x(), floor_pos.x())
            arrival_scroll = window.findChild(QScrollArea, "ArrivalScroll")
            self.assertEqual(
                arrival_scroll.horizontalScrollBarPolicy(),
                Qt.ScrollBarPolicy.ScrollBarAsNeeded,
            )
            self.assertEqual(
                arrival_scroll.verticalScrollBarPolicy(),
                Qt.ScrollBarPolicy.ScrollBarAsNeeded,
            )
            self.assertGreater(arrival_scroll.verticalScrollBar().maximum(), 0)
            arrival_scroll.verticalScrollBar().setValue(
                arrival_scroll.verticalScrollBar().maximum()
            )
            self.app.processEvents()
            previous_table = window.arrival_detail_tables["outside_waiting"]["previous"]
            previous_bottom = previous_table.mapTo(
                arrival_scroll.viewport(),
                previous_table.rect().bottomRight(),
            )
            self.assertLessEqual(previous_bottom.y(), arrival_scroll.viewport().height())
        finally:
            window.close()

    def test_arrival_dashboard_uses_raw_for_today_and_excel_only_for_previous(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            product = MilkrunProductRow(
                "현재 거래처",
                "10807763",
                "3",
                "30",
                "123",
                "현재 상품",
                "M3370492",
            )
            window._populate_milkrun_products((product,))
            button = window.raw_table.cellWidget(0, 9)
            window._configure_category_button(
                button,
                "경량",
                manual=True,
                enabled=True,
            )
            snapshot = ArrivalSequenceSnapshot(
                workbook=Path("sample.xlsm"),
                sheet_name="입차순번",
                refreshed_at=datetime(2026, 8, 9, 1, 2, 3),
                summary=ArrivalSummary(
                    departure=(("1", "2", "3"), ("4", "5", "6"), ("7", "8", "9")),
                    outside_waiting=(("10", "11", "12"), ("13", "14", "15"), ("16", "17", "18")),
                    floor_targets=(("20", "21"), ("22", "23"), ("30", "31")),
                ),
                entries=(
                    ArrivalSequenceEntry(
                        18,
                        "MBN3370492",
                        "M3370492",
                        "milkrun",
                        "1F",
                        "99",
                        "99",
                    ),
                    ArrivalSequenceEntry(
                        19,
                        "tbn003370493",
                        "T3370493",
                        "truck",
                        "",
                        "8",
                        "8",
                    ),
                ),
                floor_assignments=(
                    BookingFloorAssignment(
                        "M3370492",
                        "milkrun",
                        "2F",
                        "Raw_밀크런",
                        2,
                    ),
                ),
            )

            window._render_arrival_sequence(snapshot)

            self.assertEqual(
                window.arrival_summary_tables["floor_targets"].item(2, 1).text(),
                "31",
            )
            self.assertEqual(
                window.arrival_summary_tables["outside_waiting"].item(0, 0).text(),
                "10",
            )
            self.assertEqual(
                window.arrival_detail_tables["departure"]["first"].item(0, 1).text(),
                "3 Pallet",
            )
            self.assertEqual(
                window.arrival_detail_tables["outside_waiting"]["previous"].item(0, 1).text(),
                "8 Pallet",
            )
            floor_second = window.arrival_detail_tables["floor_targets"]["second"]
            self.assertEqual(floor_second.item(0, 0).text(), "총 팔렛트")
            self.assertEqual(floor_second.item(0, 1).text(), "3 Pallet")
            self.assertEqual(floor_second.item(1, 0).text(), "경량")
            self.assertEqual(floor_second.item(1, 1).text(), "3 Pallet")
            self.assertNotIn("previous", window.arrival_detail_tables["floor_targets"])
        finally:
            window.close()

    def test_operation_progress_shows_completed_and_remaining_percent(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            window._on_detail_progress(1, 4)
            self.assertFalse(window.operation_progress.isHidden())
            self.assertEqual(window.operation_progress.value(), 25)
            self.assertEqual(
                window.operation_progress.format(),
                "25% 완료 · 75% 남음",
            )
            self.assertEqual(window.status_label.text(), "상세 상품 조회 1/4 · 75% 남음")

            window._on_weight_progress(2, 3)
            self.assertEqual(window.operation_progress.value(), 66)
            self.assertEqual(
                window.operation_progress.format(),
                "66% 완료 · 34% 남음",
            )
            self.assertEqual(window.status_label.text(), "상품 무게 확인 2/3 · 34% 남음")

            window._set_automation_working(False)
            self.assertTrue(window.operation_progress.isHidden())
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

                self.assertEqual(first_button.text(), "양곡")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "양곡",
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
                    ("M3370492", "M3370492", "M3370492"),
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

                self.assertEqual(first_button.text(), "양곡")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "양곡",
                )
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("456").category_override,
                    "중량",
                )

                window._refresh_current_product_memory()
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "양곡")

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

    def test_open_excel_uses_close_prompt_instead_of_error_report(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            with (
                patch.object(QMessageBox, "warning") as warning,
                patch.object(window, "_show_error_dialog") as error_dialog,
            ):
                window._on_excel_close_required(
                    Path("download.csv"),
                    "연결된 입고스케줄 Excel 파일이 열려 있습니다.",
                )

            error_dialog.assert_not_called()
            warning.assert_called_once()
            self.assertEqual(warning.call_args.args[1], "Excel을 닫아 주세요")
            self.assertIn("파일을 닫은 뒤", warning.call_args.args[2])
            self.assertIn("download.csv", warning.call_args.args[2])
            self.assertEqual(window.status_label.text(), "Excel 닫기 필요")
        finally:
            window.close()

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

    def test_recent_date_snapshot_restores_tables_and_memory_without_wms(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(
                str(root / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            settings.setValue("base_date_mode", "manual")
            settings.setValue("manual_base_date", "2026-08-09")
            snapshot_path = root / "snapshots.json"
            memory_path = root / "memory.json"
            store = BookingSnapshotStore(snapshot_path)
            store.save_table(
                date(2026, 8, 9),
                "milkrun",
                (
                    MilkrunProductRow(
                        "복원 거래처",
                        "10807763",
                        "1",
                        "2",
                        "123",
                        "복원 상품 / A",
                        "M3370492",
                    ),
                ),
            )
            store.save_table(
                date(2026, 8, 9),
                "truck",
                (
                    MilkrunProductRow(
                        "트럭 거래처",
                        "PALLET_1",
                        "2",
                        "10",
                        "456",
                        "트럭 상품",
                        "T3372829",
                    ),
                ),
            )
            memory = ProductMemory(memory_path)
            record = memory.upsert_measurement("123", "복원 상품 / A", "1000", "2", "1")
            memory.set_manual_category("123", "고단")
            memory.upsert_measurement("456", "트럭 상품", "2000", "10", "2")

            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=memory_path,
                snapshot_file=snapshot_path,
            )
            try:
                self.assertEqual(window.raw_table.rowCount(), 1)
                self.assertEqual(window.raw_table.item(0, 0).text(), "복원 거래처")
                self.assertEqual(window.raw_table.item(0, 7).text(), "1000")
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "고단")
                self.assertEqual(window.truck_table.rowCount(), 1)
                self.assertEqual(window.truck_table.item(0, 7).text(), "2000")
                self.assertIsNone(window.milkrun_worker)
                self.assertIsNone(window.weight_worker)
                self.assertIn("저장 표 복원됨", window.status_label.text())
            finally:
                window.close()

    def test_completed_pipeline_automatically_saves_its_actual_base_date(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            snapshot_path = root / "snapshots.json"
            window = MainWindow(
                smoke_test=True,
                product_memory_file=root / "memory.json",
                snapshot_file=snapshot_path,
            )
            product = MilkrunProductRow(
                "자동 저장 거래처",
                "10807763",
                "1",
                "2",
                "123",
                "자동 저장 상품",
                "M3370492",
            )
            result = mock.Mock(
                booking_type="milkrun",
                base_date=date(2026, 8, 7),
                excel=mock.Mock(
                    source_file=Path("download.csv"),
                    target_workbook=Path("입고스케줄관리.xlsx"),
                    sheet_name="Raw_밀크런",
                    rows=2,
                    columns=14,
                    filtered_rows=0,
                ),
                daily_inbound=mock.Mock(
                    products=(product,),
                    unmatched_dispatches=(),
                    empty_detail_dispatches=(),
                ),
            )
            try:
                with patch.object(window, "_start_weight_lookup") as start_weight:
                    window._on_milkrun_completed(result)

                saved = BookingSnapshotStore(snapshot_path).get(date(2026, 8, 7))
                self.assertEqual(saved.milkrun_products, (product,))
                self.assertEqual(window.raw_table.item(0, 1).text(), "M3370492")
                start_weight.assert_called_once_with(
                    window.current_products,
                    retry_mode="resume",
                )
            finally:
                window.close()

    def test_full_pipeline_restart_remeasures_all_skus_after_detail_collection(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            window = MainWindow(
                smoke_test=True,
                product_memory_file=root / "memory.json",
                snapshot_file=root / "snapshots.json",
            )
            product = MilkrunProductRow(
                "거래처", "10807763", "1", "2", "123", "상품", "M3370492"
            )
            result = mock.Mock(
                booking_type="milkrun",
                base_date=date(2026, 8, 9),
                excel=mock.Mock(
                    source_file=Path("download.csv"),
                    target_workbook=Path("입고스케줄관리.xlsx"),
                    sheet_name="Raw_밀크런",
                    rows=2,
                    columns=14,
                    filtered_rows=0,
                ),
                daily_inbound=mock.Mock(
                    products=(product,),
                    unmatched_dispatches=(),
                    empty_detail_dispatches=(),
                ),
            )
            try:
                window._pending_full_pipeline_restart = True
                with patch.object(window, "_start_weight_lookup") as start_weight:
                    window._on_milkrun_completed(result)

                start_weight.assert_called_once_with(
                    window.current_products,
                    retry_mode="restart",
                )
                self.assertFalse(window._pending_full_pipeline_restart)
            finally:
                window.close()

    def test_changing_base_date_loads_snapshot_or_clears_previous_date(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(
                str(root / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            settings.setValue("base_date_mode", "manual")
            settings.setValue("manual_base_date", "2026-08-08")
            snapshot_path = root / "snapshots.json"
            store = BookingSnapshotStore(snapshot_path)
            store.save_table(
                date(2026, 8, 8),
                "milkrun",
                (MilkrunProductRow("A", "100", "1", "2", "123", "첫날", "M1"),),
            )
            store.save_table(
                date(2026, 8, 9),
                "milkrun",
                (MilkrunProductRow("B", "200", "1", "3", "456", "둘째날", "M2"),),
            )
            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=root / "memory.json",
                snapshot_file=snapshot_path,
            )
            try:
                self.assertEqual(window.raw_table.item(0, 6).text(), "첫날")

                window.manual_base_date.setDate(QDate(2026, 8, 9))
                self.assertEqual(window.raw_table.item(0, 6).text(), "둘째날")

                window.manual_base_date.setDate(QDate(2026, 8, 10))
                self.assertEqual(window.raw_table.rowCount(), 0)
                self.assertEqual(window.truck_table.rowCount(), 0)
            finally:
                window.close()

    def test_table_bundle_import_restores_date_rows_and_related_sku_memory(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source_store = BookingSnapshotStore(root / "source-snapshots.json")
            selected = date(2026, 8, 9)
            source_store.save_table(
                selected,
                "milkrun",
                (
                    MilkrunProductRow(
                        "공유 거래처",
                        "10807763",
                        "1",
                        "2",
                        "123",
                        "공유 상품",
                        "M3370492",
                    ),
                ),
            )
            source_memory = ProductMemory(root / "source-memory.json")
            source_memory.upsert_measurement("123", "공유 상품", "1500", "2", "1")
            bundle = root / "shared.json"
            source_store.export_bundle(
                selected,
                bundle,
                source_memory.export_payload({"123"}),
            )

            settings = QSettings(
                str(root / "destination-settings.ini"),
                QSettings.Format.IniFormat,
            )
            destination_memory = root / "destination-memory.json"
            destination_snapshots = root / "destination-snapshots.json"
            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=destination_memory,
                snapshot_file=destination_snapshots,
            )
            try:
                with (
                    patch.object(
                        QFileDialog,
                        "getOpenFileName",
                        return_value=(str(bundle), ""),
                    ),
                    patch.object(QMessageBox, "information") as information,
                ):
                    window.import_table_snapshot()

                self.assertEqual(window.base_date_mode.currentData(), "manual")
                self.assertEqual(window.manual_base_date.date(), QDate(2026, 8, 9))
                self.assertEqual(window.raw_table.item(0, 0).text(), "공유 거래처")
                self.assertEqual(window.raw_table.item(0, 7).text(), "1500")
                self.assertEqual(ProductMemory(destination_memory).get("123").weight_grams, Decimal("1500"))
                information.assert_called_once()

                with (
                    patch.object(
                        window,
                        "_ask_weight_retry_action",
                        return_value="resume",
                    ) as ask_retry,
                    patch.object(window, "_start_weight_lookup") as start_weight,
                ):
                    window._start_booking_download("milkrun")

                ask_retry.assert_called_once_with(
                    cached_count=1,
                    total_count=1,
                    has_checkpoint=False,
                )
                start_weight.assert_called_once_with(
                    window._products_by_booking["milkrun"],
                    retry_mode="resume",
                )

                exported = root / "re-exported.json"
                with (
                    patch.object(
                        QFileDialog,
                        "getSaveFileName",
                        return_value=(str(exported), ""),
                    ),
                    patch.object(QMessageBox, "information"),
                ):
                    window.export_table_snapshot()
                self.assertTrue(exported.is_file())
            finally:
                window.close()

    def test_table_bundle_import_prompts_for_each_duplicate_and_applies_each_choice(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            selected = date(2026, 8, 9)
            source_store = BookingSnapshotStore(root / "source-snapshots.json")
            products = (
                MilkrunProductRow("거래처", "100", "1", "2", "123", "상품 123", "M1"),
                MilkrunProductRow("거래처", "100", "1", "2", "456", "상품 456", "M1"),
            )
            source_store.save_table(selected, "milkrun", products)
            source_memory = ProductMemory(root / "source-memory.json")
            source_memory.upsert_measurement("123", "가져온 123", "1500", "200", "1")
            source_memory.set_manual_category("123", "고단")
            source_memory.upsert_measurement("456", "가져온 456", "2500", "100", "1")
            source_memory.set_manual_category("456", "양곡")
            bundle = root / "shared-duplicates.json"
            source_store.export_bundle(
                selected,
                bundle,
                source_memory.export_payload({"123", "456"}),
            )

            destination_memory = root / "destination-memory.json"
            destination = ProductMemory(destination_memory)
            destination.upsert_measurement("123", "현재 123", "1000", "100", "1")
            destination.set_manual_category("123", "양곡")
            destination.upsert_measurement("456", "현재 456", "900", "100", "1")
            destination.set_manual_category("456", "고단")
            window = MainWindow(
                smoke_test=True,
                product_memory_file=destination_memory,
                snapshot_file=root / "destination-snapshots.json",
            )
            try:
                with (
                    patch.object(
                        QFileDialog,
                        "getOpenFileName",
                        return_value=(str(bundle), ""),
                    ),
                    patch.object(
                        window,
                        "_ask_duplicate_memory_action",
                        side_effect=("overwrite", "keep"),
                    ) as ask_duplicate,
                    patch.object(QMessageBox, "information") as information,
                ):
                    window.import_table_snapshot()

                self.assertEqual(ask_duplicate.call_count, 2)
                imported = ProductMemory(destination_memory)
                self.assertEqual(imported.get("123").product_name, "가져온 123")
                self.assertEqual(imported.get("123").weight_grams, Decimal("1500"))
                self.assertEqual(imported.get("123").category_override, "고단")
                self.assertEqual(imported.get("456").product_name, "현재 456")
                self.assertEqual(imported.get("456").weight_grams, Decimal("900"))
                self.assertEqual(imported.get("456").category_override, "고단")
                message = information.call_args.args[2]
                self.assertIn("덮어쓰기 1개", message)
                self.assertIn("기존 값 유지 1개", message)
            finally:
                window.close()

    def test_table_bundle_import_does_not_prompt_for_same_duplicate_values(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            selected = date(2026, 8, 9)
            source_store = BookingSnapshotStore(root / "source-snapshots.json")
            product = MilkrunProductRow(
                "거래처", "100", "1", "2", "123", "동일 상품", "M1"
            )
            source_store.save_table(selected, "milkrun", (product,))
            source_memory = ProductMemory(root / "source-memory.json")
            source_memory.upsert_measurement("123", "동일 상품", "1500", "2", "1")
            source_memory.set_manual_category("123", "고단")
            bundle = root / "same-values.json"
            source_store.export_bundle(
                selected,
                bundle,
                source_memory.export_payload({"123"}),
            )

            destination_memory = root / "destination-memory.json"
            destination = ProductMemory(destination_memory)
            destination.upsert_measurement("123", "동일 상품", "1500", "2", "1")
            destination.set_manual_category("123", "고단")
            window = MainWindow(
                smoke_test=True,
                product_memory_file=destination_memory,
                snapshot_file=root / "destination-snapshots.json",
            )
            try:
                with (
                    patch.object(
                        QFileDialog,
                        "getOpenFileName",
                        return_value=(str(bundle), ""),
                    ),
                    patch.object(window, "_ask_duplicate_memory_action") as ask_duplicate,
                    patch.object(QMessageBox, "information") as information,
                ):
                    window.import_table_snapshot()

                ask_duplicate.assert_not_called()
                self.assertIn("기존 값 유지 1개", information.call_args.args[2])
            finally:
                window.close()

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
                settings.setValue("apply_raw_to_excel", False)
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
                    self.assertFalse(
                        worker_class.call_args.kwargs["apply_to_excel"]
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

    def test_raw_search_filters_supported_fields_and_keeps_vehicle_group_visible(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            milkrun_products = (
                MilkrunProductRow(
                    "거래처 알파",
                    "10807763",
                    "1",
                    "10",
                    "123",
                    "사과 상품",
                    "M3370492",
                ),
                MilkrunProductRow(
                    "거래처 알파",
                    "10807763",
                    "1",
                    "20",
                    "456",
                    "배 상품",
                    "M3370492",
                ),
                MilkrunProductRow(
                    "거래처 베타",
                    "10808888",
                    "1",
                    "30",
                    "789",
                    "독립 상품",
                    "M3370493",
                ),
            )
            window._populate_milkrun_products(milkrun_products)

            window.milkrun_search_input.setText("배 상품")
            self.assertFalse(window.raw_table.isRowHidden(0))
            self.assertFalse(window.raw_table.isRowHidden(1))
            self.assertTrue(window.raw_table.isRowHidden(2))

            self.assertEqual(window.raw_table.item(0, 1).text(), "M3370492")
            self.assertEqual(window.raw_table.rowSpan(0, 1), 2)

            window.milkrun_search_input.setText("M3370493 789")
            self.assertTrue(window.raw_table.isRowHidden(0))
            self.assertTrue(window.raw_table.isRowHidden(1))
            self.assertFalse(window.raw_table.isRowHidden(2))

            truck_products = (
                MilkrunProductRow(
                    "트럭 거래처",
                    "PALLET_1",
                    "1",
                    "10",
                    "1001",
                    "트럭 상품 A",
                    "T3372829",
                ),
                MilkrunProductRow(
                    "다른 거래처",
                    "PALLET_2",
                    "1",
                    "20",
                    "1002",
                    "트럭 상품 B",
                    "T3372830",
                ),
            )
            window._populate_truck_products(truck_products)
            window.truck_search_input.setText("T3372830")
            self.assertTrue(window.truck_table.isRowHidden(0))
            self.assertFalse(window.truck_table.isRowHidden(1))

            window.milkrun_search_input.clear()
            self.assertTrue(
                all(
                    not window.raw_table.isRowHidden(row)
                    for row in range(window.raw_table.rowCount())
                )
            )
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

    def test_incomplete_checkpoint_resume_skips_full_shipments_pipeline(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(
                str(root / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            settings.setValue("base_date_mode", "manual")
            settings.setValue("manual_base_date", "2026-08-09")
            memory_path = root / "memory.json"
            ProductMemory(memory_path).upsert_measurement(
                "123", "저장 상품", "1000", "2", "1"
            )
            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=memory_path,
                snapshot_file=root / "snapshots.json",
            )
            products = (
                MilkrunProductRow(
                    "거래처", "M1", "1", "2", "123", "저장 상품", "M3370492"
                ),
                MilkrunProductRow(
                    "거래처", "M1", "1", "2", "456", "미완료 상품", "M3370492"
                ),
            )
            try:
                window._populate_milkrun_products(products)
                window._remember_weight_retry_checkpoint(
                    products,
                    memory=ProductMemory(memory_path),
                )
                with (
                    patch.object(
                        window,
                        "_ask_weight_retry_action",
                        return_value="resume",
                    ) as ask_retry,
                    patch.object(window, "_start_weight_lookup") as start_weight,
                ):
                    window._start_booking_download("milkrun")

                ask_retry.assert_called_once_with(
                    cached_count=1,
                    total_count=2,
                    has_checkpoint=True,
                )
                start_weight.assert_called_once_with(products, retry_mode="resume")
                self.assertEqual(window.current_products, products)
            finally:
                window.close()

    def test_imported_or_restored_table_resume_skips_full_shipments_without_checkpoint(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(
                str(root / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            settings.setValue("base_date_mode", "manual")
            settings.setValue("manual_base_date", "2026-08-09")
            memory_path = root / "memory.json"
            ProductMemory(memory_path).upsert_measurement(
                "123", "저장 상품", "1000", "2", "1"
            )
            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=memory_path,
                snapshot_file=root / "snapshots.json",
            )
            products = (
                MilkrunProductRow(
                    "거래처", "T1", "1", "2", "123", "저장 상품", "T3370492"
                ),
                MilkrunProductRow(
                    "거래처", "T1", "1", "2", "456", "미완료 상품", "T3370492"
                ),
            )
            try:
                window._populate_truck_products(products)
                window._refresh_current_product_memory(announce=False)
                self.assertEqual(window._read_weight_retry_checkpoints(), {})

                with (
                    patch.object(
                        window,
                        "_ask_weight_retry_action",
                        return_value="resume",
                    ) as ask_retry,
                    patch.object(window, "_start_weight_lookup") as start_weight,
                ):
                    window._start_booking_download("truck")

                ask_retry.assert_called_once_with(
                    cached_count=1,
                    total_count=2,
                    has_checkpoint=False,
                )
                start_weight.assert_called_once_with(products, retry_mode="resume")
                self.assertEqual(window.current_products, products)
            finally:
                window.close()

    def test_weight_problem_choice_routes_resume_or_full_restart(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            window = MainWindow(
                smoke_test=True,
                product_memory_file=root / "memory.json",
                snapshot_file=root / "snapshots.json",
            )
            products = (
                MilkrunProductRow(
                    "거래처", "M1", "1", "2", "123", "상품", "M3370492"
                ),
            )
            try:
                window._populate_milkrun_products(products)
                ProductMemory(window.product_memory_file).upsert_measurement(
                    "123", "상품", "1000", "2", "1"
                )
                with (
                    patch.object(window, "_ask_weight_retry_action", return_value="resume"),
                    patch.object(window, "_start_weight_lookup") as start_weight,
                    patch.object(window, "_start_booking_download") as start_booking,
                ):
                    window._offer_weight_retry_after_problem()
                start_weight.assert_called_once_with(products, retry_mode="resume")
                start_booking.assert_not_called()

                with (
                    patch.object(window, "_ask_weight_retry_action", return_value="restart"),
                    patch.object(window, "_start_weight_lookup") as start_weight,
                    patch.object(window, "_start_booking_download") as start_booking,
                ):
                    window._offer_weight_retry_after_problem()
                start_weight.assert_not_called()
                start_booking.assert_called_once_with("milkrun", retry_mode="restart")
            finally:
                window.close()

    def test_successful_weight_completion_clears_persistent_retry_checkpoint(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(
                str(root / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            settings.setValue("base_date_mode", "manual")
            settings.setValue("manual_base_date", "2026-08-09")
            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=root / "memory.json",
                snapshot_file=root / "snapshots.json",
            )
            products = (
                MilkrunProductRow(
                    "거래처", "M1", "1", "2", "123", "상품", "M3370492"
                ),
            )
            try:
                window._populate_milkrun_products(products)
                key, _pending = window._remember_weight_retry_checkpoint(products)
                window._pending_weight_summary = ProductWeightSummary(
                    total_skus=1,
                    cache_hits=1,
                    wms_successes=0,
                    failures=(),
                )
                with patch.object(QMessageBox, "information"):
                    window._finalize_weight_lookup()

                self.assertNotIn(key, window._read_weight_retry_checkpoints())
            finally:
                window.close()

    def test_restart_checkpoint_resume_keeps_only_current_run_successes_completed(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = QSettings(
                str(root / "settings.ini"),
                QSettings.Format.IniFormat,
            )
            settings.setValue("base_date_mode", "manual")
            settings.setValue("manual_base_date", "2026-08-09")
            memory_path = root / "memory.json"
            memory = ProductMemory(memory_path)
            memory.upsert_measurement("123", "완료 상품", "1000", "2", "1")
            memory.upsert_measurement("456", "실패 전 기존 상품", "2000", "2", "1")
            window = MainWindow(
                smoke_test=True,
                settings=settings,
                product_memory_file=memory_path,
                snapshot_file=root / "snapshots.json",
            )
            products = (
                MilkrunProductRow("거래처", "M1", "1", "2", "123", "완료 상품", "M1"),
                MilkrunProductRow("거래처", "M1", "1", "2", "456", "실패 상품", "M1"),
            )
            try:
                window._populate_milkrun_products(products)
                _key, initial_pending = window._remember_weight_retry_checkpoint(
                    products,
                    memory=memory,
                    retry_mode="restart",
                )
                self.assertEqual(initial_pending, ("123", "456"))
                window._mark_weight_checkpoint_completed("123")

                _key, resumed_pending = window._remember_weight_retry_checkpoint(
                    products,
                    memory=memory,
                    retry_mode="resume",
                )

                self.assertEqual(resumed_pending, ("456",))
                checkpoint = window._read_weight_retry_checkpoints()[
                    window._weight_checkpoint_key()
                ]
                self.assertEqual(checkpoint["completed_sku_ids"], ["123"])
            finally:
                window.close()

    def test_partial_weight_failure_offers_retry_choice_after_error_details(self) -> None:
        window = MainWindow(smoke_test=True)
        products = (
            MilkrunProductRow(
                "거래처", "M1", "1", "2", "123", "상품", "M3370492"
            ),
        )
        try:
            window._populate_milkrun_products(products)
            failure = SkuWeightFailure(
                sku_id="123",
                product_name="상품",
                details=FailureDetails(summary="조회 실패", detail="조회 실패"),
            )
            window._pending_weight_summary = ProductWeightSummary(
                total_skus=1,
                cache_hits=0,
                wms_successes=0,
                failures=(failure,),
            )
            with (
                patch.object(window, "_show_error_dialog") as show_error,
                patch.object(window, "_offer_weight_retry_after_problem") as offer_retry,
            ):
                window._finalize_weight_lookup()

            show_error.assert_called_once()
            offer_retry.assert_called_once_with()
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
                self.assertEqual(window.raw_table.cellWidget(0, 9).text(), "양곡")
                self.assertEqual(window.raw_table.cellWidget(1, 9).text(), "양곡")
                self.assertEqual(
                    ProductMemory(window.product_memory_file).get("123").category_override,
                    "양곡",
                )
                self.assertIn(
                    "이후 데이터 조회에서도 유지",
                    window.raw_table.cellWidget(0, 9).toolTip(),
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

    def test_product_memory_dialog_filters_grain_records(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            memory = ProductMemory(Path(temp) / "memory.json")
            memory.set_manual_category("123", "양곡", "양곡 상품")
            memory.set_manual_category("456", "고단", "고단 상품")
            dialog = ProductMemoryDialog(memory)
            try:
                self.assertIn("양곡", dialog.FILTERS)
                dialog.filter_combo.setCurrentText("양곡")
                self.assertEqual(dialog.table.rowCount(), 1)
                self.assertEqual(dialog.table.item(0, 0).text(), "123")
                self.assertEqual(dialog.table.item(0, 3).text(), "양곡")
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
