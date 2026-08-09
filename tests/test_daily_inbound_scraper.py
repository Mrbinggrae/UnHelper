from __future__ import annotations

import threading
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By

from decimal import Decimal

from Modules.Excel.MilkrunExcelImporter import MilkrunExcelImportResult
from Modules.Excel.TruckExcelImporter import (
    TruckExcelImportResult,
    TruckReservationMetrics,
)
from Modules.GUI.MainWindow import MilkrunWorker
from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.Shipments.DailyInboundScraper import (
    TRUCK_DAILY_INBOUND_PROFILE,
    DailyInboundError,
    DailyInboundScraper,
    EmptyDailyInboundDetail,
)
from Modules.Shipments.MilkrunDownloader import (
    AutomationCancelled,
    MilkrunDownloadRequest,
    MilkrunDownloadResult,
)
from Modules.Shipments.TruckDownloader import (
    TruckDownloadRequest,
    TruckDownloadResult,
)


class _Label:
    def __init__(self, text: str):
        self.text = text


class _Card:
    def __init__(self, label: str = "", *, stale: bool = False):
        self.label = label
        self.stale = stale

    def find_elements(self, by, value):
        if self.stale:
            raise StaleElementReferenceException("rerendered")
        self.assert_selector(by, value)
        return [_Label(self.label)]

    @staticmethod
    def assert_selector(by, value) -> None:
        if (by, value) != (By.CSS_SELECTOR, "b"):
            raise AssertionError((by, value))


class _Driver:
    def __init__(self, cards):
        self.cards = cards

    def find_elements(self, by, value):
        if (by, value) != (By.CSS_SELECTOR, "div.booking-slot"):
            return []
        return list(self.cards)


class _Browser:
    def __init__(self, cards=()):
        self.log_messages: list[str] = []
        self.log = self.log_messages.append
        self.stop_event = threading.Event()
        self._driver = _Driver(cards)

    def _check_cancelled(self) -> None:
        if self.stop_event.is_set():
            raise AutomationCancelled("cancelled")


class _StubScraper(DailyInboundScraper):
    def _open_daily_schedule(self) -> None:
        pass

    def _select_center(self, center_name: str) -> None:
        pass

    def _set_schedule_date(self, target: date) -> None:
        pass

    def _query_schedule(self) -> None:
        pass

    def _matching_slots(self, dispatch_number: str):
        return [object()] if dispatch_number == "M3370492" else []

    def _open_detail_and_read(self, card, dispatch_number: str):
        return (
            MilkrunProductRow("거래처 A", "100", "1", "10", "SKU1", "상품 A", dispatch_number),
            MilkrunProductRow("거래처 B", "100", "1", "10", "SKU1", "상품 B", dispatch_number),
        )


class DailyInboundScraperTests(unittest.TestCase):
    def test_empty_detail_is_skipped_and_later_dispatches_continue(self) -> None:
        class PartiallyEmptyScraper(_StubScraper):
            def _matching_slots(self, dispatch_number: str):
                return [object()]

            def _open_detail_and_read(self, card, dispatch_number: str):
                if dispatch_number == "M3367934":
                    raise EmptyDailyInboundDetail(
                        "예약 상세 페이지에 상품 데이터가 없습니다."
                    )
                return (
                    MilkrunProductRow(
                        "거래처",
                        "10807763",
                        "1",
                        "10",
                        "123",
                        "상품",
                        dispatch_number,
                    ),
                )

        browser = _Browser()
        progress = []
        scraper = PartiallyEmptyScraper(
            browser,
            progress=lambda completed, total: progress.append((completed, total)),
        )

        result = scraper.run(
            ("M3367934", "M3370492"),
            center_name="안산2",
            schedule_date=date(2026, 8, 7),
        )

        self.assertEqual(len(result.products), 1)
        self.assertEqual(result.products[0].dispatch_number, "M3370492")
        self.assertEqual(result.empty_detail_dispatches, ("M3367934",))
        self.assertIn("건너뛰고 다음 예약", "\n".join(browser.log_messages))
        self.assertEqual(progress, [(0, 2), (1, 2), (2, 2)])

    def test_all_empty_details_complete_without_automation_failure(self) -> None:
        class EmptyScraper(_StubScraper):
            def _matching_slots(self, dispatch_number: str):
                return [object()]

            def _open_detail_and_read(self, card, dispatch_number: str):
                raise EmptyDailyInboundDetail(
                    "예약 상세 페이지에 상품 데이터가 없습니다."
                )

        result = EmptyScraper(_Browser()).run(
            ("M3367934",),
            center_name="안산2",
            schedule_date=date(2026, 8, 7),
        )

        self.assertEqual(result.products, ())
        self.assertEqual(result.matched_dispatches, ("M3367934",))
        self.assertEqual(result.empty_detail_dispatches, ("M3367934",))

    def test_stale_unrelated_card_does_not_discard_an_exact_match(self) -> None:
        browser = _Browser(
            (_Card(stale=True), _Card("M3370492"), _Card("T3370492"), _Card("3370492"))
        )
        scraper = DailyInboundScraper(browser)

        matches = scraper._matching_slots("M3370492")

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].label, "M3370492")

    def test_truck_profile_matches_only_explicit_t_card_with_same_number(self) -> None:
        browser = _Browser(
            (_Card(stale=True), _Card("T3370492"), _Card("M3370492"), _Card("3370492"))
        )
        scraper = DailyInboundScraper(
            browser,
            profile=TRUCK_DAILY_INBOUND_PROFILE,
        )

        matches = scraper._matching_slots("T3370492")

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].label, "T3370492")
        self.assertEqual(
            scraper._unique_dispatches(("3370492", "T3370492", "M3370492")),
            ("T3370492",),
        )

    def test_truck_profile_uses_truck_detail_href_and_source_error_message(self) -> None:
        scraper = DailyInboundScraper(
            _Browser(),
            profile=TRUCK_DAILY_INBOUND_PROFILE,
        )

        xpath = scraper._detail_link_xpath()

        self.assertIn("/app/inbound-booking/truck/detail", xpath)
        self.assertNotIn("/app/inbound-booking/milkrun/detail", xpath)
        with self.assertRaises(DailyInboundError) as raised:
            scraper.run(
                ("M3370492",),
                center_name="안산2",
                schedule_date=date(2026, 8, 8),
            )
        self.assertIn("A열 예약번호", str(raised.exception))
        self.assertIn("트럭", str(raised.exception))

    def test_truck_detail_candidate_uses_container_table_layout(self) -> None:
        browser = _Browser()
        browser._driver.execute_script = mock.Mock(
            return_value=[
                (
                    "PALLET",
                    "PALLET_007",
                    "CBN0027164455",
                    "1",
                    "17240577",
                    "죠리퐁 165g 12개입",
                    "18801111944202",
                    "40",
                    "40",
                )
            ]
        )
        scraper = DailyInboundScraper(browser, profile=TRUCK_DAILY_INBOUND_PROFILE)

        rows = scraper._detail_logical_rows()

        self.assertEqual(rows[0][4], "17240577")
        script, prefix = browser._driver.execute_script.call_args.args
        self.assertEqual(prefix, "T")
        self.assertIn("table#truckContainerList tbody", script)
        self.assertIn("row.length >= 9", script)

    def test_distinct_display_rows_are_not_deduplicated_by_sku_alone(self) -> None:
        scraper = _StubScraper(_Browser())

        result = scraper.run(
            ("M3370492", "M9999999"),
            center_name="안산2",
            schedule_date=date(2026, 8, 8),
        )

        self.assertEqual(len(result.products), 2)
        self.assertEqual(result.products[0].sku_id, result.products[1].sku_id)
        self.assertNotEqual(result.products[0].vendor_name, result.products[1].vendor_name)
        self.assertEqual(result.unmatched_dispatches, ("M9999999",))

    def test_worker_reports_partial_cancel_after_excel_was_saved(self) -> None:
        download_result = MilkrunDownloadResult(
            Path("download.csv"),
            date(2026, 8, 7),
            date(2026, 8, 8),
            "08.07-08.08",
        )
        import_result = MilkrunExcelImportResult(
            source_file=Path("download.csv"),
            target_workbook=Path("입고스케줄관리.xlsx"),
            sheet_name="Raw_밀크런",
            rows=2,
            columns=14,
            dispatch_numbers=("M3370492",),
        )

        fake_importer = mock.Mock()
        fake_importer.validate_workbook.return_value = import_result.target_workbook
        fake_importer.import_values.return_value = import_result
        fake_downloader = mock.Mock()
        fake_downloader.run.return_value = download_result
        fake_scraper = mock.Mock()
        fake_scraper.run.side_effect = AutomationCancelled("사용자가 작업을 중지했습니다.")

        worker = MilkrunWorker(
            MilkrunDownloadRequest(Path("downloads"), today=date(2026, 8, 8)),
            Path("chromedriver.exe"),
            import_result.target_workbook,
        )
        partial: list[tuple[object, str]] = []
        cancelled: list[str] = []
        worker.detail_cancelled.connect(lambda result, text: partial.append((result, text)))
        worker.cancelled.connect(cancelled.append)

        with (
            mock.patch("Modules.GUI.MainWindow.MilkrunExcelImporter", return_value=fake_importer),
            mock.patch("Modules.GUI.MainWindow.MilkrunDownloader", return_value=fake_downloader),
            mock.patch("Modules.GUI.MainWindow.DailyInboundScraper", return_value=fake_scraper),
        ):
            worker.run()

        self.assertEqual(partial, [(import_result, "사용자가 작업을 중지했습니다.")])
        self.assertEqual(cancelled, [])
        fake_downloader.close.assert_called_once()

    def test_worker_completes_header_only_import_without_daily_scraper(self) -> None:
        download_result = MilkrunDownloadResult(
            Path("download.csv"),
            date(2026, 8, 7),
            date(2026, 8, 8),
            "08.07-08.08",
        )
        import_result = MilkrunExcelImportResult(
            source_file=Path("download.csv"),
            target_workbook=Path("입고스케줄관리.xlsx"),
            sheet_name="Raw_밀크런",
            rows=1,
            columns=14,
            dispatch_numbers=(),
            filtered_rows=2,
        )
        fake_importer = mock.Mock()
        fake_importer.validate_workbook.return_value = import_result.target_workbook
        fake_importer.import_values.return_value = import_result
        fake_downloader = mock.Mock()
        fake_downloader.run.return_value = download_result
        worker = MilkrunWorker(
            MilkrunDownloadRequest(Path("downloads"), today=date(2026, 8, 8)),
            Path("chromedriver.exe"),
            import_result.target_workbook,
        )
        completed = []
        failures = []
        logs = []
        worker.completed.connect(completed.append)
        worker.detail_failed.connect(lambda *args: failures.append(args))
        worker.log_updated.connect(logs.append)

        with (
            mock.patch("Modules.GUI.MainWindow.MilkrunExcelImporter", return_value=fake_importer),
            mock.patch("Modules.GUI.MainWindow.MilkrunDownloader", return_value=fake_downloader),
            mock.patch("Modules.GUI.MainWindow.DailyInboundScraper") as scraper_class,
        ):
            worker.run()

        self.assertEqual(len(completed), 1)
        self.assertEqual(completed[0].excel, import_result)
        self.assertEqual(completed[0].daily_inbound.products, ())
        self.assertEqual(completed[0].daily_inbound.requested_dispatches, ())
        self.assertEqual(failures, [])
        self.assertTrue(any("상세 조회를 건너뜁니다" in message for message in logs))
        scraper_class.assert_not_called()
        fake_downloader.close.assert_called_once()

    def test_worker_emits_completed_only_after_all_milkrun_lists_are_read(self) -> None:
        download_result = MilkrunDownloadResult(
            Path("download.csv"),
            date(2026, 8, 7),
            date(2026, 8, 8),
            "08.07-08.08",
        )
        import_result = MilkrunExcelImportResult(
            source_file=Path("download.csv"),
            target_workbook=Path("입고스케줄관리.xlsx"),
            sheet_name="Raw_밀크런",
            rows=2,
            columns=14,
            dispatch_numbers=("M3370492", "M3370493"),
        )
        fake_importer = mock.Mock()
        fake_importer.validate_workbook.return_value = import_result.target_workbook
        fake_importer.import_values.return_value = import_result
        fake_downloader = mock.Mock()
        fake_downloader.run.return_value = download_result
        fake_scraper = mock.Mock()
        events: list[str] = []
        fake_scraper.run.side_effect = lambda *args, **kwargs: (
            events.append("all_milkrun_lists_read")
            or mock.Mock(
                products=(),
                requested_dispatches=import_result.dispatch_numbers,
                matched_dispatches=import_result.dispatch_numbers,
                unmatched_dispatches=(),
            )
        )

        worker = MilkrunWorker(
            MilkrunDownloadRequest(Path("downloads"), today=date(2026, 8, 8)),
            Path("chromedriver.exe"),
            import_result.target_workbook,
        )
        worker.completed.connect(lambda _result: events.append("wms_can_start"))

        with (
            mock.patch("Modules.GUI.MainWindow.MilkrunExcelImporter", return_value=fake_importer),
            mock.patch("Modules.GUI.MainWindow.MilkrunDownloader", return_value=fake_downloader),
            mock.patch("Modules.GUI.MainWindow.DailyInboundScraper", return_value=fake_scraper),
        ):
            worker.run()

        self.assertLess(
            events.index("all_milkrun_lists_read"),
            events.index("wms_can_start"),
        )
        fake_downloader.close.assert_called_once()

    def test_truck_worker_applies_units_and_pallets_after_all_details_are_read(self) -> None:
        download_result = TruckDownloadResult(
            Path("truck.csv"),
            date(2026, 8, 8),
            date(2026, 8, 8),
            "08.08",
        )
        metrics = TruckReservationMetrics(
            reservation_number="T3372829",
            unit_count=Decimal("2"),
            pallet_count=Decimal("1"),
            source_rows=(2,),
        )
        import_result = TruckExcelImportResult(
            source_file=Path("truck.csv"),
            target_workbook=Path("입고스케줄관리.xlsx"),
            sheet_name="Raw_트럭",
            rows=2,
            columns=19,
            dispatch_numbers=("T3372829",),
            reservation_metrics=(metrics,),
        )
        daily_result = mock.Mock(
            products=(
                MilkrunProductRow(
                    "거래처",
                    "PALLET_001",
                    "1",
                    "2",
                    "56913939",
                    "상품",
                    "T3372829",
                ),
            ),
            requested_dispatches=("T3372829",),
            matched_dispatches=("T3372829",),
            unmatched_dispatches=(),
        )
        # dataclasses.replace requires the real immutable result type.
        from Modules.Shipments.DailyInboundScraper import DailyInboundResult

        daily_result = DailyInboundResult(
            products=daily_result.products,
            requested_dispatches=daily_result.requested_dispatches,
            matched_dispatches=daily_result.matched_dispatches,
            unmatched_dispatches=(),
        )
        fake_importer = mock.Mock()
        fake_importer.validate_workbook.return_value = import_result.target_workbook
        fake_importer.import_values.return_value = import_result
        fake_downloader = mock.Mock()
        fake_downloader.run.return_value = download_result
        fake_downloader.log = mock.Mock()
        fake_downloader.stop_event = threading.Event()
        fake_scraper = mock.Mock()
        events: list[str] = []
        fake_scraper.run.side_effect = lambda *args, **kwargs: (
            events.append("all_truck_details_read") or daily_result
        )

        worker = MilkrunWorker(
            TruckDownloadRequest(Path("downloads"), today=date(2026, 8, 8)),
            Path("chromedriver.exe"),
            import_result.target_workbook,
            booking_type="truck",
        )
        completed = []
        worker.completed.connect(
            lambda result: (events.append("wms_can_start"), completed.append(result))
        )

        with (
            mock.patch("Modules.GUI.MainWindow.TruckExcelImporter", return_value=fake_importer),
            mock.patch("Modules.GUI.MainWindow.TruckDownloader", return_value=fake_downloader),
            mock.patch("Modules.GUI.MainWindow.DailyInboundScraper", return_value=fake_scraper) as scraper,
        ):
            worker.run()

        self.assertEqual(events, ["all_truck_details_read", "wms_can_start"])
        self.assertEqual(len(completed), 1)
        self.assertEqual(completed[0].booking_type, "truck")
        product = completed[0].daily_inbound.products[0]
        self.assertEqual(product.dispatch_number, "T3372829")
        self.assertEqual(product.box_count, Decimal("2"))
        self.assertEqual(product.pallet_count, Decimal("1"))
        self.assertNotIn("exclude_arrival_date", fake_importer.import_values.call_args.kwargs)
        self.assertEqual(scraper.call_args.kwargs["profile"].booking_prefix, "T")
        fake_downloader.close.assert_called_once()

    def test_truck_worker_keeps_all_skus_for_one_reservation(self) -> None:
        from Modules.Shipments.DailyInboundScraper import DailyInboundResult

        metric = TruckReservationMetrics(
            reservation_number="T3372829",
            unit_count=Decimal("20"),
            pallet_count=Decimal("2"),
            source_rows=(2,),
        )
        import_result = TruckExcelImportResult(
            source_file=Path("truck.csv"),
            target_workbook=Path("입고스케줄관리.xlsx"),
            sheet_name="Raw_트럭",
            rows=2,
            columns=19,
            dispatch_numbers=("T3372829",),
            reservation_metrics=(metric,),
        )
        daily_result = DailyInboundResult(
            products=(
                MilkrunProductRow("", "PALLET_1", "1", "12", "SKU1", "상품1", "T3372829"),
                MilkrunProductRow("", "PALLET_2", "1", "8", "SKU2", "상품2", "T3372829"),
            ),
            requested_dispatches=("T3372829",),
            matched_dispatches=("T3372829",),
            unmatched_dispatches=(),
        )
        worker = MilkrunWorker(
            TruckDownloadRequest(Path("downloads")),
            Path("chromedriver.exe"),
            import_result.target_workbook,
            booking_type="truck",
        )

        logs = []
        worker.log_updated.connect(logs.append)

        updated = worker._apply_truck_reservation_metrics(daily_result, import_result)

        self.assertEqual(len(updated.products), 2)
        self.assertEqual(
            tuple(product.dispatch_number for product in updated.products),
            ("T3372829", "T3372829"),
        )
        self.assertEqual(
            tuple(product.box_count for product in updated.products),
            (Decimal("12"), Decimal("8")),
        )
        self.assertEqual(
            tuple(product.pallet_count for product in updated.products),
            (Decimal("1"), Decimal("1")),
        )
        self.assertEqual(logs, [])

    def test_truck_detail_sheet_close_uses_truck_profile_href(self) -> None:
        browser = _Browser()
        browser._visible_elements = mock.Mock(return_value=[])
        scraper = DailyInboundScraper(browser, profile=TRUCK_DAILY_INBOUND_PROFILE)

        self.assertTrue(scraper._detail_sheet_closed())
        detail_call = browser._visible_elements.call_args_list[0]
        self.assertIn("/app/inbound-booking/truck/detail", detail_call.args[1])

    def test_fresh_result_requires_dom_change_or_successful_query_request(self) -> None:
        scraper = DailyInboundScraper(_Browser())
        base_state = {
            "started": 1,
            "active": 0,
            "successes": 1,
            "failures": 0,
            "query_successes": 0,
            "query_failures": 0,
            "result_mutations": 0,
            "query_settled_ms": 5000,
            "last_failure": "",
        }
        with (
            mock.patch.object(scraper, "_query_monitor_state", return_value=base_state),
            mock.patch.object(scraper, "_slot_signature", return_value=("M3370492",)),
            mock.patch.object(scraper, "_has_no_result_message", return_value=False),
        ):
            self.assertFalse(scraper._fresh_query_result_observed(("M3370492",)))

        successful_query = dict(base_state, query_successes=1, query_settled_ms=1000)
        with (
            mock.patch.object(scraper, "_query_monitor_state", return_value=successful_query),
            mock.patch.object(scraper, "_slot_signature", return_value=("M3370492",)),
            mock.patch.object(scraper, "_has_no_result_message", return_value=False),
        ):
            self.assertTrue(scraper._fresh_query_result_observed(("M3370492",)))

    def test_query_http_failure_is_reported(self) -> None:
        scraper = DailyInboundScraper(_Browser())
        state = {
            "started": 1,
            "active": 0,
            "successes": 0,
            "failures": 1,
            "query_successes": 0,
            "query_failures": 1,
            "result_mutations": 0,
            "query_settled_ms": 0,
            "last_failure": "/ibs/inbound-schedule · 500",
        }
        with mock.patch.object(scraper, "_query_monitor_state", return_value=state):
            self.assertIn("500", scraper._query_monitor_failure())


if __name__ == "__main__":
    unittest.main()
