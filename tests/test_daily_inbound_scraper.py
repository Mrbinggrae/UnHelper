from __future__ import annotations

import threading
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By

from Modules.Excel.MilkrunExcelImporter import MilkrunExcelImportResult
from Modules.GUI.MainWindow import MilkrunWorker
from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.Shipments.DailyInboundScraper import DailyInboundScraper
from Modules.Shipments.MilkrunDownloader import (
    AutomationCancelled,
    MilkrunDownloadRequest,
    MilkrunDownloadResult,
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

    def _matching_slots(self, order_number: str):
        return [object()] if order_number == "8789357" else []

    def _open_detail_and_read(self, card, order_number: str):
        return (
            MilkrunProductRow("거래처 A", "100", "1", "10", "SKU1", "상품 A", order_number),
            MilkrunProductRow("거래처 B", "100", "1", "10", "SKU1", "상품 B", order_number),
        )


class DailyInboundScraperTests(unittest.TestCase):
    def test_stale_unrelated_card_does_not_discard_an_exact_match(self) -> None:
        browser = _Browser((_Card(stale=True), _Card("T8789357")))
        scraper = DailyInboundScraper(browser)

        matches = scraper._matching_slots("8789357")

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].label, "T8789357")

    def test_distinct_display_rows_are_not_deduplicated_by_sku_alone(self) -> None:
        scraper = _StubScraper(_Browser())

        result = scraper.run(
            ("8789357", "9999999"),
            center_name="안산2",
            schedule_date=date(2026, 8, 8),
        )

        self.assertEqual(len(result.products), 2)
        self.assertEqual(result.products[0].sku_id, result.products[1].sku_id)
        self.assertNotEqual(result.products[0].vendor_name, result.products[1].vendor_name)
        self.assertEqual(result.unmatched_orders, ("9999999",))

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
            order_numbers=("8789357",),
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
            order_numbers=(),
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
        self.assertEqual(completed[0].daily_inbound.requested_orders, ())
        self.assertEqual(failures, [])
        self.assertTrue(any("상세 조회를 건너뜁니다" in message for message in logs))
        scraper_class.assert_not_called()
        fake_downloader.close.assert_called_once()

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
            mock.patch.object(scraper, "_slot_signature", return_value=("8789357",)),
            mock.patch.object(scraper, "_has_no_result_message", return_value=False),
        ):
            self.assertFalse(scraper._fresh_query_result_observed(("8789357",)))

        successful_query = dict(base_state, query_successes=1, query_settled_ms=1000)
        with (
            mock.patch.object(scraper, "_query_monitor_state", return_value=successful_query),
            mock.patch.object(scraper, "_slot_signature", return_value=("8789357",)),
            mock.patch.object(scraper, "_has_no_result_message", return_value=False),
        ):
            self.assertTrue(scraper._fresh_query_result_observed(("8789357",)))

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
