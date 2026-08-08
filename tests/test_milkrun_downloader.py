from __future__ import annotations

import unittest
import tempfile
from datetime import date, datetime
from pathlib import Path
from unittest import mock

from selenium.webdriver.common.by import By

from Modules.Shipments.MilkrunDownloader import (
    HistoryEntry,
    MilkrunDownloadRequest,
    MilkrunDownloader,
)
from Modules.Shipments.TruckDownloader import (
    TruckDownloadRequest,
    TruckDownloader,
)


class MilkrunDownloaderTests(unittest.TestCase):
    def test_milkrun_and_truck_resolve_distinct_date_ranges(self) -> None:
        today = date(2026, 8, 8)

        self.assertEqual(
            MilkrunDownloader._resolve_date_range(today),
            (date(2026, 8, 7), today),
        )
        self.assertEqual(TruckDownloader._resolve_date_range(today), (today, today))

    def test_reason_uses_yesterday_today_format(self) -> None:
        self.assertEqual(
            MilkrunDownloader.format_reason(date(2026, 8, 7), date(2026, 8, 8)),
            "08.07-08.08",
        )

    def test_truck_reason_uses_today_only_format(self) -> None:
        self.assertEqual(
            TruckDownloader.format_reason(date(2026, 8, 8), date(2026, 8, 8)),
            "08.08",
        )

    def test_material_date_text_supports_korean_and_numeric_values(self) -> None:
        expected = (2026, 8, 7)
        self.assertEqual(MilkrunDownloader.parse_material_date_text("2026년 8월 7일 금요일"), expected)
        self.assertEqual(MilkrunDownloader.parse_material_date_text("2026. 8. 7."), expected)

    def test_calendar_month_parser_handles_year_boundary(self) -> None:
        self.assertEqual(MilkrunDownloader._parse_calendar_month("2026년 12월"), (2026, 12))
        self.assertEqual(MilkrunDownloader._parse_calendar_month("January 2027"), (2027, 1))

    def test_latest_history_row_is_not_assumed_to_be_first(self) -> None:
        started = datetime(2026, 8, 8, 15, 12, 42)
        entries = [
            HistoryEntry(None, "밀크런 입고예약 목록", "다운로드 준비완료", "08.07-08.08", datetime(2026, 8, 8, 14, 57), 0),
            HistoryEntry(None, "밀크런 입고예약 목록", "다운로드 준비완료", "다른 사유", datetime(2026, 8, 8, 15, 14), 1),
            HistoryEntry(None, "트럭 입고예약 목록", "다운로드 준비완료", "08.07-08.08", datetime(2026, 8, 8, 15, 15), 2),
            HistoryEntry(None, "밀크런 입고예약 목록", "다운로드 준비중", "08.07-08.08", datetime(2026, 8, 8, 15, 13), 3),
        ]

        selected = MilkrunDownloader.choose_latest_history_entry(entries, "08.07-08.08", started)

        self.assertIsNotNone(selected)
        self.assertEqual(selected.index, 3)
        self.assertEqual(selected.status, "다운로드 준비중")

    def test_same_minute_duplicate_prefers_later_dom_row(self) -> None:
        started = datetime(2026, 8, 8, 15, 13, 59)
        timestamp = datetime(2026, 8, 8, 15, 13)
        entries = [
            HistoryEntry(None, "밀크런 입고예약 목록", "다운로드 준비완료", "08.07-08.08", timestamp, 1),
            HistoryEntry(None, "밀크런 입고예약 목록", "다운로드 준비완료", "08.07-08.08", timestamp, 4),
        ]
        selected = MilkrunDownloader.choose_latest_history_entry(entries, "08.07-08.08", started)
        self.assertEqual(selected.index, 4)

    def test_existing_same_minute_download_url_is_excluded(self) -> None:
        started = datetime(2026, 8, 8, 15, 13, 59)
        timestamp = datetime(2026, 8, 8, 15, 13)
        entries = [
            HistoryEntry(
                None,
                "밀크런 입고예약 목록",
                "다운로드 준비완료",
                "08.07-08.08",
                timestamp,
                9,
                "https://shipments.coupang.net/ibs/csv-donwload?uuid=old",
            ),
            HistoryEntry(
                None,
                "밀크런 입고예약 목록",
                "다운로드 준비완료",
                "08.07-08.08",
                timestamp,
                1,
                "https://shipments.coupang.net/ibs/csv-donwload?uuid=new",
            ),
        ]
        selected = MilkrunDownloader.choose_latest_history_entry(
            entries,
            "08.07-08.08",
            started,
            {"https://shipments.coupang.net/ibs/csv-donwload?uuid=old"},
        )
        self.assertEqual(selected.download_href, "https://shipments.coupang.net/ibs/csv-donwload?uuid=new")

    def test_truck_history_selection_ignores_milkrun_with_same_reason(self) -> None:
        started = datetime(2026, 8, 8, 15, 13, 59)
        entries = [
            HistoryEntry(
                None,
                "밀크런 입고예약 목록",
                "다운로드 준비완료",
                "08.08",
                datetime(2026, 8, 8, 15, 14),
                5,
                "https://shipments.coupang.net/ibs/csv-donwload?uuid=milkrun",
            ),
            HistoryEntry(
                None,
                "트럭 입고예약 목록",
                "다운로드 준비중",
                "08.08",
                datetime(2026, 8, 8, 15, 13),
                2,
                "https://shipments.coupang.net/ibs/csv-donwload?uuid=truck",
            ),
        ]

        selected = TruckDownloader.choose_latest_history_entry(entries, "08.08", started)

        self.assertIsNotNone(selected)
        self.assertEqual(selected.download_type, "트럭 입고예약 목록")
        self.assertEqual(selected.download_href, "https://shipments.coupang.net/ibs/csv-donwload?uuid=truck")

    def test_truck_opens_exact_booking_list_link(self) -> None:
        downloader = TruckDownloader(Path("chromedriver.exe"), log=lambda _message: None)
        expected_xpath = (
            "//a[@href='/app/inbound-booking/truck/list' "
            "and .//span[normalize-space()='트럭 입고예약 목록']]"
        )

        with (
            mock.patch.object(downloader, "_click_locator") as click_locator,
            mock.patch.object(downloader, "_wait_document_ready"),
            mock.patch.object(downloader, "_wait"),
        ):
            downloader._open_booking_list()

        self.assertEqual(click_locator.call_count, 2)
        self.assertEqual(
            click_locator.call_args_list[1],
            mock.call(
                By.XPATH,
                expected_xpath,
                "트럭 입고예약 목록",
                timeout=60,
            ),
        )

    def test_center_selection_prefers_center_code_control(self) -> None:
        downloader = MilkrunDownloader(Path("chromedriver.exe"), log=lambda _message: None)
        center_select = mock.Mock()
        center_select.get_attribute.return_value = "false"
        center_select.text = "안산2"
        center_option = mock.Mock()
        fake_driver = mock.Mock()
        fake_driver.find_elements.return_value = [center_option]
        downloader.driver = fake_driver

        def visible_elements(_by, selector):
            if selector == "mat-select[formcontrolname='centerCode']":
                return [center_select]
            if selector == "div[role='listbox'] mat-option":
                return [center_option]
            if selector == "div[role='listbox']":
                return []
            return []

        with (
            mock.patch.object(downloader, "_visible_elements", side_effect=visible_elements) as visible,
            mock.patch.object(
                downloader,
                "_wait",
                side_effect=lambda _timeout, condition, _label: condition(),
            ),
            mock.patch.object(downloader, "_click_element"),
        ):
            downloader._select_center("안산2")

        self.assertEqual(
            visible.call_args_list[0],
            mock.call(By.CSS_SELECTOR, "mat-select[formcontrolname='centerCode']"),
        )
        self.assertNotIn(
            mock.call(By.CSS_SELECTOR, "mat-select[role='combobox'], mat-select"),
            visible.call_args_list,
        )

    def test_wait_for_download_rejects_unrelated_file_type(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            unrelated = root / "unrelated.log"
            unrelated.write_text("not a download", encoding="utf-8")
            downloader = MilkrunDownloader(root / "chromedriver.exe", log=lambda _message: None)
            downloader.DOWNLOAD_TIMEOUT_SECONDS = 3

            with self.assertRaisesRegex(RuntimeError, "지원하지 않습니다"):
                downloader._wait_for_download(root, {})

    def test_wait_for_download_accepts_stable_supported_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            expected = root / "milkrun.csv"
            expected.write_text("name,value\nitem,1\n", encoding="utf-8")
            downloader = MilkrunDownloader(root / "chromedriver.exe", log=lambda _message: None)
            downloader.DOWNLOAD_TIMEOUT_SECONDS = 3

            selected = downloader._wait_for_download(root, {})

            self.assertEqual(selected, expected)

    def test_wait_for_download_ignores_file_disappearing_between_listing_and_stat(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            expected = root / "milkrun.csv"
            expected.write_text("name,value\nitem,1\n", encoding="utf-8")
            downloader = MilkrunDownloader(root / "chromedriver.exe", log=lambda _message: None)
            downloader.DOWNLOAD_TIMEOUT_SECONDS = 3
            original_stat = Path.stat
            expected_stat_calls = 0

            def flaky_stat(path, *args, **kwargs):
                nonlocal expected_stat_calls
                if path == expected:
                    expected_stat_calls += 1
                    if expected_stat_calls == 2:
                        raise FileNotFoundError(str(path))
                return original_stat(path, *args, **kwargs)

            with mock.patch.object(Path, "stat", new=flaky_stat):
                selected = downloader._wait_for_download(root, {})

            self.assertEqual(selected, expected)

    def test_move_staged_download_avoids_overwriting_existing_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            download_dir = Path(temp)
            staging = download_dir / f"{MilkrunDownloader.STAGING_PREFIX}test"
            staging.mkdir()
            source = staging / "milkrun.csv"
            source.write_text("new", encoding="utf-8")
            existing = download_dir / "milkrun.csv"
            existing.write_text("old", encoding="utf-8")

            moved = MilkrunDownloader._move_staged_download(source, download_dir)

            self.assertNotEqual(moved, existing)
            self.assertEqual(existing.read_text(encoding="utf-8"), "old")
            self.assertEqual(moved.read_text(encoding="utf-8"), "new")

    def test_move_staged_download_retries_atomic_name_collision(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            download_dir = Path(temp)
            staging = download_dir / f"{MilkrunDownloader.STAGING_PREFIX}test"
            staging.mkdir()
            source = staging / "milkrun.csv"
            source.write_text("new", encoding="utf-8")
            original_rename = MilkrunDownloader._rename_no_replace
            calls = 0

            def collide_once(source_path: Path, destination: Path) -> None:
                nonlocal calls
                calls += 1
                if calls == 1:
                    destination.write_text("racer", encoding="utf-8")
                    raise FileExistsError(str(destination))
                original_rename(source_path, destination)

            with mock.patch.object(
                MilkrunDownloader,
                "_rename_no_replace",
                side_effect=collide_once,
            ):
                moved = MilkrunDownloader._move_staged_download(source, download_dir)

            self.assertGreaterEqual(calls, 2)
            self.assertEqual((download_dir / "milkrun.csv").read_text(encoding="utf-8"), "racer")
            self.assertNotEqual(moved, download_dir / "milkrun.csv")
            self.assertEqual(moved.read_text(encoding="utf-8"), "new")

    def test_success_can_keep_authenticated_browser_open_for_follow_up(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            downloader = MilkrunDownloader(root / "chromedriver.exe", log=lambda _message: None)
            fake_driver = object()

            def create_staged_file(staging_dir: Path, _before) -> Path:
                staged = staging_dir / "milkrun.csv"
                staged.write_text("a,b\n1,2\n", encoding="utf-8")
                return staged

            no_op_methods = (
                "_open_and_wait_for_login",
                "_open_booking_list",
                "_set_date_range",
                "_select_center",
                "_click_button_text",
                "_wait_for_query_complete",
                "_wait_for_text_download_button",
                "_click_text_download",
                "_submit_download_reason",
                "_close_request_confirmation",
                "_open_download_history",
                "_download_latest_history_file",
            )
            patches = [mock.patch.object(downloader, name) for name in no_op_methods]
            for patcher in patches:
                patcher.start()
            try:
                with (
                    mock.patch.object(downloader, "_build_driver", return_value=fake_driver),
                    mock.patch.object(downloader, "_result_table_signature", return_value=""),
                    mock.patch.object(downloader, "_snapshot_history_download_hrefs", return_value=set()),
                    mock.patch.object(downloader, "_download_snapshot", return_value={}),
                    mock.patch.object(downloader, "_wait_for_download", side_effect=create_staged_file),
                    mock.patch.object(downloader, "close") as close,
                ):
                    result = downloader.run(
                        MilkrunDownloadRequest(download_dir=root, today=date(2026, 8, 8)),
                        keep_browser_open=True,
                    )

                self.assertEqual(result.file_path, root / "milkrun.csv")
                self.assertIs(downloader.driver, fake_driver)
                close.assert_not_called()
                self.assertFalse(
                    any(path.name.startswith(downloader.STAGING_PREFIX) for path in root.iterdir())
                )
            finally:
                for patcher in reversed(patches):
                    patcher.stop()

    def test_truck_run_uses_today_for_calendar_reason_and_csv_download(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            downloader = TruckDownloader(root / "chromedriver.exe", log=lambda _message: None)
            fake_driver = object()

            def create_staged_file(staging_dir: Path, _before) -> Path:
                staged = staging_dir / "TRUCK_BOOKING_LIST_20260808.csv"
                staged.write_text("reservation,value\n3370492,1\n", encoding="utf-8")
                return staged

            with (
                mock.patch.object(downloader, "_build_driver", return_value=fake_driver),
                mock.patch.object(downloader, "_open_and_wait_for_login"),
                mock.patch.object(downloader, "_open_booking_list"),
                mock.patch.object(downloader, "_set_date_range") as set_date_range,
                mock.patch.object(downloader, "_select_center"),
                mock.patch.object(downloader, "_result_table_signature", return_value=""),
                mock.patch.object(downloader, "_click_button_text"),
                mock.patch.object(downloader, "_wait_for_query_complete"),
                mock.patch.object(downloader, "_wait_for_text_download_button"),
                mock.patch.object(downloader, "_click_text_download"),
                mock.patch.object(downloader, "_snapshot_history_download_hrefs", return_value=set()),
                mock.patch.object(downloader, "_submit_download_reason") as submit_reason,
                mock.patch.object(downloader, "_close_request_confirmation"),
                mock.patch.object(downloader, "_open_download_history"),
                mock.patch.object(downloader, "_download_snapshot", return_value={}),
                mock.patch.object(downloader, "_download_latest_history_file"),
                mock.patch.object(downloader, "_wait_for_download", side_effect=create_staged_file),
                mock.patch.object(downloader, "close") as close,
            ):
                result = downloader.run(
                    TruckDownloadRequest(download_dir=root, today=date(2026, 8, 8)),
                    keep_browser_open=True,
                )

            set_date_range.assert_called_once_with(date(2026, 8, 8), date(2026, 8, 8))
            submit_reason.assert_called_once_with("08.08")
            self.assertEqual(result.reason, "08.08")
            self.assertEqual(result.file_path.name, "TRUCK_BOOKING_LIST_20260808.csv")
            self.assertIs(downloader.driver, fake_driver)
            close.assert_not_called()
            self.assertFalse(
                any(path.name.startswith(downloader.STAGING_PREFIX) for path in root.iterdir())
            )


if __name__ == "__main__":
    unittest.main()
