from __future__ import annotations

import unittest
from datetime import date, datetime

from Modules.Shipments.MilkrunDownloader import HistoryEntry, MilkrunDownloader


class MilkrunDownloaderTests(unittest.TestCase):
    def test_reason_uses_yesterday_today_format(self) -> None:
        self.assertEqual(
            MilkrunDownloader.format_reason(date(2026, 8, 7), date(2026, 8, 8)),
            "08.07-08.08",
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


if __name__ == "__main__":
    unittest.main()
