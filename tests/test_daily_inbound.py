from __future__ import annotations

import unittest

from Modules.Shipments.DailyInbound import (
    extract_dispatch_numbers,
    normalize_dispatch_number,
    normalize_milkrun_card_number,
    parse_detail_table_cells,
)


class DailyInboundTests(unittest.TestCase):
    def test_dispatch_number_adds_m_prefix_and_rejects_other_card_types(self) -> None:
        self.assertEqual(normalize_dispatch_number(" 3370492 "), "M3370492")
        self.assertEqual(normalize_dispatch_number("m3370492"), "M3370492")
        self.assertEqual(normalize_dispatch_number("M3370492"), "M3370492")
        self.assertEqual(normalize_dispatch_number("3,370,492"), "M3370492")
        self.assertEqual(normalize_dispatch_number(3370492), "M3370492")
        self.assertEqual(normalize_dispatch_number(3370492.0), "M3370492")
        for invalid in (True, "T3370492", "I3370492", "3370492/1", 3370492.5, "배차번호"):
            with self.subTest(invalid=invalid):
                self.assertEqual(normalize_dispatch_number(invalid), "")

        self.assertEqual(normalize_milkrun_card_number("M3370492"), "M3370492")
        for non_m_card in ("3370492", "T3370492", "I3370492"):
            with self.subTest(non_m_card=non_m_card):
                self.assertEqual(normalize_milkrun_card_number(non_m_card), "")

    def test_extracts_unique_dispatch_numbers_from_first_source_column(self) -> None:
        header = ("배차번호", "센터", "입고일")
        first = ("3370492", "안산2", "2026-08-08")
        duplicate = (3370492.0, "안산2", "2026-08-08")
        second = ("M3370510", "안산2", "2026-08-08")

        self.assertEqual(
            extract_dispatch_numbers((header, first, duplicate, second)),
            ("M3370492", "M3370510"),
        )

    def test_rowspan_detail_rows_inherit_vendor_and_counts(self) -> None:
        rows = (
            (
                "주식회사 한티앤에스\n(A00556074)",
                "10813478",
                "10",
                "720",
                "138716974 (RG)",
                "",
                "72246083",
                "한경희 폴더블 선풍기 / 아이보리",
                "8809964240765",
                "504",
            ),
            ("", "72246115", "한경희 폴더블 선풍기 / 차콜", "8809964240772", "216"),
            (
                "엘제이디(LJD)\n(A01723626)",
                "10799314",
                "1",
                "80",
                "138563655 (RG)",
                "",
                "75259427",
                "바유이 전동드릴",
                "NEW01KRGJDZ102",
                "80",
            ),
        )

        result = parse_detail_table_cells(rows, dispatch_number="M3370492")

        self.assertEqual(len(result), 3)
        self.assertEqual(result[0].vendor_name, "주식회사 한티앤에스")
        self.assertEqual(result[1].milkrun_number, "10813478")
        self.assertEqual(result[1].sku_id, "72246115")
        self.assertEqual(result[2].pallet_count, "1")
        self.assertEqual(result[2].sku_name, "바유이 전동드릴")
        self.assertEqual(result[2].dispatch_number, "M3370492")

    def test_parser_keeps_one_pallet_and_two_boxes_in_their_own_fields(self) -> None:
        rows = (
            (
                "거래처 (A00001)",
                "10813478",
                "1",
                "2",
                "shipment",
                "",
                "123",
                "상품",
                "barcode",
                "2",
            ),
        )

        result = parse_detail_table_cells(rows)

        self.assertEqual(result[0].pallet_count, "1")
        self.assertEqual(result[0].box_count, "2")


if __name__ == "__main__":
    unittest.main()
