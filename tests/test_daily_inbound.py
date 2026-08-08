from __future__ import annotations

import unittest

from Modules.Shipments.DailyInbound import (
    extract_order_numbers,
    normalize_order_number,
    parse_detail_table_cells,
)


class DailyInboundTests(unittest.TestCase):
    def test_order_number_uses_first_slash_part_and_matches_t_card(self) -> None:
        self.assertEqual(normalize_order_number(" 1234567/7654321 "), "1234567")
        self.assertEqual(normalize_order_number("T1234567"), "1234567")
        self.assertEqual(normalize_order_number("1,234,567"), "1234567")
        self.assertEqual(normalize_order_number(1234567.0), "1234567")
        self.assertEqual(normalize_order_number("발주번호"), "")

    def test_extracts_unique_order_numbers_from_fourteenth_source_column(self) -> None:
        header = tuple(f"h{index}" for index in range(1, 15))
        first = tuple([""] * 13 + ["8789357/8789357"])
        duplicate = tuple([""] * 13 + [8789357.0])
        second = tuple([""] * 13 + ["T8827836"])

        self.assertEqual(
            extract_order_numbers((header, first, duplicate, second)),
            ("8789357", "8827836"),
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

        result = parse_detail_table_cells(rows, order_number="T8789357")

        self.assertEqual(len(result), 3)
        self.assertEqual(result[0].vendor_name, "주식회사 한티앤에스")
        self.assertEqual(result[1].milkrun_number, "10813478")
        self.assertEqual(result[1].sku_id, "72246115")
        self.assertEqual(result[2].pallet_count, "1")
        self.assertEqual(result[2].sku_name, "바유이 전동드릴")
        self.assertEqual(result[2].order_number, "8789357")


if __name__ == "__main__":
    unittest.main()
