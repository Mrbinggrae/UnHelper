from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path

from Modules.Common.BookingSnapshotStore import BookingSnapshotStore
from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.WMS.ProductMemory import ProductMemory


def _product(sku_id: str, dispatch_number: str, *, milkrun_number: str = "100"):
    return MilkrunProductRow(
        vendor_name=f"거래처 {sku_id}",
        milkrun_number=milkrun_number,
        pallet_count="1",
        box_count="2",
        sku_id=sku_id,
        sku_name=f"상품 / {sku_id}",
        dispatch_number=dispatch_number,
    )


class BookingSnapshotStoreTests(unittest.TestCase):
    def test_keeps_both_tables_for_a_date_and_only_two_recent_dates(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            store = BookingSnapshotStore(Path(temp) / "snapshots.json")
            first = date(2026, 8, 7)
            second = date(2026, 8, 8)
            third = date(2026, 8, 9)

            store.save_table(first, "milkrun", (_product("101", "M1"),))
            store.save_table(second, "truck", (_product("202", "T2"),))
            store.save_table(second, "milkrun", (_product("203", "M2"),))

            combined = store.get(second)
            self.assertEqual(combined.truck_products[0].sku_id, "202")
            self.assertEqual(combined.milkrun_products[0].sku_id, "203")

            store.save_table(third, "milkrun", (_product("303", "M3"),))

            self.assertEqual(store.dates(), (third, second))
            self.assertIsNone(store.get(first))

    def test_bundle_contains_only_referenced_product_memory_and_round_trips_slash(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            store = BookingSnapshotStore(root / "snapshots.json")
            selected = date(2026, 8, 9)
            product = _product("123", "M3370492", milkrun_number="10/20")
            store.save_table(selected, "milkrun", (product,))

            source_memory = ProductMemory(root / "source-memory.json")
            source_memory.upsert_measurement("123", "상품 / A", "1000", "2", "1")
            source_memory.upsert_measurement("999", "내보내지 않을 상품", "2000", "2", "1")
            bundle = root / "shared.json"
            store.export_bundle(
                selected,
                bundle,
                source_memory.export_payload({"123"}),
            )

            # v0.1.14 briefly stored an order_number field. New versions read
            # and discard it so existing local/share files remain usable.
            legacy_payload = json.loads(bundle.read_text(encoding="utf-8"))
            legacy_payload["snapshot"]["tables"]["milkrun"][0]["order_number"] = "138123"
            bundle.write_text(
                json.dumps(legacy_payload, ensure_ascii=False),
                encoding="utf-8",
            )

            snapshot, payload = BookingSnapshotStore.read_bundle(bundle)
            records = ProductMemory.validate_payload(payload)
            self.assertEqual(snapshot.base_date, selected)
            self.assertEqual(snapshot.milkrun_products[0].milkrun_number, "10/20")
            self.assertFalse(hasattr(snapshot.milkrun_products[0], "order_number"))
            self.assertEqual([record.sku_id for record in records], ["123"])

            destination = ProductMemory(root / "destination-memory.json")
            summary = destination.import_records(records)
            self.assertEqual((summary.added, summary.skipped), (1, 0))
            self.assertEqual(destination.get("123").weight_grams, source_memory.get("123").weight_grams)

    def test_invalid_bundle_is_rejected_before_use(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "invalid.json"
            path.write_text(json.dumps({"type": "not-unhelper"}), encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "공유 파일"):
                BookingSnapshotStore.read_bundle(path)


if __name__ == "__main__":
    unittest.main()
