from __future__ import annotations

import json
import tempfile
import unittest
from dataclasses import replace
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
            legacy_payload["version"] = 1
            legacy_row = legacy_payload["snapshot"]["tables"]["milkrun"][0]
            legacy_row.pop("container_pallets")
            legacy_row.pop("detail_unavailable")
            legacy_row["order_number"] = "138123"
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

    def test_truck_container_pallets_round_trip_and_legacy_v1_rows_still_load(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            store = BookingSnapshotStore(root / "snapshots.json")
            selected = date(2026, 8, 10)
            product = MilkrunProductRow(
                vendor_name="거래처",
                milkrun_number="PALLET_001",
                pallet_count="2",
                box_count="20",
                sku_id="123",
                sku_name="상품",
                dispatch_number="T3372829",
                container_pallets=(("barcode:cbn-1", "2"),),
            )
            store.save_table(selected, "truck", (product,))

            restored = store.get(selected)

            self.assertEqual(
                restored.truck_products[0].container_pallets,
                (("barcode:cbn-1", "2"),),
            )

            legacy_payload = json.loads(store.path.read_text(encoding="utf-8"))
            legacy_payload["version"] = 1
            legacy_row = legacy_payload["entries"][0]["tables"]["truck"][0]
            legacy_row.pop("container_pallets")
            legacy_row.pop("detail_unavailable")
            store.path.write_text(
                json.dumps(legacy_payload, ensure_ascii=False),
                encoding="utf-8",
            )

            legacy_restored = store.get(selected)

            self.assertEqual(legacy_restored.truck_products[0].container_pallets, ())
            self.assertFalse(legacy_restored.truck_products[0].detail_unavailable)

    def test_v2_detail_unavailable_round_trips_through_store_and_bundle(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            store = BookingSnapshotStore(root / "snapshots.json")
            selected = date(2026, 8, 10)
            placeholder = MilkrunProductRow(
                vendor_name="원본 거래처",
                milkrun_number="",
                pallet_count="16",
                box_count="864",
                sku_id="",
                sku_name="일별 입고 카드 미조회 · 다운로드 원본 합계로 표시",
                dispatch_number="T8886709",
                detail_unavailable=True,
            )
            store.save_table(selected, "truck", (placeholder,))

            restored = store.get(selected)

            self.assertEqual(restored.truck_products, (placeholder,))
            bundle = store.export_bundle(
                selected,
                root / "shared.json",
                ProductMemory(root / "memory.json").export_payload(),
            )
            shared_snapshot, _payload = BookingSnapshotStore.read_bundle(bundle)
            self.assertEqual(shared_snapshot.truck_products, (placeholder,))

    def test_v2_detail_unavailable_requires_truck_placeholder_invariants(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            store = BookingSnapshotStore(root / "snapshots.json")
            selected = date(2026, 8, 10)
            store.save_table(
                selected,
                "truck",
                (
                    MilkrunProductRow(
                        "거래처",
                        "",
                        "16",
                        "864",
                        "",
                        "상세 미확인",
                        "T8886709",
                        detail_unavailable=True,
                    ),
                ),
            )
            original_payload = json.loads(store.path.read_text(encoding="utf-8"))
            cases = (
                ("milkrun dispatch", "dispatch_number", "M8886709", "T로 시작"),
                ("bare dispatch", "dispatch_number", "8886709", "T로 시작"),
                ("unknown dispatch", "dispatch_number", "X8886709", "T로 시작"),
                ("non-empty SKU", "sku_id", "123", "SKU ID"),
                (
                    "container allocation",
                    "container_pallets",
                    [["barcode:cbn-1", "1"]],
                    "컨테이너 팔렛트",
                ),
            )

            for label, field, value, error_pattern in cases:
                with self.subTest(label=label):
                    invalid_payload = json.loads(json.dumps(original_payload))
                    invalid_payload["entries"][0]["tables"]["truck"][0][field] = value
                    store.path.write_text(
                        json.dumps(invalid_payload, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    with self.assertRaisesRegex(ValueError, error_pattern):
                        store.get(selected)

            invalid_payload = json.loads(json.dumps(original_payload))
            truck_row = invalid_payload["entries"][0]["tables"]["truck"].pop()
            invalid_payload["entries"][0]["tables"]["milkrun"].append(truck_row)
            store.path.write_text(
                json.dumps(invalid_payload, ensure_ascii=False),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(ValueError, "트럭 표"):
                store.get(selected)

    def test_save_table_rejects_invalid_detail_unavailable_rows_immediately(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            store = BookingSnapshotStore(Path(temp) / "snapshots.json")
            selected = date(2026, 8, 10)
            placeholder = MilkrunProductRow(
                "거래처",
                "",
                "16",
                "864",
                "",
                "상세 미확인",
                "T8886709",
                detail_unavailable=True,
            )
            cases = (
                (
                    "wrong table",
                    "milkrun",
                    placeholder,
                    "트럭 표",
                ),
                (
                    "wrong prefix",
                    "truck",
                    replace(placeholder, dispatch_number="M8886709"),
                    "T로 시작",
                ),
                (
                    "non-empty SKU",
                    "truck",
                    replace(placeholder, sku_id="123"),
                    "SKU ID",
                ),
                (
                    "non-empty containers",
                    "truck",
                    replace(
                        placeholder,
                        container_pallets=(("barcode:cbn-1", "1"),),
                    ),
                    "컨테이너 팔렛트",
                ),
            )

            for label, booking_type, product, error_pattern in cases:
                with self.subTest(label=label):
                    with self.assertRaisesRegex(ValueError, error_pattern):
                        store.save_table(selected, booking_type, (product,))

    def test_store_and_bundle_versions_require_an_exact_integer(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            store = BookingSnapshotStore(root / "snapshots.json")
            selected = date(2026, 8, 10)
            store.save_table(selected, "truck", (_product("123", "T8886709"),))
            store_payload = json.loads(store.path.read_text(encoding="utf-8"))
            bundle = store.export_bundle(
                selected,
                root / "shared.json",
                ProductMemory(root / "memory.json").export_payload(),
            )
            bundle_payload = json.loads(bundle.read_text(encoding="utf-8"))

            for invalid_version in (True, 1.0, "2"):
                with self.subTest(source="store", version=invalid_version):
                    invalid_payload = json.loads(json.dumps(store_payload))
                    invalid_payload["version"] = invalid_version
                    store.path.write_text(
                        json.dumps(invalid_payload, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    with self.assertRaisesRegex(ValueError, "RAW 표 저장 버전"):
                        store.get(selected)

                with self.subTest(source="bundle", version=invalid_version):
                    invalid_payload = json.loads(json.dumps(bundle_payload))
                    invalid_payload["version"] = invalid_version
                    bundle.write_text(
                        json.dumps(invalid_payload, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    with self.assertRaisesRegex(ValueError, "RAW 표 공유 버전"):
                        BookingSnapshotStore.read_bundle(bundle)

    def test_truck_container_pallet_count_rejects_extreme_decimal_exponent(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            store = BookingSnapshotStore(root / "snapshots.json")
            selected = date(2026, 8, 10)
            store.save_table(
                selected,
                "truck",
                (
                    MilkrunProductRow(
                        "거래처",
                        "PALLET_001",
                        "1",
                        "1",
                        "123",
                        "상품",
                        "T3372829",
                        (("barcode:cbn-1", "1"),),
                    ),
                ),
            )
            payload = json.loads(store.path.read_text(encoding="utf-8"))
            payload["entries"][0]["tables"]["truck"][0]["container_pallets"][0][1] = (
                "1e999999999"
            )
            store.path.write_text(json.dumps(payload), encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "범위의 정수"):
                store.get(selected)

    def test_invalid_bundle_is_rejected_before_use(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "invalid.json"
            path.write_text(json.dumps({"type": "not-unhelper"}), encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "공유 파일"):
                BookingSnapshotStore.read_bundle(path)


if __name__ == "__main__":
    unittest.main()
