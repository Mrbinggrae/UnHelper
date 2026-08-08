from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.WMS.ProductMemory import ProductMemory
from Modules.WMS.ProductWeightCrawler import ProductWeightLookup, WMSWeightError
from Modules.WMS.ProductWeightWorker import ProductWeightWorker


def _product(sku_id: str, name: str = "상품 A/B") -> MilkrunProductRow:
    return MilkrunProductRow("거래처", "M1", "2", "100", sku_id, name)


def _truck_product(
    reservation_number: str,
    sku_id: str,
    *,
    name: str = "트럭 상품",
    pallet_count: str = "1",
    unit_count: str = "300",
) -> MilkrunProductRow:
    return MilkrunProductRow(
        "거래처",
        reservation_number,
        pallet_count,
        unit_count,
        sku_id,
        name,
        dispatch_number=reservation_number,
    )


def _milkrun_booking_product(
    milkrun_number: str,
    sku_id: str,
    *,
    dispatch_number: str = "M3370492",
    name: str = "밀크런 상품",
    pallet_count: str = "1",
    box_count: str = "300",
) -> MilkrunProductRow:
    return MilkrunProductRow(
        "거래처",
        milkrun_number,
        pallet_count,
        box_count,
        sku_id,
        name,
        dispatch_number=dispatch_number,
    )


class FakeCrawler:
    instances: list["FakeCrawler"] = []
    failures: set[str] = set()

    def __init__(self, *_args, **_kwargs):
        self.lookups: list[str] = []
        self.started = False
        self.closed = False
        self.evidence = 0
        type(self).instances.append(self)

    def start(self) -> None:
        self.started = True

    def lookup(self, sku_id: str) -> ProductWeightLookup:
        self.lookups.append(sku_id)
        if sku_id in type(self).failures:
            raise WMSWeightError("상품 무게 없음")
        return ProductWeightLookup(sku_id, f"WMS {sku_id}", Decimal("1000"))

    def save_failure_evidence(self, *_args) -> tuple:
        self.evidence += 1
        return ()

    def close(self) -> None:
        self.closed = True


class ProductWeightWorkerTests(unittest.TestCase):
    def setUp(self) -> None:
        FakeCrawler.instances.clear()
        FakeCrawler.failures.clear()

    def _worker(
        self,
        root: Path,
        products,
        *,
        wms_id="id",
        password="pw",
        quantity_label="박스",
    ) -> ProductWeightWorker:
        return ProductWeightWorker(
            products,
            root / "memory.json",
            root / "chromedriver.exe",
            wms_id,
            password,
            evidence_dir=root,
            quantity_label=quantity_label,
            crawler_factory=FakeCrawler,
        )

    def test_cache_hit_skips_wms_and_recalculates_current_pallet(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement("123", "상품 A/B", "1000", "80", "2")
            worker = self._worker(root, (_product("123"),), wms_id="", password="")
            records = []
            summaries = []
            worker.record_ready.connect(lambda record, cache_hit: records.append((record, cache_hit)))
            worker.completed.connect(summaries.append)

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(len(records), 1)
            self.assertTrue(records[0][1])
            self.assertEqual(records[0][0].boxes_per_pallet, Decimal("50"))
            self.assertEqual(summaries[0].cache_hits, 1)
            self.assertFalse(summaries[0].failures)

    def test_cache_hit_recalculates_one_pallet_two_boxes_as_two(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement("123", "상품 A/B", "1000", "80", "2")
            memory.set_manual_category("123", "고단")
            product = MilkrunProductRow("거래처", "M1", "1", "2", "123", "상품 A/B")
            worker = self._worker(root, (product,))
            records = []
            worker.record_ready.connect(lambda record, cache_hit: records.append((record, cache_hit)))

            worker.run()

            self.assertEqual(len(records), 1)
            self.assertTrue(records[0][1])
            self.assertEqual(records[0][0].boxes_per_pallet, Decimal("2"))
            self.assertEqual(records[0][0].category_override, "고단")

    def test_cache_hit_recalculates_category_without_wms_and_keeps_only_high_override(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement("123", "상품 A/B", "1000", "300", "1")
            memory.set_manual_category("123", "중량")
            product = MilkrunProductRow("거래처", "M1", "1", "2", "123", "상품 A/B")
            worker = self._worker(root, (product,), wms_id="", password="")
            records = []
            logs = []
            worker.record_ready.connect(lambda record, cache_hit: records.append((record, cache_hit)))
            worker.log_updated.connect(logs.append)

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(len(records), 1)
            record, cache_hit = records[0]
            self.assertTrue(cache_hit)
            self.assertEqual(record.boxes_per_pallet, Decimal("2"))
            self.assertEqual(record.automatic_category, "경량")
            self.assertIsNone(record.category_override)
            self.assertTrue(any("저장된 WMS 무게" in line for line in logs))

    def test_invalid_current_pallet_counts_still_use_cached_weight_without_opening_wms(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement(
                "123", "상품", "1000", "2", "1"
            )
            memory.set_manual_category("123", "중량")
            product = MilkrunProductRow("거래처", "M1", "0", "2", "123", "상품")
            worker = self._worker(root, (product,))
            records = []
            logs = []
            worker.record_ready.connect(lambda record, cache_hit: records.append((record, cache_hit)))
            worker.log_updated.connect(logs.append)

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(len(records), 1)
            self.assertTrue(records[0][1])
            self.assertIsNone(records[0][0].category_override)
            self.assertIsNone(ProductMemory(root / "memory.json").get("123").category_override)
            self.assertTrue(any("계산하지 못했습니다" in line for line in logs))

    def test_duplicate_sku_is_looked_up_once_and_saved_with_daily_full_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            products = (_product("123", "상품 A/\n상품 B"), _product("123", "다른 표시"))
            worker = self._worker(root, products)
            summaries = []
            worker.completed.connect(summaries.append)

            worker.run()

            crawler = FakeCrawler.instances[0]
            self.assertEqual(crawler.lookups, ["123"])
            self.assertTrue(crawler.closed)
            record = ProductMemory(root / "memory.json").get("123")
            self.assertEqual(record.product_name, "상품 A/ 상품 B")
            self.assertEqual(record.weight_grams, Decimal("1000"))
            self.assertEqual(summaries[0].wms_successes, 1)

    def test_multi_sku_truck_wms_results_persist_each_sku_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            products = (
                _truck_product("T12345", "123"),
                _truck_product("T12345", "456"),
            )
            worker = self._worker(root, products, quantity_label="유닛")

            worker.run()

            crawler = FakeCrawler.instances[0]
            self.assertEqual(crawler.lookups, ["123", "456"])
            for sku_id in ("123", "456"):
                record = ProductMemory(root / "memory.json").get(sku_id)
                self.assertEqual(record.weight_grams, Decimal("1000"))
                self.assertEqual(record.automatic_category, "중량")
                self.assertIsNone(record.category_override)
                self.assertEqual(record.boxes_per_pallet, Decimal("300"))
                self.assertEqual(record.pallet_weight_kg, Decimal("300"))

    def test_invalid_sku_in_multi_truck_does_not_disable_valid_sku_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            products = (
                _truck_product("T12345", "123"),
                _truck_product("T12345", "BAD/SKU"),
            )
            worker = self._worker(root, products, quantity_label="유닛")

            worker.run()

            record = ProductMemory(root / "memory.json").get("123")
            self.assertIsNotNone(record)
            self.assertEqual(record.weight_grams, Decimal("1000"))
            self.assertEqual(record.automatic_category, "중량")
            self.assertIsNone(record.category_override)
            self.assertEqual(record.boxes_per_pallet, Decimal("300"))
            self.assertEqual(record.pallet_weight_kg, Decimal("300"))

    def test_multi_sku_truck_cache_hit_preserves_global_categories_without_reopening_wms(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            for sku_id in ("123", "456"):
                memory.upsert_measurement(sku_id, f"상품 {sku_id}", "1000", "300", "1")
                memory.set_manual_category(sku_id, "고단")
            products = (
                _truck_product("T12345", "123"),
                _truck_product("T12345", "456"),
            )
            worker = self._worker(
                root,
                products,
                wms_id="",
                password="",
                quantity_label="유닛",
            )
            records = []
            worker.record_ready.connect(lambda record, cache_hit: records.append((record, cache_hit)))

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(len(records), 2)
            self.assertTrue(all(cache_hit for _, cache_hit in records))
            for sku_id in ("123", "456"):
                record = ProductMemory(root / "memory.json").get(sku_id)
                self.assertEqual(record.weight_grams, Decimal("1000"))
                self.assertEqual(record.effective_category, "고단")
                self.assertEqual(record.boxes_per_pallet, Decimal("300"))
                self.assertEqual(record.pallet_weight_kg, Decimal("300"))

    def test_multi_sku_truck_wms_miss_updates_weight_but_preserves_manual_placeholder(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            ProductMemory(root / "memory.json").set_manual_category(
                "123", "고단", "기존 수동 상품"
            )
            products = (
                _truck_product("T12345", "123", name="조회 상품"),
                _truck_product("T12345", "456", name="신규 상품"),
            )
            worker = self._worker(root, products, quantity_label="유닛")

            worker.run()

            preserved = ProductMemory(root / "memory.json").get("123")
            created = ProductMemory(root / "memory.json").get("456")
            self.assertEqual(preserved.weight_grams, Decimal("1000"))
            self.assertEqual(preserved.category_override, "고단")
            self.assertEqual(preserved.effective_category, "고단")
            self.assertEqual(preserved.boxes_per_pallet, Decimal("300"))
            self.assertEqual(preserved.pallet_weight_kg, Decimal("300"))
            self.assertEqual(created.weight_grams, Decimal("1000"))
            self.assertEqual(created.effective_category, "중량")
            self.assertEqual(created.boxes_per_pallet, Decimal("300"))
            self.assertEqual(created.pallet_weight_kg, Decimal("300"))

    def test_same_sku_in_multi_and_single_truck_keeps_high_and_prior_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement("123", "공유 상품", "1000", "80", "2")
            memory.set_manual_category("123", "고단")
            memory.upsert_measurement("456", "다중 전용", "2000", "20", "1")
            products = (
                _truck_product("T12345", "123", name="공유 상품"),
                _truck_product("T12345", "456", name="다중 전용"),
                _truck_product("T67890", "123", name="공유 상품", unit_count="2"),
            )
            worker = self._worker(
                root,
                products,
                wms_id="",
                password="",
                quantity_label="유닛",
            )

            worker.run()

            record = ProductMemory(root / "memory.json").get("123")
            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(record.category_override, "고단")
            self.assertEqual(record.boxes_per_pallet, Decimal("300"))
            self.assertEqual(record.pallet_weight_kg, Decimal("300"))

    def test_multi_sku_truck_missing_credentials_does_not_delete_manual_placeholder(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.set_manual_category("123", "고단", "수동 상품")
            memory.upsert_weight_only("456", "캐시 상품", "1000")
            products = (
                _truck_product("T12345", "123", name="수동 상품"),
                _truck_product("T12345", "456", name="캐시 상품"),
            )
            worker = self._worker(
                root,
                products,
                wms_id="",
                password="",
                quantity_label="유닛",
            )

            worker.run()

            placeholder = ProductMemory(root / "memory.json").get("123")
            self.assertEqual(FakeCrawler.instances, [])
            self.assertIsNone(placeholder.weight_grams)
            self.assertEqual(placeholder.category_override, "고단")

    def test_single_sku_truck_keeps_existing_recalculation_and_high_override(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement("123", "상품", "1000", "80", "2")
            memory.set_manual_category("123", "고단")
            products = (
                _truck_product("T12345", "123", unit_count="2"),
                _truck_product("T12345", "123", unit_count="2"),
            )
            worker = self._worker(root, products, quantity_label="유닛")
            records = []
            worker.record_ready.connect(lambda record, cache_hit: records.append((record, cache_hit)))

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(len(records), 1)
            record, cache_hit = records[0]
            self.assertTrue(cache_hit)
            self.assertEqual(record.boxes_per_pallet, Decimal("2"))
            self.assertEqual(record.category_override, "고단")

    def test_multi_sku_milkrun_cache_hit_recalculates_each_sku_and_keeps_high(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            for sku_id in ("123", "456"):
                memory.upsert_measurement(sku_id, f"상품 {sku_id}", "1000", "80", "2")
                memory.set_manual_category(sku_id, "고단")
            products = (
                MilkrunProductRow("거래처", "M12345", "1", "2", "123", "상품 123", "M12345"),
                MilkrunProductRow("거래처", "M12345", "1", "2", "456", "상품 456", "M12345"),
            )
            worker = self._worker(root, products)
            before = (root / "memory.json").read_bytes()

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertNotEqual((root / "memory.json").read_bytes(), before)
            for sku_id in ("123", "456"):
                record = ProductMemory(root / "memory.json").get(sku_id)
                self.assertEqual(record.boxes_per_pallet, Decimal("2"))
                self.assertEqual(record.pallet_weight_kg, Decimal("2"))
                self.assertEqual(record.category_override, "고단")

    def test_multi_sku_milkrun_wms_results_persist_each_sku_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            products = (
                _milkrun_booking_product("10813478", "123"),
                _milkrun_booking_product("10813478", "456"),
            )
            worker = self._worker(root, products)

            worker.run()

            crawler = FakeCrawler.instances[0]
            self.assertEqual(crawler.lookups, ["123", "456"])
            for sku_id in ("123", "456"):
                record = ProductMemory(root / "memory.json").get(sku_id)
                self.assertEqual(record.weight_grams, Decimal("1000"))
                self.assertEqual(record.automatic_category, "중량")
                self.assertIsNone(record.category_override)
                self.assertEqual(record.boxes_per_pallet, Decimal("300"))
                self.assertEqual(record.pallet_weight_kg, Decimal("300"))

    def test_milkrun_groups_are_separated_within_the_same_outer_m_card(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            products = (
                _milkrun_booking_product("10813478", "123", box_count="2"),
                _milkrun_booking_product("10813479", "456", box_count="2"),
            )
            worker = self._worker(root, products)

            worker.run()

            self.assertEqual(FakeCrawler.instances[0].lookups, ["123", "456"])
            for sku_id in ("123", "456"):
                record = ProductMemory(root / "memory.json").get(sku_id)
                self.assertEqual(record.boxes_per_pallet, Decimal("2"))
                self.assertEqual(record.pallet_weight_kg, Decimal("2"))
                self.assertEqual(record.automatic_category, "경량")

    def test_invalid_sku_in_multi_milkrun_does_not_disable_valid_sku_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            products = (
                _milkrun_booking_product("10813478", "123"),
                _milkrun_booking_product("10813478", "BAD/SKU"),
            )
            worker = self._worker(root, products)
            failures = []
            worker.sku_failed.connect(failures.append)

            worker.run()

            self.assertEqual([failure.sku_id for failure in failures], ["BAD/SKU"])
            self.assertEqual(FakeCrawler.instances[0].lookups, ["123"])
            record = ProductMemory(root / "memory.json").get("123")
            self.assertEqual(record.weight_grams, Decimal("1000"))
            self.assertEqual(record.automatic_category, "중량")
            self.assertIsNone(record.category_override)
            self.assertEqual(record.boxes_per_pallet, Decimal("300"))
            self.assertEqual(record.pallet_weight_kg, Decimal("300"))

    def test_multi_sku_milkrun_wms_miss_replaces_stale_light_heavy_placeholder(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            ProductMemory(root / "memory.json").set_manual_category(
                "123", "중량", "기존 수동 상품"
            )
            products = (
                _milkrun_booking_product("10813478", "123", name="조회 상품"),
                _milkrun_booking_product("10813478", "456", name="신규 상품"),
            )
            worker = self._worker(root, products)

            worker.run()

            preserved = ProductMemory(root / "memory.json").get("123")
            created = ProductMemory(root / "memory.json").get("456")
            self.assertEqual(preserved.weight_grams, Decimal("1000"))
            self.assertIsNone(preserved.category_override)
            self.assertEqual(preserved.automatic_category, "중량")
            self.assertEqual(preserved.boxes_per_pallet, Decimal("300"))
            self.assertEqual(preserved.pallet_weight_kg, Decimal("300"))
            self.assertEqual(created.weight_grams, Decimal("1000"))
            self.assertEqual(created.effective_category, "중량")
            self.assertEqual(created.boxes_per_pallet, Decimal("300"))
            self.assertEqual(created.pallet_weight_kg, Decimal("300"))

    def test_same_sku_in_multi_and_single_milkrun_uses_first_current_row_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            memory.upsert_measurement("123", "공유 상품", "1000", "80", "2")
            memory.set_manual_category("123", "중량")
            memory.upsert_weight_only("456", "다중 전용", "2000")
            products = (
                _milkrun_booking_product("10813479", "123", box_count="2"),
                _milkrun_booking_product("10813478", "123"),
                _milkrun_booking_product("10813478", "456"),
            )
            worker = self._worker(root, products, wms_id="", password="")
            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            record = ProductMemory(root / "memory.json").get("123")
            self.assertIsNone(record.category_override)
            self.assertEqual(record.automatic_category, "경량")
            self.assertEqual(record.boxes_per_pallet, Decimal("2"))
            self.assertEqual(record.pallet_weight_kg, Decimal("2"))

    def test_multi_sku_truck_uses_normal_calculation_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            memory = ProductMemory(root / "memory.json")
            for sku_id in ("123", "456"):
                memory.upsert_measurement(sku_id, f"상품 {sku_id}", "1000", "80", "2")
                memory.set_manual_category(sku_id, "고단")
            products = (
                _truck_product("T12345", "123", unit_count="2"),
                _truck_product("T12345", "456", unit_count="2"),
            )
            worker = self._worker(
                root,
                products,
                quantity_label="유닛",
            )

            worker.run()

            for sku_id in ("123", "456"):
                record = ProductMemory(root / "memory.json").get(sku_id)
                self.assertEqual(record.boxes_per_pallet, Decimal("2"))
                self.assertEqual(record.category_override, "고단")

    def test_partial_sku_failure_keeps_success_and_reports_failure(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            FakeCrawler.failures.add("456")
            worker = self._worker(root, (_product("123"), _product("456")))
            failures = []
            summaries = []
            worker.sku_failed.connect(failures.append)
            worker.completed.connect(summaries.append)

            worker.run()

            self.assertIsNotNone(ProductMemory(root / "memory.json").get("123"))
            self.assertIsNone(ProductMemory(root / "memory.json").get("456"))
            self.assertEqual([failure.sku_id for failure in failures], ["456"])
            self.assertEqual(summaries[0].wms_successes, 1)
            self.assertEqual(len(summaries[0].failures), 1)
            self.assertEqual(FakeCrawler.instances[0].evidence, 1)

    def test_missing_credentials_preserves_manual_placeholder_and_does_not_open_wms(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            ProductMemory(root / "memory.json").set_manual_category("123", "고단", "상품 A/B")
            worker = self._worker(root, (_product("123"),), wms_id="", password="")
            records = []
            summaries = []
            worker.record_ready.connect(lambda record, cache_hit: records.append(record))
            worker.completed.connect(summaries.append)

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(records[0].category_override, "고단")
            self.assertEqual(len(summaries[0].failures), 1)

    def test_pre_cancel_emits_cancelled_without_starting_browser(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            worker = self._worker(root, (_product("123"),))
            messages = []
            worker.cancelled.connect(messages.append)
            worker.request_cancel()

            worker.run()

            self.assertEqual(FakeCrawler.instances, [])
            self.assertEqual(messages, ["사용자가 WMS 무게 조회를 중지했습니다."])


if __name__ == "__main__":
    unittest.main()
