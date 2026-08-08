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

    def _worker(self, root: Path, products, *, wms_id="id", password="pw") -> ProductWeightWorker:
        return ProductWeightWorker(
            products,
            root / "memory.json",
            root / "chromedriver.exe",
            wms_id,
            password,
            evidence_dir=root,
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
