from __future__ import annotations

import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from PySide6.QtCore import QThread, Signal

from Modules.Common.ErrorReport import FailureDetails
from Modules.Shipments.DailyInbound import MilkrunProductRow
from Modules.WMS.ProductMemory import ProductMemory, ProductMemoryRecord, normalize_product_name, normalize_sku_id
from Modules.WMS.ProductWeightCrawler import ProductWeightCrawler


@dataclass(frozen=True, slots=True)
class SkuWeightFailure:
    sku_id: str
    product_name: str
    details: FailureDetails


@dataclass(frozen=True, slots=True)
class ProductWeightSummary:
    total_skus: int
    cache_hits: int
    wms_successes: int
    failures: tuple[SkuWeightFailure, ...]


class ProductWeightWorker(QThread):
    log_updated = Signal(str)
    record_ready = Signal(object, bool)
    sku_failed = Signal(object)
    completed = Signal(object)
    failed = Signal(object)
    cancelled = Signal(str)

    def __init__(
        self,
        products: Iterable[MilkrunProductRow],
        memory_path: str | Path,
        driver_path: str | Path,
        wms_id: str,
        wms_password: str,
        *,
        evidence_dir: str | Path,
        crawler_factory: Callable[..., ProductWeightCrawler] = ProductWeightCrawler,
    ) -> None:
        super().__init__()
        self.products = tuple(products)
        self.memory_path = Path(memory_path)
        self.driver_path = Path(driver_path)
        self.wms_id = str(wms_id or "")
        self.wms_password = str(wms_password or "")
        self.evidence_dir = Path(evidence_dir)
        self.crawler_factory = crawler_factory
        self.stop_event = threading.Event()
        self.crawler: ProductWeightCrawler | None = None

    def request_cancel(self) -> None:
        self.stop_event.set()

    def run(self) -> None:
        try:
            unique_products, invalid_failures = self._unique_products(self.products)
            failures = list(invalid_failures)
            for failure in invalid_failures:
                self.sku_failed.emit(failure)

            memory = ProductMemory(self.memory_path)
            misses: list[MilkrunProductRow] = []
            cache_hits = 0
            wms_successes = 0

            for product in unique_products:
                self._check_cancelled()
                record = memory.get(product.sku_id)
                if record is None or record.weight_grams is None:
                    if record is not None:
                        self.record_ready.emit(record, False)
                    misses.append(product)
                    continue
                cache_hits += 1
                try:
                    record = memory.update_calculation(
                        product.sku_id,
                        product.box_count,
                        product.pallet_count,
                    )
                except (TypeError, ValueError):
                    # The WMS weight remains a valid cache hit even if today's
                    # shipment counts cannot produce a pallet calculation.
                    pass
                self.log_updated.emit(f"SKU {record.sku_id}는 저장된 WMS 무게를 사용합니다.")
                self.record_ready.emit(record, True)

            if not misses:
                self.completed.emit(
                    ProductWeightSummary(len(unique_products), cache_hits, 0, tuple(failures))
                )
                return

            if not self.wms_id.strip() or not self.wms_password:
                message = "미저장 SKU를 조회할 WMS ID와 비밀번호가 설정되지 않았습니다."
                for product in misses:
                    failure = SkuWeightFailure(
                        sku_id=self._safe_sku(product.sku_id),
                        product_name=normalize_product_name(product.sku_name),
                        details=FailureDetails(summary=message, detail=message),
                    )
                    failures.append(failure)
                    self.sku_failed.emit(failure)
                self.completed.emit(
                    ProductWeightSummary(len(unique_products), cache_hits, 0, tuple(failures))
                )
                return

            self.log_updated.emit(f"저장되지 않은 SKU {len(misses)}개를 WMS에서 조회합니다.")
            self.crawler = self.crawler_factory(
                self.driver_path,
                self.wms_id,
                self.wms_password,
                stop_event=self.stop_event,
                log=self.log_updated.emit,
            )
            self.crawler.start()

            evidence_saved = False
            for product in misses:
                self._check_cancelled()
                sku_id = self._safe_sku(product.sku_id)
                try:
                    lookup = self.crawler.lookup(sku_id)
                    product_name = normalize_product_name(product.sku_name) or lookup.product_name
                    try:
                        record = memory.upsert_measurement(
                            lookup.sku_id,
                            product_name,
                            lookup.weight_grams,
                            product.box_count,
                            product.pallet_count,
                        )
                    except (TypeError, ValueError):
                        record = memory.upsert_weight(
                            lookup.sku_id,
                            product_name,
                            lookup.weight_grams,
                        )
                    wms_successes += 1
                    self.record_ready.emit(record, False)
                except Exception as exc:
                    if self.stop_event.is_set():
                        raise
                    details = FailureDetails.from_exception(exc)
                    failure = SkuWeightFailure(
                        sku_id=sku_id,
                        product_name=normalize_product_name(product.sku_name),
                        details=details,
                    )
                    failures.append(failure)
                    self.sku_failed.emit(failure)
                    self.log_updated.emit(f"[WMS 조회 실패] SKU {sku_id}: {details.summary}")
                    if not evidence_saved:
                        self.crawler.save_failure_evidence(self.evidence_dir, exc)
                        evidence_saved = True

            self.completed.emit(
                ProductWeightSummary(
                    total_skus=len(unique_products),
                    cache_hits=cache_hits,
                    wms_successes=wms_successes,
                    failures=tuple(failures),
                )
            )
        except Exception as exc:
            if self.stop_event.is_set():
                self.cancelled.emit("사용자가 WMS 무게 조회를 중지했습니다.")
            else:
                if self.crawler is not None:
                    self.crawler.save_failure_evidence(self.evidence_dir, exc)
                self.failed.emit(FailureDetails.from_exception(exc))
        finally:
            if self.crawler is not None:
                self.crawler.close()
            self.crawler = None

    def _check_cancelled(self) -> None:
        if self.stop_event.is_set():
            raise RuntimeError("사용자가 WMS 무게 조회를 중지했습니다.")

    @staticmethod
    def _safe_sku(value: object) -> str:
        try:
            return normalize_sku_id(value)
        except ValueError:
            return normalize_product_name(value) or "알 수 없음"

    @classmethod
    def _unique_products(
        cls,
        products: Iterable[MilkrunProductRow],
    ) -> tuple[tuple[MilkrunProductRow, ...], tuple[SkuWeightFailure, ...]]:
        seen: set[str] = set()
        unique: list[MilkrunProductRow] = []
        failures: list[SkuWeightFailure] = []
        for product in products:
            try:
                sku_id = normalize_sku_id(product.sku_id)
            except ValueError as exc:
                details = FailureDetails.from_exception(exc)
                failures.append(
                    SkuWeightFailure(
                        sku_id=normalize_product_name(product.sku_id) or "알 수 없음",
                        product_name=normalize_product_name(product.sku_name),
                        details=details,
                    )
                )
                continue
            if sku_id in seen:
                continue
            seen.add(sku_id)
            if sku_id == product.sku_id:
                unique.append(product)
            else:
                unique.append(
                    MilkrunProductRow(
                        vendor_name=product.vendor_name,
                        milkrun_number=product.milkrun_number,
                        pallet_count=product.pallet_count,
                        box_count=product.box_count,
                        sku_id=sku_id,
                        sku_name=product.sku_name,
                        dispatch_number=product.dispatch_number,
                    )
                )
        return tuple(unique), tuple(failures)
