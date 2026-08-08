from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .MilkrunDownloader import MilkrunDownloader


@dataclass(frozen=True)
class TruckDownloadRequest:
    download_dir: Path
    center_name: str = "안산2"
    today: date | None = None


@dataclass(frozen=True)
class TruckDownloadResult:
    file_path: Path
    start_date: date
    end_date: date
    reason: str


class TruckDownloader(MilkrunDownloader):
    """Download today's Truck inbound booking list with the shared browser flow."""

    BOOKING_LIST_HREF = "/app/inbound-booking/truck/list"
    BOOKING_LIST_LABEL = "트럭 입고예약 목록"
    DOWNLOAD_TYPE = "트럭 입고예약 목록"

    @classmethod
    def _resolve_date_range(cls, target_date: date) -> tuple[date, date]:
        return target_date, target_date

    @staticmethod
    def format_reason(start_date: date, end_date: date) -> str:
        return f"{end_date:%m.%d}"

    def run(
        self,
        request: TruckDownloadRequest,
        *,
        keep_browser_open: bool = False,
    ) -> TruckDownloadResult:
        result = super().run(request, keep_browser_open=keep_browser_open)
        return TruckDownloadResult(
            file_path=result.file_path,
            start_date=result.start_date,
            end_date=result.end_date,
            reason=result.reason,
        )
