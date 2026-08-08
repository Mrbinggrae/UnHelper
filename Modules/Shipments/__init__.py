"""Coupang Shipments browser automation."""

from .DailyInbound import (
    MilkrunProductRow,
    extract_booking_numbers,
    extract_dispatch_numbers,
    extract_truck_reservation_numbers,
    normalize_booking_card_number,
    normalize_booking_number,
    normalize_dispatch_number,
    normalize_milkrun_card_number,
    normalize_truck_card_number,
    normalize_truck_reservation_number,
    parse_detail_table_cells,
)
from .DailyInboundScraper import (
    MILKRUN_DAILY_INBOUND_PROFILE,
    TRUCK_DAILY_INBOUND_PROFILE,
    DailyInboundError,
    DailyInboundProfile,
    DailyInboundResult,
    DailyInboundScraper,
)

from .MilkrunDownloader import (
    AutomationCancelled,
    HistoryEntry,
    MilkrunDownloadRequest,
    MilkrunDownloadResult,
    MilkrunDownloader,
)
from .TruckDownloader import (
    TruckDownloadRequest,
    TruckDownloadResult,
    TruckDownloader,
)

__all__ = [
    "AutomationCancelled",
    "HistoryEntry",
    "MilkrunDownloadRequest",
    "MilkrunDownloadResult",
    "MilkrunDownloader",
    "TruckDownloadRequest",
    "TruckDownloadResult",
    "TruckDownloader",
    "MilkrunProductRow",
    "extract_booking_numbers",
    "extract_dispatch_numbers",
    "extract_truck_reservation_numbers",
    "normalize_booking_card_number",
    "normalize_booking_number",
    "normalize_dispatch_number",
    "normalize_milkrun_card_number",
    "normalize_truck_card_number",
    "normalize_truck_reservation_number",
    "parse_detail_table_cells",
    "DailyInboundError",
    "DailyInboundProfile",
    "DailyInboundResult",
    "DailyInboundScraper",
    "MILKRUN_DAILY_INBOUND_PROFILE",
    "TRUCK_DAILY_INBOUND_PROFILE",
]
