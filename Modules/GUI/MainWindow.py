from __future__ import annotations

import json
import threading
from dataclasses import dataclass, replace
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from PySide6.QtCore import QDate, QSettings, QThread, QTimer, Qt, QUrl, Signal
from PySide6.QtGui import QBrush, QCloseEvent, QColor, QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QLayout,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from Modules.Common.Credentials import (
    CredentialError,
    WMSCredentials,
    WMSCredentialStore,
)
from Modules.Common.BookingSnapshotStore import (
    BookingDateSnapshot,
    BookingSnapshotStore,
)
from Modules.Common.paths import (
    booking_snapshot_path,
    chromedriver_path,
    default_download_dir,
    product_memory_path,
)
from Modules.Common.version import CURRENT_VERSION
from Modules.Common.ErrorReport import FailureDetails
from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportCancelled,
    ExcelImportError,
    ExcelWorkbookOpenError,
    MilkrunExcelImportResult,
    MilkrunExcelImporter,
)
from Modules.Excel.TruckExcelImporter import (
    TruckExcelImportResult,
    TruckExcelImporter,
)
from Modules.Excel.ArrivalSequenceReader import (
    ArrivalSequenceReader,
    ArrivalSequenceSnapshot,
    RawBookingAggregate,
    build_floor_target_breakdowns,
    build_arrival_vehicles,
    build_status_pallet_breakdowns,
)
from Modules.GUI.Dialogs import ErrorReportDialog, UpdateHistoryDialog
from Modules.GUI.ProductMemoryDialog import ProductMemoryDialog
from Modules.GUI.Theme import COLORS
from Modules.Shipments.MilkrunDownloader import (
    AutomationCancelled,
    MilkrunDownloadRequest,
    MilkrunDownloader,
)
from Modules.Shipments.TruckDownloader import (
    TruckDownloadRequest,
    TruckDownloader,
)
from Modules.Shipments.DailyInboundScraper import (
    DailyInboundError,
    DailyInboundResult,
    DailyInboundScraper,
    TRUCK_DAILY_INBOUND_PROFILE,
)
from Modules.Shipments.DailyInbound import normalize_booking_number
from Modules.WMS.ProductMemory import (
    AUTOMATIC_CATEGORIES,
    GRAIN_CATEGORY,
    HIGH_CATEGORY,
    MANUAL_CATEGORIES,
    PERSISTENT_MANUAL_CATEGORIES,
    ProductMemory,
    ProductMemoryRecord,
    calculate_boxes_per_pallet,
    calculate_pallet_measurement,
    normalize_product_name,
    normalize_sku_id,
)
from Modules.WMS.ProductWeightWorker import (
    ProductWeightSummary,
    ProductWeightWorker,
    SkuWeightFailure,
)


@dataclass(frozen=True)
class MilkrunPipelineResult:
    excel: MilkrunExcelImportResult | TruckExcelImportResult
    daily_inbound: DailyInboundResult
    booking_type: str = "milkrun"
    base_date: date | None = None


class ArrivalSequenceWorker(QThread):
    log_updated = Signal(str)
    completed = Signal(object)
    failed = Signal(object)
    cancelled = Signal(str)

    def __init__(self, workbook_path: Path):
        super().__init__()
        self.workbook_path = workbook_path
        self.stop_event = threading.Event()

    def run(self) -> None:
        try:
            reader = ArrivalSequenceReader(log=self.log_updated.emit)
            result = reader.read(
                self.workbook_path,
                cancel_requested=self.stop_event.is_set,
            )
            if self.stop_event.is_set():
                raise ExcelImportCancelled("사용자가 입차순번 새로고침을 중지했습니다.")
            self.completed.emit(result)
        except ExcelImportCancelled as exc:
            self.cancelled.emit(str(exc))
        except Exception as exc:
            self.failed.emit(FailureDetails.from_exception(exc))

    def request_cancel(self) -> None:
        self.stop_event.set()


class MilkrunWorker(QThread):
    log_updated = Signal(str)
    detail_progress = Signal(int, int)
    completed = Signal(object)
    excel_failed = Signal(object, object)
    excel_close_required = Signal(object, str)
    detail_failed = Signal(object, object)
    detail_cancelled = Signal(object, str)
    failed = Signal(object)
    cancelled = Signal(str)

    def __init__(
        self,
        request: MilkrunDownloadRequest | TruckDownloadRequest,
        driver_path: Path,
        target_workbook: Path,
        *,
        booking_type: str = "milkrun",
        apply_to_excel: bool = True,
    ):
        super().__init__()
        self.request = request
        self.driver_path = driver_path
        self.target_workbook = target_workbook
        self.booking_type = booking_type
        self.apply_to_excel = apply_to_excel
        self.stop_event = threading.Event()
        self.downloader: MilkrunDownloader | None = None

    def run(self) -> None:
        download_result = None
        import_result = None
        try:
            is_truck = self.booking_type == "truck"
            booking_label = "트럭" if is_truck else "Milkrun"
            importer = (
                TruckExcelImporter(
                    log=self.log_updated.emit,
                    reject_open_target=True,
                )
                if is_truck
                else MilkrunExcelImporter(
                    log=self.log_updated.emit,
                    reject_open_target=True,
                )
            )
            if self.apply_to_excel:
                self.log_updated.emit("연결된 Excel 파일을 확인합니다.")
                importer.validate_workbook(self.target_workbook)
            else:
                self.log_updated.emit(
                    f"연결된 Excel의 '{importer.TARGET_SHEET}' 시트 반영을 제외했습니다."
                )
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")

            downloader_class = TruckDownloader if is_truck else MilkrunDownloader
            self.downloader = downloader_class(
                self.driver_path,
                log=self.log_updated.emit,
                stop_event=self.stop_event,
            )
            download_result = self.downloader.run(self.request, keep_browser_open=True)
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")
            if self.apply_to_excel:
                self.log_updated.emit("다운로드 데이터를 연결된 Excel 파일에 반영합니다.")
            else:
                self.log_updated.emit(
                    "다운로드 데이터를 앱 조회용으로 읽습니다. 연결된 Excel은 변경하지 않습니다."
                )
            import_options = {}
            if not self.apply_to_excel:
                import_options["apply_to_target"] = False
            if is_truck:
                import_result = importer.import_values(
                    download_result.file_path,
                    self.target_workbook,
                    cancel_requested=self.stop_event.is_set,
                    **import_options,
                )
            else:
                import_result = importer.import_values(
                    download_result.file_path,
                    self.target_workbook,
                    cancel_requested=self.stop_event.is_set,
                    exclude_arrival_date=download_result.start_date,
                    **import_options,
                )
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")
            if import_result.rows == 1 and not import_result.dispatch_numbers:
                self.log_updated.emit(
                    f"기준일에 반영할 {booking_label} 데이터가 없어 일별 입고 상세 조회를 건너뜁니다."
                )
                daily_result = DailyInboundResult(
                    products=(),
                    requested_dispatches=(),
                    matched_dispatches=(),
                    unmatched_dispatches=(),
                )
            else:
                if is_truck:
                    self.log_updated.emit(
                        "다운로드 첫 시트 A열의 예약번호를 T 접두사로 변환해 기준일 일별 입고 상세를 조회합니다."
                    )
                else:
                    self.log_updated.emit(
                        "다운로드 첫 시트 A열의 배차번호를 M 접두사로 변환해 기준일 일별 입고 상세를 조회합니다."
                    )
                scraper_kwargs = {
                    "evidence_dir": self.request.download_dir,
                    "progress": self.detail_progress.emit,
                }
                if is_truck:
                    scraper_kwargs["profile"] = TRUCK_DAILY_INBOUND_PROFILE
                daily_result = DailyInboundScraper(
                    self.downloader,
                    **scraper_kwargs,
                ).run(
                    import_result.dispatch_numbers,
                    center_name=self.request.center_name,
                    schedule_date=download_result.end_date,
                )
                if is_truck:
                    daily_result = self._apply_truck_reservation_metrics(
                        daily_result,
                        import_result,
                    )
            self.completed.emit(
                MilkrunPipelineResult(
                    import_result,
                    daily_result,
                    self.booking_type,
                    download_result.base_date,
                )
            )
        except (AutomationCancelled, ExcelImportCancelled) as exc:
            if import_result is not None:
                self.detail_cancelled.emit(import_result, str(exc))
            else:
                self.cancelled.emit(str(exc))
        except ExcelWorkbookOpenError as exc:
            downloaded_file = download_result.file_path if download_result is not None else None
            self.excel_close_required.emit(downloaded_file, str(exc))
        except ExcelImportError as exc:
            failure = FailureDetails.from_exception(exc)
            if download_result is not None:
                self.excel_failed.emit(download_result.file_path, failure)
            else:
                self.failed.emit(failure)
        except DailyInboundError as exc:
            if self.downloader is not None and not exc.evidence_captured:
                self.downloader.save_failure_snapshot(self.request.download_dir, exc)
            self.detail_failed.emit(import_result, FailureDetails.from_exception(exc))
        except Exception as exc:
            if self.downloader is not None:
                self.downloader.save_failure_snapshot(self.request.download_dir, exc)
            failure = FailureDetails.from_exception(exc)
            if import_result is not None:
                self.detail_failed.emit(import_result, failure)
            else:
                self.failed.emit(failure)
        finally:
            if self.downloader is not None:
                self.downloader.close()
            self.downloader = None

    def request_cancel(self) -> None:
        self.stop_event.set()

    def _apply_truck_reservation_metrics(
        self,
        daily_result: DailyInboundResult,
        import_result: TruckExcelImportResult,
    ) -> DailyInboundResult:
        metrics = import_result.metrics_by_reservation
        products = []
        detail_units: dict[str, Decimal] = {}
        detail_pallets: dict[str, Decimal] = {}
        for product in daily_result.products:
            metric = metrics.get(product.dispatch_number)
            if metric is None:
                raise DailyInboundError(
                    f"예약번호 {product.dispatch_number}의 유닛 수와 팔렛트 수를 "
                    "다운로드 파일에서 찾지 못했습니다."
                )
            try:
                unit_count = Decimal(str(product.box_count).strip().replace(",", ""))
                pallet_count = Decimal(
                    str(product.pallet_count).strip().replace(",", "")
                )
            except Exception as exc:
                raise DailyInboundError(
                    f"예약번호 {product.dispatch_number}의 SKU {product.sku_id} "
                    "컨테이너 수량 또는 총 수량을 읽지 못했습니다."
                ) from exc
            if (
                not unit_count.is_finite()
                or not pallet_count.is_finite()
                or unit_count <= 0
                or pallet_count <= 0
            ):
                raise DailyInboundError(
                    f"예약번호 {product.dispatch_number}의 SKU {product.sku_id} "
                    "컨테이너 수량 또는 총 수량이 0 이하입니다."
                )
            products.append(
                replace(
                    product,
                    vendor_name=metric.vendor_name or product.vendor_name,
                    box_count=unit_count,
                    pallet_count=pallet_count,
                )
            )
            detail_units[product.dispatch_number] = (
                detail_units.get(product.dispatch_number, Decimal("0")) + unit_count
            )
            detail_pallets[product.dispatch_number] = (
                detail_pallets.get(product.dispatch_number, Decimal("0")) + pallet_count
            )

        grouped_products: dict[str, list] = {}
        group_order: list[str] = []
        for product in products:
            if product.dispatch_number not in grouped_products:
                grouped_products[product.dispatch_number] = []
                group_order.append(product.dispatch_number)
            grouped_products[product.dispatch_number].append(product)

        for reservation_number in detail_units:
            metric = metrics[reservation_number]
            if (
                detail_units.get(reservation_number) != metric.unit_count
                or detail_pallets.get(reservation_number) != metric.pallet_count
            ):
                self.log_updated.emit(
                    f"예약번호 {reservation_number}의 컨테이너 상세 합계가 다운로드 "
                    "M/N 합계와 다릅니다. SKU별 계산에는 상세 표의 컨테이너 수량과 "
                    "총 수량을 사용합니다."
                )

        ordered_products = tuple(
            product
            for reservation_number in group_order
            for product in grouped_products[reservation_number]
        )
        return replace(daily_result, products=ordered_products)


class UpdateCheckWorker(QThread):
    update_available = Signal(object)
    no_update = Signal()
    failed = Signal(object)

    def __init__(self, use_prerelease: bool):
        super().__init__()
        self.use_prerelease = use_prerelease

    def run(self) -> None:
        try:
            from Modules.Common.AutoUpdater import AutoUpdater

            has_update, info = AutoUpdater("UnHelper", self.use_prerelease).check_for_update()
            if has_update and info:
                self.update_available.emit(info)
            else:
                self.no_update.emit()
        except Exception as exc:
            self.failed.emit(FailureDetails.from_exception(exc))


class ReleaseRestoreWorker(QThread):
    available = Signal(object)
    unavailable = Signal(str)
    failed = Signal(object)

    def run(self) -> None:
        try:
            from Modules.Common.AutoUpdater import AutoUpdater

            has_restore, info, message = AutoUpdater("UnHelper", False).check_for_release_restore()
            if has_restore and info:
                self.available.emit(info)
            else:
                self.unavailable.emit(message)
        except Exception as exc:
            self.failed.emit(FailureDetails.from_exception(exc))


class UpdateDownloadWorker(QThread):
    progress = Signal(int)
    completed = Signal(str)
    failed = Signal(object)

    def __init__(self, info):
        super().__init__()
        self.info = info

    def run(self) -> None:
        try:
            from Modules.Common.AutoUpdater import AutoUpdater

            path = AutoUpdater("UnHelper").download_patch(self.info, self.progress.emit)
            self.completed.emit(str(path))
        except Exception as exc:
            self.failed.emit(FailureDetails.from_exception(exc))


class UpdateDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.allow_close = True

    def reject(self) -> None:
        if self.allow_close:
            super().reject()

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.allow_close:
            event.accept()
        else:
            event.ignore()


def _open_product_memory_with_recovery(
    path: str | Path,
    parent: QWidget | None,
) -> ProductMemory | None:
    memory_path = Path(path)
    try:
        return ProductMemory(memory_path)
    except ValueError as exc:
        answer = QMessageBox.question(
            parent,
            "저장된 상품 메모리 복구",
            "저장된 상품 무게/분류 파일이 손상되었거나 현재 버전에서 읽을 수 없습니다.\n\n"
            f"오류: {exc}\n\n"
            "기존 파일을 같은 폴더에 백업한 뒤 저장 목록을 모두 초기화하시겠습니까?\n"
            "초기화하면 표의 SKU 무게를 WMS에서 다시 측정합니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return None

        try:
            backup = ProductMemory.quarantine_and_reset(memory_path)
            memory = ProductMemory(memory_path)
        except Exception as recovery_exc:
            ErrorReportDialog(
                "상품 메모리 복구 실패",
                FailureDetails.from_exception(recovery_exc),
                context={"category": "상품 분류 메모리 복구"},
                parent=parent,
            ).exec()
            return None

        QMessageBox.information(
            parent,
            "상품 메모리 복구 완료",
            "손상된 파일을 백업하고 저장 목록을 초기화했습니다.\n"
            f"백업 파일: {backup}\n\n"
            "현재 표의 SKU 무게는 WMS에서 다시 측정합니다.",
        )
        return memory
    except Exception as exc:
        ErrorReportDialog(
            "상품 메모리 열기 실패",
            FailureDetails.from_exception(exc),
            context={"category": "상품 분류 메모리"},
            parent=parent,
        ).exec()
        return None


class SettingsDialog(QDialog):
    beta_changed = Signal(bool)
    update_requested = Signal()
    history_requested = Signal()
    product_memory_changed = Signal()

    def __init__(
        self,
        settings: QSettings,
        parent=None,
        *,
        memory_path: str | Path | None = None,
    ):
        super().__init__(parent)
        self.settings = settings
        self.memory_path = Path(memory_path) if memory_path else product_memory_path()
        self.credential_store = WMSCredentialStore(settings)
        self._credential_load_error: CredentialError | None = None
        self.setWindowTitle("UnHelper 설정")
        self.setObjectName("SettingsDialog")
        self.resize(700, 650)
        self.setMinimumWidth(640)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(26, 24, 26, 22)
        layout.setSpacing(12)

        title = QLabel("설정")
        title.setObjectName("DialogTitle")
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setObjectName("SettingsScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.viewport().setObjectName("SettingsViewport")
        scroll_content = QWidget()
        scroll_content.setObjectName("SettingsContent")
        content_layout = QVBoxLayout(scroll_content)
        content_layout.setContentsMargins(0, 0, 6, 0)
        content_layout.setSpacing(12)

        file_card = QFrame()
        file_card.setProperty("card", True)
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(18, 16, 18, 16)
        file_layout.setSpacing(9)
        file_title = QLabel("파일 및 데이터")
        file_title.setObjectName("SectionTitle")
        file_layout.addWidget(file_title)

        download_label = QLabel("다운로드 파일 저장 폴더")
        download_label.setObjectName("FieldLabel")
        file_layout.addWidget(download_label)
        download_row = QHBoxLayout()
        self.download_path = QLineEdit(
            str(self.settings.value("download_dir", str(default_download_dir())))
        )
        browse = QPushButton("찾아보기")
        browse.clicked.connect(self._browse)
        download_row.addWidget(self.download_path, 1)
        download_row.addWidget(browse)
        file_layout.addLayout(download_row)

        excel_label = QLabel("Milkrun·트럭 데이터를 반영할 Excel 파일")
        excel_label.setObjectName("FieldLabel")
        file_layout.addWidget(excel_label)
        excel_row = QHBoxLayout()
        self.excel_path = QLineEdit(str(self.settings.value("milkrun_excel_path", "")))
        self.excel_path.setReadOnly(True)
        self.excel_path.setPlaceholderText("파일명에 '입고스케줄관리'가 포함된 Excel을 연결해 주세요")
        excel_browse = QPushButton("연결")
        excel_clear = QPushButton("해제")
        excel_browse.clicked.connect(self._browse_excel)
        excel_clear.clicked.connect(self.excel_path.clear)
        excel_row.addWidget(self.excel_path, 1)
        excel_row.addWidget(excel_browse)
        excel_row.addWidget(excel_clear)
        file_layout.addLayout(excel_row)

        excel_help = QLabel(
            "Milkrun은 Raw_밀크런!C1:P1000, 트럭은 Raw_트럭!C1:U1000의 "
            "값만 지운 뒤 각각 C1부터 값으로 붙여넣습니다."
        )
        excel_help.setWordWrap(True)
        excel_help.setObjectName("HelpText")
        file_layout.addWidget(excel_help)
        content_layout.addWidget(file_card)

        wms_card = QFrame()
        wms_card.setProperty("card", True)
        wms_layout = QVBoxLayout(wms_card)
        wms_layout.setContentsMargins(18, 16, 18, 16)
        wms_layout.setSpacing(9)
        wms_label = QLabel("WMS 로그인")
        wms_label.setObjectName("SectionTitle")
        wms_layout.addWidget(wms_label)
        wms_help = QLabel("저장된 계정은 Windows 사용자 계정으로 암호화되어 보관됩니다.")
        wms_help.setObjectName("HelpText")
        wms_layout.addWidget(wms_help)
        try:
            credentials = self.credential_store.load()
        except CredentialError as exc:
            credentials = WMSCredentials(
                wms_id=str(self.settings.value("wms_id", "")),
                password="",
            )
            self._credential_load_error = exc

        wms_id_row = QHBoxLayout()
        wms_id_label = QLabel("ID")
        wms_id_label.setObjectName("FieldLabel")
        wms_id_label.setMinimumWidth(74)
        wms_id_row.addWidget(wms_id_label)
        self.wms_id_input = QLineEdit(credentials.wms_id)
        self.wms_id_input.setPlaceholderText("WMS ID")
        wms_id_row.addWidget(self.wms_id_input, 1)
        wms_layout.addLayout(wms_id_row)

        wms_password_row = QHBoxLayout()
        wms_password_label = QLabel("Password")
        wms_password_label.setObjectName("FieldLabel")
        wms_password_label.setMinimumWidth(74)
        wms_password_row.addWidget(wms_password_label)
        self.wms_password_input = QLineEdit(credentials.password)
        self.wms_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.wms_password_input.setPlaceholderText("WMS Password")
        self.show_wms_password = QCheckBox("표시")
        self.show_wms_password.toggled.connect(
            lambda checked: self.wms_password_input.setEchoMode(
                QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
            )
        )
        wms_password_row.addWidget(self.wms_password_input, 1)
        wms_password_row.addWidget(self.show_wms_password)
        wms_layout.addLayout(wms_password_row)
        content_layout.addWidget(wms_card)

        app_card = QFrame()
        app_card.setProperty("card", True)
        app_layout = QVBoxLayout(app_card)
        app_layout.setContentsMargins(18, 16, 18, 16)
        app_layout.setSpacing(10)
        app_title = QLabel("상품 메모리 및 업데이트")
        app_title.setObjectName("SectionTitle")
        app_layout.addWidget(app_title)
        product_memory_row = QHBoxLayout()
        product_memory_help = QLabel("저장된 상품은 다음 실행에서 WMS 무게 조회를 생략합니다.")
        product_memory_help.setObjectName("HelpText")
        self.product_memory_button = QPushButton("저장된 상품 분류 목록")
        self.product_memory_button.clicked.connect(self._show_product_memory)
        product_memory_row.addWidget(product_memory_help, 1)
        product_memory_row.addWidget(self.product_memory_button)
        app_layout.addLayout(product_memory_row)

        self.beta_checkbox = QCheckBox("Beta(테스트 릴리즈) 업데이트 받기")
        self.beta_checkbox.setChecked(self.settings.value("use_prerelease", False, type=bool))
        self.beta_checkbox.setToolTip("끄면 최신 정식 릴리즈로 복구할 수 있습니다.")
        app_layout.addWidget(self.beta_checkbox)

        version = QLabel(f"현재 버전: v{CURRENT_VERSION}")
        version.setObjectName("MutedText")
        app_layout.addWidget(version)
        content_layout.addWidget(app_card)
        content_layout.addStretch(1)
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll, 1)

        buttons = QHBoxLayout()
        update = QPushButton("업데이트 확인")
        self.update_history_button = QPushButton("업데이트 내역")
        close = QPushButton("저장하고 닫기")
        close.setObjectName("PrimaryButton")
        update.clicked.connect(self._request_update)
        self.update_history_button.clicked.connect(self.history_requested.emit)
        close.clicked.connect(self.accept)
        buttons.addWidget(update)
        buttons.addWidget(self.update_history_button)
        buttons.addStretch(1)
        buttons.addWidget(close)
        layout.addLayout(buttons)

        if self._credential_load_error is not None:
            QTimer.singleShot(0, self._show_credential_load_error)

    def _browse(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "다운로드 폴더 선택",
            self.download_path.text(),
        )
        if selected:
            self.download_path.setText(selected)

    def _browse_excel(self) -> None:
        current = Path(self.excel_path.text()).expanduser() if self.excel_path.text() else Path.home()
        initial = current.parent if current.is_file() or current.suffix else current
        selected, _filter = QFileDialog.getOpenFileName(
            self,
            "입고 스케줄 Excel 파일 연결",
            str(initial),
            "Excel 통합 문서 (*.xlsx *.xlsm *.xlsb *.xls)",
        )
        if selected:
            self.excel_path.setText(selected)

    def _persist(self) -> tuple[bool, bool] | None:
        previous_beta = self.settings.value("use_prerelease", False, type=bool)
        next_beta = self.beta_checkbox.isChecked()
        path = self.download_path.text().strip() or str(default_download_dir())
        excel_path = self.excel_path.text().strip()
        if excel_path:
            try:
                excel_path = str(MilkrunExcelImporter.validate_target_path(excel_path))
            except ExcelImportError as exc:
                QMessageBox.warning(self, "Excel 파일 확인", str(exc))
                return None
        try:
            self.credential_store.save(
                WMSCredentials(
                    wms_id=self.wms_id_input.text().strip(),
                    password=self.wms_password_input.text(),
                )
            )
        except CredentialError as exc:
            self._show_settings_error("WMS 계정 저장 실패", exc, "WMS 계정 설정")
            return None
        self.settings.setValue("download_dir", path)
        self.settings.setValue("milkrun_excel_path", excel_path)
        self.settings.setValue("use_prerelease", next_beta)
        self.settings.sync()
        return previous_beta != next_beta, next_beta

    def _show_product_memory(self) -> None:
        try:
            memory = _open_product_memory_with_recovery(self.memory_path, self)
            if memory is None:
                return
            dialog = ProductMemoryDialog(memory, self)
            dialog.memory_changed.connect(self.product_memory_changed.emit)
            dialog.exec()
            # Recovery can replace a corrupt cache with an empty one without a
            # ProductMemoryDialog edit, so always reconcile the current RAW table
            # once the list closes as well.
            self.product_memory_changed.emit()
        except Exception as exc:
            self._show_settings_error("저장된 상품 목록 열기 실패", exc, "상품 분류 메모리")

    def _show_credential_load_error(self) -> None:
        if self._credential_load_error is None:
            return
        self._show_settings_error(
            "WMS 비밀번호 불러오기 실패",
            self._credential_load_error,
            "WMS 계정 설정",
        )

    def _show_settings_error(self, title: str, exc: Exception, category: str) -> None:
        ErrorReportDialog(
            title,
            FailureDetails.from_exception(exc),
            context={"category": category},
            parent=self,
        ).exec()

    def _request_update(self) -> None:
        persisted = self._persist()
        if persisted is None:
            return
        changed, next_beta = persisted
        super().accept()
        if changed:
            self.beta_changed.emit(next_beta)
        else:
            self.update_requested.emit()

    def accept(self) -> None:
        persisted = self._persist()
        if persisted is None:
            return
        changed, next_beta = persisted
        super().accept()
        if changed:
            self.beta_changed.emit(next_beta)


class MainWindow(QMainWindow):
    WEIGHT_RETRY_CHECKPOINTS_KEY = "weight_retry_checkpoints_v1"
    MAX_WEIGHT_RETRY_CHECKPOINTS = 8

    def __init__(
        self,
        smoke_test: bool = False,
        *,
        settings: QSettings | None = None,
        product_memory_file: str | Path | None = None,
        snapshot_file: str | Path | None = None,
    ):
        super().__init__()
        self.setObjectName("MainWindow")
        self.smoke_test = smoke_test
        self.settings = settings if settings is not None else QSettings("Mrbinggrae", "UnHelper")
        self._base_date_load_error = ""
        self.milkrun_worker: MilkrunWorker | None = None
        self.weight_worker: ProductWeightWorker | None = None
        self.arrival_worker: ArrivalSequenceWorker | None = None
        self._arrival_snapshot: ArrivalSequenceSnapshot | None = None
        self._arrival_auto_refreshed = False
        self.product_memory_file = (
            Path(product_memory_file) if product_memory_file else product_memory_path()
        )
        self.booking_snapshot_file = (
            Path(snapshot_file) if snapshot_file else booking_snapshot_path()
        )
        self._snapshot_restore_enabled = not smoke_test or snapshot_file is not None
        self.current_products = ()
        self.current_pipeline_result: MilkrunPipelineResult | None = None
        self._products_by_booking: dict[str, tuple] = {"milkrun": (), "truck": ()}
        self._pipeline_results_by_booking: dict[str, MilkrunPipelineResult | None] = {
            "milkrun": None,
            "truck": None,
        }
        self._active_booking_type = "milkrun"
        self._milkrun_group_categories: dict[str, str] = {}
        self._truck_group_categories: dict[str, str] = {}
        self._session_manual_category_skus: set[str] = set()
        self._weight_records: dict[str, ProductMemoryRecord] = {}
        self._weight_failures: dict[str, SkuWeightFailure] = {}
        self._weight_row_errors: dict[str, str] = {}
        self._credential_load_failure: FailureDetails | None = None
        self._pending_weight_summary: ProductWeightSummary | None = None
        self._pending_weight_failure: FailureDetails | None = None
        self._pending_weight_cancel = ""
        self._weight_finalize_pending = False
        self._active_weight_checkpoint_key = ""
        self._pending_full_pipeline_restart = False
        self.update_check_worker: UpdateCheckWorker | None = None
        self.restore_worker: ReleaseRestoreWorker | None = None
        self.update_download_worker: UpdateDownloadWorker | None = None
        self.update_dialog: QDialog | None = None
        self.current_update_info = None
        self.manual_update_check = False
        self._closing_after_cancel = False
        self._automation_cancel_requested = False
        self._closing_after_workers = False

        self.setWindowTitle(f"UnHelper v{CURRENT_VERSION}")
        self.resize(1400, 820)
        self.setMinimumSize(1024, 650)
        self._build_ui()
        if self._snapshot_restore_enabled:
            self._restore_snapshot_for_selected_date(announce=True, clear_if_missing=False)
        if not smoke_test:
            QTimer.singleShot(1200, self.check_for_update)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("Root")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(root)

        header = QFrame()
        header.setObjectName("Header")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(28, 15, 28, 15)
        header_layout.setSpacing(14)
        brand_mark = QLabel("UH")
        brand_mark.setObjectName("BrandMark")
        brand_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_mark.setFixedSize(46, 46)
        header_layout.addWidget(brand_mark)
        title_col = QVBoxLayout()
        title_col.setSpacing(1)
        title = QLabel("UnHelper")
        title.setObjectName("Title")
        version = QLabel(f"v{CURRENT_VERSION} · Shipments/WMS 자동화 도우미")
        version.setObjectName("Version")
        title_col.addWidget(title)
        title_col.addWidget(version)
        header_layout.addLayout(title_col)
        header_layout.addStretch(1)

        base_date_panel = QFrame()
        base_date_panel.setObjectName("BaseDatePanel")
        base_date_layout = QHBoxLayout(base_date_panel)
        base_date_layout.setContentsMargins(10, 5, 10, 5)
        base_date_layout.setSpacing(8)
        base_date_label = QLabel("기준일")
        base_date_label.setObjectName("BaseDateLabel")
        self.base_date_mode = QComboBox()
        self.base_date_mode.setObjectName("BaseDateMode")
        self.base_date_mode.addItem("자동 (실행일)", "auto")
        self.base_date_mode.addItem("수동", "manual")
        self.base_date_mode.setMinimumWidth(122)
        self.manual_base_date = QDateEdit()
        self.manual_base_date.setObjectName("ManualBaseDate")
        self.manual_base_date.setCalendarPopup(True)
        self.manual_base_date.setDisplayFormat("yyyy-MM-dd")
        self.manual_base_date.setMinimumWidth(112)
        self.manual_base_date.setToolTip(
            "Milkrun은 기준일 전날~기준일, 트럭은 기준일 하루만 조회합니다."
        )
        self._load_base_date_controls()
        self.base_date_mode.currentIndexChanged.connect(self._on_base_date_changed)
        self.manual_base_date.dateChanged.connect(self._on_base_date_changed)
        base_date_layout.addWidget(base_date_label)
        base_date_layout.addWidget(self.base_date_mode)
        base_date_layout.addWidget(self.manual_base_date)
        header_layout.addWidget(base_date_panel)

        self.settings_button = QPushButton("설정")
        self.settings_button.setObjectName("SettingsButton")
        self.settings_button.setFixedSize(82, 40)
        self.settings_button.setToolTip("설정")
        self.settings_button.clicked.connect(self.show_settings)
        header_layout.addWidget(self.settings_button)
        layout.addWidget(header)

        self.main_tabs = QTabWidget()
        self.main_tabs.setObjectName("MainTabs")
        self.main_tabs.tabBar().setObjectName("MainTabBar")
        self.main_tabs.addTab(self._build_arrival_tabs(), "입차순번")
        self.main_tabs.addTab(self._build_raw_tabs(), "RAW")
        self.main_tabs.setCurrentIndex(1)
        self.main_tabs.currentChanged.connect(self._on_main_tab_changed)
        layout.addWidget(self.main_tabs, 1)

        footer = QFrame()
        footer.setObjectName("Footer")
        footer_layout = QVBoxLayout(footer)
        footer_layout.setContentsMargins(28, 10, 28, 16)
        footer_layout.setSpacing(8)
        status_row = QHBoxLayout()
        status_dot = QLabel("●")
        status_dot.setObjectName("StatusDot")
        self.status_label = QLabel("대기 중")
        self.status_label.setObjectName("Status")
        self.operation_progress = QProgressBar()
        self.operation_progress.setObjectName("OperationProgress")
        self.operation_progress.setRange(0, 100)
        self.operation_progress.setValue(0)
        self.operation_progress.setFormat("0% 완료 · 100% 남음")
        self.operation_progress.setFixedWidth(170)
        self.operation_progress.setVisible(False)
        self.log_toggle_button = QPushButton()
        self.log_toggle_button.setObjectName("LogToggleButton")
        self.log_toggle_button.setCheckable(True)
        self.log_toggle_button.clicked.connect(self._toggle_log_view)
        self.import_table_button = QPushButton("표 가져오기")
        self.import_table_button.setToolTip(
            "다른 사용자가 내보낸 기준일 RAW 표와 SKU 무게·분류를 가져옵니다."
        )
        self.import_table_button.clicked.connect(self.import_table_snapshot)
        self.export_table_button = QPushButton("표 내보내기")
        self.export_table_button.setToolTip(
            "선택한 기준일의 Milkrun·트럭 표와 관련 SKU 무게·분류를 내보냅니다."
        )
        self.export_table_button.clicked.connect(self.export_table_snapshot)
        self.open_folder_button = QPushButton("다운로드 폴더 열기")
        self.open_folder_button.clicked.connect(self.open_download_folder)
        status_row.addWidget(status_dot)
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.operation_progress)
        status_row.addWidget(self.log_toggle_button)
        status_row.addWidget(self.import_table_button)
        status_row.addWidget(self.export_table_button)
        status_row.addWidget(self.open_folder_button)
        footer_layout.addLayout(status_row)
        self.log_view = QPlainTextEdit()
        self.log_view.setObjectName("LogView")
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumHeight(112)
        self.log_view.setPlaceholderText("진행 로그")
        footer_layout.addWidget(self.log_view)
        self._set_log_expanded(
            self.settings.value("log_expanded", True, type=bool),
            persist=False,
        )
        layout.addWidget(footer)

    def _toggle_log_view(self, checked: bool = False) -> None:
        self._set_log_expanded(bool(checked), persist=True)

    def _set_log_expanded(self, expanded: bool, *, persist: bool) -> None:
        self.log_view.setVisible(expanded)
        self.log_toggle_button.setChecked(expanded)
        self.log_toggle_button.setText("로그 접기 ▲" if expanded else "로그 펼치기 ▼")
        self.log_toggle_button.setToolTip(
            "진행 로그를 숨깁니다." if expanded else "진행 로그를 표시합니다."
        )
        if persist:
            self.settings.setValue("log_expanded", expanded)
            self.settings.sync()

    def _build_arrival_tabs(self) -> QWidget:
        page = QScrollArea()
        page.setObjectName("ArrivalScroll")
        page.setWidgetResizable(True)
        page.setFrameShape(QFrame.Shape.NoFrame)
        page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        page.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        content = QWidget()
        content.setObjectName("Page")
        # Keep the cards from being compressed on short laptop displays.  The
        # surrounding QScrollArea provides vertical scrolling instead.
        content.setMinimumHeight(460)
        # Three summary cards stay on one row like the operations-board
        # reference.  Narrow windows can reach every card via horizontal
        # scrolling instead of squeezing labels and values until they clip.
        content.setMinimumWidth(935)
        outer = QVBoxLayout(content)
        outer.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
        outer.setContentsMargins(28, 16, 28, 18)
        outer.setSpacing(14)

        header = QHBoxLayout()
        heading_column = QVBoxLayout()
        heading_column.setSpacing(2)
        title = QLabel("입차순번 현황")
        title.setObjectName("SectionTitle")
        self.arrival_file_label = QLabel("연결된 Excel의 입차순번 시트를 읽습니다.")
        self.arrival_file_label.setObjectName("SectionDescription")
        heading_column.addWidget(title)
        heading_column.addWidget(self.arrival_file_label)
        header.addLayout(heading_column)
        header.addStretch(1)
        self.arrival_updated_label = QLabel("아직 새로고침하지 않았습니다.")
        self.arrival_updated_label.setObjectName("MutedText")
        header.addWidget(self.arrival_updated_label)
        self.arrival_refresh_button = QPushButton("새로고침")
        self.arrival_refresh_button.setObjectName("PrimaryButton")
        self.arrival_refresh_button.setToolTip(
            "연결된 Excel이 열려 있으면 현재 값을 읽고, 닫혀 있으면 읽기 전용으로 엽니다."
        )
        self.arrival_refresh_button.clicked.connect(self.refresh_arrival_sequence)
        header.addWidget(self.arrival_refresh_button)
        outer.addLayout(header)

        cards = QHBoxLayout()
        cards.setSpacing(12)
        self.arrival_summary_tables: dict[str, QTableWidget] = {}
        self.arrival_detail_tables: dict[str, dict[str, QTableWidget]] = {}
        cards.addWidget(
            self._build_arrival_summary_card(
                "outside_waiting",
                "외부대기",
                ("T", "M", "이관"),
                ("1F", "2F", "전일자"),
            ),
            1,
            Qt.AlignmentFlag.AlignTop,
        )
        cards.addWidget(
            self._build_arrival_summary_card(
                "departure",
                "출차",
                ("T", "M", "이관"),
                ("1F", "2F", "전일자"),
            ),
            1,
            Qt.AlignmentFlag.AlignTop,
        )
        cards.addWidget(
            self._build_arrival_summary_card(
                "floor_targets",
                "각층 목표치",
                ("T", "M"),
                ("1F", "2F", "합계"),
            ),
            1,
            Qt.AlignmentFlag.AlignTop,
        )
        outer.addLayout(cards)

        outer.addStretch(1)
        page.setWidget(content)
        return page

    def _build_arrival_summary_card(
        self,
        key: str,
        title: str,
        columns: tuple[str, ...],
        rows: tuple[str, ...],
    ) -> QFrame:
        card = QFrame()
        card.setObjectName("ArrivalCard")
        card.setProperty("card", True)
        card.setMinimumWidth(285)
        card.setMinimumHeight(560)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(8)
        label = QLabel(f"{title} · 차량 대수")
        label.setObjectName("ArrivalCardTitle")
        layout.addWidget(label)
        table = QTableWidget(len(rows), len(columns))
        table.setObjectName("ArrivalSummaryTable")
        table.setHorizontalHeaderLabels(columns)
        table.setVerticalHeaderLabels(rows)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        table.verticalHeader().setMinimumSectionSize(30)
        table.verticalHeader().setDefaultSectionSize(30)
        table.verticalHeader().setMinimumWidth(64)
        table.horizontalHeader().setFixedHeight(30)
        table.horizontalHeader().setDefaultSectionSize(70)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        for row in range(len(rows)):
            for column in range(len(columns)):
                item = QTableWidgetItem("-")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table.setItem(row, column, item)
        table.setMinimumHeight(122)
        table.setFixedHeight(122)
        self.arrival_summary_tables[key] = table
        layout.addWidget(table)

        detail_panel = QFrame()
        detail_panel.setObjectName("ArrivalDetailPanel")
        detail_layout = QVBoxLayout(detail_panel)
        detail_layout.setContentsMargins(4, 4, 4, 2)
        detail_layout.setSpacing(8)
        detail_tables: dict[str, QTableWidget] = {}

        def add_detail_section(
            section_key: str,
            heading_text: str,
            row_labels: tuple[str, ...],
        ) -> None:
            heading = QLabel(heading_text)
            heading.setObjectName("ArrivalDetailHeading")
            heading.setAlignment(Qt.AlignmentFlag.AlignCenter)
            detail_layout.addWidget(heading)

            detail_table = QTableWidget(len(row_labels), 2)
            detail_table.setObjectName("ArrivalDetailTable")
            detail_table.horizontalHeader().setVisible(False)
            detail_table.verticalHeader().setVisible(False)
            detail_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            detail_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
            detail_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            detail_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            detail_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            detail_table.verticalHeader().setMinimumSectionSize(31)
            detail_table.verticalHeader().setDefaultSectionSize(31)
            detail_table.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.ResizeMode.Stretch
            )
            detail_table.horizontalHeader().setSectionResizeMode(
                1, QHeaderView.ResizeMode.Fixed
            )
            detail_table.setColumnWidth(1, 104)
            for row_index, row_label in enumerate(row_labels):
                label_item = QTableWidgetItem(row_label)
                label_item.setTextAlignment(
                    Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
                )
                label_item.setBackground(QBrush(QColor(COLORS["raised"])))
                label_item.setForeground(QBrush(QColor(COLORS["secondary"])))
                label_font = label_item.font()
                label_font.setBold(True)
                label_item.setFont(label_font)
                detail_table.setItem(row_index, 0, label_item)

                value_item = QTableWidgetItem("0 Pallet")
                value_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                detail_table.setItem(row_index, 1, value_item)
            table_height = (len(row_labels) * 31) + 2
            detail_table.setMinimumHeight(table_height)
            detail_table.setFixedHeight(table_height)
            detail_layout.addWidget(detail_table)
            detail_tables[section_key] = detail_table

        add_detail_section("first", "1F 상세 정보", ("총 팔렛트",))
        add_detail_section(
            "second",
            "2F 상세 정보",
            ("총 팔렛트", "경량", "중량", "고단", "양곡", "미분류"),
        )
        if key != "floor_targets":
            add_detail_section("previous", "전일자", ("총 팔렛트",))

        detail_layout.addStretch(1)
        self.arrival_detail_tables[key] = detail_tables
        layout.addWidget(detail_panel)
        return card

    def _on_main_tab_changed(self, index: int) -> None:
        if index != 0 or self._arrival_auto_refreshed:
            return
        self.refresh_arrival_sequence(automatic=True)

    def refresh_arrival_sequence(
        self,
        _checked: bool = False,
        *,
        automatic: bool = False,
    ) -> None:
        if self._automation_worker_running():
            if not automatic:
                QMessageBox.information(
                    self,
                    "작업 진행 중",
                    "현재 작업이 끝난 뒤 입차순번을 새로고침해 주세요.",
                )
            return
        if not any(self._products_by_booking.values()):
            QMessageBox.information(
                self,
                "RAW 데이터 필요",
                "입차순번 현황을 계산할 RAW 데이터가 없습니다.\n\n"
                "먼저 RAW 탭에서 Milkrun 또는 트럭의 '데이터 얻기'를 실행하거나 "
                "저장된 표를 가져온 뒤 다시 새로고침해 주세요.",
            )
            return

        configured_workbook = str(
            self.settings.value("milkrun_excel_path", "")
        ).strip()
        if not configured_workbook:
            QMessageBox.warning(
                self,
                "Excel 파일 연결 필요",
                "먼저 설정에서 입고스케줄관리 Excel 파일을 연결해 주세요.",
            )
            return
        try:
            workbook_path = ArrivalSequenceReader.validate_target_path(
                configured_workbook
            )
        except ExcelImportError as exc:
            QMessageBox.warning(self, "Excel 파일 확인", str(exc))
            return

        self.arrival_file_label.setText(f"연결된 Excel: {workbook_path.name}")
        self.arrival_updated_label.setText("Excel 값을 읽는 중...")
        self.append_log("연결된 Excel의 입차순번 시트를 새로고침합니다.")
        self.status_label.setText("입차순번 새로고침 중")
        self.arrival_worker = ArrivalSequenceWorker(workbook_path)
        self.arrival_worker.log_updated.connect(self.append_log)
        self.arrival_worker.completed.connect(self._on_arrival_sequence_completed)
        self.arrival_worker.failed.connect(self._on_arrival_sequence_failed)
        self.arrival_worker.cancelled.connect(self._on_arrival_sequence_cancelled)
        self.arrival_worker.finished.connect(self._on_arrival_sequence_finished)
        self._set_automation_working(True)
        self.arrival_worker.start()

    @staticmethod
    def _canonical_raw_booking(value: object, booking_type: str) -> str:
        prefix = "T" if booking_type == "truck" else "M"
        normalized = normalize_booking_number(value, prefix=prefix)
        if normalized:
            digits = normalized[1:].lstrip("0") or "0"
            return f"{prefix}{digits}"
        candidate = str(value or "").strip().replace(" ", "").upper()
        if candidate.startswith(prefix) and candidate[1:].isdigit():
            digits = candidate[1:].lstrip("0") or "0"
            return f"{prefix}{digits}"
        return ""

    def _raw_booking_aggregates(self) -> dict[str, RawBookingAggregate]:
        mutable: dict[str, dict[str, object]] = {}
        valid_categories = {"경량", "중량", "고단", GRAIN_CATEGORY, "?"}
        for booking_type in ("milkrun", "truck"):
            products = self._products_by_booking.get(booking_type, ())
            table = self._table_for_booking(booking_type)
            multi_groups = self._booking_multi_sku_groups(products, booking_type)
            group_categories = self._group_categories_for_booking(booking_type)
            row_to_group = {
                row_index: group_key
                for group_key, rows in multi_groups.items()
                for row_index in rows
            }
            for row_index, product in enumerate(products):
                booking_key = self._canonical_raw_booking(
                    product.dispatch_number,
                    booking_type,
                )
                if not booking_key:
                    continue
                state = mutable.setdefault(
                    booking_key,
                    {
                        "vendors": [],
                        "pallets": Decimal("0"),
                        "categories": {},
                        "missing": 0,
                    },
                )
                vendor = normalize_product_name(product.vendor_name)
                vendors = state["vendors"]
                if vendor and vendor not in vendors:
                    vendors.append(vendor)

                try:
                    pallets = Decimal(str(product.pallet_count).replace(",", ""))
                    if not pallets.is_finite() or pallets < 0:
                        raise ValueError("invalid pallet count")
                except (InvalidOperation, TypeError, ValueError):
                    state["missing"] += 1
                    continue

                state["pallets"] += pallets
                group_key = row_to_group.get(row_index)
                if group_key is not None:
                    category = group_categories.get(group_key, "?")
                else:
                    button = table.cellWidget(row_index, 9)
                    category = button.text() if isinstance(button, QPushButton) else "?"
                if category not in valid_categories:
                    category = "?"
                categories = state["categories"]
                categories[category] = categories.get(category, Decimal("0")) + pallets

        return {
            booking_key: RawBookingAggregate(
                booking_key=booking_key,
                vendor_names=tuple(state["vendors"]),
                pallet_count=state["pallets"],
                category_pallets=tuple(
                    (category, state["categories"].get(category, Decimal("0")))
                    for category in ("경량", "중량", "고단", GRAIN_CATEGORY, "?")
                    if state["categories"].get(category, Decimal("0")) != 0
                ),
                missing_pallet_rows=int(state["missing"]),
            )
            for booking_key, state in mutable.items()
        }

    def _on_arrival_sequence_completed(
        self,
        snapshot: ArrivalSequenceSnapshot,
    ) -> None:
        self._arrival_snapshot = snapshot
        self._arrival_auto_refreshed = True
        self._render_arrival_sequence(snapshot)
        stamp = snapshot.refreshed_at.strftime("%Y-%m-%d %H:%M:%S")
        self.arrival_updated_label.setText(f"마지막 새로고침 {stamp}")
        self.status_label.setText(f"입차순번 새로고침 완료 · 차량 {len(snapshot.entries)}대")
        self.append_log(f"입차순번 현황을 차량 {len(snapshot.entries)}대로 갱신했습니다.")

    def _render_arrival_sequence(self, snapshot: ArrivalSequenceSnapshot) -> None:
        summaries = {
            "outside_waiting": snapshot.summary.outside_waiting,
            "departure": snapshot.summary.departure,
            "floor_targets": snapshot.summary.floor_targets,
        }
        for key, values in summaries.items():
            table = self.arrival_summary_tables[key]
            for row_index in range(table.rowCount()):
                for column_index in range(table.columnCount()):
                    try:
                        value = values[row_index][column_index]
                    except IndexError:
                        value = ""
                    item = table.item(row_index, column_index)
                    item.setText(str(value).strip() or "-")

        raw_bookings = self._raw_booking_aggregates()
        vehicles = build_arrival_vehicles(snapshot, raw_bookings)
        floor_breakdowns = build_floor_target_breakdowns(snapshot, raw_bookings)

        def pallet_value_text(value: Decimal) -> str:
            return f"{self._format_decimal(value)} Pallet"

        def status_breakdown_note(breakdown) -> str:
            notes: list[str] = []
            if breakdown.missing_pallet_vehicles:
                notes.append(f"팔렛트 수 미입력 {breakdown.missing_pallet_vehicles}대")
            if breakdown.unmapped_bookings:
                notes.append("층 미매핑 " + ", ".join(breakdown.unmapped_bookings))
            notes.extend(breakdown.notes)
            return " · ".join(notes)

        def render_detail_tables(
            key: str,
            first,
            second,
            previous=None,
            *,
            first_note: str = "",
            second_note: str = "",
            previous_note: str = "",
        ) -> None:
            tables = self.arrival_detail_tables[key]
            section_values = {
                "first": ((first.pallet_count,), first_note),
                "second": (
                    (
                        second.pallet_count,
                        second.categories.get("경량", Decimal("0")),
                        second.categories.get("중량", Decimal("0")),
                        second.categories.get("고단", Decimal("0")),
                        second.categories.get(GRAIN_CATEGORY, Decimal("0")),
                        second.categories.get("?", Decimal("0")),
                    ),
                    second_note,
                ),
            }
            if previous is not None and "previous" in tables:
                section_values["previous"] = ((previous.pallet_count,), previous_note)

            for section_key, (values, note) in section_values.items():
                table = tables[section_key]
                for row_index, value in enumerate(values):
                    table.item(row_index, 1).setText(pallet_value_text(value))
                    table.item(row_index, 0).setToolTip(note)
                    table.item(row_index, 1).setToolTip(note)

        for key, status in (("outside_waiting", "외부대기"), ("departure", "출차")):
            breakdowns = build_status_pallet_breakdowns(vehicles, status=status)
            first, second, previous = breakdowns
            render_detail_tables(
                key,
                first,
                second,
                previous,
                first_note=status_breakdown_note(first),
                second_note=status_breakdown_note(second),
                previous_note=status_breakdown_note(previous),
            )

        floor_notes: list[str] = []
        for row_index, breakdown in enumerate(floor_breakdowns):
            notes: list[str] = []
            if breakdown.missing_pallet_rows:
                notes.append(f"팔렛트 수 미입력 {breakdown.missing_pallet_rows}행")
            if breakdown.unmapped_bookings:
                notes.append("RAW 미매칭 " + ", ".join(breakdown.unmapped_bookings))
            if row_index == 0 and breakdown.unassigned_raw_bookings:
                missing_keys = breakdown.unassigned_raw_bookings
                preview = ", ".join(missing_keys[:10])
                if len(missing_keys) > 10:
                    preview += f" 외 {len(missing_keys) - 10}건"
                notes.append("층 미매핑 " + preview)
            floor_notes.append(" · ".join(notes))
        first_floor, second_floor = floor_breakdowns
        render_detail_tables(
            "floor_targets",
            first_floor,
            second_floor,
            first_note=floor_notes[0],
            second_note=floor_notes[1],
        )

    def _refresh_arrival_from_current_raw(self) -> None:
        """Recalculate pallet details without re-reading the linked workbook."""

        if self._arrival_snapshot is not None:
            self._render_arrival_sequence(self._arrival_snapshot)


    def _on_arrival_sequence_failed(self, failure: FailureDetails | object) -> None:
        details = FailureDetails.coerce(failure)
        self.arrival_updated_label.setText("새로고침 실패")
        self.status_label.setText("입차순번 새로고침 실패")
        self.append_log(f"[입차순번 오류] {details.summary}")
        if not self._closing_after_cancel:
            self._show_error_dialog(
                "입차순번 새로고침 실패",
                details,
                category="입차순번 Excel 읽기",
            )

    def _on_arrival_sequence_cancelled(self, message: str) -> None:
        self.arrival_updated_label.setText("새로고침 취소")
        self.status_label.setText("입차순번 새로고침 취소")
        self.append_log(message)

    def _on_arrival_sequence_finished(self) -> None:
        worker = self.arrival_worker
        self.arrival_worker = None
        if worker is not None:
            worker.deleteLater()
        self._set_automation_working(self._automation_worker_running())
        if self._closing_after_cancel and not self._automation_worker_running():
            QTimer.singleShot(0, self.close)

    def _build_raw_tabs(self) -> QTabWidget:
        self._raw_excel_apply_checkboxes: list[QCheckBox] = []
        self._apply_raw_to_excel = self.settings.value(
            "apply_raw_to_excel",
            True,
            type=bool,
        )
        self.raw_tabs = QTabWidget()
        self.raw_tabs.setObjectName("RawTabs")
        self.raw_tabs.tabBar().setObjectName("SubTabBar")
        self.raw_tabs.addTab(self._build_raw_truck(), "트럭")
        self.raw_tabs.addTab(self._build_raw_milkrun(), "Milkrun")
        self.raw_tabs.setCurrentIndex(1)
        return self.raw_tabs

    def _build_raw_milkrun(self) -> QWidget:
        return self._build_raw_booking_page("milkrun")

    def _build_raw_truck(self) -> QWidget:
        return self._build_raw_booking_page("truck")

    def _build_raw_booking_page(self, booking_type: str) -> QWidget:
        is_truck = booking_type == "truck"
        page = QWidget()
        page.setObjectName("Page")
        outer = QVBoxLayout(page)
        outer.setContentsMargins(28, 12, 28, 14)
        outer.setSpacing(12)

        section_row = QHBoxLayout()
        section_title = QLabel("트럭 데이터" if is_truck else "Milkrun 데이터")
        section_title.setObjectName("SectionTitle")
        section_description = QLabel(
            (
                "선택한 기준일의 트럭 예약과 WMS 상품 무게를 예약번호 기준으로 정리합니다."
                if is_truck
                else "Shipments 예약 정보와 WMS 상품 무게를 선택한 기준일로 정리합니다."
            )
        )
        section_description.setObjectName("SectionDescription")
        section_row.addWidget(section_title)
        section_row.addStretch(1)
        search_bar = QFrame()
        search_bar.setObjectName("TableSearchBar")
        search_layout = QHBoxLayout(search_bar)
        search_layout.setContentsMargins(12, 6, 12, 6)
        search_layout.setSpacing(8)
        search_label = QLabel("검색")
        search_label.setObjectName("FieldLabel")
        search_input = QLineEdit()
        search_input.setObjectName("TableSearchInput")
        search_input.setClearButtonEnabled(True)
        search_input.setPlaceholderText(
            "거래처 이름, 예약번호, SKU ID, SKU 명 검색"
            if is_truck
            else "거래처 이름, 배차번호, SKU ID, SKU 명 검색"
        )
        search_input.textChanged.connect(
            lambda text, kind=booking_type: self._filter_booking_table(kind, text)
        )
        search_layout.addWidget(search_label)
        search_layout.addWidget(search_input, 1)
        search_bar.setMinimumWidth(420)
        search_bar.setMaximumWidth(620)
        section_row.addWidget(search_bar, 0, Qt.AlignmentFlag.AlignVCenter)
        outer.addLayout(section_row)
        outer.addWidget(section_description)

        data_card = QFrame()
        data_card.setObjectName("DataCard")
        data_layout = QVBoxLayout(data_card)
        data_layout.setContentsMargins(0, 0, 0, 0)
        data_layout.setSpacing(0)
        table = QTableWidget(0, 10)
        table.setObjectName("RawTable")
        table.setHorizontalHeaderLabels(
            (
                [
                    "거래처 이름",
                    "예약번호",
                    "팔렛트 수",
                    "유닛 수",
                    "팔렛트당 유닛",
                    "SKU ID",
                    "SKU 명",
                    "상품 무게(g)",
                    "1팔렛트 무게(kg)",
                    "분류",
                ]
                if is_truck
                else [
                "거래처 이름",
                "배차번호",
                "팔렛트 수",
                "유닛 수",
                "팔렛트당 유닛",
                "SKU ID",
                "SKU 명",
                "상품 무게(g)",
                "1팔렛트 무게(kg)",
                "분류",
                ]
            )
        )
        table.setWordWrap(False)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in range(1, 9):
            table.horizontalHeader().setSectionResizeMode(
                column,
                QHeaderView.ResizeMode.Stretch
                if column == 6
                else QHeaderView.ResizeMode.ResizeToContents,
            )
        table.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeMode.Fixed)
        table.setColumnWidth(9, 84)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(42)
        table.horizontalHeader().setMinimumHeight(42)
        table.setAlternatingRowColors(True)
        data_layout.addWidget(table, 1)
        outer.addWidget(data_card, 1)

        actions = QHBoxLayout()
        apply_excel_checkbox = QCheckBox("연결된 Excel RAW 시트에 반영")
        apply_excel_checkbox.setObjectName("ApplyRawToExcelCheckbox")
        apply_excel_checkbox.setChecked(self._apply_raw_to_excel)
        apply_excel_checkbox.setToolTip(
            "체크를 끄면 다운로드·상품 상세·WMS 무게 조회는 진행하지만 "
            "연결된 Excel의 RAW 시트는 지우거나 저장하지 않습니다."
        )
        apply_excel_checkbox.toggled.connect(self._on_apply_raw_to_excel_toggled)
        self._raw_excel_apply_checkboxes.append(apply_excel_checkbox)
        actions.addWidget(apply_excel_checkbox)
        actions.addStretch(1)
        get_button = QPushButton("데이터 얻기")
        get_button.setObjectName("PrimaryButton")
        get_button.setMinimumWidth(210)
        get_button.clicked.connect(
            self.start_truck_download if is_truck else self.start_milkrun_download
        )
        stop_button = QPushButton("작업 중지")
        stop_button.setObjectName("StopButton")
        stop_button.setVisible(False)
        stop_button.clicked.connect(self.cancel_milkrun_download)
        actions.addWidget(get_button)
        actions.addWidget(stop_button)
        actions.addStretch(1)
        outer.addLayout(actions)

        if is_truck:
            self.truck_table = table
            self.truck_search_input = search_input
            self.truck_get_data_button = get_button
            self.truck_stop_button = stop_button
            self.truck_apply_excel_checkbox = apply_excel_checkbox
        else:
            # Keep the original public attribute names for the Milkrun tests and
            # existing integrations while Truck owns a separate table/state.
            self.raw_table = table
            self.milkrun_search_input = search_input
            self.get_data_button = get_button
            self.stop_button = stop_button
            self.milkrun_apply_excel_checkbox = apply_excel_checkbox
        return page

    def _on_apply_raw_to_excel_toggled(self, checked: bool) -> None:
        self._apply_raw_to_excel = bool(checked)
        for checkbox in self._raw_excel_apply_checkboxes:
            if checkbox.isChecked() == self._apply_raw_to_excel:
                continue
            checkbox.blockSignals(True)
            checkbox.setChecked(self._apply_raw_to_excel)
            checkbox.blockSignals(False)
        self.settings.setValue("apply_raw_to_excel", self._apply_raw_to_excel)
        self.settings.sync()

    @staticmethod
    def _placeholder(message: str) -> QWidget:
        page = QWidget()
        page.setObjectName("Page")
        layout = QVBoxLayout(page)
        label = QLabel(message)
        label.setObjectName("PlaceholderText")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
        return page

    def _table_for_booking(self, booking_type: str) -> QTableWidget:
        return self.truck_table if booking_type == "truck" else self.raw_table

    @staticmethod
    def _booking_label(booking_type: str) -> str:
        return "트럭" if booking_type == "truck" else "Milkrun"

    def start_milkrun_download(self) -> None:
        self._start_booking_download("milkrun")

    def start_truck_download(self) -> None:
        self._start_booking_download("truck")

    def _load_base_date_controls(self) -> None:
        saved_mode = str(self.settings.value("base_date_mode", "auto")).strip().lower()
        saved_date = QDate.fromString(
            str(self.settings.value("manual_base_date", "")),
            Qt.DateFormat.ISODate,
        )
        invalid_mode = saved_mode not in {"auto", "manual"}
        invalid_manual_date = saved_mode == "manual" and not saved_date.isValid()
        if invalid_mode or invalid_manual_date:
            self.base_date_mode.addItem("선택 필요", "invalid")
            self.base_date_mode.setCurrentIndex(self.base_date_mode.count() - 1)
            self._base_date_load_error = (
                "저장된 기준일 선택 방식을 읽을 수 없습니다. "
                if invalid_mode
                else "저장된 수동 기준일을 읽을 수 없습니다. "
            ) + "메인 화면에서 자동 또는 수동 기준일을 다시 선택해 주세요."
        else:
            self.base_date_mode.setCurrentIndex(1 if saved_mode == "manual" else 0)
        self.manual_base_date.setDate(
            saved_date if saved_date.isValid() else QDate.currentDate()
        )
        self._sync_base_date_controls()

    def _on_base_date_changed(self, *_args) -> None:
        mode = self.base_date_mode.currentData()
        self._sync_base_date_controls()
        if mode not in {"auto", "manual"}:
            return
        self._base_date_load_error = ""
        invalid_index = self.base_date_mode.findData("invalid")
        if invalid_index >= 0:
            self.base_date_mode.blockSignals(True)
            self.base_date_mode.removeItem(invalid_index)
            self.base_date_mode.blockSignals(False)
        self.settings.setValue("base_date_mode", mode)
        self.settings.setValue(
            "manual_base_date",
            self.manual_base_date.date().toString(Qt.DateFormat.ISODate),
        )
        self.settings.sync()
        if self._snapshot_restore_enabled and not self._automation_worker_running():
            self._restore_snapshot_for_selected_date(
                announce=True,
                clear_if_missing=True,
            )

    def _sync_base_date_controls(self) -> None:
        working = self._automation_worker_running()
        self.base_date_mode.setEnabled(not working)
        self.manual_base_date.setEnabled(
            not working and self.base_date_mode.currentData() == "manual"
        )

    def _configured_base_date(self) -> date | None:
        mode = self.base_date_mode.currentData()
        if mode not in {"auto", "manual"}:
            raise ValueError(
                self._base_date_load_error
                or "메인 화면에서 자동 또는 수동 기준일을 선택해 주세요."
            )
        if mode == "auto":
            return None
        selected = self.manual_base_date.date()
        if not selected.isValid():
            raise ValueError("메인 화면에서 올바른 수동 기준일을 선택해 주세요.")
        return date(selected.year(), selected.month(), selected.day())

    def _selected_snapshot_date(self) -> date | None:
        mode = self.base_date_mode.currentData()
        if mode == "auto":
            return date.today()
        if mode != "manual":
            return None
        selected = self.manual_base_date.date()
        if not selected.isValid():
            return None
        return date(selected.year(), selected.month(), selected.day())

    def _clear_booking_snapshot_view(self) -> None:
        for booking_type in ("milkrun", "truck"):
            table = self._table_for_booking(booking_type)
            table.clearSpans()
            table.setRowCount(0)
            self._products_by_booking[booking_type] = ()
            self._pipeline_results_by_booking[booking_type] = None
            self._group_categories_for_booking(booking_type).clear()
        self.current_products = ()
        self.current_pipeline_result = None
        self._weight_records.clear()
        self._weight_failures.clear()
        self._weight_row_errors.clear()

    def _restore_booking_snapshot(
        self,
        snapshot: BookingDateSnapshot,
        *,
        announce: bool,
    ) -> None:
        self._clear_booking_snapshot_view()
        self._populate_booking_products(snapshot.truck_products, "truck")
        self._populate_booking_products(snapshot.milkrun_products, "milkrun")
        self._active_booking_type = (
            "truck" if self.raw_tabs.currentIndex() == 0 else "milkrun"
        )
        self.current_products = self._products_by_booking[self._active_booking_type]
        self.current_pipeline_result = None
        if any(self._products_by_booking.values()):
            self._refresh_current_product_memory(announce=False)
        if announce:
            self.append_log(
                f"저장된 {snapshot.base_date:%Y-%m-%d} RAW 표를 복원했습니다: "
                f"Milkrun {len(snapshot.milkrun_products)}행 · "
                f"트럭 {len(snapshot.truck_products)}행"
            )
            self.status_label.setText(
                f"{snapshot.base_date:%Y-%m-%d} 저장 표 복원됨"
            )

    def _restore_snapshot_for_selected_date(
        self,
        *,
        announce: bool,
        clear_if_missing: bool,
    ) -> bool:
        selected_date = self._selected_snapshot_date()
        if selected_date is None:
            return False
        try:
            snapshot = BookingSnapshotStore(self.booking_snapshot_file).get(selected_date)
        except Exception as exc:
            if clear_if_missing:
                self._clear_booking_snapshot_view()
            self.append_log(f"[저장 표 읽기 오류] {exc}")
            self.status_label.setText("저장 표 읽기 실패")
            return False
        if snapshot is None:
            if clear_if_missing:
                self._clear_booking_snapshot_view()
                self.append_log(
                    f"{selected_date:%Y-%m-%d}에 저장된 RAW 표가 없습니다."
                )
                self.status_label.setText("선택한 기준일의 저장 표 없음")
            return False
        self._restore_booking_snapshot(snapshot, announce=announce)
        return True

    def _save_booking_snapshot(
        self,
        base_date: date,
        booking_type: str,
        products,
    ) -> None:
        try:
            BookingSnapshotStore(self.booking_snapshot_file).save_table(
                base_date,
                booking_type,
                tuple(products),
            )
            self.append_log(
                f"{base_date:%Y-%m-%d} {self._booking_label(booking_type)} 표를 "
                "로컬에 저장했습니다. 최근 2개 기준일만 유지합니다."
            )
        except Exception as exc:
            self.append_log(f"[RAW 표 자동 저장 오류] {exc}")
            QMessageBox.warning(
                self,
                "RAW 표 저장 실패",
                "표 표시는 완료했지만 재실행 복원용 파일을 저장하지 못했습니다.\n\n"
                f"{exc}",
            )

    def _set_manual_snapshot_date(self, base_date: date) -> None:
        self.base_date_mode.blockSignals(True)
        self.manual_base_date.blockSignals(True)
        try:
            self.base_date_mode.setCurrentIndex(
                self.base_date_mode.findData("manual")
            )
            self.manual_base_date.setDate(
                QDate(base_date.year, base_date.month, base_date.day)
            )
        finally:
            self.manual_base_date.blockSignals(False)
            self.base_date_mode.blockSignals(False)
        self._base_date_load_error = ""
        self.settings.setValue("base_date_mode", "manual")
        self.settings.setValue("manual_base_date", base_date.isoformat())
        self.settings.sync()
        self._sync_base_date_controls()

    def export_table_snapshot(self) -> None:
        selected_date = self._selected_snapshot_date()
        if selected_date is None:
            QMessageBox.warning(
                self,
                "기준일 확인",
                "내보낼 기준일을 자동 또는 수동으로 먼저 선택해 주세요.",
            )
            return
        store = BookingSnapshotStore(self.booking_snapshot_file)
        try:
            snapshot = store.get(selected_date)
            if snapshot is None:
                QMessageBox.information(
                    self,
                    "내보낼 표 없음",
                    f"{selected_date:%Y-%m-%d}에 저장된 RAW 표가 없습니다.",
                )
                return
            sku_ids = {
                product.sku_id
                for booking_type in ("milkrun", "truck")
                for product in snapshot.products_for(booking_type)
            }
            memory = ProductMemory(self.product_memory_file)
            product_payload = memory.export_payload(sku_ids)
        except Exception as exc:
            self._show_error_dialog(
                "RAW 표 내보내기 실패",
                FailureDetails.from_exception(exc),
                category="RAW 표 저장·공유",
            )
            return

        default_path = default_download_dir() / (
            f"UnHelper_RAW_표_{selected_date:%Y%m%d}.json"
        )
        path, _selected_filter = QFileDialog.getSaveFileName(
            self,
            "RAW 표 내보내기",
            str(default_path),
            "UnHelper RAW 표 (*.json)",
        )
        if not path:
            return
        destination = Path(path)
        if destination.suffix.lower() != ".json":
            destination = destination.with_suffix(".json")
        try:
            exported = store.export_bundle(
                selected_date,
                destination,
                product_payload,
            )
        except Exception as exc:
            self._show_error_dialog(
                "RAW 표 내보내기 실패",
                FailureDetails.from_exception(exc),
                category="RAW 표 저장·공유",
            )
            return
        QMessageBox.information(
            self,
            "RAW 표 내보내기 완료",
            f"{selected_date:%Y-%m-%d} 표를 저장했습니다.\n\n{exported}",
        )

    @staticmethod
    def _memory_record_description(record: ProductMemoryRecord) -> str:
        weight = (
            f"{record.weight_grams}g"
            if record.weight_grams is not None
            else "미측정"
        )
        category = record.effective_category or "미분류"
        return (
            f"상품명: {record.product_name or '-'}\n"
            f"무게: {weight}\n"
            f"분류: {category}"
        )

    @staticmethod
    def _memory_records_equivalent(
        existing: ProductMemoryRecord,
        incoming: ProductMemoryRecord,
    ) -> bool:
        return (
            existing.sku_id == incoming.sku_id
            and existing.product_name == incoming.product_name
            and existing.weight_grams == incoming.weight_grams
            and existing.automatic_category == incoming.automatic_category
            and existing.category_override == incoming.category_override
            and existing.boxes_per_pallet == incoming.boxes_per_pallet
            and existing.pallet_weight_kg == incoming.pallet_weight_kg
        )

    def _ask_duplicate_memory_action(
        self,
        existing: ProductMemoryRecord,
        incoming: ProductMemoryRecord,
        *,
        index: int,
        total: int,
    ) -> str:
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Icon.Question)
        dialog.setWindowTitle("중복 SKU 상품 메모리 확인")
        dialog.setText(
            f"SKU {incoming.sku_id}가 이미 저장되어 있습니다. ({index}/{total})\n"
            "가져온 값으로 덮어쓸까요?"
        )
        dialog.setInformativeText(
            "[현재 저장값]\n"
            f"{self._memory_record_description(existing)}\n\n"
            "[가져올 값]\n"
            f"{self._memory_record_description(incoming)}"
        )
        overwrite_button = dialog.addButton(
            "덮어쓰기",
            QMessageBox.ButtonRole.AcceptRole,
        )
        keep_button = dialog.addButton(
            "기존 유지",
            QMessageBox.ButtonRole.RejectRole,
        )
        cancel_button = dialog.addButton(
            "가져오기 취소",
            QMessageBox.ButtonRole.DestructiveRole,
        )
        dialog.setDefaultButton(keep_button)
        dialog.setEscapeButton(cancel_button)
        dialog.exec()
        clicked = dialog.clickedButton()
        if clicked is overwrite_button:
            return "overwrite"
        if clicked is keep_button:
            return "keep"
        return "cancel"

    def import_table_snapshot(self) -> None:
        path, _selected_filter = QFileDialog.getOpenFileName(
            self,
            "RAW 표 가져오기",
            str(default_download_dir()),
            "UnHelper RAW 표 (*.json)",
        )
        if not path:
            return
        try:
            snapshot, product_payload = BookingSnapshotStore.read_bundle(path)
            records = ProductMemory.validate_payload(product_payload)
            memory = ProductMemory(self.product_memory_file)
            store = BookingSnapshotStore(self.booking_snapshot_file)
            duplicates = tuple(
                (existing, record)
                for record in records
                if (existing := memory.get(record.sku_id)) is not None
                and not self._memory_records_equivalent(existing, record)
            )
            overwrite_sku_ids: set[str] = set()
            for index, (existing, incoming) in enumerate(duplicates, start=1):
                action = self._ask_duplicate_memory_action(
                    existing,
                    incoming,
                    index=index,
                    total=len(duplicates),
                )
                if action == "cancel":
                    self.append_log("RAW 표 가져오기를 사용자가 취소했습니다.")
                    return
                if action == "overwrite":
                    overwrite_sku_ids.add(incoming.sku_id)
            store.save_snapshot(snapshot)
            summary = memory.import_records(
                records,
                overwrite_sku_ids=overwrite_sku_ids,
            )
            self._set_manual_snapshot_date(snapshot.base_date)
            self._restore_snapshot_for_selected_date(
                announce=True,
                clear_if_missing=True,
            )
        except Exception as exc:
            self._show_error_dialog(
                "RAW 표 가져오기 실패",
                FailureDetails.from_exception(exc),
                category="RAW 표 저장·공유",
            )
            return
        QMessageBox.information(
            self,
            "RAW 표 가져오기 완료",
            f"{snapshot.base_date:%Y-%m-%d} 표를 가져왔습니다.\n"
            f"상품 메모리 추가 {summary.added}개 · "
            f"덮어쓰기 {summary.overwritten}개 · 기존 값 유지 {summary.skipped}개",
        )

    def _start_booking_download(
        self,
        booking_type: str,
        *,
        retry_mode: str | None = None,
    ) -> None:
        if self._automation_worker_running():
            return
        is_truck = booking_type == "truck"
        booking_label = "트럭" if is_truck else "Milkrun"
        driver = chromedriver_path()
        if not driver.is_file():
            QMessageBox.critical(
                self,
                "ChromeDriver 없음",
                f"ChromeDriver를 찾을 수 없습니다.\n{driver}",
            )
            return
        try:
            base_date = self._configured_base_date()
        except ValueError as exc:
            QMessageBox.warning(self, "기준일 확인", str(exc))
            self.base_date_mode.setFocus()
            return
        effective_base_date = base_date or date.today()
        checkpoint_key = self._weight_checkpoint_key(
            booking_type=booking_type,
            base_date=effective_base_date,
        )
        checkpoints = self._read_weight_retry_checkpoints()
        existing_products = self._products_by_booking.get(booking_type, ())
        if retry_mode is None and existing_products:
            sku_ids = self._weight_checkpoint_skus(existing_products)
            checkpoint = checkpoints.get(checkpoint_key)
            if checkpoint is not None:
                cached_count = self._checkpoint_completed_count(
                    checkpoint,
                    existing_products,
                )
            else:
                cached_count = self._stored_weight_count(existing_products)
            retry_mode = self._ask_weight_retry_action(
                cached_count=cached_count,
                total_count=len(sku_ids),
                has_checkpoint=checkpoint is not None,
            )
            if retry_mode == "cancel":
                self.status_label.setText(
                    "미완료 WMS 무게 측정은 다음에 이어서 진행할 수 있습니다"
                    if checkpoint is not None
                    else "현재 RAW 표를 그대로 유지했습니다"
                )
                return
            if retry_mode == "resume":
                self._active_booking_type = booking_type
                self.current_products = tuple(existing_products)
                self.current_pipeline_result = self._pipeline_results_by_booking.get(
                    booking_type
                )
                self._automation_cancel_requested = False
                self._start_weight_lookup(self.current_products, retry_mode="resume")
                return

        configured_workbook = str(self.settings.value("milkrun_excel_path", "")).strip()
        if not configured_workbook:
            QMessageBox.warning(
                self,
                "Excel 파일 연결 필요",
                "먼저 설정에서 Raw_밀크런 및 Raw_트럭 시트가 있는 Excel 파일을 연결해 주세요.",
            )
            self.show_settings()
            return
        try:
            if is_truck:
                from Modules.Excel.TruckExcelImporter import TruckExcelImporter

                target_workbook = TruckExcelImporter.validate_target_path(configured_workbook)
            else:
                target_workbook = MilkrunExcelImporter.validate_target_path(configured_workbook)
        except ExcelImportError as exc:
            QMessageBox.warning(self, "Excel 파일 확인", str(exc))
            return

        download_dir = Path(
            str(self.settings.value("download_dir", str(default_download_dir())))
        ).expanduser()
        if is_truck:
            from Modules.Shipments.TruckDownloader import TruckDownloadRequest

            request = TruckDownloadRequest(
                download_dir=download_dir,
                center_name="안산2",
                base_date=base_date,
            )
        else:
            request = MilkrunDownloadRequest(
                download_dir=download_dir,
                center_name="안산2",
                base_date=base_date,
            )
        self.log_view.clear()
        table = self._table_for_booking(booking_type)
        table.setRowCount(0)
        self._active_booking_type = booking_type
        self._products_by_booking[booking_type] = ()
        self._pipeline_results_by_booking[booking_type] = None
        self._group_categories_for_booking(booking_type).clear()
        self._session_manual_category_skus.clear()
        self._automation_cancel_requested = False
        self.current_products = ()
        self.current_pipeline_result = None
        self._pending_weight_summary = None
        self._pending_weight_failure = None
        self._pending_weight_cancel = ""
        self._credential_load_failure = None
        self._weight_finalize_pending = False
        self._pending_full_pipeline_restart = retry_mode == "restart"
        apply_to_excel = self._apply_raw_to_excel
        if apply_to_excel:
            self.append_log(f"{booking_label} 텍스트 다운로드 및 Excel 반영 작업을 시작합니다.")
        else:
            self.append_log(
                f"{booking_label} 텍스트 다운로드 작업을 시작합니다. "
                "연결된 Excel RAW 시트 반영은 제외합니다."
            )
        self.append_log(
            f"조회 기준일: {base_date:%Y-%m-%d} (수동)"
            if base_date is not None
            else "조회 기준일: 실행일 자동"
        )
        self.append_log(f"연결된 Excel: {target_workbook}")
        self._set_automation_working(True)
        self.milkrun_worker = MilkrunWorker(
            request,
            driver,
            target_workbook,
            booking_type=booking_type,
            apply_to_excel=apply_to_excel,
        )
        self.milkrun_worker.log_updated.connect(self.append_log)
        self.milkrun_worker.detail_progress.connect(self._on_detail_progress)
        self.milkrun_worker.completed.connect(self._on_milkrun_completed)
        self.milkrun_worker.excel_failed.connect(self._on_milkrun_excel_failed)
        self.milkrun_worker.excel_close_required.connect(self._on_excel_close_required)
        self.milkrun_worker.detail_failed.connect(self._on_milkrun_detail_failed)
        self.milkrun_worker.detail_cancelled.connect(self._on_milkrun_detail_cancelled)
        self.milkrun_worker.failed.connect(self._on_milkrun_failed)
        self.milkrun_worker.cancelled.connect(self._on_milkrun_cancelled)
        self.milkrun_worker.finished.connect(self._on_milkrun_finished)
        self.milkrun_worker.start()

    def cancel_milkrun_download(self) -> None:
        requested = (
            self.milkrun_worker is not None
            or self.weight_worker is not None
            or self.arrival_worker is not None
        )
        if self.milkrun_worker is not None:
            self.milkrun_worker.request_cancel()
        if self.weight_worker is not None:
            self.weight_worker.request_cancel()
        if self.arrival_worker is not None:
            self.arrival_worker.request_cancel()
        if requested:
            self._automation_cancel_requested = True
            self.append_log("작업 중지를 요청했습니다.")
            self.status_label.setText("작업 중지 중...")
            self.stop_button.setEnabled(False)
            self.truck_stop_button.setEnabled(False)

    def _on_milkrun_completed(self, result) -> None:
        if self._closing_after_cancel:
            self.append_log("종료 요청이 처리 중이므로 WMS 무게 조회를 시작하지 않습니다.")
            return
        if self._automation_cancel_requested:
            self.append_log("작업 중지 요청이 처리되어 WMS 무게 조회를 시작하지 않습니다.")
            self.status_label.setText("작업 취소됨")
            return
        booking_type = getattr(result, "booking_type", self._active_booking_type)
        self._active_booking_type = booking_type
        excel = result.excel
        daily = result.daily_inbound
        self.current_pipeline_result = result
        self._pipeline_results_by_booking[booking_type] = result
        self._populate_booking_products(daily.products, booking_type)
        snapshot_date = getattr(result, "base_date", None) or self._selected_snapshot_date()
        if snapshot_date is not None:
            self._save_booking_snapshot(
                snapshot_date,
                booking_type,
                daily.products,
            )
        self.append_log(f"다운로드 완료: {excel.source_file}")
        if getattr(excel, "target_updated", True):
            self.append_log(
                f"Excel 반영 완료: {excel.target_workbook} · {excel.sheet_name}!C1 · "
                f"{excel.rows}행 × {excel.columns}열"
            )
        else:
            self.append_log(
                "Excel RAW 반영 제외: 연결된 파일은 변경하지 않았습니다. "
                f"다운로드 데이터 {excel.rows}행 × {excel.columns}열은 앱 조회에만 사용했습니다."
            )
        if excel.filtered_rows:
            self.append_log(
                f"입고일이 기준일 전날인 행 {excel.filtered_rows}개를 제외했습니다."
            )
        self.append_log(f"일별 입고 상세 표시 완료: {len(daily.products)}개 상품")
        if daily.unmatched_dispatches:
            number_label = "예약번호" if booking_type == "truck" else "배차번호"
            self.append_log(
                f"기준일 카드에서 찾지 못한 {number_label}: "
                + ", ".join(daily.unmatched_dispatches)
            )
        empty_details = getattr(daily, "empty_detail_dispatches", ())
        if empty_details:
            number_label = "예약번호" if booking_type == "truck" else "배차번호"
            self.append_log(
                f"상품 데이터가 없어 건너뛴 {number_label}: " + ", ".join(empty_details)
            )
        self.status_label.setText("일별 입고 표 완료 · WMS 무게 확인 준비")
        weight_retry_mode = (
            "restart" if self._pending_full_pipeline_restart else "resume"
        )
        self._pending_full_pipeline_restart = False
        self._start_weight_lookup(
            self.current_products,
            retry_mode=weight_retry_mode,
        )

    def _populate_milkrun_products(self, products) -> None:
        self._active_booking_type = "milkrun"
        self._populate_booking_products(products, "milkrun")

    def _populate_truck_products(self, products) -> None:
        self._active_booking_type = "truck"
        self._populate_booking_products(products, "truck")

    @staticmethod
    def _group_sku_key(value: object) -> str:
        try:
            return normalize_sku_id(value)
        except ValueError:
            return f"invalid:{str(value).strip()}"

    @staticmethod
    def _booking_group_key(product, booking_type: str) -> str:
        if booking_type == "truck":
            return str(product.dispatch_number or "").strip()
        # The daily-detail table rowspans vendor/Milkrun/pallet/box values.
        # ``milkrun_number`` identifies that inner quantity group, while
        # ``dispatch_number`` identifies the outer M schedule card and can
        # contain several independent Milkrun groups.
        return normalize_product_name(product.milkrun_number)

    @staticmethod
    def _display_booking_number(product, booking_type: str) -> str:
        if booking_type == "truck":
            return str(product.dispatch_number or "").strip()
        return normalize_product_name(
            product.dispatch_number or product.milkrun_number
        )

    def _group_categories_for_booking(self, booking_type: str) -> dict[str, str]:
        return (
            self._truck_group_categories
            if booking_type == "truck"
            else self._milkrun_group_categories
        )

    @classmethod
    def _booking_multi_sku_groups(
        cls,
        products,
        booking_type: str,
    ) -> dict[str, tuple[int, ...]]:
        # Both detail tables now provide per-SKU quantities: Milkrun exposes
        # pallet/box columns and Truck exposes PALLET container count/total
        # quantity. No booking group needs the legacy weight-only/manual mode.
        return {}

    @classmethod
    def _visual_multi_sku_groups(cls, products, booking_type: str) -> dict[str, tuple[int, ...]]:
        rows_by_group: dict[str, list[int]] = {}
        skus_by_group: dict[str, set[str]] = {}
        for row_index, product in enumerate(products):
            group_key = cls._booking_group_key(product, booking_type)
            if not group_key:
                continue
            rows_by_group.setdefault(group_key, []).append(row_index)
            skus_by_group.setdefault(group_key, set()).add(
                cls._group_sku_key(product.sku_id)
            )
        return {
            group_key: tuple(rows_by_group[group_key])
            for group_key, sku_ids in skus_by_group.items()
            if len(sku_ids) > 1
        }

    @staticmethod
    def _apply_group_row_tint(
        table: QTableWidget,
        groups: dict[str, tuple[int, ...]],
    ) -> None:
        """Give each multi-SKU vehicle group a subtle alternating dark tint."""
        row_colors = (QColor(COLORS["group_blue"]), QColor(COLORS["group_violet"]))
        key_colors = (
            QColor(COLORS["group_blue_key"]),
            QColor(COLORS["group_violet_key"]),
        )
        for group_index, rows in enumerate(groups.values()):
            row_brush = QBrush(row_colors[group_index % len(row_colors)])
            key_brush = QBrush(key_colors[group_index % len(key_colors)])
            for row_index in rows:
                for column_index in range(table.columnCount()):
                    item = table.item(row_index, column_index)
                    if item is not None:
                        item.setBackground(
                            key_brush if column_index == 1 else row_brush
                        )

    @classmethod
    def _booking_multi_sku_ids(cls, products, booking_type: str) -> set[str]:
        groups = cls._booking_multi_sku_groups(products, booking_type)
        sku_ids: set[str] = set()
        for rows in groups.values():
            for row_index in rows:
                try:
                    sku_ids.add(normalize_sku_id(products[row_index].sku_id))
                except ValueError:
                    continue
        return sku_ids

    @classmethod
    def _ordered_booking_products(cls, products, booking_type: str) -> tuple:
        """Keep every row while making identical booking groups contiguous."""
        grouped: dict[tuple[str, object], list] = {}
        group_order: list[tuple[str, object]] = []
        for row_index, product in enumerate(products):
            group_key = cls._booking_group_key(product, booking_type)
            identity: tuple[str, object] = (
                ("group", group_key) if group_key else ("row", row_index)
            )
            if identity not in grouped:
                grouped[identity] = []
                group_order.append(identity)
            grouped[identity].append(product)
        return tuple(product for identity in group_order for product in grouped[identity])

    @classmethod
    def _truck_multi_sku_groups(cls, products) -> dict[str, tuple[int, ...]]:
        return cls._booking_multi_sku_groups(products, "truck")

    @classmethod
    def _truck_multi_sku_ids(cls, products) -> set[str]:
        return cls._booking_multi_sku_ids(products, "truck")

    def _populate_booking_products(self, products, booking_type: str) -> None:
        self.current_products = self._ordered_booking_products(tuple(products), booking_type)
        self._products_by_booking[booking_type] = self.current_products
        self._group_categories_for_booking(booking_type).clear()
        self._weight_records.clear()
        self._weight_failures.clear()
        self._weight_row_errors.clear()
        table = self._table_for_booking(booking_type)
        table.clearSpans()
        table.setRowCount(len(self.current_products))
        multi_sku_groups = self._booking_multi_sku_groups(
            self.current_products,
            booking_type,
        )
        visual_multi_sku_groups = self._visual_multi_sku_groups(
            self.current_products,
            booking_type,
        )
        multi_sku_rows = {
            row_index: reservation_number
            for reservation_number, rows in multi_sku_groups.items()
            for row_index in rows
        }
        for row_index, product in enumerate(self.current_products):
            try:
                boxes_per_pallet = self._format_decimal(
                    calculate_boxes_per_pallet(product.box_count, product.pallet_count),
                    3,
                )
            except (TypeError, ValueError):
                boxes_per_pallet = "-"
            values = (
                product.vendor_name,
                self._display_booking_number(product, booking_type),
                product.pallet_count,
                product.box_count,
                boxes_per_pallet,
                product.sku_id,
                normalize_product_name(product.sku_name),
                "-",
                "?" if row_index in multi_sku_rows else "-",
            )
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if column_index in (0, 6):
                    item.setToolTip(str(value))
                table.setItem(row_index, column_index, item)
            if row_index not in multi_sku_rows:
                category_button = QPushButton("?")
                category_button.setObjectName("CategoryButton")
                category_button.setProperty("classification", "?")
                category_button.setEnabled(False)
                category_button.setToolTip("WMS 무게 확인 후 분류를 변경할 수 있습니다.")
                category_button.clicked.connect(
                    lambda _checked=False, sku_id=product.sku_id, kind=booking_type: self._cycle_category(
                        sku_id,
                        kind,
                    )
                )
                table.setCellWidget(row_index, 9, category_button)

        for group_key, rows in multi_sku_groups.items():
            first_row = rows[0]
            if rows != tuple(range(first_row, first_row + len(rows))):
                group_label = "예약번호" if booking_type == "truck" else "밀크런 번호"
                raise ValueError(
                    f"{group_label} {group_key}의 SKU 행이 연속되지 않아 표를 병합할 수 없습니다."
                )
            table.setSpan(first_row, 2, len(rows), 1)
            table.setSpan(first_row, 0, len(rows), 1)
            table.setSpan(first_row, 1, len(rows), 1)
            table.setSpan(first_row, 9, len(rows), 1)
            category_button = QPushButton("?")
            category_button.setObjectName("CategoryButton")
            category_button.clicked.connect(
                lambda _checked=False, kind=booking_type, key=group_key: self._cycle_booking_group_category(
                    kind,
                    key,
                )
            )
            self._configure_booking_group_button(
                category_button,
                self._group_categories_for_booking(booking_type).get(group_key),
                booking_type=booking_type,
                enabled=False,
            )
            table.setCellWidget(first_row, 9, category_button)

        for group_key, rows in visual_multi_sku_groups.items():
            first_row = rows[0]
            if rows != tuple(range(first_row, first_row + len(rows))):
                group_label = "예약번호" if booking_type == "truck" else "밀크런 번호"
                raise ValueError(
                    f"{group_label} {group_key}의 SKU 행이 연속되지 않아 "
                    "표를 병합할 수 없습니다."
                )
            # Vehicle identity is shared, while pallet/unit/box calculation
            # and category cells remain independent for every SKU.
            table.setSpan(first_row, 0, len(rows), 1)
            displayed_numbers = {
                self._display_booking_number(self.current_products[row], booking_type)
                for row in rows
            }
            if len(displayed_numbers) == 1:
                table.setSpan(first_row, 1, len(rows), 1)

        self._apply_group_row_tint(table, visual_multi_sku_groups)
        search_input = self._search_input_for_booking(booking_type)
        self._filter_booking_table(booking_type, search_input.text())

    def _search_input_for_booking(self, booking_type: str) -> QLineEdit:
        return (
            self.truck_search_input
            if booking_type == "truck"
            else self.milkrun_search_input
        )

    def _filter_booking_table(self, booking_type: str, query: str) -> None:
        table = self._table_for_booking(booking_type)
        products = self._products_by_booking.get(booking_type, ())
        tokens = tuple(normalize_product_name(query).casefold().split())
        if not tokens:
            for row_index in range(table.rowCount()):
                table.setRowHidden(row_index, False)
            return

        visible_rows: set[int] = set()
        for row_index, product in enumerate(products):
            booking_number = self._display_booking_number(product, booking_type)
            searchable = normalize_product_name(
                " ".join(
                    (
                        str(product.vendor_name or ""),
                        str(booking_number or ""),
                        str(product.sku_id or ""),
                        str(product.sku_name or ""),
                    )
                )
            ).casefold()
            if all(token in searchable for token in tokens):
                visible_rows.add(row_index)

        # Vehicle identity cells are merged for multi-SKU groups. If one SKU
        # matches, keep the entire group visible so those spans stay valid and
        # the result still reads as one vehicle.
        for rows in self._visual_multi_sku_groups(products, booking_type).values():
            if visible_rows.intersection(rows):
                visible_rows.update(rows)

        for row_index in range(table.rowCount()):
            table.setRowHidden(row_index, row_index not in visible_rows)

    @staticmethod
    def _weight_checkpoint_skus(products) -> tuple[str, ...]:
        sku_ids: set[str] = set()
        for product in products:
            try:
                sku_ids.add(normalize_sku_id(product.sku_id))
            except ValueError:
                continue
        return tuple(sorted(sku_ids))

    def _weight_checkpoint_key(
        self,
        *,
        booking_type: str | None = None,
        base_date: date | None = None,
    ) -> str:
        kind = booking_type or self._active_booking_type
        pipeline = self._pipeline_results_by_booking.get(kind)
        selected_date = (
            base_date
            or getattr(pipeline, "base_date", None)
            or self._selected_snapshot_date()
        )
        date_text = selected_date.isoformat() if selected_date is not None else "unknown"
        return f"{kind}|{date_text}"

    def _read_weight_retry_checkpoints(self) -> dict[str, dict]:
        raw = self.settings.value(self.WEIGHT_RETRY_CHECKPOINTS_KEY, "")
        if not raw:
            return {}
        try:
            payload = json.loads(str(raw))
            if not isinstance(payload, dict):
                raise ValueError("체크포인트 루트가 객체가 아닙니다.")
            checkpoints: dict[str, dict] = {}
            for key, value in payload.items():
                if not isinstance(key, str) or not isinstance(value, dict):
                    raise ValueError("체크포인트 항목 형식이 올바르지 않습니다.")
                sku_ids = value.get("sku_ids")
                completed_sku_ids = value.get("completed_sku_ids", [])
                started_at = value.get("started_at")
                if (
                    not isinstance(sku_ids, list)
                    or not all(isinstance(sku_id, str) for sku_id in sku_ids)
                    or not isinstance(completed_sku_ids, list)
                    or not all(
                        isinstance(sku_id, str) for sku_id in completed_sku_ids
                    )
                    or not isinstance(started_at, str)
                ):
                    raise ValueError("체크포인트 값 형식이 올바르지 않습니다.")
                checkpoints[key] = {
                    "sku_ids": list(dict.fromkeys(sku_ids)),
                    "completed_sku_ids": list(dict.fromkeys(completed_sku_ids)),
                    "started_at": started_at,
                }
            return checkpoints
        except (TypeError, ValueError, json.JSONDecodeError):
            self.settings.remove(self.WEIGHT_RETRY_CHECKPOINTS_KEY)
            self.settings.sync()
            return {}

    def _write_weight_retry_checkpoints(self, checkpoints: dict[str, dict]) -> None:
        if checkpoints:
            ordered = dict(
                sorted(
                    checkpoints.items(),
                    key=lambda item: str(item[1].get("started_at", "")),
                    reverse=True,
                )[: self.MAX_WEIGHT_RETRY_CHECKPOINTS]
            )
            self.settings.setValue(
                self.WEIGHT_RETRY_CHECKPOINTS_KEY,
                json.dumps(ordered, ensure_ascii=False, separators=(",", ":")),
            )
        else:
            self.settings.remove(self.WEIGHT_RETRY_CHECKPOINTS_KEY)
        self.settings.sync()

    def _remember_weight_retry_checkpoint(
        self,
        products,
        *,
        memory: ProductMemory | None = None,
        retry_mode: str = "resume",
    ) -> tuple[str, tuple[str, ...]]:
        key = self._weight_checkpoint_key()
        checkpoints = self._read_weight_retry_checkpoints()
        target_skus = self._weight_checkpoint_skus(products)
        previous = checkpoints.get(key)
        if retry_mode == "restart":
            completed_skus: set[str] = set()
        elif previous is not None:
            completed_skus = set(previous.get("completed_sku_ids", ()))
            completed_skus.intersection_update(target_skus)
        elif memory is not None:
            completed_skus = {
                sku_id
                for sku_id in target_skus
                if (record := memory.get(sku_id)) is not None
                and record.weight_grams is not None
            }
        else:
            completed_skus = set()
        checkpoints[key] = {
            "sku_ids": list(target_skus),
            "completed_sku_ids": sorted(completed_skus),
            "started_at": datetime.now().isoformat(timespec="seconds"),
        }
        self._write_weight_retry_checkpoints(checkpoints)
        self._active_weight_checkpoint_key = key
        pending_skus = tuple(
            sku_id for sku_id in target_skus if sku_id not in completed_skus
        )
        return key, pending_skus

    def _mark_weight_checkpoint_completed(self, sku_id: str) -> None:
        key = self._active_weight_checkpoint_key
        if not key:
            return
        checkpoints = self._read_weight_retry_checkpoints()
        checkpoint = checkpoints.get(key)
        if checkpoint is None or sku_id not in checkpoint.get("sku_ids", ()):
            return
        completed = set(checkpoint.get("completed_sku_ids", ()))
        if sku_id in completed:
            return
        completed.add(sku_id)
        checkpoint["completed_sku_ids"] = sorted(completed)
        checkpoints[key] = checkpoint
        self._write_weight_retry_checkpoints(checkpoints)

    @staticmethod
    def _checkpoint_completed_count(
        checkpoint: dict,
        products,
    ) -> int:
        target_skus = set(MainWindow._weight_checkpoint_skus(products))
        completed_skus = set(checkpoint.get("completed_sku_ids", ()))
        return len(target_skus.intersection(completed_skus))

    def _stored_weight_count(self, products) -> int:
        target_skus = set(self._weight_checkpoint_skus(products))
        return sum(
            1
            for sku_id in target_skus
            if (record := self._weight_records.get(sku_id)) is not None
            and record.weight_grams is not None
        )

    def _clear_weight_retry_checkpoint(self, key: str | None = None) -> None:
        checkpoint_key = key or self._active_weight_checkpoint_key
        if not checkpoint_key:
            return
        checkpoints = self._read_weight_retry_checkpoints()
        if checkpoint_key in checkpoints:
            del checkpoints[checkpoint_key]
            self._write_weight_retry_checkpoints(checkpoints)
        if checkpoint_key == self._active_weight_checkpoint_key:
            self._active_weight_checkpoint_key = ""

    def _ask_weight_retry_action(
        self,
        *,
        cached_count: int,
        total_count: int,
        has_checkpoint: bool = True,
    ) -> str:
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Icon.Question)
        dialog.setWindowTitle(
            "WMS 무게 측정 다시 진행"
            if has_checkpoint
            else "저장된 RAW 표에서 진행"
        )
        dialog.setText(
            "이 기준일의 WMS 무게 측정이 완료되지 않았습니다."
            if has_checkpoint
            else "이 기준일의 저장된 RAW 표가 있습니다."
        )
        dialog.setInformativeText(
            f"현재 표 SKU {total_count}개 중 저장된 무게 {cached_count}개를 확인했습니다.\n\n"
            "이어서 진행: 현재 표를 유지하고 저장된 무게는 다시 측정하지 않으며 "
            "미완료 SKU만 조회합니다.\n"
            "처음부터 다시: Shipments 조회, 파일 다운로드, Excel 반영, 상세 상품 수집과 "
            "WMS 무게 측정을 모두 처음부터 다시 실행합니다. 재측정에 실패해도 기존 "
            "저장값은 먼저 삭제하지 않습니다."
        )
        resume_button = dialog.addButton("이어서 진행", QMessageBox.ButtonRole.AcceptRole)
        restart_button = dialog.addButton(
            "처음부터 다시",
            QMessageBox.ButtonRole.DestructiveRole,
        )
        cancel_button = dialog.addButton("나중에", QMessageBox.ButtonRole.RejectRole)
        dialog.setDefaultButton(resume_button)
        dialog.setEscapeButton(cancel_button)
        dialog.exec()
        clicked = dialog.clickedButton()
        if clicked is restart_button:
            return "restart"
        if clicked is resume_button:
            return "resume"
        return "cancel"

    def _start_weight_lookup(self, products, *, retry_mode: str | None = None) -> None:
        if self.weight_worker and self.weight_worker.isRunning():
            return
        self._credential_load_failure = None
        if self._closing_after_cancel or self._automation_cancel_requested:
            self.append_log("중지 요청이 처리 중이므로 WMS 무게 조회를 시작하지 않습니다.")
            return
        products = tuple(products)
        if not products:
            self._pending_weight_summary = ProductWeightSummary(
                total_skus=0,
                cache_hits=0,
                wms_successes=0,
                failures=(),
            )
            self._pending_weight_failure = None
            self._pending_weight_cancel = ""
            self._weight_finalize_pending = True
            self.status_label.setText("기준일 표시 상품 없음 · WMS 조회 생략")
            self.append_log("기준일에 표시할 상품이 없어 WMS 무게 조회를 건너뜁니다.")
            self._finalize_weight_if_ready()
            return
        memory = _open_product_memory_with_recovery(self.product_memory_file, self)
        if memory is None:
            message = (
                "상품 메모리를 준비하지 못해 WMS 무게 조회를 시작하지 않았습니다. "
                "복구를 취소했다면 설정에서 저장 목록을 다시 열어 복구할 수 있습니다."
            )
            self._pending_weight_failure = FailureDetails(summary=message, detail=message)
            self._weight_finalize_pending = True
            self.status_label.setText("부분 완료 · 상품 메모리 복구 취소")
            self.append_log(message)
            self._finalize_weight_if_ready()
            return
        retry_mode = retry_mode if retry_mode in {"resume", "restart"} else "resume"
        try:
            credentials = WMSCredentialStore(self.settings).load()
        except CredentialError as exc:
            credentials = WMSCredentials(
                wms_id=str(self.settings.value("wms_id", "")),
                password="",
            )
            self._credential_load_failure = FailureDetails.from_exception(exc)

        download_dir = Path(
            str(self.settings.value("download_dir", str(default_download_dir())))
        ).expanduser()
        self._pending_weight_summary = None
        self._pending_weight_failure = None
        self._pending_weight_cancel = ""
        _checkpoint_key, pending_skus = self._remember_weight_retry_checkpoint(
            products,
            memory=memory,
            retry_mode=retry_mode,
        )
        self.weight_worker = ProductWeightWorker(
            products,
            self.product_memory_file,
            chromedriver_path(),
            credentials.wms_id,
            credentials.password,
            evidence_dir=download_dir,
            quantity_label="유닛",
            force_refresh_sku_ids=pending_skus,
        )
        self.weight_worker.log_updated.connect(self.append_log)
        self.weight_worker.progress_updated.connect(self._on_weight_progress)
        self.weight_worker.record_ready.connect(self._on_weight_record_ready)
        self.weight_worker.sku_failed.connect(self._on_weight_sku_failed)
        self.weight_worker.completed.connect(self._on_weight_completed)
        self.weight_worker.failed.connect(self._on_weight_failed)
        self.weight_worker.cancelled.connect(self._on_weight_cancelled)
        self.weight_worker.finished.connect(self._on_weight_finished)
        self.status_label.setText("WMS 상품 무게 확인 중")
        if retry_mode == "restart":
            self.append_log("현재 표의 모든 SKU 무게를 WMS에서 처음부터 다시 측정합니다.")
        else:
            self.append_log("저장된 무게를 유지하고 미완료 SKU부터 이어서 확인합니다.")
        self._set_automation_working(True)
        self.weight_worker.start()

    def _set_operation_progress(
        self,
        stage: str,
        completed: int,
        total: int,
    ) -> None:
        if total <= 0:
            self.operation_progress.setVisible(False)
            return
        completed = min(max(int(completed), 0), int(total))
        completed_percent = (completed * 100) // total
        remaining_percent = 100 - completed_percent
        self.operation_progress.setValue(completed_percent)
        self.operation_progress.setFormat(
            f"{completed_percent}% 완료 · {remaining_percent}% 남음"
        )
        self.operation_progress.setVisible(True)
        self.status_label.setText(
            f"{stage} {completed}/{total} · {remaining_percent}% 남음"
        )

    def _on_detail_progress(self, completed: int, total: int) -> None:
        self._set_operation_progress("상세 상품 조회", completed, total)

    def _on_weight_progress(self, completed: int, total: int) -> None:
        self._set_operation_progress("상품 무게 확인", completed, total)

    def _offer_weight_retry_after_problem(self) -> None:
        if not self.current_products or self._automation_worker_running():
            return
        memory = _open_product_memory_with_recovery(self.product_memory_file, self)
        if memory is None:
            return
        sku_ids = self._weight_checkpoint_skus(self.current_products)
        checkpoint = self._read_weight_retry_checkpoints().get(
            self._weight_checkpoint_key(),
            {},
        )
        cached_count = self._checkpoint_completed_count(
            checkpoint,
            self.current_products,
        )
        action = self._ask_weight_retry_action(
            cached_count=cached_count,
            total_count=len(sku_ids),
        )
        if action == "resume":
            self._start_weight_lookup(self.current_products, retry_mode="resume")
        elif action == "restart":
            self._start_booking_download(
                self._active_booking_type,
                retry_mode="restart",
            )

    def _on_weight_record_ready(self, record: ProductMemoryRecord, cache_hit: bool) -> None:
        self._weight_records[record.sku_id] = record
        if record.weight_grams is not None:
            self._mark_weight_checkpoint_completed(record.sku_id)
        self._render_weight_record(record, self._active_booking_type)
        source = "저장 정보" if cache_hit else "WMS"
        self.append_log(f"SKU {record.sku_id} 무게 반영: {source}")

    def _render_weight_record(
        self,
        record: ProductMemoryRecord,
        booking_type: str | None = None,
    ) -> None:
        booking_type = booking_type or self._active_booking_type
        products = self._products_by_booking.get(booking_type, ())
        table = self._table_for_booking(booking_type)
        multi_sku_groups = self._booking_multi_sku_groups(products, booking_type)
        multi_sku_group_skus = self._booking_multi_sku_ids(products, booking_type)
        group_categories = self._group_categories_for_booking(booking_type)
        self._weight_row_errors.pop(record.sku_id, None)
        for row_index, product in enumerate(products):
            try:
                row_sku = normalize_sku_id(product.sku_id)
            except ValueError:
                continue
            if row_sku != record.sku_id:
                continue

            group_key = self._booking_group_key(product, booking_type)
            if group_key in multi_sku_groups:
                if record.weight_grams is not None:
                    self._set_table_text(
                        row_index,
                        7,
                        self._format_decimal(record.weight_grams),
                        table=table,
                    )
                else:
                    self._set_table_text(row_index, 7, "-", table=table)
                self._clear_table_tooltip(row_index, 7, table=table)
                self._set_table_text(row_index, 8, "?", table=table)
                first_row = multi_sku_groups[group_key][0]
                button = table.cellWidget(first_row, 9)
                if isinstance(button, QPushButton):
                    self._configure_booking_group_button(
                        button,
                        group_categories.get(group_key),
                        booking_type=booking_type,
                        enabled=not self._automation_worker_running(),
                    )
                continue

            display_override = record.category_override
            if (
                display_override in AUTOMATIC_CATEGORIES
                and record.sku_id in multi_sku_group_skus
                and record.sku_id not in self._session_manual_category_skus
            ):
                # A multi-SKU booking row must not mutate the global SKU memory.
                # If the same SKU also has a single-SKU reservation, ignore a
                # stale saved light/heavy override for this run and show the
                # current local pallet calculation. High-stack still persists.
                display_override = None
            category = display_override or "?"
            error_text = ""
            try:
                boxes_per_pallet = calculate_boxes_per_pallet(
                    product.box_count,
                    product.pallet_count,
                )
                self._set_table_text(
                    row_index,
                    4,
                    self._format_decimal(boxes_per_pallet, 3),
                    table=table,
                )
            except (TypeError, ValueError) as exc:
                self._set_table_text(row_index, 4, "-", table=table)
                error_text = str(exc)
                self._weight_row_errors[record.sku_id] = error_text

            if record.weight_grams is not None:
                self._set_table_text(
                    row_index,
                    7,
                    self._format_decimal(record.weight_grams),
                    table=table,
                )
                if not error_text:
                    _, pallet_weight_kg, automatic_category = calculate_pallet_measurement(
                        record.weight_grams,
                        product.box_count,
                        product.pallet_count,
                    )
                    self._set_table_text(
                        row_index,
                        8,
                        self._format_decimal(pallet_weight_kg, 3),
                        table=table,
                    )
                    category = display_override or automatic_category
                else:
                    self._set_table_text(row_index, 8, "-", table=table)
            else:
                self._set_table_text(row_index, 7, "-", table=table)
                self._set_table_text(row_index, 8, "-", table=table)

            button = table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                self._configure_category_button(
                    button,
                    category,
                    manual=display_override is not None,
                    enabled=not self._automation_worker_running(),
                    error_text=error_text,
                    quantity_label="유닛",
                )

    def _on_weight_sku_failed(self, failure: SkuWeightFailure) -> None:
        self._weight_failures[failure.sku_id] = failure
        self.append_log(f"[WMS 조회 실패] SKU {failure.sku_id}: {failure.details.summary}")
        products = self._products_by_booking.get(self._active_booking_type, ())
        table = self._table_for_booking(self._active_booking_type)
        multi_sku_groups = self._booking_multi_sku_groups(
            products,
            self._active_booking_type,
        )
        for row_index, product in enumerate(products):
            try:
                matches = normalize_sku_id(product.sku_id) == normalize_sku_id(failure.sku_id)
            except ValueError:
                matches = str(product.sku_id).strip() == failure.sku_id
            if not matches:
                continue
            group_key = self._booking_group_key(product, self._active_booking_type)
            if group_key in multi_sku_groups:
                weight_item = table.item(row_index, 7)
                if weight_item is None:
                    self._set_table_text(row_index, 7, "-", table=table)
                    weight_item = table.item(row_index, 7)
                if weight_item is not None:
                    weight_item.setToolTip(failure.details.summary)
                continue
            button = table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                button.setToolTip(failure.details.summary)

    def _on_weight_completed(self, summary: ProductWeightSummary) -> None:
        if self._automation_cancel_requested:
            self._pending_weight_cancel = "사용자가 WMS 무게 조회를 중지했습니다."
            self.status_label.setText("WMS 무게 조회 취소 처리 중")
            return
        self._pending_weight_summary = summary
        self.status_label.setText("WMS 무게 결과 정리 중")

    def _on_weight_failed(self, failure: FailureDetails | object) -> None:
        self._pending_weight_failure = FailureDetails.coerce(failure)
        self.append_log(f"[WMS 무게 조회 오류] {self._pending_weight_failure.summary}")

    def _on_weight_cancelled(self, message: str) -> None:
        self._pending_weight_cancel = message
        self.append_log(message)

    def _on_weight_finished(self) -> None:
        worker = self.weight_worker
        self.weight_worker = None
        if worker is not None:
            worker.deleteLater()
        self._set_category_buttons_enabled(not self._automation_worker_running())
        self._set_automation_working(self._automation_worker_running())
        if not self._automation_worker_running() and self.current_products:
            self._set_category_buttons_enabled(True)
        self._refresh_arrival_from_current_raw()
        if self._closing_after_cancel:
            if not self._automation_worker_running():
                QTimer.singleShot(0, self.close)
            return
        self._weight_finalize_pending = True
        self._finalize_weight_if_ready()

    def _finalize_weight_if_ready(self) -> None:
        if not self._weight_finalize_pending or self._automation_worker_running():
            return
        self._weight_finalize_pending = False
        self._finalize_weight_lookup()

    def _finalize_weight_lookup(self) -> None:
        booking_label = self._booking_label(self._active_booking_type)
        number_label = "예약번호" if self._active_booking_type == "truck" else "배차번호"
        if self._pending_weight_cancel:
            self.status_label.setText("부분 완료 · WMS 무게 조회 취소")
            QMessageBox.information(
                self,
                "WMS 무게 조회 취소",
                f"{booking_label} 다운로드, Excel 반영과 일별 입고 표시는 완료됐습니다.\n"
                "이미 확인된 상품 무게는 저장됐으며 나머지는 다음 실행에서 다시 조회합니다.",
            )
            return

        problems: list[str] = []
        details: list[str] = []
        if self._credential_load_failure is not None:
            problems.append("저장된 WMS 비밀번호를 불러오지 못했습니다.")
            details.append(self._credential_load_failure.detail)
        if self._pending_weight_failure is not None:
            problems.append(self._pending_weight_failure.summary)
            details.append(self._pending_weight_failure.detail)
        summary = self._pending_weight_summary
        retryable_weight_problem = bool(
            self._credential_load_failure is not None
            or self._pending_weight_failure is not None
            or (summary is not None and summary.failures)
        )
        if summary is not None:
            for failure in summary.failures:
                problems.append(f"SKU {failure.sku_id}: {failure.details.summary}")
                details.append(
                    f"SKU {failure.sku_id} / {failure.product_name}\n{failure.details.detail}"
                )
        for sku_id, message in self._weight_row_errors.items():
            problems.append(f"SKU {sku_id} 팔렛트 계산: {message}")
            details.append(f"SKU {sku_id} 팔렛트 계산\n{message}")

        if problems:
            self.status_label.setText(f"부분 완료 · WMS/계산 오류 {len(problems)}건")
            failure = FailureDetails(
                summary=(
                    f"{booking_label} 다운로드, Excel 반영과 일별 입고 표시는 완료됐지만 "
                    f"상품 무게 확인 중 {len(problems)}건의 문제가 발생했습니다.\n\n"
                    + "\n".join(problems[:20])
                ),
                detail="\n\n".join(details),
            )
            self._show_error_dialog(
                "WMS 상품 무게 조회 일부 실패",
                failure,
                category=f"{booking_label} WMS 무게 조회",
            )
            if retryable_weight_problem and self.current_products:
                self._offer_weight_retry_after_problem()
            elif not retryable_weight_problem:
                self._clear_weight_retry_checkpoint()
            return

        cache_hits = summary.cache_hits if summary is not None else 0
        wms_successes = summary.wms_successes if summary is not None else 0
        product_count = len(self.current_products)
        manual_groups = self._booking_multi_sku_groups(
            self.current_products,
            self._active_booking_type,
        )
        unmatched = (
            self.current_pipeline_result.daily_inbound.unmatched_dispatches
            if self.current_pipeline_result is not None
            else ()
        )
        empty_details = (
            getattr(
                self.current_pipeline_result.daily_inbound,
                "empty_detail_dispatches",
                (),
            )
            if self.current_pipeline_result is not None
            else ()
        )
        if manual_groups:
            self.status_label.setText(
                f"완료 · 상품 {product_count}개 · 수동 분류 필요 {len(manual_groups)}건"
            )
        else:
            self.status_label.setText(f"완료 · 상품 {product_count}개 · 무게 분류 완료")
        self._clear_weight_retry_checkpoint()
        completion_text = (
            (
                "WMS SKU별 유닛 무게 확인을 완료했습니다."
                if self._active_booking_type == "truck"
                else "WMS SKU별 상품 무게 확인을 완료했습니다."
            )
            if manual_groups
            else "WMS 무게 분류를 완료했습니다."
        )
        message = (
            f"{booking_label} 다운로드, Excel 값 반영, 일별 입고 상세와 {completion_text}\n\n"
            f"표시 상품: {product_count}개\n"
            f"저장된 무게 사용: {cache_hits}개\n"
            f"WMS 신규 조회: {wms_successes}개"
        )
        if manual_groups:
            group_label = "예약" if self._active_booking_type == "truck" else "밀크런"
            quantity_label = "유닛"
            message += (
                f"\n\n다중 SKU {group_label} 수동 분류: {len(manual_groups)}건\n"
                f"해당 {group_label}은 SKU별 {quantity_label} 무게만 저장했습니다. "
                f"표의 병합된 분류 버튼을 눌러 {group_label} 전체를 수동 분류해 주세요. "
                "수동 값은 상품 메모리에 저장되지 않습니다."
            )
        if unmatched:
            message += f"\n\n기준일 카드에서 찾지 못한 {number_label}: " + ", ".join(unmatched)
        if empty_details:
            message += (
                f"\n\n상품 데이터가 없어 건너뛴 {number_label}: "
                + ", ".join(empty_details)
            )
        if unmatched or empty_details:
            QMessageBox.warning(self, f"{booking_label} 작업 완료", message)
        elif manual_groups:
            QMessageBox.warning(self, f"{booking_label} 작업 완료", message)
        else:
            QMessageBox.information(self, f"{booking_label} 작업 완료", message)

    def _cycle_category(
        self,
        sku_value: object,
        booking_type: str | None = None,
    ) -> None:
        if self._automation_worker_running():
            return
        booking_type = booking_type or self._active_booking_type
        try:
            sku_id = normalize_sku_id(sku_value)
            memory = ProductMemory(self.product_memory_file)
            record = memory.get(sku_id)
            override = record.category_override if record is not None else None
            if (
                override in AUTOMATIC_CATEGORIES
                and sku_id in self._booking_multi_sku_ids(
                    self._products_by_booking.get(booking_type, ()),
                    booking_type,
                )
                and sku_id not in self._session_manual_category_skus
            ):
                override = None
            product_name = ""
            for product in self._products_by_booking.get(booking_type, ()):
                try:
                    product_sku_id = normalize_sku_id(product.sku_id)
                except ValueError:
                    # An unrelated malformed row is already reported as an
                    # individual WMS failure. It must not disable classification
                    # editing for a valid SKU elsewhere in the table.
                    continue
                if product_sku_id == sku_id:
                    product_name = normalize_product_name(product.sku_name)
                    break
            if override is None:
                # Enter manual mode at a deterministic first choice so every
                # category remains reachable even when the automatic result is
                # already 중량.
                next_category = "경량"
            elif override == "경량":
                next_category = "중량"
            elif override == "중량":
                next_category = "고단"
            elif override == HIGH_CATEGORY:
                next_category = GRAIN_CATEGORY
            else:
                next_category = None
            updated = memory.set_manual_category(sku_id, next_category, product_name)
        except Exception as exc:
            self._show_error_dialog(
                "상품 분류 저장 실패",
                FailureDetails.from_exception(exc),
                category="상품 분류 메모리",
            )
            return

        if updated is None:
            self._session_manual_category_skus.discard(sku_id)
            self._weight_records.pop(sku_id, None)
            for kind in ("milkrun", "truck"):
                self._render_unknown_sku(sku_id, kind)
            self.append_log(f"SKU {sku_id} 수동 분류를 해제했습니다.")
        else:
            if updated.category_override is None:
                self._session_manual_category_skus.discard(sku_id)
            else:
                self._session_manual_category_skus.add(sku_id)
            self._weight_records[sku_id] = updated
            for kind in ("milkrun", "truck"):
                self._render_weight_record(updated, kind)
            source = "자동" if updated.category_override is None else "수동"
            self.append_log(f"SKU {sku_id} 분류 변경: {updated.effective_category or '?'} ({source})")
        self._refresh_arrival_from_current_raw()

    def _cycle_booking_group_category(self, booking_type: str, group_key: str) -> None:
        if self._automation_worker_running():
            return
        categories = self._group_categories_for_booking(booking_type)
        current = categories.get(group_key)
        if current is None:
            next_category = "경량"
        elif current == "경량":
            next_category = "중량"
        elif current == "중량":
            next_category = "고단"
        elif current == HIGH_CATEGORY:
            next_category = GRAIN_CATEGORY
        else:
            next_category = None

        if next_category is None:
            categories.pop(group_key, None)
        else:
            categories[group_key] = next_category

        products = self._products_by_booking.get(booking_type, ())
        rows = self._booking_multi_sku_groups(products, booking_type).get(group_key, ())
        if rows:
            table = self._table_for_booking(booking_type)
            button = table.cellWidget(rows[0], 9)
            if isinstance(button, QPushButton):
                self._configure_booking_group_button(
                    button,
                    next_category,
                    booking_type=booking_type,
                    enabled=True,
                )
        group_label = "예약번호" if booking_type == "truck" else "밀크런 번호"
        self.append_log(
            f"{group_label} {group_key} 표 분류 변경: {next_category or '?'} "
            "(현재 표에만 적용·상품 메모리에 저장하지 않음)"
        )
        self._refresh_arrival_from_current_raw()

    def _cycle_truck_group_category(self, reservation_number: str) -> None:
        self._cycle_booking_group_category("truck", reservation_number)

    def _render_unknown_sku(
        self,
        sku_id: str,
        booking_type: str | None = None,
    ) -> None:
        booking_type = booking_type or self._active_booking_type
        products = self._products_by_booking.get(booking_type, ())
        table = self._table_for_booking(booking_type)
        multi_sku_groups = self._booking_multi_sku_groups(products, booking_type)
        group_categories = self._group_categories_for_booking(booking_type)
        for row_index, product in enumerate(products):
            try:
                if normalize_sku_id(product.sku_id) != sku_id:
                    continue
            except ValueError:
                continue
            group_key = self._booking_group_key(product, booking_type)
            try:
                boxes_per_pallet = calculate_boxes_per_pallet(
                    product.box_count,
                    product.pallet_count,
                )
                self._set_table_text(
                    row_index,
                    4,
                    self._format_decimal(boxes_per_pallet, 3),
                    table=table,
                )
            except (TypeError, ValueError):
                self._set_table_text(row_index, 4, "-", table=table)
            self._set_table_text(row_index, 7, "-", table=table)
            self._clear_table_tooltip(row_index, 7, table=table)
            self._set_table_text(
                row_index,
                8,
                "?" if group_key in multi_sku_groups else "-",
                table=table,
            )
            if group_key in multi_sku_groups:
                first_row = multi_sku_groups[group_key][0]
                button = table.cellWidget(first_row, 9)
                if isinstance(button, QPushButton):
                    self._configure_booking_group_button(
                        button,
                        group_categories.get(group_key),
                        booking_type=booking_type,
                        enabled=not self._automation_worker_running(),
                    )
                continue
            button = table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                self._configure_category_button(button, "?", manual=False, enabled=True)

    def _displayed_category_for_sku(
        self,
        sku_id: str,
        booking_type: str | None = None,
    ) -> str:
        booking_type = booking_type or self._active_booking_type
        products = self._products_by_booking.get(booking_type, ())
        table = self._table_for_booking(booking_type)
        multi_sku_groups = self._booking_multi_sku_groups(products, booking_type)
        group_categories = self._group_categories_for_booking(booking_type)
        for row_index, product in enumerate(products):
            try:
                if normalize_sku_id(product.sku_id) != sku_id:
                    continue
            except ValueError:
                continue
            group_key = self._booking_group_key(product, booking_type)
            if group_key in multi_sku_groups:
                return group_categories.get(group_key, "?")
            button = table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                return button.text()
        return "?"

    def _set_category_buttons_enabled(self, enabled: bool) -> None:
        for table in (self.raw_table, self.truck_table):
            for row_index in range(table.rowCount()):
                button = table.cellWidget(row_index, 9)
                if isinstance(button, QPushButton):
                    button.setEnabled(enabled)

    @staticmethod
    def _configure_category_button(
        button: QPushButton,
        category: str,
        *,
        manual: bool,
        enabled: bool,
        error_text: str = "",
        quantity_label: str = "유닛",
    ) -> None:
        display = category if category in MANUAL_CATEGORIES else "?"
        button.setText(display)
        button.setProperty("classification", display)
        button.setEnabled(enabled)
        if error_text:
            tooltip = f"팔렛트 무게 계산 오류: {error_text}\n클릭해 수동 분류할 수 있습니다."
        elif manual and display in PERSISTENT_MANUAL_CATEGORIES:
            tooltip = f"{display} 수동 분류입니다. 이후 데이터 조회에서도 유지됩니다."
        elif manual:
            tooltip = (
                "현재 표시의 수동 분류입니다. 다음 데이터 조회에서는 "
                f"팔렛트당 {quantity_label} 수로 경량/중량을 다시 계산합니다."
            )
        else:
            tooltip = "무게 기준 자동 분류입니다. 클릭하면 수동 분류로 변경됩니다."
        button.setToolTip(tooltip)
        button.style().unpolish(button)
        button.style().polish(button)

    @classmethod
    def _configure_booking_group_button(
        cls,
        button: QPushButton,
        category: str | None,
        *,
        booking_type: str,
        enabled: bool,
    ) -> None:
        cls._configure_category_button(
            button,
            category or "?",
            manual=category is not None,
            enabled=enabled,
            quantity_label="유닛",
        )
        group_label = "예약" if booking_type == "truck" else "밀크런 번호"
        button.setToolTip(
            f"한 {group_label}에 서로 다른 SKU가 여러 개이므로 팔렛트 무게를 "
            f"자동 계산하지 않습니다. 클릭해 {group_label} 전체를 수동 분류하며, "
            "이 값은 현재 표에만 적용되고 상품 메모리에 저장되지 않습니다."
        )

    @classmethod
    def _configure_truck_group_button(
        cls,
        button: QPushButton,
        category: str | None,
        *,
        enabled: bool,
    ) -> None:
        cls._configure_booking_group_button(
            button,
            category,
            booking_type="truck",
            enabled=enabled,
        )

    def _set_table_text(
        self,
        row: int,
        column: int,
        value: str,
        *,
        table: QTableWidget | None = None,
    ) -> None:
        table = table or self._table_for_booking(self._active_booking_type)
        item = table.item(row, column)
        if item is None:
            item = QTableWidgetItem()
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(row, column, item)
        item.setText(value)

    @staticmethod
    def _clear_table_tooltip(
        row: int,
        column: int,
        *,
        table: QTableWidget,
    ) -> None:
        item = table.item(row, column)
        if item is not None:
            item.setToolTip("")

    @staticmethod
    def _format_decimal(value: Decimal, digits: int | None = None) -> str:
        text = f"{value:.{digits}f}" if digits is not None else format(value, "f")
        return text.rstrip("0").rstrip(".") if "." in text else text

    def _on_milkrun_failed(self, failure: FailureDetails | object) -> None:
        details = FailureDetails.coerce(failure)
        booking_label = self._booking_label(self._active_booking_type)
        self.append_log(f"[오류] {details.summary}")
        self.status_label.setText("작업 실패")
        if not self._closing_after_cancel:
            self._show_error_dialog(
                f"{booking_label} 작업 실패",
                details,
                category=f"{booking_label} 자동화",
            )

    def _on_milkrun_excel_failed(
        self,
        downloaded_file: Path,
        failure: FailureDetails | object,
    ) -> None:
        details = FailureDetails.coerce(failure)
        booking_label = self._booking_label(self._active_booking_type)
        self.append_log(f"다운로드 완료: {downloaded_file}")
        self.append_log(f"[Excel 반영 오류] {details.summary}")
        self.status_label.setText("부분 완료 · Excel 반영 실패")
        if not self._closing_after_cancel:
            partial_failure = FailureDetails(
                summary=(
                    f"{booking_label} 파일은 정상적으로 내려받았지만 연결된 Excel에 반영하지 못했습니다.\n"
                    f"다운로드 파일: {downloaded_file}\n\n{details.summary}"
                ),
                detail=f"다운로드 파일: {downloaded_file}\n\n{details.detail}",
            )
            self._show_error_dialog(
                "Excel 반영 실패",
                partial_failure,
                category=f"{booking_label} Excel 반영",
            )

    def _on_excel_close_required(
        self,
        downloaded_file: Path | None,
        message: str,
    ) -> None:
        if downloaded_file is not None:
            self.append_log(f"다운로드 완료: {downloaded_file}")
        self.append_log("[Excel 닫기 필요] 연결된 입고스케줄 파일이 열려 있습니다.")
        self.status_label.setText("Excel 닫기 필요")
        if self._closing_after_cancel:
            return
        detail = message
        if downloaded_file is not None:
            detail += f"\n\n다운로드 파일은 보관되어 있습니다.\n{downloaded_file}"
        QMessageBox.warning(
            self,
            "Excel을 닫아 주세요",
            detail + "\n\n파일을 닫은 뒤 데이터 얻기를 다시 눌러 주세요.",
        )

    def _on_milkrun_detail_failed(
        self,
        import_result,
        failure: FailureDetails | object,
    ) -> None:
        details = FailureDetails.coerce(failure)
        booking_label = self._booking_label(self._active_booking_type)
        if import_result is not None:
            self.append_log(
                f"Excel 반영 완료: {import_result.target_workbook} · "
                f"{import_result.sheet_name}!C1"
            )
        self.append_log(f"[일별 입고 상세 오류] {details.summary}")
        self.status_label.setText("부분 완료 · Excel 반영 완료 · 일별 상세 실패")
        if not self._closing_after_cancel:
            prefix = (
                f"{booking_label} 파일 다운로드와 Excel 값 반영은 완료되었지만 "
                "일별 입고 상세를 가져오지 못했습니다."
            )
            partial_failure = FailureDetails(
                summary=f"{prefix}\n\n{details.summary}",
                detail=f"{prefix}\n\n{details.detail}",
            )
            self._show_error_dialog(
                "일별 입고 상세 조회 실패",
                partial_failure,
                category=f"{booking_label} 일별 입고 상세",
            )

    def _on_milkrun_detail_cancelled(self, import_result, message: str) -> None:
        booking_label = self._booking_label(self._active_booking_type)
        self.append_log(
            f"Excel 반영 완료: {import_result.target_workbook} · "
            f"{import_result.sheet_name}!C1 · {import_result.rows}행 × {import_result.columns}열"
        )
        self.append_log("일별 입고 상세 조회를 사용자가 중지했습니다.")
        self.status_label.setText("부분 완료 · Excel 반영 완료 · 일별 상세 취소")
        if not self._closing_after_cancel:
            QMessageBox.information(
                self,
                "일별 입고 상세 조회 취소",
                f"작업을 중지했지만 그 전에 {booking_label} 다운로드와 Excel 값 반영은 완료되었습니다.\n\n"
                f"대상: {import_result.target_workbook}\n"
                f"범위: {import_result.sheet_name}!C1 "
                f"({import_result.rows}행 × {import_result.columns}열)\n\n"
                f"{message}",
            )

    def _on_milkrun_cancelled(self, message: str) -> None:
        self.append_log(message)
        self.status_label.setText("작업 취소됨")

    def _on_milkrun_finished(self) -> None:
        worker = self.milkrun_worker
        self.milkrun_worker = None
        if worker is not None:
            worker.deleteLater()
        self._set_automation_working(self._automation_worker_running())
        if not self._automation_worker_running() and self.current_products:
            self._set_category_buttons_enabled(True)
        if self._closing_after_cancel and not self._automation_worker_running():
            QTimer.singleShot(0, self.close)
        elif not self._closing_after_cancel:
            self._finalize_weight_if_ready()

    def _automation_worker_running(self) -> bool:
        # Treat the worker as active until its queued ``finished`` slot clears
        # the reference.  QThread.run() may already have returned while a
        # preceding completed signal is still waiting in the GUI event queue;
        # considering that gap idle could close the window and then start WMS
        # from the delayed completion callback.
        return (
            self.milkrun_worker is not None
            or self.weight_worker is not None
            or self.arrival_worker is not None
        )

    def _set_automation_working(self, working: bool) -> None:
        self.get_data_button.setEnabled(not working)
        self.truck_get_data_button.setEnabled(not working)
        self.settings_button.setEnabled(not working)
        self.import_table_button.setEnabled(not working)
        self.export_table_button.setEnabled(not working)
        self.base_date_mode.setEnabled(not working)
        self.manual_base_date.setEnabled(
            not working and self.base_date_mode.currentData() == "manual"
        )
        self.arrival_refresh_button.setEnabled(not working)
        for checkbox in self._raw_excel_apply_checkboxes:
            checkbox.setEnabled(not working)
        for button in (self.stop_button, self.truck_stop_button):
            button.setVisible(working)
            button.setEnabled(working)
        if working:
            self.status_label.setText("작업 중 · 로그인 화면이면 브라우저에서 직접 인증해 주세요")
        else:
            self.operation_progress.setVisible(False)
            if self.status_label.text().startswith("작업 중"):
                self.status_label.setText("대기 중")

    def append_log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_view.appendPlainText(f"[{stamp}] {message}")
        if self._automation_worker_running():
            self.status_label.setText(message)

    def show_settings(self) -> None:
        dialog = SettingsDialog(
            self.settings,
            self,
            memory_path=self.product_memory_file,
        )
        dialog.update_requested.connect(lambda: self.check_for_update(manual=True))
        dialog.beta_changed.connect(self._on_beta_changed)
        dialog.history_requested.connect(self.show_update_history)
        dialog.product_memory_changed.connect(self._refresh_current_product_memory)
        dialog.exec()

    def _refresh_current_product_memory(self, announce: bool = True) -> None:
        if not any(self._products_by_booking.values()):
            return
        try:
            memory = ProductMemory(self.product_memory_file)
            self._weight_records.clear()
            self._weight_row_errors.clear()
            for booking_type in ("milkrun", "truck"):
                seen: set[str] = set()
                for product in self._products_by_booking.get(booking_type, ()):
                    try:
                        sku_id = normalize_sku_id(product.sku_id)
                    except ValueError:
                        continue
                    if sku_id in seen:
                        continue
                    seen.add(sku_id)
                    record = memory.get(sku_id)
                    if record is None:
                        self._render_unknown_sku(sku_id, booking_type)
                        continue
                    self._weight_records[sku_id] = record
                    self._render_weight_record(record, booking_type)
            if announce:
                self.append_log("설정에서 변경한 상품 분류 메모리를 RAW 표에 반영했습니다.")
            self._refresh_arrival_from_current_raw()
        except Exception as exc:
            self._show_error_dialog(
                "상품 분류 새로고침 실패",
                FailureDetails.from_exception(exc),
                category="상품 분류 메모리",
            )

    def show_update_history(self) -> None:
        UpdateHistoryDialog(self).exec()

    def _show_error_dialog(
        self,
        title: str,
        failure: FailureDetails | object,
        *,
        category: str,
    ) -> None:
        context = {
            "category": category,
            "log_tail": self.log_view.toPlainText(),
        }
        ErrorReportDialog(title, failure, context=context, parent=self).exec()

    def _on_beta_changed(self, enabled: bool) -> None:
        if enabled:
            self.append_log("Beta 업데이트 채널을 사용합니다.")
            self.check_for_update(manual=True)
            return
        answer = QMessageBox.question(
            self,
            "정식 릴리즈로 되돌리기",
            "Beta 참여를 해제했습니다. 최신 정식 릴리즈를 확인해 복구하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer == QMessageBox.StandardButton.Yes:
            self._check_release_restore()

    def check_for_update(self, manual: bool = False) -> None:
        if self._update_worker_running():
            if manual:
                QMessageBox.information(self, "업데이트", "다른 업데이트 작업이 진행 중입니다.")
            return
        self.manual_update_check = manual
        if manual:
            self.append_log("업데이트를 확인합니다.")
        use_prerelease = self.settings.value("use_prerelease", False, type=bool)
        worker = UpdateCheckWorker(use_prerelease)
        self.update_check_worker = worker
        worker.update_available.connect(self._show_update_dialog)
        worker.no_update.connect(self._on_no_update)
        worker.failed.connect(self._on_update_check_failed)
        worker.finished.connect(lambda: self._on_worker_finished("update_check_worker", worker))
        worker.start()

    def _on_no_update(self) -> None:
        if self._closing_after_workers:
            return
        if self.manual_update_check:
            QMessageBox.information(self, "업데이트", "현재 선택한 채널의 최신 버전입니다.")
        self.manual_update_check = False

    def _on_update_check_failed(self, failure: FailureDetails | object) -> None:
        details = FailureDetails.coerce(failure)
        self.append_log(f"[업데이트 확인 오류] {details.summary}")
        if self.manual_update_check and not self._closing_after_workers:
            self._show_error_dialog(
                "업데이트 확인 실패",
                details,
                category="앱 업데이트 확인",
            )
        self.manual_update_check = False

    def _check_release_restore(self) -> None:
        if self._update_worker_running():
            QMessageBox.information(self, "정식 릴리즈", "다른 업데이트 작업이 진행 중입니다.")
            return
        self.append_log("최신 정식 릴리즈를 확인합니다.")
        worker = ReleaseRestoreWorker()
        self.restore_worker = worker
        worker.available.connect(self._show_update_dialog)
        worker.unavailable.connect(self._on_restore_unavailable)
        worker.failed.connect(self._on_restore_failed)
        worker.finished.connect(lambda: self._on_worker_finished("restore_worker", worker))
        worker.start()

    def _on_restore_unavailable(self, message: str) -> None:
        self.append_log(message)
        if not self._closing_after_workers:
            QMessageBox.information(self, "정식 릴리즈", message)

    def _on_restore_failed(self, failure: FailureDetails | object) -> None:
        details = FailureDetails.coerce(failure)
        self.append_log(f"[정식 릴리즈 확인 오류] {details.summary}")
        if not self._closing_after_workers:
            self._show_error_dialog(
                "정식 릴리즈 확인 실패",
                details,
                category="정식 릴리즈 복구 확인",
            )

    def _automation_blocks_update(self, *, notify: bool) -> bool:
        if not self._automation_worker_running():
            return False
        message = "Shipments/WMS 작업 중에는 업데이트를 표시하거나 적용할 수 없습니다."
        self.append_log(message)
        if notify and not self._closing_after_workers:
            QMessageBox.warning(self, "작업 실행 중", message)
        return True

    def _show_update_dialog(self, info) -> None:
        from Modules.Common.AutoUpdater import AutoUpdater

        self.manual_update_check = False
        if self._closing_after_workers:
            return
        if self._automation_blocks_update(notify=False):
            return
        if self.update_dialog and self.update_dialog.isVisible():
            return
        self.current_update_info = info
        dialog = UpdateDialog(self)
        dialog.setObjectName("UpdateDialog")
        dialog.setWindowTitle("정식 릴리즈 복구" if getattr(info, "is_release_restore", False) else "업데이트")
        dialog.setMinimumWidth(520)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)
        action = "정식 릴리즈로 복구" if getattr(info, "is_release_restore", False) else "새 업데이트"
        title = QLabel(f"{action}: v{info.version}")
        title.setObjectName("DialogTitle")
        layout.addWidget(title)
        description = QLabel("변경 내용을 확인한 뒤 원하는 시점에 업데이트를 적용할 수 있습니다.")
        description.setObjectName("HelpText")
        layout.addWidget(description)
        changelog = QPlainTextEdit()
        changelog.setObjectName("DocumentView")
        changelog.setReadOnly(True)
        changelog.setPlainText(info.changelog or "변경사항 없음")
        changelog.setMaximumHeight(180)
        layout.addWidget(changelog)
        mode = "델타" if getattr(info, "patch_mode", "full") == "delta" else "전체"
        size = AutoUpdater.format_size(info.patch_size) if info.patch_size else "알 수 없음"
        patch_info = QLabel(f"패치: {mode} · {size}")
        patch_info.setObjectName("MutedText")
        layout.addWidget(patch_info)
        self.update_progress = QProgressBar()
        self.update_progress.setVisible(False)
        layout.addWidget(self.update_progress)
        self.update_status = QLabel("")
        layout.addWidget(self.update_status)
        buttons = QHBoxLayout()
        later = QPushButton("나중에")
        self.update_later_button = later
        self.apply_update_button = QPushButton("지금 적용")
        self.apply_update_button.setObjectName("PrimaryButton")
        later.clicked.connect(dialog.reject)
        self.apply_update_button.clicked.connect(lambda: self._download_update(info))
        buttons.addStretch(1)
        buttons.addWidget(later)
        buttons.addWidget(self.apply_update_button)
        layout.addLayout(buttons)
        self.update_dialog = dialog
        dialog.exec()

    def _download_update(self, info) -> None:
        if self._automation_blocks_update(notify=True):
            return
        if self.update_download_worker and self.update_download_worker.isRunning():
            return
        self.apply_update_button.setEnabled(False)
        self.update_later_button.setEnabled(False)
        if isinstance(self.update_dialog, UpdateDialog):
            self.update_dialog.allow_close = False
        self.update_progress.setVisible(True)
        self.update_status.setText("패치 다운로드 중...")
        worker = UpdateDownloadWorker(info)
        self.update_download_worker = worker
        worker.progress.connect(self.update_progress.setValue)
        worker.completed.connect(lambda path: self._apply_downloaded_update(path, info))
        worker.failed.connect(self._on_update_download_failed)
        worker.finished.connect(lambda: self._on_worker_finished("update_download_worker", worker))
        worker.start()

    def _apply_downloaded_update(self, zip_path: str, info) -> None:
        if self._automation_blocks_update(notify=True):
            if hasattr(self, "update_status"):
                self.update_status.setText("Shipments/WMS 작업 완료 후 업데이트를 다시 적용해 주세요.")
            if hasattr(self, "apply_update_button"):
                self.apply_update_button.setEnabled(True)
            if hasattr(self, "update_later_button"):
                self.update_later_button.setEnabled(True)
            if isinstance(self.update_dialog, UpdateDialog):
                self.update_dialog.allow_close = True
            return

        from Modules.Common.AutoUpdater import AutoUpdater

        self.update_progress.setValue(100)
        self.update_status.setText("업데이트 적용 준비 중...")
        try:
            success = AutoUpdater("UnHelper").apply_update(
                Path(zip_path),
                info.version,
                info.manifest,
            )
        except Exception as exc:
            self._on_update_download_failed(FailureDetails.from_exception(exc))
            return
        if success:
            if isinstance(self.update_dialog, UpdateDialog):
                self.update_dialog.allow_close = True
            QMessageBox.information(self, "업데이트", "앱을 종료한 뒤 패치를 적용하고 자동 재시작합니다.")
            self.close()
        else:
            self._on_update_download_failed("업데이트 적용 준비에 실패했습니다.")

    def _on_update_download_failed(self, failure: FailureDetails | object) -> None:
        details = FailureDetails.coerce(failure)
        self.update_status.setText(f"오류: {details.summary}")
        self.apply_update_button.setEnabled(True)
        self.update_later_button.setEnabled(True)
        if isinstance(self.update_dialog, UpdateDialog):
            self.update_dialog.allow_close = True
        if not self._closing_after_workers:
            self._show_error_dialog(
                "업데이트 실패",
                details,
                category="앱 업데이트 다운로드/적용",
            )

    def _update_worker_running(self) -> bool:
        return any(
            worker is not None and worker.isRunning()
            for worker in (
                self.update_check_worker,
                self.restore_worker,
                self.update_download_worker,
            )
        )

    def _on_worker_finished(self, attribute: str, worker: QThread) -> None:
        if getattr(self, attribute, None) is worker:
            setattr(self, attribute, None)
        worker.deleteLater()
        if self._closing_after_workers and not self._update_worker_running():
            QTimer.singleShot(0, self.close)

    def open_download_folder(self) -> None:
        path = Path(str(self.settings.value("download_dir", str(default_download_dir())))).expanduser()
        path.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve())))

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._automation_worker_running():
            answer = QMessageBox.question(
                self,
                "작업 중",
                "진행 중인 Shipments/WMS 작업을 중지하고 종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            self._closing_after_cancel = True
            self.cancel_milkrun_download()
            event.ignore()
            return
        if self._update_worker_running():
            self._closing_after_workers = True
            self.setEnabled(False)
            self.hide()
            event.ignore()
            return
        event.accept()
