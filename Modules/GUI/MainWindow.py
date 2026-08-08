from __future__ import annotations

import threading
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import QSettings, QThread, QTimer, Qt, QUrl, Signal
from PySide6.QtGui import QCloseEvent, QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
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
from Modules.Common.paths import chromedriver_path, default_download_dir, product_memory_path
from Modules.Common.version import CURRENT_VERSION
from Modules.Common.ErrorReport import FailureDetails
from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportCancelled,
    ExcelImportError,
    MilkrunExcelImportResult,
    MilkrunExcelImporter,
)
from Modules.GUI.Dialogs import ErrorReportDialog, UpdateHistoryDialog
from Modules.GUI.ProductMemoryDialog import ProductMemoryDialog
from Modules.Shipments.MilkrunDownloader import (
    AutomationCancelled,
    MilkrunDownloadRequest,
    MilkrunDownloader,
)
from Modules.Shipments.DailyInboundScraper import (
    DailyInboundError,
    DailyInboundResult,
    DailyInboundScraper,
)
from Modules.WMS.ProductMemory import (
    MANUAL_CATEGORIES,
    ProductMemory,
    ProductMemoryRecord,
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
    excel: MilkrunExcelImportResult
    daily_inbound: DailyInboundResult


class MilkrunWorker(QThread):
    log_updated = Signal(str)
    completed = Signal(object)
    excel_failed = Signal(object, object)
    detail_failed = Signal(object, object)
    detail_cancelled = Signal(object, str)
    failed = Signal(object)
    cancelled = Signal(str)

    def __init__(
        self,
        request: MilkrunDownloadRequest,
        driver_path: Path,
        target_workbook: Path,
    ):
        super().__init__()
        self.request = request
        self.driver_path = driver_path
        self.target_workbook = target_workbook
        self.stop_event = threading.Event()
        self.downloader: MilkrunDownloader | None = None

    def run(self) -> None:
        download_result = None
        import_result = None
        try:
            importer = MilkrunExcelImporter(log=self.log_updated.emit)
            self.log_updated.emit("연결된 Excel 파일을 확인합니다.")
            importer.validate_workbook(self.target_workbook)
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")

            self.downloader = MilkrunDownloader(
                self.driver_path,
                log=self.log_updated.emit,
                stop_event=self.stop_event,
            )
            download_result = self.downloader.run(self.request, keep_browser_open=True)
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")
            self.log_updated.emit("다운로드 데이터를 연결된 Excel 파일에 반영합니다.")
            import_result = importer.import_values(
                download_result.file_path,
                self.target_workbook,
                cancel_requested=self.stop_event.is_set,
                exclude_arrival_date=download_result.start_date,
            )
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")
            if import_result.rows == 1 and not import_result.order_numbers:
                self.log_updated.emit(
                    "오늘 반영할 Milkrun 데이터가 없어 일별 입고 상세 조회를 건너뜁니다."
                )
                daily_result = DailyInboundResult(
                    products=(),
                    requested_orders=(),
                    matched_orders=(),
                    unmatched_orders=(),
                )
            else:
                self.log_updated.emit("Excel P열의 발주번호로 오늘 일별 입고 상세를 조회합니다.")
                daily_result = DailyInboundScraper(
                    self.downloader,
                    evidence_dir=self.request.download_dir,
                ).run(
                    import_result.order_numbers,
                    center_name=self.request.center_name,
                    schedule_date=download_result.end_date,
                )
            self.completed.emit(MilkrunPipelineResult(import_result, daily_result))
        except (AutomationCancelled, ExcelImportCancelled) as exc:
            if import_result is not None:
                self.detail_cancelled.emit(import_result, str(exc))
            else:
                self.cancelled.emit(str(exc))
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
        self.setMinimumWidth(560)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 22, 22, 22)
        layout.setSpacing(14)

        title = QLabel("설정")
        title.setStyleSheet("font-size: 18pt; font-weight: 800;")
        layout.addWidget(title)

        download_label = QLabel("다운로드 파일 저장 폴더")
        layout.addWidget(download_label)
        download_row = QHBoxLayout()
        self.download_path = QLineEdit(
            str(self.settings.value("download_dir", str(default_download_dir())))
        )
        browse = QPushButton("찾아보기")
        browse.clicked.connect(self._browse)
        download_row.addWidget(self.download_path, 1)
        download_row.addWidget(browse)
        layout.addLayout(download_row)

        excel_label = QLabel("Milkrun 데이터를 반영할 Excel 파일")
        layout.addWidget(excel_label)
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
        layout.addLayout(excel_row)

        excel_help = QLabel(
            "다운로드 완료 후 Raw_밀크런!C1:P1000의 값만 지우고, C1부터 값으로 붙여넣습니다."
        )
        excel_help.setWordWrap(True)
        excel_help.setStyleSheet("color: #A1A1AA;")
        layout.addWidget(excel_help)

        wms_label = QLabel("WMS 로그인 정보")
        layout.addWidget(wms_label)
        try:
            credentials = self.credential_store.load()
        except CredentialError as exc:
            credentials = WMSCredentials(
                wms_id=str(self.settings.value("wms_id", "")),
                password="",
            )
            self._credential_load_error = exc

        wms_id_row = QHBoxLayout()
        wms_id_row.addWidget(QLabel("ID"))
        self.wms_id_input = QLineEdit(credentials.wms_id)
        self.wms_id_input.setPlaceholderText("WMS ID")
        wms_id_row.addWidget(self.wms_id_input, 1)
        layout.addLayout(wms_id_row)

        wms_password_row = QHBoxLayout()
        wms_password_row.addWidget(QLabel("Password"))
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
        layout.addLayout(wms_password_row)

        product_memory_row = QHBoxLayout()
        product_memory_help = QLabel("저장된 상품은 다음 실행에서 WMS 무게 조회를 생략합니다.")
        product_memory_help.setStyleSheet("color: #A1A1AA;")
        self.product_memory_button = QPushButton("저장된 상품 분류 목록")
        self.product_memory_button.clicked.connect(self._show_product_memory)
        product_memory_row.addWidget(product_memory_help, 1)
        product_memory_row.addWidget(self.product_memory_button)
        layout.addLayout(product_memory_row)

        self.beta_checkbox = QCheckBox("Beta(테스트 릴리즈) 업데이트 받기")
        self.beta_checkbox.setChecked(self.settings.value("use_prerelease", False, type=bool))
        self.beta_checkbox.setToolTip("끄면 최신 정식 릴리즈로 복구할 수 있습니다.")
        layout.addWidget(self.beta_checkbox)

        version = QLabel(f"현재 버전: v{CURRENT_VERSION}")
        version.setStyleSheet("color: #A1A1AA;")
        layout.addWidget(version)

        buttons = QHBoxLayout()
        update = QPushButton("업데이트 확인")
        self.update_history_button = QPushButton("업데이트 내역")
        close = QPushButton("저장하고 닫기")
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
            "Milkrun Excel 파일 연결",
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
    def __init__(self, smoke_test: bool = False):
        super().__init__()
        self.smoke_test = smoke_test
        self.settings = QSettings("Mrbinggrae", "UnHelper")
        self.milkrun_worker: MilkrunWorker | None = None
        self.weight_worker: ProductWeightWorker | None = None
        self.product_memory_file = product_memory_path()
        self.current_products = ()
        self.current_pipeline_result: MilkrunPipelineResult | None = None
        self._weight_records: dict[str, ProductMemoryRecord] = {}
        self._weight_failures: dict[str, SkuWeightFailure] = {}
        self._weight_row_errors: dict[str, str] = {}
        self._credential_load_failure: FailureDetails | None = None
        self._pending_weight_summary: ProductWeightSummary | None = None
        self._pending_weight_failure: FailureDetails | None = None
        self._pending_weight_cancel = ""
        self._weight_finalize_pending = False
        self.update_check_worker: UpdateCheckWorker | None = None
        self.restore_worker: ReleaseRestoreWorker | None = None
        self.update_download_worker: UpdateDownloadWorker | None = None
        self.update_dialog: QDialog | None = None
        self.current_update_info = None
        self.manual_update_check = False
        self._closing_after_cancel = False
        self._closing_after_workers = False

        self.setWindowTitle(f"UnHelper v{CURRENT_VERSION}")
        self.resize(1345, 760)
        self.setMinimumSize(980, 620)
        self._build_ui()
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
        header_layout.setContentsMargins(28, 14, 28, 14)
        title_col = QVBoxLayout()
        title = QLabel("UnHelper")
        title.setObjectName("Title")
        version = QLabel(f"v{CURRENT_VERSION} · Shipments/WMS 자동화 도우미")
        version.setObjectName("Version")
        title_col.addWidget(title)
        title_col.addWidget(version)
        header_layout.addLayout(title_col)
        header_layout.addStretch(1)
        self.settings_button = QPushButton("⚙")
        self.settings_button.setObjectName("SettingsButton")
        self.settings_button.setFixedSize(70, 70)
        self.settings_button.setToolTip("설정")
        self.settings_button.clicked.connect(self.show_settings)
        header_layout.addWidget(self.settings_button)
        layout.addWidget(header)

        self.main_tabs = QTabWidget()
        self.main_tabs.setObjectName("MainTabs")
        self.main_tabs.addTab(self._build_arrival_tabs(), "입차순번")
        self.main_tabs.addTab(self._build_raw_tabs(), "RAW")
        self.main_tabs.setCurrentIndex(1)
        layout.addWidget(self.main_tabs, 1)

        footer = QWidget()
        footer_layout = QVBoxLayout(footer)
        footer_layout.setContentsMargins(28, 4, 28, 14)
        footer_layout.setSpacing(6)
        status_row = QHBoxLayout()
        self.status_label = QLabel("대기 중")
        self.status_label.setObjectName("Status")
        self.open_folder_button = QPushButton("다운로드 폴더 열기")
        self.open_folder_button.clicked.connect(self.open_download_folder)
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.open_folder_button)
        footer_layout.addLayout(status_row)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumHeight(105)
        self.log_view.setPlaceholderText("진행 로그")
        footer_layout.addWidget(self.log_view)
        layout.addWidget(footer)

    def _build_arrival_tabs(self) -> QTabWidget:
        tabs = QTabWidget()
        tabs.addTab(self._placeholder("입차순번 트럭 기능은 다음 단계에서 연결됩니다."), "트럭")
        tabs.addTab(self._placeholder("입차순번 Milkrun 기능은 다음 단계에서 연결됩니다."), "Milkrun")
        return tabs

    def _build_raw_tabs(self) -> QTabWidget:
        self.raw_tabs = QTabWidget()
        self.raw_tabs.addTab(self._placeholder("RAW 트럭 기능은 다음 단계에서 연결됩니다."), "트럭")
        self.raw_tabs.addTab(self._build_raw_milkrun(), "Milkrun")
        self.raw_tabs.setCurrentIndex(1)
        return self.raw_tabs

    def _build_raw_milkrun(self) -> QWidget:
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(36, 14, 28, 4)
        outer.setSpacing(10)

        content = QHBoxLayout()
        content.setSpacing(18)
        self.raw_table = QTableWidget(0, 10)
        self.raw_table.setHorizontalHeaderLabels(
            [
                "거래처 이름",
                "밀크런 번호",
                "팔렛트 수",
                "박스 수",
                "팔렛트당 박스",
                "SKU ID",
                "SKU 명",
                "상품 무게(g)",
                "1팔렛트 무게(kg)",
                "분류",
            ]
        )
        self.raw_table.setWordWrap(False)
        self.raw_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.raw_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.raw_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        self.raw_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(9, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.verticalHeader().setVisible(False)
        self.raw_table.setAlternatingRowColors(True)
        content.addWidget(self.raw_table, 1)
        outer.addLayout(content, 1)

        actions = QHBoxLayout()
        actions.addStretch(1)
        self.get_data_button = QPushButton("데이터 얻기")
        self.get_data_button.setObjectName("PrimaryButton")
        self.get_data_button.setMinimumWidth(280)
        self.get_data_button.clicked.connect(self.start_milkrun_download)
        self.stop_button = QPushButton("작업 중지")
        self.stop_button.setObjectName("StopButton")
        self.stop_button.setVisible(False)
        self.stop_button.clicked.connect(self.cancel_milkrun_download)
        actions.addWidget(self.get_data_button)
        actions.addWidget(self.stop_button)
        actions.addStretch(1)
        outer.addLayout(actions)
        return page

    @staticmethod
    def _placeholder(message: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        label = QLabel(message)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("color: #71717A; font-size: 13pt;")
        layout.addWidget(label)
        return page

    def start_milkrun_download(self) -> None:
        if self._automation_worker_running():
            return
        driver = chromedriver_path()
        if not driver.is_file():
            QMessageBox.critical(self, "ChromeDriver 없음", f"ChromeDriver를 찾을 수 없습니다.\n{driver}")
            return

        configured_workbook = str(self.settings.value("milkrun_excel_path", "")).strip()
        if not configured_workbook:
            QMessageBox.warning(
                self,
                "Excel 파일 연결 필요",
                "먼저 설정에서 Raw_밀크런 시트가 있는 Excel 파일을 연결해 주세요.",
            )
            self.show_settings()
            return
        try:
            target_workbook = MilkrunExcelImporter.validate_target_path(configured_workbook)
        except ExcelImportError as exc:
            QMessageBox.warning(self, "Excel 파일 확인", str(exc))
            return

        download_dir = Path(
            str(self.settings.value("download_dir", str(default_download_dir())))
        ).expanduser()
        request = MilkrunDownloadRequest(download_dir=download_dir, center_name="안산2")
        self.log_view.clear()
        self.raw_table.setRowCount(0)
        self.current_products = ()
        self.current_pipeline_result = None
        self._pending_weight_summary = None
        self._pending_weight_failure = None
        self._pending_weight_cancel = ""
        self._credential_load_failure = None
        self._weight_finalize_pending = False
        self.append_log("Milkrun 텍스트 다운로드 및 Excel 반영 작업을 시작합니다.")
        self.append_log(f"연결된 Excel: {target_workbook}")
        self._set_automation_working(True)
        self.milkrun_worker = MilkrunWorker(request, driver, target_workbook)
        self.milkrun_worker.log_updated.connect(self.append_log)
        self.milkrun_worker.completed.connect(self._on_milkrun_completed)
        self.milkrun_worker.excel_failed.connect(self._on_milkrun_excel_failed)
        self.milkrun_worker.detail_failed.connect(self._on_milkrun_detail_failed)
        self.milkrun_worker.detail_cancelled.connect(self._on_milkrun_detail_cancelled)
        self.milkrun_worker.failed.connect(self._on_milkrun_failed)
        self.milkrun_worker.cancelled.connect(self._on_milkrun_cancelled)
        self.milkrun_worker.finished.connect(self._on_milkrun_finished)
        self.milkrun_worker.start()

    def cancel_milkrun_download(self) -> None:
        requested = False
        if self.milkrun_worker and self.milkrun_worker.isRunning():
            self.milkrun_worker.request_cancel()
            requested = True
        if self.weight_worker and self.weight_worker.isRunning():
            self.weight_worker.request_cancel()
            requested = True
        if requested:
            self.append_log("작업 중지를 요청했습니다.")
            self.status_label.setText("작업 중지 중...")
            self.stop_button.setEnabled(False)

    def _on_milkrun_completed(self, result) -> None:
        if self._closing_after_cancel:
            self.append_log("종료 요청이 처리 중이므로 WMS 무게 조회를 시작하지 않습니다.")
            return
        excel = result.excel
        daily = result.daily_inbound
        self.current_pipeline_result = result
        self._populate_milkrun_products(daily.products)
        self.append_log(f"다운로드 완료: {excel.source_file}")
        self.append_log(
            f"Excel 반영 완료: {excel.target_workbook} · {excel.sheet_name}!C1 · "
            f"{excel.rows}행 × {excel.columns}열"
        )
        if excel.filtered_rows:
            self.append_log(f"입고일이 어제인 행 {excel.filtered_rows}개를 제외했습니다.")
        self.append_log(f"일별 입고 상세 표시 완료: {len(daily.products)}개 상품")
        if daily.unmatched_orders:
            self.append_log(
                "오늘 카드에서 찾지 못한 발주번호: "
                + ", ".join(f"T{value}" for value in daily.unmatched_orders)
            )
        self.status_label.setText("일별 입고 표 완료 · WMS 무게 확인 준비")
        self._start_weight_lookup(daily.products)

    def _populate_milkrun_products(self, products) -> None:
        self.current_products = tuple(products)
        self._weight_records.clear()
        self._weight_failures.clear()
        self._weight_row_errors.clear()
        self.raw_table.setRowCount(len(self.current_products))
        for row_index, product in enumerate(self.current_products):
            values = (
                product.vendor_name,
                product.milkrun_number,
                product.pallet_count,
                product.box_count,
                "-",
                product.sku_id,
                normalize_product_name(product.sku_name),
                "-",
                "-",
            )
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if column_index == 6:
                    item.setToolTip(str(value))
                self.raw_table.setItem(row_index, column_index, item)
            category_button = QPushButton("?")
            category_button.setObjectName("CategoryButton")
            category_button.setProperty("classification", "?")
            category_button.setEnabled(False)
            category_button.setToolTip("WMS 무게 확인 후 분류를 변경할 수 있습니다.")
            category_button.clicked.connect(
                lambda _checked=False, sku_id=product.sku_id: self._cycle_category(sku_id)
            )
            self.raw_table.setCellWidget(row_index, 9, category_button)

    def _start_weight_lookup(self, products) -> None:
        if self.weight_worker and self.weight_worker.isRunning():
            return
        self._credential_load_failure = None
        if self._closing_after_cancel:
            self.append_log("종료 요청이 처리 중이므로 WMS 무게 조회를 시작하지 않습니다.")
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
            self.status_label.setText("오늘 표시할 상품 없음 · WMS 조회 생략")
            self.append_log("오늘 표시할 상품이 없어 WMS 무게 조회를 건너뜁니다.")
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
        self.weight_worker = ProductWeightWorker(
            products,
            self.product_memory_file,
            chromedriver_path(),
            credentials.wms_id,
            credentials.password,
            evidence_dir=download_dir,
        )
        self.weight_worker.log_updated.connect(self.append_log)
        self.weight_worker.record_ready.connect(self._on_weight_record_ready)
        self.weight_worker.sku_failed.connect(self._on_weight_sku_failed)
        self.weight_worker.completed.connect(self._on_weight_completed)
        self.weight_worker.failed.connect(self._on_weight_failed)
        self.weight_worker.cancelled.connect(self._on_weight_cancelled)
        self.weight_worker.finished.connect(self._on_weight_finished)
        self.status_label.setText("WMS 상품 무게 확인 중")
        self.append_log("표의 SKU별 상품 무게와 팔렛트 분류를 확인합니다.")
        self._set_automation_working(True)
        self.weight_worker.start()

    def _on_weight_record_ready(self, record: ProductMemoryRecord, cache_hit: bool) -> None:
        self._weight_records[record.sku_id] = record
        self._render_weight_record(record)
        source = "저장 정보" if cache_hit else "WMS"
        self.append_log(f"SKU {record.sku_id} 무게 반영: {source}")

    def _render_weight_record(self, record: ProductMemoryRecord) -> None:
        self._weight_row_errors.pop(record.sku_id, None)
        for row_index, product in enumerate(self.current_products):
            try:
                row_sku = normalize_sku_id(product.sku_id)
            except ValueError:
                continue
            if row_sku != record.sku_id:
                continue

            category = record.category_override or "?"
            error_text = ""
            if record.weight_grams is not None:
                self._set_table_text(row_index, 7, self._format_decimal(record.weight_grams))
                try:
                    boxes_per_pallet, pallet_weight_kg, automatic_category = calculate_pallet_measurement(
                        record.weight_grams,
                        product.box_count,
                        product.pallet_count,
                    )
                    self._set_table_text(row_index, 4, self._format_decimal(boxes_per_pallet, 3))
                    self._set_table_text(row_index, 8, self._format_decimal(pallet_weight_kg, 3))
                    category = record.category_override or automatic_category
                except (TypeError, ValueError) as exc:
                    self._set_table_text(row_index, 4, "-")
                    self._set_table_text(row_index, 8, "-")
                    error_text = str(exc)
                    self._weight_row_errors[record.sku_id] = error_text
            else:
                self._set_table_text(row_index, 7, "-")
                self._set_table_text(row_index, 4, "-")
                self._set_table_text(row_index, 8, "-")

            button = self.raw_table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                self._configure_category_button(
                    button,
                    category,
                    manual=record.category_override is not None,
                    enabled=not self._automation_worker_running(),
                    error_text=error_text,
                )

    def _on_weight_sku_failed(self, failure: SkuWeightFailure) -> None:
        self._weight_failures[failure.sku_id] = failure
        self.append_log(f"[WMS 조회 실패] SKU {failure.sku_id}: {failure.details.summary}")
        for row_index, product in enumerate(self.current_products):
            try:
                matches = normalize_sku_id(product.sku_id) == normalize_sku_id(failure.sku_id)
            except ValueError:
                matches = str(product.sku_id).strip() == failure.sku_id
            if not matches:
                continue
            button = self.raw_table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                button.setToolTip(failure.details.summary)

    def _on_weight_completed(self, summary: ProductWeightSummary) -> None:
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
        if self._pending_weight_cancel:
            self.status_label.setText("부분 완료 · WMS 무게 조회 취소")
            QMessageBox.information(
                self,
                "WMS 무게 조회 취소",
                "Milkrun 다운로드, Excel 반영과 일별 입고 표시는 완료됐습니다.\n"
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
                    "Milkrun 다운로드, Excel 반영과 일별 입고 표시는 완료됐지만 "
                    f"상품 무게 확인 중 {len(problems)}건의 문제가 발생했습니다.\n\n"
                    + "\n".join(problems[:20])
                ),
                detail="\n\n".join(details),
            )
            self._show_error_dialog(
                "WMS 상품 무게 조회 일부 실패",
                failure,
                category="Milkrun WMS 무게 조회",
            )
            return

        cache_hits = summary.cache_hits if summary is not None else 0
        wms_successes = summary.wms_successes if summary is not None else 0
        product_count = len(self.current_products)
        unmatched = (
            self.current_pipeline_result.daily_inbound.unmatched_orders
            if self.current_pipeline_result is not None
            else ()
        )
        self.status_label.setText(f"완료 · 상품 {product_count}개 · 무게 분류 완료")
        message = (
            "Milkrun 다운로드, Excel 값 반영, 일별 입고 상세와 WMS 무게 분류를 완료했습니다.\n\n"
            f"표시 상품: {product_count}개\n"
            f"저장된 무게 사용: {cache_hits}개\n"
            f"WMS 신규 조회: {wms_successes}개"
        )
        if unmatched:
            message += "\n\n오늘 카드에서 찾지 못한 발주번호: " + ", ".join(
                f"T{value}" for value in unmatched
            )
            QMessageBox.warning(self, "Milkrun 작업 완료", message)
        else:
            QMessageBox.information(self, "Milkrun 작업 완료", message)

    def _cycle_category(self, sku_value: object) -> None:
        if self._automation_worker_running():
            return
        try:
            sku_id = normalize_sku_id(sku_value)
            memory = ProductMemory(self.product_memory_file)
            record = memory.get(sku_id)
            override = record.category_override if record is not None else None
            product_name = ""
            for product in self.current_products:
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
            self._weight_records.pop(sku_id, None)
            self._render_unknown_sku(sku_id)
            self.append_log(f"SKU {sku_id} 수동 분류를 해제했습니다.")
        else:
            self._weight_records[sku_id] = updated
            self._render_weight_record(updated)
            source = "자동" if updated.category_override is None else "수동"
            self.append_log(f"SKU {sku_id} 분류 변경: {updated.effective_category or '?'} ({source})")

    def _render_unknown_sku(self, sku_id: str) -> None:
        for row_index, product in enumerate(self.current_products):
            try:
                if normalize_sku_id(product.sku_id) != sku_id:
                    continue
            except ValueError:
                continue
            self._set_table_text(row_index, 4, "-")
            self._set_table_text(row_index, 7, "-")
            self._set_table_text(row_index, 8, "-")
            button = self.raw_table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                self._configure_category_button(button, "?", manual=False, enabled=True)

    def _displayed_category_for_sku(self, sku_id: str) -> str:
        for row_index, product in enumerate(self.current_products):
            try:
                if normalize_sku_id(product.sku_id) != sku_id:
                    continue
            except ValueError:
                continue
            button = self.raw_table.cellWidget(row_index, 9)
            if isinstance(button, QPushButton):
                return button.text()
        return "?"

    def _set_category_buttons_enabled(self, enabled: bool) -> None:
        for row_index in range(self.raw_table.rowCount()):
            button = self.raw_table.cellWidget(row_index, 9)
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
    ) -> None:
        display = category if category in MANUAL_CATEGORIES else "?"
        button.setText(display)
        button.setProperty("classification", display)
        button.setEnabled(enabled)
        if error_text:
            tooltip = f"팔렛트 무게 계산 오류: {error_text}\n클릭해 수동 분류할 수 있습니다."
        elif manual:
            tooltip = "수동 분류입니다. 클릭하면 다음 분류로 변경됩니다."
        else:
            tooltip = "무게 기준 자동 분류입니다. 클릭하면 수동 분류로 변경됩니다."
        button.setToolTip(tooltip)
        button.style().unpolish(button)
        button.style().polish(button)

    def _set_table_text(self, row: int, column: int, value: str) -> None:
        item = self.raw_table.item(row, column)
        if item is None:
            item = QTableWidgetItem()
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.raw_table.setItem(row, column, item)
        item.setText(value)

    @staticmethod
    def _format_decimal(value: Decimal, digits: int | None = None) -> str:
        text = f"{value:.{digits}f}" if digits is not None else format(value, "f")
        return text.rstrip("0").rstrip(".") if "." in text else text

    def _on_milkrun_failed(self, failure: FailureDetails | object) -> None:
        details = FailureDetails.coerce(failure)
        self.append_log(f"[오류] {details.summary}")
        self.status_label.setText("작업 실패")
        if not self._closing_after_cancel:
            self._show_error_dialog("Milkrun 작업 실패", details, category="Milkrun 자동화")

    def _on_milkrun_excel_failed(
        self,
        downloaded_file: Path,
        failure: FailureDetails | object,
    ) -> None:
        details = FailureDetails.coerce(failure)
        self.append_log(f"다운로드 완료: {downloaded_file}")
        self.append_log(f"[Excel 반영 오류] {details.summary}")
        self.status_label.setText("부분 완료 · Excel 반영 실패")
        if not self._closing_after_cancel:
            partial_failure = FailureDetails(
                summary=(
                    "Milkrun 파일은 정상적으로 내려받았지만 연결된 Excel에 반영하지 못했습니다.\n"
                    f"다운로드 파일: {downloaded_file}\n\n{details.summary}"
                ),
                detail=f"다운로드 파일: {downloaded_file}\n\n{details.detail}",
            )
            self._show_error_dialog("Excel 반영 실패", partial_failure, category="Milkrun Excel 반영")

    def _on_milkrun_detail_failed(
        self,
        import_result,
        failure: FailureDetails | object,
    ) -> None:
        details = FailureDetails.coerce(failure)
        if import_result is not None:
            self.append_log(
                f"Excel 반영 완료: {import_result.target_workbook} · "
                f"{import_result.sheet_name}!C1"
            )
        self.append_log(f"[일별 입고 상세 오류] {details.summary}")
        self.status_label.setText("부분 완료 · Excel 반영 완료 · 일별 상세 실패")
        if not self._closing_after_cancel:
            prefix = (
                "Milkrun 파일 다운로드와 Excel 값 반영은 완료되었지만 "
                "일별 입고 상세를 가져오지 못했습니다."
            )
            partial_failure = FailureDetails(
                summary=f"{prefix}\n\n{details.summary}",
                detail=f"{prefix}\n\n{details.detail}",
            )
            self._show_error_dialog(
                "일별 입고 상세 조회 실패",
                partial_failure,
                category="Milkrun 일별 입고 상세",
            )

    def _on_milkrun_detail_cancelled(self, import_result, message: str) -> None:
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
                "작업을 중지했지만 그 전에 Milkrun 다운로드와 Excel 값 반영은 완료되었습니다.\n\n"
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
        return any(
            worker is not None and worker.isRunning()
            for worker in (self.milkrun_worker, self.weight_worker)
        )

    def _set_automation_working(self, working: bool) -> None:
        self.get_data_button.setEnabled(not working)
        self.settings_button.setEnabled(not working)
        self.stop_button.setVisible(working)
        self.stop_button.setEnabled(working)
        if working:
            self.status_label.setText("작업 중 · 로그인 화면이면 브라우저에서 직접 인증해 주세요")
        elif self.status_label.text().startswith("작업 중"):
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

    def _refresh_current_product_memory(self) -> None:
        if not self.current_products:
            return
        try:
            memory = ProductMemory(self.product_memory_file)
            self._weight_records.clear()
            self._weight_row_errors.clear()
            seen: set[str] = set()
            for product in self.current_products:
                try:
                    sku_id = normalize_sku_id(product.sku_id)
                except ValueError:
                    continue
                if sku_id in seen:
                    continue
                seen.add(sku_id)
                record = memory.get(sku_id)
                if record is None:
                    self._render_unknown_sku(sku_id)
                    continue
                self._weight_records[sku_id] = record
                self._render_weight_record(record)
            self.append_log("설정에서 변경한 상품 분류 메모리를 현재 표에 반영했습니다.")
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
        message = "Milkrun/WMS 작업 중에는 업데이트를 표시하거나 적용할 수 없습니다."
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
        dialog.setWindowTitle("정식 릴리즈 복구" if getattr(info, "is_release_restore", False) else "업데이트")
        dialog.setMinimumWidth(480)
        layout = QVBoxLayout(dialog)
        action = "정식 릴리즈로 복구" if getattr(info, "is_release_restore", False) else "새 업데이트"
        title = QLabel(f"{action}: v{info.version}")
        title.setStyleSheet("font-size: 16pt; font-weight: 800;")
        layout.addWidget(title)
        changelog = QPlainTextEdit()
        changelog.setReadOnly(True)
        changelog.setPlainText(info.changelog or "변경사항 없음")
        changelog.setMaximumHeight(180)
        layout.addWidget(changelog)
        mode = "델타" if getattr(info, "patch_mode", "full") == "delta" else "전체"
        size = AutoUpdater.format_size(info.patch_size) if info.patch_size else "알 수 없음"
        layout.addWidget(QLabel(f"패치: {mode} · {size}"))
        self.update_progress = QProgressBar()
        self.update_progress.setVisible(False)
        layout.addWidget(self.update_progress)
        self.update_status = QLabel("")
        layout.addWidget(self.update_status)
        buttons = QHBoxLayout()
        later = QPushButton("나중에")
        self.update_later_button = later
        self.apply_update_button = QPushButton("지금 적용")
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
                self.update_status.setText("Milkrun/WMS 작업 완료 후 업데이트를 다시 적용해 주세요.")
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
                "진행 중인 Milkrun/WMS 작업을 중지하고 종료하시겠습니까?",
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
