from __future__ import annotations

import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QSettings, QThread, QTimer, Qt, QUrl, Signal
from PySide6.QtGui import QCloseEvent, QDesktopServices
from PySide6.QtWidgets import (
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

from Modules.Common.paths import chromedriver_path, default_download_dir
from Modules.Common.version import CURRENT_VERSION
from Modules.Common.ErrorReport import FailureDetails
from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportCancelled,
    ExcelImportError,
    MilkrunExcelImportResult,
    MilkrunExcelImporter,
)
from Modules.GUI.Dialogs import ErrorReportDialog, UpdateHistoryDialog
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
            )
            if self.stop_event.is_set():
                raise AutomationCancelled("사용자가 작업을 중지했습니다.")
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


class SettingsDialog(QDialog):
    beta_changed = Signal(bool)
    update_requested = Signal()
    history_requested = Signal()

    def __init__(self, settings: QSettings, parent=None):
        super().__init__(parent)
        self.settings = settings
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
        self.settings.setValue("download_dir", path)
        self.settings.setValue("milkrun_excel_path", excel_path)
        self.settings.setValue("use_prerelease", next_beta)
        self.settings.sync()
        return previous_beta != next_beta, next_beta

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
        version = QLabel(f"v{CURRENT_VERSION} · Shipments 자동화 도우미")
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
        self.raw_table = QTableWidget(0, 6)
        self.raw_table.setHorizontalHeaderLabels(
            ["거래처 이름", "밀크런 번호", "팔렛트 수", "박스 수", "SKU ID", "SKU 명"]
        )
        self.raw_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.raw_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.raw_table.verticalHeader().setVisible(False)
        self.raw_table.setAlternatingRowColors(True)
        content.addWidget(self.raw_table, 6)

        categories = QVBoxLayout()
        categories.setSpacing(7)
        for _ in range(3):
            row = QHBoxLayout()
            for text in ("고단", "중량", "경량", "?"):
                button = QPushButton(text)
                button.setObjectName("CategoryButton")
                button.setEnabled(False)
                row.addWidget(button)
            categories.addLayout(row)
        categories.addStretch(1)
        content.addLayout(categories, 2)
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
        if self.milkrun_worker and self.milkrun_worker.isRunning():
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
        if self.milkrun_worker and self.milkrun_worker.isRunning():
            self.append_log("작업 중지를 요청했습니다.")
            self.status_label.setText("작업 중지 중...")
            self.stop_button.setEnabled(False)
            self.milkrun_worker.request_cancel()

    def _on_milkrun_completed(self, result) -> None:
        excel = result.excel
        daily = result.daily_inbound
        self._populate_milkrun_products(daily.products)
        self.append_log(f"다운로드 완료: {excel.source_file}")
        self.append_log(
            f"Excel 반영 완료: {excel.target_workbook} · {excel.sheet_name}!C1 · "
            f"{excel.rows}행 × {excel.columns}열"
        )
        self.append_log(f"일별 입고 상세 표시 완료: {len(daily.products)}개 상품")
        missing_text = ""
        if daily.unmatched_orders:
            missing_text = "\n\n오늘 카드에서 찾지 못한 발주번호: " + ", ".join(
                f"T{value}" for value in daily.unmatched_orders
            )
            self.status_label.setText(
                f"완료 · 상품 {len(daily.products)}개 · 일부 발주 미조회"
            )
        else:
            self.status_label.setText(f"완료 · 상품 {len(daily.products)}개")
        if not self._closing_after_cancel:
            message = (
                "Milkrun 파일을 내려받고 연결된 Excel에 값을 반영한 뒤 "
                "일별 입고 상세를 표시했습니다.\n\n"
                f"다운로드: {excel.source_file}\n"
                f"대상: {excel.target_workbook}\n"
                f"범위: {excel.sheet_name}!C1 ({excel.rows}행 × {excel.columns}열)\n"
                f"표시 상품: {len(daily.products)}개{missing_text}"
            )
            show_message = QMessageBox.warning if daily.unmatched_orders else QMessageBox.information
            show_message(
                self,
                "Milkrun 작업 완료",
                message,
            )

    def _populate_milkrun_products(self, products) -> None:
        self.raw_table.setRowCount(len(products))
        for row_index, product in enumerate(products):
            values = (
                product.vendor_name,
                product.milkrun_number,
                product.pallet_count,
                product.box_count,
                product.sku_id,
                product.sku_name,
            )
            for column_index, value in enumerate(values):
                self.raw_table.setItem(row_index, column_index, QTableWidgetItem(str(value)))

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
        self._set_automation_working(False)
        self.milkrun_worker = None
        if self._closing_after_cancel:
            QTimer.singleShot(0, self.close)

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
        if self.milkrun_worker and self.milkrun_worker.isRunning():
            self.status_label.setText(message)

    def show_settings(self) -> None:
        dialog = SettingsDialog(self.settings, self)
        dialog.update_requested.connect(lambda: self.check_for_update(manual=True))
        dialog.beta_changed.connect(self._on_beta_changed)
        dialog.history_requested.connect(self.show_update_history)
        dialog.exec()

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

    def _show_update_dialog(self, info) -> None:
        from Modules.Common.AutoUpdater import AutoUpdater

        self.manual_update_check = False
        if self._closing_after_workers:
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
        if self.milkrun_worker and self.milkrun_worker.isRunning():
            QMessageBox.warning(self, "작업 실행 중", "Milkrun 작업을 먼저 끝내거나 중지해 주세요.")
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
        if self.milkrun_worker and self.milkrun_worker.isRunning():
            answer = QMessageBox.question(
                self,
                "작업 중",
                "진행 중인 Milkrun 다운로드/Excel 반영 작업을 중지하고 종료하시겠습니까?",
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
