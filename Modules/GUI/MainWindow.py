from __future__ import annotations

import threading
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
from Modules.Shipments.MilkrunDownloader import (
    AutomationCancelled,
    MilkrunDownloadRequest,
    MilkrunDownloader,
)


class MilkrunWorker(QThread):
    log_updated = Signal(str)
    completed = Signal(object)
    failed = Signal(str)
    cancelled = Signal(str)

    def __init__(self, request: MilkrunDownloadRequest, driver_path: Path):
        super().__init__()
        self.request = request
        self.driver_path = driver_path
        self.stop_event = threading.Event()
        self.downloader: MilkrunDownloader | None = None

    def run(self) -> None:
        self.downloader = MilkrunDownloader(
            self.driver_path,
            log=self.log_updated.emit,
            stop_event=self.stop_event,
        )
        try:
            result = self.downloader.run(self.request)
            self.completed.emit(result)
        except AutomationCancelled as exc:
            self.cancelled.emit(str(exc))
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            self.downloader = None

    def request_cancel(self) -> None:
        self.stop_event.set()


class UpdateCheckWorker(QThread):
    update_available = Signal(object)
    no_update = Signal()
    failed = Signal(str)

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
            self.failed.emit(str(exc))


class ReleaseRestoreWorker(QThread):
    available = Signal(object)
    unavailable = Signal(str)

    def run(self) -> None:
        try:
            from Modules.Common.AutoUpdater import AutoUpdater

            has_restore, info, message = AutoUpdater("UnHelper", False).check_for_release_restore()
            if has_restore and info:
                self.available.emit(info)
            else:
                self.unavailable.emit(message)
        except Exception as exc:
            self.unavailable.emit(str(exc))


class UpdateDownloadWorker(QThread):
    progress = Signal(int)
    completed = Signal(str)
    failed = Signal(str)

    def __init__(self, info):
        super().__init__()
        self.info = info

    def run(self) -> None:
        try:
            from Modules.Common.AutoUpdater import AutoUpdater

            path = AutoUpdater("UnHelper").download_patch(self.info, self.progress.emit)
            self.completed.emit(str(path))
        except Exception as exc:
            self.failed.emit(str(exc))


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

        download_label = QLabel("텍스트 파일 저장 폴더")
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

        self.beta_checkbox = QCheckBox("Beta(테스트 릴리즈) 업데이트 받기")
        self.beta_checkbox.setChecked(self.settings.value("use_prerelease", False, type=bool))
        self.beta_checkbox.setToolTip("끄면 최신 정식 릴리즈로 복구할 수 있습니다.")
        layout.addWidget(self.beta_checkbox)

        version = QLabel(f"현재 버전: v{CURRENT_VERSION}")
        version.setStyleSheet("color: #A1A1AA;")
        layout.addWidget(version)

        buttons = QHBoxLayout()
        update = QPushButton("업데이트 확인")
        close = QPushButton("저장하고 닫기")
        update.clicked.connect(self._request_update)
        close.clicked.connect(self.accept)
        buttons.addWidget(update)
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

    def _persist(self) -> tuple[bool, bool]:
        previous_beta = self.settings.value("use_prerelease", False, type=bool)
        next_beta = self.beta_checkbox.isChecked()
        path = self.download_path.text().strip() or str(default_download_dir())
        self.settings.setValue("download_dir", path)
        self.settings.setValue("use_prerelease", next_beta)
        self.settings.sync()
        return previous_beta != next_beta, next_beta

    def _request_update(self) -> None:
        changed, next_beta = self._persist()
        super().accept()
        if changed:
            self.beta_changed.emit(next_beta)
        else:
            self.update_requested.emit()

    def accept(self) -> None:
        changed, next_beta = self._persist()
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
        self.raw_table = QTableWidget(9, 3)
        self.raw_table.setHorizontalHeaderLabels(["거래처", "상품이름", "팔렛 수량"])
        self.raw_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.raw_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.raw_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.raw_table.verticalHeader().setVisible(False)
        self.raw_table.setAlternatingRowColors(True)
        self.raw_table.setItem(0, 0, QTableWidgetItem("30차 정도"))
        content.addWidget(self.raw_table, 4)

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

        download_dir = Path(
            str(self.settings.value("download_dir", str(default_download_dir())))
        ).expanduser()
        request = MilkrunDownloadRequest(download_dir=download_dir, center_name="안산2")
        self.log_view.clear()
        self.append_log("Milkrun 텍스트 다운로드 작업을 시작합니다.")
        self._set_automation_working(True)
        self.milkrun_worker = MilkrunWorker(request, driver)
        self.milkrun_worker.log_updated.connect(self.append_log)
        self.milkrun_worker.completed.connect(self._on_milkrun_completed)
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
        self.append_log(f"완료: {result.file_path}")
        self.status_label.setText(f"완료 · {result.file_path.name}")
        QMessageBox.information(
            self,
            "다운로드 완료",
            f"Milkrun 텍스트 파일을 내려받았습니다.\n\n{result.file_path}",
        )

    def _on_milkrun_failed(self, message: str) -> None:
        self.append_log(f"[오류] {message}")
        self.status_label.setText("작업 실패")
        QMessageBox.critical(self, "Milkrun 다운로드 실패", message)

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
        dialog.exec()

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

    def _on_update_check_failed(self, message: str) -> None:
        self.append_log(f"[업데이트 확인 오류] {message}")
        if self.manual_update_check and not self._closing_after_workers:
            QMessageBox.warning(self, "업데이트 확인 실패", message)
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
        worker.finished.connect(lambda: self._on_worker_finished("restore_worker", worker))
        worker.start()

    def _on_restore_unavailable(self, message: str) -> None:
        self.append_log(message)
        if not self._closing_after_workers:
            QMessageBox.information(self, "정식 릴리즈", message)

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
        success = AutoUpdater("UnHelper").apply_update(
            Path(zip_path),
            info.version,
            info.manifest,
        )
        if success:
            if isinstance(self.update_dialog, UpdateDialog):
                self.update_dialog.allow_close = True
            QMessageBox.information(self, "업데이트", "앱을 종료한 뒤 패치를 적용하고 자동 재시작합니다.")
            self.close()
        else:
            self._on_update_download_failed("업데이트 적용 준비에 실패했습니다.")

    def _on_update_download_failed(self, message: str) -> None:
        self.update_status.setText(f"오류: {message}")
        self.apply_update_button.setEnabled(True)
        self.update_later_button.setEnabled(True)
        if isinstance(self.update_dialog, UpdateDialog):
            self.update_dialog.allow_close = True
        if not self._closing_after_workers:
            QMessageBox.warning(self, "업데이트 실패", message)

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
                "진행 중인 브라우저 작업을 중지하고 종료하시겠습니까?",
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
