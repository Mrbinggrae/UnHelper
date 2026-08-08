from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStyle,
    QVBoxLayout,
)

from Modules.Common.ErrorReport import (
    FailureDetails,
    build_error_report,
)
from Modules.Common.paths import bundled_root
from Modules.GUI.BugReportWorker import BugReportWorker


class ErrorReportDialog(QDialog):
    def __init__(
        self,
        title: str,
        failure: FailureDetails | object,
        *,
        context: dict[str, object] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self.failure = FailureDetails.coerce(failure)
        self.report_context = dict(context or {})
        self.report = build_error_report(title, self.failure, context=self.report_context)
        self.report_title = title
        self._report_worker: BugReportWorker | None = None
        self._report_in_progress = False

        self.setWindowTitle(title)
        self.setMinimumSize(680, 440)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)

        heading_row = QHBoxLayout()
        icon = QLabel()
        icon.setPixmap(
            self.style()
            .standardIcon(QStyle.StandardPixmap.SP_MessageBoxCritical)
            .pixmap(42, 42)
        )
        heading = QLabel("작업 중 오류가 발생했습니다.")
        heading.setStyleSheet("font-size: 16pt; font-weight: 800;")
        heading_row.addWidget(icon)
        heading_row.addWidget(heading, 1)
        layout.addLayout(heading_row)

        self.summary_label = QLabel(self.failure.summary)
        self.summary_label.setWordWrap(True)
        self.summary_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.summary_label)

        details_label = QLabel("오류 상세 내용")
        details_label.setStyleSheet("font-weight: 700;")
        layout.addWidget(details_label)
        self.detail_view = QPlainTextEdit()
        self.detail_view.setReadOnly(True)
        self.detail_view.setPlainText(self.failure.detail or self.failure.summary)
        layout.addWidget(self.detail_view, 1)

        self.action_status = QLabel(
            "신고 버튼을 누르면 개인정보와 인증정보를 가린 보고서를 자동으로 전송합니다. "
            "사용자의 GitHub 로그인은 필요하지 않습니다."
        )
        self.action_status.setWordWrap(True)
        self.action_status.setStyleSheet("color: #A1A1AA;")
        layout.addWidget(self.action_status)

        buttons = QHBoxLayout()
        self.copy_button = QPushButton("오류 내용 복사")
        self.report_button = QPushButton("GitHub 오류 자동 신고")
        self.close_button = QPushButton("닫기")
        self.copy_button.clicked.connect(self.copy_report)
        self.report_button.clicked.connect(self.submit_github_issue)
        self.close_button.clicked.connect(self.accept)
        buttons.addWidget(self.copy_button)
        buttons.addWidget(self.report_button)
        buttons.addStretch(1)
        buttons.addWidget(self.close_button)
        layout.addLayout(buttons)

    def copy_report(self) -> None:
        clipboard = QGuiApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(self.report)
        self.action_status.setText("개인정보와 인증정보를 가린 오류 내용이 클립보드에 복사되었습니다.")

    def submit_github_issue(self) -> None:
        if self._report_in_progress:
            return

        self._report_in_progress = True
        self.report_button.setEnabled(False)
        self.copy_button.setEnabled(False)
        self.close_button.setEnabled(False)
        self.action_status.setText("오류 신고를 안전하게 전송하고 있습니다...")

        worker = BugReportWorker(
            self.report_title,
            self.failure,
            self.report,
            self.report_context,
        )
        self._report_worker = worker
        worker.report_finished.connect(self._on_report_finished)
        worker.finished.connect(self._on_report_worker_finished)
        worker.start()

    def _on_report_finished(self, success: bool, message: str, issue_url: str) -> None:
        del issue_url
        self.action_status.setText(message)
        if success:
            QMessageBox.information(self, "신고 완료", message)
        else:
            QMessageBox.warning(
                self,
                "신고 실패",
                f"오류 신고를 전송하지 못했습니다.\n\n{message}\n\n"
                "오류 내용 복사 버튼으로 내용을 보관한 뒤 관리자에게 전달해 주세요.",
            )

    def _on_report_worker_finished(self) -> None:
        worker = self._report_worker
        self._report_worker = None
        self._report_in_progress = False
        self.report_button.setEnabled(True)
        self.copy_button.setEnabled(True)
        self.close_button.setEnabled(True)
        if worker is not None:
            worker.deleteLater()

    def accept(self) -> None:
        if self._report_in_progress:
            self.action_status.setText("오류 신고 전송이 끝날 때까지 잠시 기다려 주세요.")
            return
        super().accept()

    def reject(self) -> None:
        if self._report_in_progress:
            self.action_status.setText("오류 신고 전송이 끝날 때까지 잠시 기다려 주세요.")
            return
        super().reject()


class UpdateHistoryDialog(QDialog):
    def __init__(self, parent=None, *, history_path: Path | None = None):
        super().__init__(parent)
        self.history_path = history_path or (bundled_root() / "UPDATE_HISTORY.txt")
        self.setWindowTitle("업데이트 내역")
        self.resize(620, 520)

        layout = QVBoxLayout(self)
        self.history_view = QPlainTextEdit()
        self.history_view.setReadOnly(True)
        self.history_view.setPlainText(self._read_history())
        layout.addWidget(self.history_view, 1)
        close_button = QPushButton("닫기")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button, alignment=Qt.AlignmentFlag.AlignRight)

    def _read_history(self) -> str:
        try:
            return self.history_path.read_text(encoding="utf-8")
        except OSError as exc:
            return f"업데이트 내역을 불러올 수 없습니다.\n\n{exc}"
