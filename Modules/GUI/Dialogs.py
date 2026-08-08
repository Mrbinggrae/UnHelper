from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices, QGuiApplication
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
    build_github_issue_url,
)
from Modules.Common.paths import bundled_root


class ErrorReportDialog(QDialog):
    def __init__(
        self,
        title: str,
        failure: FailureDetails | object,
        *,
        context: dict[str, object] | None = None,
        parent=None,
        open_url: Callable[[QUrl], bool] | None = None,
    ):
        super().__init__(parent)
        self.failure = FailureDetails.coerce(failure)
        self.report = build_error_report(title, self.failure, context=context)
        self.issue_url = build_github_issue_url(title, self.report)
        self._open_url = open_url or QDesktopServices.openUrl

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
            "신고 버튼을 누르면 개인정보와 인증정보를 가린 보고서를 복사하고 GitHub 작성 화면을 엽니다."
        )
        self.action_status.setWordWrap(True)
        self.action_status.setStyleSheet("color: #A1A1AA;")
        layout.addWidget(self.action_status)

        buttons = QHBoxLayout()
        self.copy_button = QPushButton("오류 내용 복사")
        self.report_button = QPushButton("GitHub 이슈 신고")
        close_button = QPushButton("닫기")
        self.copy_button.clicked.connect(self.copy_report)
        self.report_button.clicked.connect(self.open_github_issue)
        close_button.clicked.connect(self.accept)
        buttons.addWidget(self.copy_button)
        buttons.addWidget(self.report_button)
        buttons.addStretch(1)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)

    def copy_report(self) -> None:
        clipboard = QGuiApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(self.report)
        self.action_status.setText("개인정보와 인증정보를 가린 오류 내용이 클립보드에 복사되었습니다.")

    def open_github_issue(self) -> None:
        self.copy_report()
        if self._open_url(QUrl(self.issue_url)):
            self.action_status.setText(
                "GitHub 이슈 작성 화면을 열었습니다. 내용을 확인한 뒤 제출해 주세요. "
                "본문이 잘렸다면 클립보드의 전체 내용을 붙여넣을 수 있습니다."
            )
            return
        self.action_status.setText(
            "브라우저를 열지 못했습니다. 복사된 오류 내용을 UnHelper GitHub Issues에 붙여넣어 주세요."
        )


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
