from __future__ import annotations

from PySide6.QtCore import QThread, Signal

from Modules.Common.ErrorReport import FailureDetails, sanitize_report_text
from Modules.Common.GitHubIssueReporter import GitHubIssueReporter


class BugReportWorker(QThread):
    """Submit a GitHub issue without blocking the Qt event loop."""

    report_finished = Signal(bool, str, str)

    def __init__(
        self,
        title: str,
        failure: FailureDetails | object,
        report: str,
        context: dict[str, object] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self.title = title
        self.failure = FailureDetails.coerce(failure)
        self.report = report
        self.context = dict(context or {})

    def run(self) -> None:
        try:
            result = GitHubIssueReporter().report_error(
                self.title,
                self.failure.detail or self.failure.summary,
                self.context,
                report=self.report,
            )
            if result.created:
                message = f"새 오류 신고가 접수되었습니다. (#{result.number})"
            else:
                message = f"같은 오류 신고에 재발생 내역을 추가했습니다. (#{result.number})"
            self.report_finished.emit(True, message, result.url)
        except Exception as exc:
            message = sanitize_report_text(str(exc)).strip() or type(exc).__name__
            self.report_finished.emit(False, message, "")
