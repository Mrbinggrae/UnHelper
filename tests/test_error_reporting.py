from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings, QUrl
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QApplication

from Modules.Common.ErrorReport import (
    FailureDetails,
    build_error_report,
    build_github_issue_url,
    sanitize_report_text,
)
from Modules.Excel.MilkrunExcelImporter import ExcelImportError, MilkrunExcelImporter
from Modules.GUI.Dialogs import ErrorReportDialog, UpdateHistoryDialog
from Modules.GUI.MainWindow import MainWindow, MilkrunWorker, SettingsDialog
from Modules.Shipments.MilkrunDownloader import MilkrunDownloadRequest


class ErrorReportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_failure_details_preserves_traceback(self) -> None:
        try:
            raise RuntimeError("수집 실패")
        except RuntimeError as exc:
            failure = FailureDetails.from_exception(exc)

        self.assertEqual(failure.summary, "수집 실패")
        self.assertIn("Traceback (most recent call last)", failure.detail)
        self.assertIn("RuntimeError: 수집 실패", failure.detail)

    def test_report_masks_credentials_and_user_profile(self) -> None:
        home = str(Path.home())
        raw = (
            f'File "{home}\\project\\worker.py"\n'
            "password=plain-secret token:abc123 Authorization: Bearer ghp_realvalue"
        )
        sanitized = sanitize_report_text(raw)

        self.assertNotIn(home, sanitized)
        self.assertIn("%USERPROFILE%", sanitized)
        self.assertNotIn("plain-secret", sanitized)
        self.assertNotIn("abc123", sanitized)
        self.assertNotIn("ghp_realvalue", sanitized)

    def test_issue_url_targets_unhelper_and_prefills_sanitized_report(self) -> None:
        failure = FailureDetails("실패", "token=do-not-send")
        report = build_error_report("Excel 반영 실패", failure)
        issue_url = build_github_issue_url("Excel 반영 실패", report)
        parsed = urlparse(issue_url)
        query = parse_qs(parsed.query)

        self.assertEqual(parsed.netloc, "github.com")
        self.assertEqual(parsed.path, "/Mrbinggrae/UnHelper/issues/new")
        self.assertEqual(query["title"], ["[UnHelper] Excel 반영 실패"])
        self.assertNotIn("do-not-send", query["body"][0])

    def test_error_dialog_copies_report_and_opens_prefilled_issue(self) -> None:
        opened = []
        dialog = ErrorReportDialog(
            "Milkrun 작업 실패",
            FailureDetails("로그인 확인 실패", "Traceback\ntoken=hidden"),
            open_url=lambda url: opened.append(url.toString()) or True,
        )
        try:
            dialog.report_button.click()
            clipboard = QGuiApplication.clipboard()
            self.assertIsNotNone(clipboard)
            self.assertEqual(clipboard.text(), dialog.report)
            self.assertNotIn("hidden", dialog.report)
            self.assertEqual([QUrl(value) for value in opened], [QUrl(dialog.issue_url)])
            self.assertIn("GitHub 이슈 작성 화면", dialog.action_status.text())
        finally:
            dialog.close()

    def test_update_history_dialog_reads_bundled_style_history_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            history = Path(temp) / "UPDATE_HISTORY.txt"
            history.write_text("## [v9.9.9]\n- 테스트 변경", encoding="utf-8")
            dialog = UpdateHistoryDialog(history_path=history)
            try:
                self.assertIn("[v9.9.9]", dialog.history_view.toPlainText())
                self.assertTrue(dialog.history_view.isReadOnly())
            finally:
                dialog.close()

    def test_settings_exposes_update_history_action(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            settings = QSettings(str(Path(temp) / "settings.ini"), QSettings.Format.IniFormat)
            dialog = SettingsDialog(settings)
            emitted = []
            dialog.history_requested.connect(lambda: emitted.append(True))
            try:
                dialog.update_history_button.click()
                self.assertEqual(emitted, [True])
            finally:
                dialog.close()

    def test_milkrun_worker_emits_traceback_for_excel_preflight_failure(self) -> None:
        worker = MilkrunWorker(
            MilkrunDownloadRequest(download_dir=Path.cwd()),
            Path("chromedriver.exe"),
            Path("입고스케줄관리.xlsx"),
        )
        failures = []
        worker.failed.connect(failures.append)
        with patch.object(
            MilkrunExcelImporter,
            "validate_workbook",
            side_effect=ExcelImportError("Raw_밀크런 시트 없음"),
        ):
            worker.run()

        self.assertEqual(len(failures), 1)
        self.assertIsInstance(failures[0], FailureDetails)
        self.assertIn("Traceback (most recent call last)", failures[0].detail)

    def test_main_window_routes_milkrun_failure_to_report_dialog(self) -> None:
        window = MainWindow(smoke_test=True)
        try:
            with patch("Modules.GUI.MainWindow.ErrorReportDialog") as dialog_class:
                window._on_milkrun_failed(FailureDetails("오류 요약", "상세 traceback"))
            dialog_class.return_value.exec.assert_called_once_with()
            self.assertEqual(window.status_label.text(), "작업 실패")
        finally:
            window.close()

    def test_manual_update_failure_uses_reportable_error_dialog(self) -> None:
        window = MainWindow(smoke_test=True)
        window.manual_update_check = True
        try:
            with patch("Modules.GUI.MainWindow.ErrorReportDialog") as dialog_class:
                window._on_update_check_failed(
                    FailureDetails("네트워크 확인 실패", "상세 update traceback")
                )

            dialog_class.return_value.exec.assert_called_once_with()
            self.assertFalse(window.manual_update_check)
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
