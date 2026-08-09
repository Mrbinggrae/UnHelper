from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from Modules.Common.ErrorReport import (
    FailureDetails,
    build_error_report,
    build_github_issue_url,
    sanitize_report_text,
)
from Modules.Common.GitHubIssueReporter import (
    FINGERPRINT_MARKER,
    GitHubAPIError,
    GitHubIssueReporter,
    IssueReportResult,
    encode_token_value,
    load_github_issue_token,
)
from Modules.Common.Credentials import WMSCredentialStore
from Modules.Excel.MilkrunExcelImporter import (
    ExcelImportError,
    ExcelWorkbookOpenError,
    MilkrunExcelImporter,
)
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
            "password=plain-secret token:abc123 Authorization: Bearer ghp_realvalue "
            "wms_id=worker123 wms_pw=another-secret"
        )
        sanitized = sanitize_report_text(raw)

        self.assertNotIn(home, sanitized)
        self.assertIn("%USERPROFILE%", sanitized)
        self.assertNotIn("plain-secret", sanitized)
        self.assertNotIn("abc123", sanitized)
        self.assertNotIn("ghp_realvalue", sanitized)
        self.assertNotIn("worker123", sanitized)
        self.assertNotIn("another-secret", sanitized)

    def test_report_masks_quoted_and_whitespace_credentials_and_authorization_line(self) -> None:
        raw = (
            'wms_password="alpha beta, gamma" token=token with spaces\n'
            "user_password='delta epsilon,zeta'\n"
            "diagnostic prefix Authorization: Custom quoted token, with tail\n"
            "safe diagnostic line"
        )

        sanitized = sanitize_report_text(raw)

        for secret_fragment in (
            "alpha",
            "beta",
            "gamma",
            "token with spaces",
            "delta",
            "epsilon",
            "zeta",
            "diagnostic prefix",
            "Custom quoted token",
            "with tail",
        ):
            self.assertNotIn(secret_fragment, sanitized)
        self.assertIn("wms_password=***", sanitized)
        self.assertIn("user_password=***", sanitized)
        self.assertIn("Authorization: ***", sanitized)
        self.assertIn("safe diagnostic line", sanitized)

    def test_report_masks_json_and_python_mapping_credentials(self) -> None:
        raw = (
            '{"password": "alpha beta", "token": "abc123", "safe": "visible"}\n'
            "{'wms_password': 'delta epsilon', 'wms_id': 'worker123', 'count': 2}"
        )

        sanitized = sanitize_report_text(raw)

        for secret in ("alpha beta", "abc123", "delta epsilon", "worker123"):
            self.assertNotIn(secret, sanitized)
        self.assertIn('"password": "***"', sanitized)
        self.assertIn('"token": "***"', sanitized)
        self.assertIn("'wms_password': '***'", sanitized)
        self.assertIn("'wms_id': '***'", sanitized)
        self.assertIn('"safe": "visible"', sanitized)
        self.assertIn("'count': 2", sanitized)

    def test_issue_url_targets_unhelper_without_putting_report_in_query(self) -> None:
        failure = FailureDetails("실패", "token=do-not-send")
        report = build_error_report("Excel 반영 실패", failure)
        issue_url = build_github_issue_url("Excel 반영 실패", report)
        parsed = urlparse(issue_url)
        query = parse_qs(parsed.query)

        self.assertEqual(parsed.netloc, "github.com")
        self.assertEqual(parsed.path, "/Mrbinggrae/UnHelper/issues/new")
        self.assertEqual(query["title"], ["[UnHelper] Excel 반영 실패"])
        self.assertNotIn("body", query)
        self.assertNotIn("do-not-send", issue_url)

    def test_very_long_report_still_builds_a_short_issue_url(self) -> None:
        issue_url = build_github_issue_url("자동화 오류", "한글 traceback " * 10_000)

        self.assertLess(len(issue_url), 1_000)
        self.assertNotIn("body=", issue_url)

    def test_bug_report_token_can_be_loaded_from_encoded_environment_value(self) -> None:
        encoded = encode_token_value("test-fine-grained-token")
        with patch.dict(
            os.environ,
            {"UNHELPER_GITHUB_ISSUE_TOKEN": encoded},
            clear=False,
        ):
            self.assertEqual(load_github_issue_token(), "test-fine-grained-token")

    def test_bug_report_token_can_be_loaded_from_gitignored_data_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            token_path = Path(temp) / "bug_report_token.dat"
            token_path.write_text(encode_token_value("file-token"), encoding="utf-8")
            with (
                patch.dict(
                    os.environ,
                    {
                        "UNHELPER_GITHUB_ISSUE_TOKEN": "",
                        "GITHUB_ISSUE_TOKEN": "",
                    },
                    clear=False,
                ),
                patch(
                    "Modules.Common.GitHubIssueReporter._candidate_token_paths",
                    return_value=[token_path],
                ),
            ):
                self.assertEqual(load_github_issue_token(), "file-token")

    def test_reporter_creates_sanitized_issue_with_fingerprint(self) -> None:
        reporter = GitHubIssueReporter(token="test-token")
        calls = []

        def request_json(method, path, payload=None):
            calls.append((method, path, payload))
            if method == "GET":
                return []
            return {"number": 31, "html_url": "https://example.invalid/issues/31"}

        with (
            patch.object(reporter, "_request_json", side_effect=request_json),
            patch("Modules.Common.GitHubIssueReporter.time.sleep"),
        ):
            result = reporter.report_error(
                "Excel 반영 실패",
                "Traceback\ntoken=secret-value",
                {"category": "Milkrun"},
                report="## 오류\n\ntoken=secret-value\n안전한 진단 내용",
            )

        self.assertTrue(result.created)
        self.assertEqual(result.number, 31)
        self.assertEqual([call[0] for call in calls], ["GET", "GET", "POST"])
        payload = calls[-1][2]
        self.assertIn(f"<!-- {FINGERPRINT_MARKER}:", payload["body"])
        self.assertNotIn("secret-value", payload["body"])
        self.assertIn("안전한 진단 내용", payload["body"])
        self.assertEqual(payload["labels"], ["bug", "auto-report"])

    def test_reporter_adds_comment_to_matching_open_issue(self) -> None:
        reporter = GitHubIssueReporter(token="test-token")
        title = "동일 오류"
        error_msg = 'File "C:/one.py", line 123\nRuntimeError: busy'
        context = {"category": "Excel"}
        fingerprint = reporter._fingerprint(title, error_msg, context)
        marker = f"<!-- {FINGERPRINT_MARKER}: {fingerprint} -->"
        calls = []

        def request_json(method, path, payload=None):
            calls.append((method, path, payload))
            if method == "GET":
                return [
                    {
                        "number": 7,
                        "html_url": "https://example.invalid/issues/7",
                        "body": marker,
                    }
                ]
            return {"id": 1}

        with patch.object(reporter, "_request_json", side_effect=request_json):
            result = reporter.report_error(
                title,
                error_msg,
                context,
                report="재발생 오류 내용",
            )

        self.assertFalse(result.created)
        self.assertEqual(result.number, 7)
        self.assertIn("/issues/7/comments", calls[-1][1])
        self.assertIn("## 오류 재발생", calls[-1][2]["body"])

    def test_reporter_reuses_recent_creation_when_issue_list_is_stale(self) -> None:
        first_reporter = GitHubIssueReporter(token="test-token")
        second_reporter = GitHubIssueReporter(token="test-token")
        calls = []

        def request_json(_reporter, method, path, payload=None):
            calls.append((method, path, payload))
            if method == "GET":
                return []
            if path.endswith("/issues"):
                return {
                    "number": 41,
                    "html_url": "https://example.invalid/issues/41",
                }
            return {"id": 1}

        with (
            patch.object(
                GitHubIssueReporter,
                "_request_json",
                autospec=True,
                side_effect=request_json,
            ),
            patch("Modules.Common.GitHubIssueReporter.time.sleep"),
        ):
            first = first_reporter.report_error(
                "생성 직후 중복 방지",
                "RuntimeError: immediate recurrence",
                {"category": "Consistency regression"},
                report="첫 오류",
            )
            second = second_reporter.report_error(
                "생성 직후 중복 방지",
                "RuntimeError: immediate recurrence",
                {"category": "Consistency regression"},
                report="두 번째 오류",
            )

        self.assertTrue(first.created)
        self.assertFalse(second.created)
        self.assertEqual((first.number, second.number), (41, 41))
        self.assertEqual(sum(method == "GET" for method, _path, _payload in calls), 2)
        self.assertEqual(
            sum(
                method == "POST" and path.endswith("/issues")
                for method, path, _payload in calls
            ),
            1,
        )
        self.assertTrue(calls[-1][1].endswith("/issues/41/comments"))

    def test_reporter_get_requests_disable_http_caches(self) -> None:
        reporter = GitHubIssueReporter(token="test-token")

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            @staticmethod
            def read():
                return b"[]"

        with patch(
            "Modules.Common.GitHubIssueReporter.request.urlopen",
            return_value=Response(),
        ) as urlopen:
            self.assertEqual(reporter._request_json("GET", "/repos/test/issues"), [])

        api_request = urlopen.call_args.args[0]
        headers = {key.lower(): value for key, value in api_request.header_items()}
        self.assertEqual(headers["cache-control"], "no-cache")
        self.assertEqual(headers["pragma"], "no-cache")

    def test_reporter_retries_without_optional_labels_on_422(self) -> None:
        reporter = GitHubIssueReporter(token="test-token")
        responses = [
            [],
            [],
            GitHubAPIError(422, "Validation Failed"),
            {"number": 9},
        ]
        payloads = []

        def request_json(method, path, payload=None):
            payloads.append(payload.copy() if payload else payload)
            value = responses.pop(0)
            if isinstance(value, Exception):
                raise value
            return value

        with (
            patch.object(reporter, "_request_json", side_effect=request_json),
            patch("Modules.Common.GitHubIssueReporter.time.sleep"),
        ):
            result = reporter.report_error("오류", "상세", report="보고서")

        self.assertTrue(result.created)
        self.assertIn("labels", payloads[2])
        self.assertNotIn("labels", payloads[3])

    def test_error_dialog_submits_report_without_opening_browser(self) -> None:
        class SuccessfulReporter:
            def report_error(self, title, error_msg, context, *, report=None):
                self.args = (title, error_msg, context, report)
                return IssueReportResult(True, 17, "https://example.invalid/issues/17", "abc")

        dialog = ErrorReportDialog(
            "Milkrun 작업 실패",
            FailureDetails("로그인 확인 실패", "Traceback\ntoken=hidden"),
        )
        try:
            with (
                patch(
                    "Modules.GUI.BugReportWorker.GitHubIssueReporter",
                    SuccessfulReporter,
                ),
                patch("Modules.GUI.Dialogs.QMessageBox.information") as information,
            ):
                dialog.report_button.click()
                worker = dialog._report_worker
                self.assertIsNotNone(worker)
                self.assertTrue(worker.wait(5_000))
                self.app.processEvents()
                self.app.processEvents()

            self.assertNotIn("hidden", dialog.report)
            self.assertIn("#17", dialog.action_status.text())
            self.assertTrue(dialog.report_button.isEnabled())
            information.assert_called_once()
        finally:
            dialog.close()

    def test_error_dialog_shows_report_failure_and_keeps_copy_fallback(self) -> None:
        class FailingReporter:
            def __init__(self):
                raise RuntimeError("토큰 없음")

        dialog = ErrorReportDialog("작업 실패", FailureDetails("실패", "상세 오류"))
        try:
            with (
                patch(
                    "Modules.GUI.BugReportWorker.GitHubIssueReporter",
                    FailingReporter,
                ),
                patch("Modules.GUI.Dialogs.QMessageBox.warning") as warning,
            ):
                dialog.report_button.click()
                worker = dialog._report_worker
                self.assertIsNotNone(worker)
                self.assertTrue(worker.wait(5_000))
                self.app.processEvents()
                self.app.processEvents()

            self.assertIn("토큰 없음", dialog.action_status.text())
            self.assertTrue(dialog.copy_button.isEnabled())
            warning.assert_called_once()
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

    def test_milkrun_worker_routes_open_excel_to_close_prompt_signal(self) -> None:
        worker = MilkrunWorker(
            MilkrunDownloadRequest(download_dir=Path.cwd()),
            Path("chromedriver.exe"),
            Path("입고스케줄관리.xlsx"),
        )
        prompts = []
        failures = []
        worker.excel_close_required.connect(
            lambda downloaded_file, message: prompts.append((downloaded_file, message))
        )
        worker.failed.connect(failures.append)

        with (
            patch.object(
                MilkrunExcelImporter,
                "validate_workbook",
                side_effect=ExcelWorkbookOpenError("Excel을 닫아 주세요."),
            ),
            patch("Modules.GUI.MainWindow.MilkrunDownloader") as downloader_class,
        ):
            worker.run()

        self.assertEqual(prompts, [(None, "Excel을 닫아 주세요.")])
        self.assertEqual(failures, [])
        downloader_class.assert_not_called()

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
