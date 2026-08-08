"""Login-free GitHub Issues reporter for UnHelper.

The GitHub token is deliberately loaded from runtime state.  It must never be
committed to the repository; release builds may include the separately-created
``bug_report_token.dat`` file when it is present beside the spec file.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from urllib import error, parse, request

from Modules.Common.ErrorReport import (
    FailureDetails,
    build_error_report,
    sanitize_report_text,
)
from Modules.Common.paths import bundled_root, project_root


ISSUE_GITHUB_OWNER = "Mrbinggrae"
ISSUE_GITHUB_REPO = "UnHelper"
TOKEN_ENV_VARS = (
    "UNHELPER_GITHUB_ISSUE_TOKEN",
    "GITHUB_ISSUE_TOKEN",
)
TOKEN_FILE_NAME = "bug_report_token.dat"
DEFAULT_LABELS = ("bug", "auto-report")
FINGERPRINT_MARKER = "unhelper-bug-report-fingerprint"
ISSUE_LOOKUP_RETRY_SECONDS = 0.25
RECENT_ISSUE_TTL_SECONDS = 30.0
_OBFUSCATION_KEY = hashlib.sha256(b"UnHelper GitHub Issue Reporter v1").digest()
_REPORT_LOCK = threading.Lock()
_RECENT_ISSUES: dict[
    tuple[str, str, str], tuple[float, dict[str, object]]
] = {}


@dataclass(frozen=True)
class IssueReportResult:
    created: bool
    number: int
    url: str
    fingerprint: str


class GitHubAPIError(RuntimeError):
    def __init__(self, status: int, message: str):
        label = str(status) if status else "network"
        super().__init__(f"GitHub API 오류({label}): {message}")
        self.status = status
        self.message = message


class MissingGitHubTokenError(RuntimeError):
    pass


class GitHubIssueReporter:
    def __init__(
        self,
        *,
        owner: str = ISSUE_GITHUB_OWNER,
        repo: str = ISSUE_GITHUB_REPO,
        token: str | None = None,
        timeout: int = 15,
    ):
        self.owner = owner
        self.repo = repo
        self.token = (token or load_github_issue_token()).strip()
        self.timeout = timeout
        if not self.token:
            raise MissingGitHubTokenError(
                "오류 신고용 GitHub 토큰이 설정되지 않았습니다. 관리자에게 문의해 주세요."
            )

    def report_error(
        self,
        title: str,
        error_msg: str,
        context: dict[str, object] | None = None,
        *,
        report: str | None = None,
    ) -> IssueReportResult:
        # Serialize the check/create pair so simultaneous reporters in this
        # process cannot both create an issue for the same fingerprint.
        with _REPORT_LOCK:
            return self._report_error(title, error_msg, context, report=report)

    def _report_error(
        self,
        title: str,
        error_msg: str,
        context: dict[str, object] | None = None,
        *,
        report: str | None = None,
    ) -> IssueReportResult:
        context = context or {}
        fingerprint = self._fingerprint(title, error_msg, context)
        safe_report = sanitize_report_text(
            report
            or build_error_report(
                title,
                FailureDetails(summary=title, detail=error_msg),
                context=context,
            )
        )

        existing = self._find_existing_issue(fingerprint)
        if existing:
            comment = self._build_comment_body(safe_report, fingerprint)
            self._request_json(
                "POST",
                f"/repos/{self.owner}/{self.repo}/issues/{existing['number']}/comments",
                {"body": comment},
            )
            return IssueReportResult(
                created=False,
                number=int(existing["number"]),
                url=str(existing.get("html_url") or ""),
                fingerprint=fingerprint,
            )

        payload = {
            "title": self._format_issue_title(title, context),
            "body": self._build_issue_body(safe_report, fingerprint),
            "labels": list(DEFAULT_LABELS),
        }
        try:
            issue = self._request_json(
                "POST", f"/repos/{self.owner}/{self.repo}/issues", payload
            )
        except GitHubAPIError as exc:
            # A fine-grained token with Issues write permission can create an
            # issue even when a requested label does not exist.  Retry without
            # optional labels for that common validation failure.
            if exc.status != 422:
                raise
            payload.pop("labels", None)
            issue = self._request_json(
                "POST", f"/repos/{self.owner}/{self.repo}/issues", payload
            )

        if not isinstance(issue, dict):
            raise GitHubAPIError(0, "GitHub 이슈 생성 응답이 올바르지 않습니다.")
        self._remember_recent_issue(fingerprint, issue)
        return IssueReportResult(
            created=True,
            number=int(issue["number"]),
            url=str(issue.get("html_url") or ""),
            fingerprint=fingerprint,
        )

    def _find_existing_issue(self, fingerprint: str) -> dict[str, object] | None:
        recent = self._recent_issue(fingerprint)
        if recent is not None:
            return recent

        marker = f"<!-- {FINGERPRINT_MARKER}: {fingerprint} -->"
        for attempt in range(2):
            for page in range(1, 4):
                query = parse.urlencode(
                    {"state": "open", "per_page": "100", "page": str(page)}
                )
                issues = self._request_json(
                    "GET", f"/repos/{self.owner}/{self.repo}/issues?{query}"
                )
                if not isinstance(issues, list) or not issues:
                    break
                for issue in issues:
                    if not isinstance(issue, dict) or "pull_request" in issue:
                        continue
                    if marker in str(issue.get("body") or ""):
                        self._remember_recent_issue(fingerprint, issue)
                        return issue
            if attempt == 0:
                time.sleep(ISSUE_LOOKUP_RETRY_SECONDS)
        return None

    def _recent_issue(self, fingerprint: str) -> dict[str, object] | None:
        key = (self.owner.casefold(), self.repo.casefold(), fingerprint)
        cached = _RECENT_ISSUES.get(key)
        if cached is None:
            return None
        expires_at, issue = cached
        if time.monotonic() >= expires_at:
            _RECENT_ISSUES.pop(key, None)
            return None
        return issue

    def _remember_recent_issue(
        self, fingerprint: str, issue: dict[str, object]
    ) -> None:
        key = (self.owner.casefold(), self.repo.casefold(), fingerprint)
        _RECENT_ISSUES[key] = (
            time.monotonic() + RECENT_ISSUE_TTL_SECONDS,
            issue,
        )

    def _request_json(
        self,
        method: str,
        path: str,
        payload: dict[str, object] | None = None,
    ) -> object:
        url = f"https://api.github.com{path}"
        data = None
        headers = {
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {self.token}",
            "User-Agent": "UnHelper-BugReporter",
            "X-GitHub-Api-Version": "2022-11-28",
        }
        if method.upper() == "GET":
            headers["Cache-Control"] = "no-cache"
            headers["Pragma"] = "no-cache"
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"

        api_request = request.Request(url, data=data, headers=headers, method=method)
        try:
            with request.urlopen(api_request, timeout=self.timeout) as response:
                raw = response.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise GitHubAPIError(exc.code, _extract_github_error_message(detail)) from exc
        except (error.URLError, TimeoutError, OSError) as exc:
            reason = getattr(exc, "reason", exc)
            raise GitHubAPIError(
                0, sanitize_report_text(str(reason)) or "네트워크 연결 실패"
            ) from exc

    @staticmethod
    def _build_issue_body(report: str, fingerprint: str) -> str:
        marker = f"<!-- {FINGERPRINT_MARKER}: {fingerprint} -->"
        return f"{marker}\n{_clip(report, 60_000)}"

    @staticmethod
    def _build_comment_body(report: str, fingerprint: str) -> str:
        return "\n".join(
            [
                "## 오류 재발생",
                "",
                f"- Fingerprint: `{fingerprint}`",
                "",
                _clip(report, 56_000),
            ]
        )

    @staticmethod
    def _format_issue_title(title: str, context: dict[str, object]) -> str:
        category = sanitize_report_text(context.get("category") or "Milkrun")
        category = re.sub(r"\s+", " ", category).strip()[:30] or "Milkrun"
        clean_title = re.sub(r"\s+", " ", sanitize_report_text(title)).strip()
        return f"[UnHelper][{category}] {clean_title or '오류 신고'}"[:120]

    @staticmethod
    def _fingerprint(
        title: str, error_msg: str, context: dict[str, object]
    ) -> str:
        normalized = _normalize_error_for_fingerprint(error_msg)
        raw = "\n".join(
            [
                "UnHelper",
                str(context.get("category") or "Milkrun"),
                sanitize_report_text(title),
                normalized,
            ]
        )
        return hashlib.sha256(raw.encode("utf-8", errors="replace")).hexdigest()[:16]


def load_github_issue_token() -> str:
    for env_name in TOKEN_ENV_VARS:
        value = os.environ.get(env_name, "").strip()
        if value:
            return _decode_token_value(value)

    for token_path in _candidate_token_paths():
        if not token_path.is_file():
            continue
        try:
            value = token_path.read_text(encoding="utf-8").strip()
        except (OSError, UnicodeDecodeError):
            continue
        if value:
            try:
                return _decode_token_value(value)
            except (ValueError, UnicodeError):
                continue
    return ""


def encode_token_value(token: str) -> str:
    token_bytes = token.strip().encode("utf-8")
    masked = bytes(
        byte ^ _OBFUSCATION_KEY[index % len(_OBFUSCATION_KEY)]
        for index, byte in enumerate(token_bytes)
    )
    return "v1:" + base64.urlsafe_b64encode(masked).decode("ascii")


def _decode_token_value(value: str) -> str:
    value = value.strip()
    if value.startswith("v1:"):
        try:
            masked = base64.urlsafe_b64decode(value[3:].encode("ascii"))
        except (ValueError, UnicodeError) as exc:
            raise ValueError("잘못된 오류 신고 토큰 파일입니다.") from exc
        token_bytes = bytes(
            byte ^ _OBFUSCATION_KEY[index % len(_OBFUSCATION_KEY)]
            for index, byte in enumerate(masked)
        )
        return token_bytes.decode("utf-8").strip()
    if value.startswith("plain:"):
        return value[6:].strip()
    return value


def _candidate_token_paths() -> list[Path]:
    executable_dir = Path(sys.executable).resolve().parent
    candidates = (
        project_root() / TOKEN_FILE_NAME,
        bundled_root() / TOKEN_FILE_NAME,
        executable_dir / TOKEN_FILE_NAME,
        executable_dir / "_internal" / TOKEN_FILE_NAME,
        Path.cwd() / TOKEN_FILE_NAME,
    )
    unique: list[Path] = []
    seen: set[str] = set()
    for path in candidates:
        key = str(path.resolve(strict=False)).casefold()
        if key not in seen:
            unique.append(path)
            seen.add(key)
    return unique


def _normalize_error_for_fingerprint(error_msg: str) -> str:
    lines: list[str] = []
    for line in sanitize_report_text(error_msg).splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        stripped = re.sub(r'File ".*?", line \d+', 'File "FILE", line N', stripped)
        stripped = re.sub(r"0x[0-9a-fA-F]+", "0xADDR", stripped)
        stripped = re.sub(
            r"\b\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}(?:[.,]\d+)?(?:[+-]\d{2}:?\d{2})?\b",
            "DATETIME",
            stripped,
        )
        lines.append(stripped)
        if sum(len(item) for item in lines) > 5_000:
            break
    return "\n".join(lines)


def _extract_github_error_message(detail: str) -> str:
    try:
        data = json.loads(detail)
        message = data.get("message")
        errors = data.get("errors")
        if errors:
            return sanitize_report_text(f"{message}: {errors}")
        if message:
            return sanitize_report_text(str(message))
    except (AttributeError, json.JSONDecodeError):
        pass
    return sanitize_report_text(detail[:1_000])


def _clip(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + "\n...[길이 제한으로 일부 생략됨]"
