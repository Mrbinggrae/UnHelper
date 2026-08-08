from __future__ import annotations

import platform
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode

from Modules.Common.version import CURRENT_VERSION


GITHUB_ISSUES_URL = "https://github.com/Mrbinggrae/UnHelper/issues/new"
_REPORT_LIMIT = 50_000


@dataclass(frozen=True)
class FailureDetails:
    """A user-facing summary paired with the original diagnostic traceback."""

    summary: str
    detail: str

    @classmethod
    def from_exception(cls, exc: BaseException) -> "FailureDetails":
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).strip()
        return cls(summary=str(exc).strip() or type(exc).__name__, detail=detail)

    @classmethod
    def coerce(cls, value: object) -> "FailureDetails":
        if isinstance(value, cls):
            return value
        text = str(value).strip() or "알 수 없는 오류가 발생했습니다."
        return cls(summary=text, detail=text)


def sanitize_report_text(value: object) -> str:
    """Remove common credentials and the local Windows user path from a report."""

    text = str(value)
    home = str(Path.home())
    if home:
        text = re.sub(re.escape(home), "%USERPROFILE%", text, flags=re.IGNORECASE)

    key_names = (
        r"wms[_-]?(?:password|passwd|pwd|pw|id)|"
        r"user[_-]?(?:password|passwd|pwd|pw|id)|username|"
        r"password|passwd|pwd|pw|token|secret|api[_-]?key|session|cookie"
    )

    # Treat an Authorization header as sensitive through the end of its line.
    # Replacing the whole line also protects malformed headers and values that
    # contain spaces, commas, quotes, or an unexpected authentication scheme.
    text = re.sub(
        r"(?im)^[^\r\n]*\bAuthorization\s*:[^\r\n]*$",
        "Authorization: ***",
        text,
    )

    # JSON and Python diagnostics commonly quote mapping keys, for example
    # ``{"password": "value"}`` or ``{'wms_id': 'worker'}``.  Mask these
    # separately while preserving the surrounding syntax so the next key can
    # still be recognized and redacted.
    quoted_mapping_pattern = re.compile(
        rf"(?P<prefix>(?P<key_quote>['\"])(?:{key_names})(?P=key_quote)\s*:\s*)"
        rf"(?:"
        rf"(?P<value_quote>['\"])(?:\\.|(?!(?P=value_quote)).)*(?P=value_quote)"
        rf"|(?P<bare_value>[^,}}\]\r\n]*)"
        rf")",
        flags=re.IGNORECASE,
    )

    def mask_quoted_mapping(match: re.Match[str]) -> str:
        quote = match.group("value_quote") or ""
        return f"{match.group('prefix')}{quote}***{quote}"

    text = quoted_mapping_pattern.sub(mask_quoted_mapping, text)

    # A credential value may legitimately contain whitespace and punctuation.
    # Stop only at the next recognized credential assignment or at the line
    # boundary; masking just the first whitespace-delimited token can leak the
    # remainder of a password into a GitHub issue.
    assignment_pattern = re.compile(
        rf"(?P<prefix>\b(?:{key_names})\b\s*[:=]\s*)"
        rf"(?P<value>.*?)"
        rf"(?=(?:\s+\b(?:{key_names})\b\s*[:=])|$)",
        flags=re.IGNORECASE | re.MULTILINE,
    )
    text = assignment_pattern.sub(lambda match: f"{match.group('prefix')}***", text)

    replacements = (
        (r"\b(ghp_|gho_|ghu_|ghs_|github_pat_)[A-Za-z0-9_\-]+", r"\1***"),
        (r"([?&](?:access_)?token=)[^&#\s]+", r"\1***"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text


def build_error_report(
    title: str,
    failure: FailureDetails | object,
    *,
    context: dict[str, object] | None = None,
) -> str:
    details = FailureDetails.coerce(failure)
    context = context or {}
    log_tail = str(context.get("log_tail") or "").strip()
    category = str(context.get("category") or "Milkrun")
    lines = [
        "## 오류 신고",
        "",
        f"- App: UnHelper",
        f"- Version: v{CURRENT_VERSION}",
        f"- Category: {sanitize_report_text(category)}",
        f"- Time: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"- OS: {sanitize_report_text(platform.platform())}",
        f"- Python: {platform.python_version()}",
        f"- Frozen: {bool(getattr(sys, 'frozen', False))}",
        "",
        "## 오류 제목",
        "",
        sanitize_report_text(title),
        "",
        "## 오류 내용",
        "",
        "```text",
        sanitize_report_text(details.detail or details.summary),
        "```",
    ]
    if log_tail:
        lines.extend(
            [
                "",
                "## 최근 로그",
                "",
                "```text",
                sanitize_report_text(log_tail[-12_000:]),
                "```",
            ]
        )
    report = "\n".join(lines).strip()
    if len(report) > _REPORT_LIMIT:
        report = report[:_REPORT_LIMIT] + "\n...[길이 제한으로 일부 생략됨]"
    return report


def build_github_issue_url(title: str, report: str) -> str:
    """Build a short issue URL; the full report is copied to the clipboard.

    GitHub rejects very long ``issues/new?body=...`` URLs with its
    ``Whoa there!`` page.  Tracebacks and Korean text expand considerably
    after percent-encoding, so even a character-limited body is not a safe
    browser URL.  Keep ``report`` in the public signature for callers that
    build the report and URL together, but never place it in the URL.
    """

    issue_title = re.sub(r"\s+", " ", sanitize_report_text(title)).strip()
    issue_title = f"[UnHelper] {issue_title or '오류 신고'}"[:120]
    return f"{GITHUB_ISSUES_URL}?{urlencode({'title': issue_title})}"
