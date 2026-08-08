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
_ISSUE_BODY_LIMIT = 5_500


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

    key_names = r"password|passwd|pwd|token|secret|api[_-]?key|session|cookie"
    replacements = (
        (rf"\b({key_names})\b\s*[:=]\s*([^\s,;]+)", r"\1=***"),
        (r"(Authorization\s*:\s*(?:Bearer|Basic)\s+)[^\s]+", r"\1***"),
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
    issue_title = re.sub(r"\s+", " ", sanitize_report_text(title)).strip()
    issue_title = f"[UnHelper] {issue_title or '오류 신고'}"[:120]
    sanitized = sanitize_report_text(report)
    if len(sanitized) > _ISSUE_BODY_LIMIT:
        sanitized = (
            sanitized[:_ISSUE_BODY_LIMIT]
            + "\n\n---\n전체 오류 내용은 클립보드에 복사되어 있습니다. 이 아래에 붙여넣어 주세요."
        )
    return f"{GITHUB_ISSUES_URL}?{urlencode({'title': issue_title, 'body': sanitized})}"
