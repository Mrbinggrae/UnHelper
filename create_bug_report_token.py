"""Create the local, gitignored token file used by the bug reporter."""

import re
from getpass import getpass
from pathlib import Path

from Modules.Common.GitHubIssueReporter import encode_token_value


FINE_GRAINED_TOKEN_PATTERN = re.compile(r"github_pat_[A-Za-z0-9_]+")


def main() -> int:
    token = getpass("GitHub Issue token: ").strip()
    if not token:
        print("토큰이 입력되지 않았습니다.")
        return 1
    if FINE_GRAINED_TOKEN_PATTERN.fullmatch(token) is None:
        print("전체 fine-grained PAT를 입력해 주세요. 토큰은 'github_pat_'로 시작해야 합니다.")
        return 2

    target = Path(__file__).resolve().parent / "bug_report_token.dat"
    target.write_text(encode_token_value(token), encoding="utf-8")
    print(f"토큰 파일 생성 완료: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
