from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any
from urllib import error, request


PROJECT_ROOT = Path(__file__).resolve().parent
SHARED_DRIVER = Path(
    os.environ.get(
        "UNHELPER_SHARED_DRIVER",
        r"C:\Users\mrbin\Python\deadline\마감\chromedriver.exe",
    )
)
GITHUB_API = "https://api.github.com/repos/Mrbinggrae/UnHelper/releases"
APP_NAME = "UnHelper"
APP_PREFIX = "UnHelper"
DIST_DIR = PROJECT_ROOT / "dist" / APP_NAME


def version_key(value: str) -> tuple[int, ...]:
    numbers = re.findall(r"\d+", str(value))
    return tuple(int(number) for number in numbers[:4]) or (0,)


def read_version() -> str:
    history_path = PROJECT_ROOT / "UPDATE_HISTORY.txt"
    with history_path.open("r", encoding="utf-8") as handle:
        for line in handle:
            match = re.search(r"\[v(\d+(?:\.\d+)+)\]", line)
            if match:
                return match.group(1)
    raise RuntimeError("UPDATE_HISTORY.txt에서 버전을 읽을 수 없습니다.")


def project_python() -> Path:
    venv_python = PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"
    return venv_python if venv_python.is_file() else Path(sys.executable)


def run(command: list[str], label: str) -> None:
    print()
    print("=" * 60)
    print(label)
    print("=" * 60)
    result = subprocess.run(command, cwd=PROJECT_ROOT, check=False)
    if result.returncode:
        raise RuntimeError(f"{label} failed (exit code {result.returncode})")
    print(f"[OK] {label}")


def chromedriver_version(path: Path) -> str:
    if not path.is_file():
        raise FileNotFoundError(f"ChromeDriver를 찾을 수 없습니다: {path}")

    result = subprocess.run(
        [str(path), "--version"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
        check=False,
    )
    output = f"{result.stdout}\n{result.stderr}"
    match = re.search(r"ChromeDriver\s+(\d+(?:\.\d+)+)", output)
    if result.returncode or not match:
        raise RuntimeError(f"ChromeDriver 버전을 확인하지 못했습니다: {path}")
    return match.group(1)


def _file_product_version(path: Path) -> str | None:
    """Read a Windows executable's product version without launching Chrome."""
    escaped_path = str(path).replace("'", "''")
    command = (
        f"(Get-Item -LiteralPath '{escaped_path}').VersionInfo.ProductVersion"
    )
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=15,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    value = result.stdout.strip()
    return value if result.returncode == 0 and re.match(r"^\d+(?:\.\d+)+$", value) else None


def _registry_chrome_versions() -> list[str]:
    try:
        import winreg
    except ImportError:
        return []

    versions: list[str] = []
    key_names = (
        r"Software\Google\Chrome\BLBeacon",
        r"Software\WOW6432Node\Google\Chrome\BLBeacon",
    )
    hives = (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE)
    views = (0, winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY)
    for hive in hives:
        for key_name in key_names:
            for view in views:
                try:
                    with winreg.OpenKey(hive, key_name, 0, winreg.KEY_READ | view) as key:
                        value = str(winreg.QueryValueEx(key, "version")[0]).strip()
                    if re.match(r"^\d+(?:\.\d+)+$", value):
                        versions.append(value)
                except OSError:
                    continue
    return versions


def installed_chrome_version() -> str:
    override = os.environ.get("UNHELPER_CHROME_VERSION", "").strip()
    if override:
        if not re.match(r"^\d+(?:\.\d+)+$", override):
            raise RuntimeError("UNHELPER_CHROME_VERSION 형식이 올바르지 않습니다.")
        return override

    candidates = _registry_chrome_versions()
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "")
    chrome_paths = [
        Path(local_app_data) / "Google/Chrome/Application/chrome.exe",
        Path(program_files) / "Google/Chrome/Application/chrome.exe",
        Path(program_files_x86) / "Google/Chrome/Application/chrome.exe",
    ]
    for chrome_path in chrome_paths:
        if chrome_path.is_file():
            value = _file_product_version(chrome_path)
            if value:
                candidates.append(value)

    if not candidates:
        raise RuntimeError(
            "설치된 Chrome 버전을 확인하지 못했습니다. Chrome 설치 상태를 확인하거나 "
            "UNHELPER_CHROME_VERSION을 지정해 주세요."
        )
    # A stale per-user BLBeacon entry can coexist with a newer system install.
    return max(candidates, key=version_key)


def copy_and_verify_chromedriver() -> str:
    driver_version = chromedriver_version(SHARED_DRIVER)
    chrome_version = installed_chrome_version()
    if version_key(driver_version)[0] != version_key(chrome_version)[0]:
        raise RuntimeError(
            "공용 ChromeDriver와 설치된 Chrome의 메이저 버전이 다릅니다: "
            f"ChromeDriver {driver_version}, Chrome {chrome_version}"
        )

    target = PROJECT_ROOT / "chromedriver.exe"
    shutil.copy2(SHARED_DRIVER, target)
    copied_version = chromedriver_version(target)
    if copied_version != driver_version:
        raise RuntimeError(
            f"ChromeDriver 복사 검증 실패: source {driver_version}, copied {copied_version}"
        )
    print(
        f"[OK] shared ChromeDriver {copied_version} / Chrome {chrome_version} "
        f"(major {version_key(chrome_version)[0]})"
    )
    return copied_version


def _github_headers() -> dict[str, str]:
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "UnHelper-Release-Builder",
        "Cache-Control": "no-cache",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


def fetch_previous_manifest(current_version: str) -> Path | None:
    """Download the newest applicable release manifest for delta generation."""
    output_dir = PROJECT_ROOT / "release"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{APP_PREFIX}_prev_manifest.json"
    manifest_name = f"{APP_PREFIX}_manifest.json"

    try:
        api_request = request.Request(GITHUB_API, headers=_github_headers())
        with request.urlopen(api_request, timeout=20) as response:
            releases = json.loads(response.read().decode("utf-8"))
        if not isinstance(releases, list):
            return None

        current_key = version_key(current_version)
        candidates: list[tuple[tuple[int, ...], str, str]] = []
        for release in releases:
            if not isinstance(release, dict) or release.get("draft"):
                continue
            tag_version = str(release.get("tag_name", "")).lstrip("vV")
            if not tag_version or version_key(tag_version) >= current_key:
                continue
            for asset in release.get("assets", []):
                if not isinstance(asset, dict):
                    continue
                download_url = asset.get("browser_download_url")
                if asset.get("name") == manifest_name and download_url:
                    candidates.append(
                        (
                            version_key(tag_version),
                            str(release.get("published_at", "")),
                            str(download_url),
                        )
                    )
                    break

        if not candidates:
            return None
        candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)

        manifest_request = request.Request(candidates[0][2], headers=_github_headers())
        with request.urlopen(manifest_request, timeout=20) as response:
            payload = response.read()
        manifest = json.loads(payload.decode("utf-8"))
        if not isinstance(manifest, dict) or not isinstance(manifest.get("files"), list):
            raise ValueError("이전 릴리즈 매니페스트 형식이 올바르지 않습니다.")
        if manifest.get("app_name") not in (None, APP_NAME):
            raise ValueError("이전 릴리즈 매니페스트의 앱 이름이 다릅니다.")
        manifest_version = str(manifest.get("version", ""))
        if not manifest_version or version_key(manifest_version) >= current_key:
            raise ValueError("이전 릴리즈 매니페스트 버전이 현재 버전보다 낮지 않습니다.")

        output_path.write_bytes(payload)
        print(
            f"[OK] previous manifest: v{manifest.get('version', 'unknown')} "
            f"-> {output_path.relative_to(PROJECT_ROOT)}"
        )
        return output_path
    except (error.URLError, TimeoutError, OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"[WARN] previous manifest download skipped: {exc}")
        return None


def find_previous_manifest(current_version: str) -> Path | None:
    downloaded = fetch_previous_manifest(current_version)
    if downloaded:
        return downloaded

    fallback = PROJECT_ROOT / "release" / f"{APP_PREFIX}_manifest.json"
    if fallback.is_file():
        try:
            manifest = json.loads(fallback.read_text(encoding="utf-8"))
            fallback_version = str(manifest.get("version", ""))
            if (
                isinstance(manifest, dict)
                and isinstance(manifest.get("files"), list)
                and fallback_version
                and version_key(fallback_version) < version_key(current_version)
            ):
                print(
                    f"[WARN] using local v{fallback_version} manifest as delta base: "
                    f"{fallback.relative_to(PROJECT_ROOT)}"
                )
                return fallback
            print("[WARN] local manifest is not older than the build; delta omitted")
        except (OSError, AttributeError, json.JSONDecodeError):
            print("[WARN] local manifest is invalid; delta omitted")
    return None


def find_makensis() -> str:
    candidates = (
        "makensis",
        r"C:\Program Files (x86)\NSIS\makensis.exe",
        r"C:\Program Files\NSIS\makensis.exe",
    )
    for candidate in candidates:
        try:
            result = subprocess.run(
                [candidate, "/VERSION"],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
        except (FileNotFoundError, subprocess.SubprocessError):
            continue
        if result.returncode == 0:
            return candidate
    raise RuntimeError("NSIS makensis.exe를 찾을 수 없습니다.")


def build_installer(version: str) -> None:
    makensis = find_makensis()
    run(
        [makensis, f"/DAPP_VERSION={version}", "UnHelper.nsi"],
        f"NSIS installer v{version}",
    )


def build_unhelper(version: str, with_installer: bool) -> None:
    copy_and_verify_chromedriver()
    previous_manifest = find_previous_manifest(version)
    python = project_python()

    run(
        [str(python), "-m", "PyInstaller", "UnHelper.spec", "--clean", "--noconfirm"],
        "PyInstaller UnHelper (onedir)",
    )

    manifest_command = [
        str(python),
        "build_manifest.py",
        "--app",
        APP_NAME,
        "--dist-dir",
        str(DIST_DIR.relative_to(PROJECT_ROOT)),
        "--version",
        version,
    ]
    if previous_manifest:
        manifest_command.extend(["--prev-manifest", str(previous_manifest)])
    run(manifest_command, "Manifest, full patch, and delta patch")

    # Initial installer builds keep a local manifest so same-version repacks can
    # be detected without hashing every installed file.
    generated_manifest = PROJECT_ROOT / "release" / f"{APP_PREFIX}_manifest.json"
    if generated_manifest.is_file():
        shutil.copy2(generated_manifest, DIST_DIR / generated_manifest.name)

    if with_installer:
        build_installer(version)


def print_release_assets(with_installer: bool) -> None:
    print()
    print("GitHub Release upload assets:")
    names = [
        f"{APP_PREFIX}_manifest.json",
        f"{APP_PREFIX}_patch.zip",
        f"{APP_PREFIX}_delta_patch.zip",
    ]
    if with_installer:
        names.append(f"{APP_PREFIX}_Setup.exe")
    for name in names:
        path = PROJECT_ROOT / "release" / name
        if path.is_file():
            print(f"  - {path.relative_to(PROJECT_ROOT)} ({path.stat().st_size:,} bytes)")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build UnHelper release assets",
        epilog=(
            "examples:\n"
            "  python build_release.py un\n"
            "  python build_release.py un --installer"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("target", choices=["un"], help="build target")
    parser.add_argument("--installer", action="store_true", help="build NSIS installer too")
    args = parser.parse_args()

    try:
        version = read_version()
        print(f"Version: {version}")
        build_unhelper(version, args.installer)
        print_release_assets(args.installer)
    except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"[FAIL] {exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
