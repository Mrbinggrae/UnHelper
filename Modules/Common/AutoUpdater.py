from __future__ import annotations

import hashlib
import json
import logging
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable
from urllib import error, request

from .version import CURRENT_VERSION, GITHUB_OWNER, GITHUB_REPO


logger = logging.getLogger(__name__)
ProgressCallback = Callable[[int], None]


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


@dataclass(slots=True)
class UpdateInfo:
    version: str
    changelog: str
    download_url: str
    patch_size: int = 0
    manifest: dict[str, Any] = field(default_factory=dict)
    is_same_version_update: bool = False
    expected_hash: str | None = None
    patch_mode: str = "full"
    is_release_restore: bool = False


class AutoUpdater:
    """TC Helper-compatible GitHub Releases updater for UnHelper."""

    RELEASES_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases"

    def __init__(self, app_name: str = "UnHelper", use_prerelease: bool = False):
        self.app_name = app_name
        self.app_prefix = app_name.replace(" ", "_")
        self.use_prerelease = use_prerelease
        self.app_dir = get_app_dir()
        self.temp_dir = Path(tempfile.gettempdir()) / f"{self.app_prefix}_update"
        self._current_version = CURRENT_VERSION

    def check_for_update(self) -> tuple[bool, UpdateInfo | None]:
        """Find a newer release or a changed patch with the current version."""
        try:
            releases = self._fetch_json(self.RELEASES_API, self._api_headers())
            if not isinstance(releases, list):
                return False, None

            candidates = self._release_candidates(
                releases,
                include_prerelease=self.use_prerelease,
                include_delta=True,
            )
            for candidate in candidates:
                release_version = candidate["version"]
                is_newer = self._is_newer(release_version, self._current_version)
                if not is_newer and release_version != self._current_version:
                    continue

                manifest = self._fetch_manifest(candidate.get("manifest_url"))
                try:
                    self._validate_release_manifest(candidate, manifest)
                except ValueError as exc:
                    logger.warning("[%s] invalid release %s: %s", self.app_name, release_version, exc)
                    continue
                is_same_version_update = (
                    release_version == self._current_version
                    and self._same_version_update_available(manifest)
                )
                if not is_newer and not is_same_version_update:
                    continue

                selected_patch = self._select_patch(candidate, manifest)
                info = UpdateInfo(
                    version=release_version,
                    changelog=str(candidate["release"].get("body") or "변경사항 없음"),
                    download_url=str(selected_patch["url"]),
                    patch_size=int(selected_patch["size"] or 0),
                    manifest=manifest,
                    is_same_version_update=is_same_version_update,
                    expected_hash=selected_patch["sha256"],
                    patch_mode=str(selected_patch["mode"]),
                )
                logger.info(
                    "[%s] update found: %s -> %s (%s%s)",
                    self.app_name,
                    self._current_version,
                    release_version,
                    info.patch_mode,
                    ", same-version" if is_same_version_update else "",
                )
                return True, info
            return False, None
        except error.URLError as exc:
            logger.warning("Update check failed (network): %s", exc)
            raise RuntimeError(f"네트워크 오류로 업데이트를 확인하지 못했습니다: {exc}") from exc
        except Exception as exc:
            logger.warning("Update check failed: %s", exc)
            if isinstance(exc, RuntimeError):
                raise
            raise RuntimeError(f"업데이트 확인 중 오류가 발생했습니다: {exc}") from exc

    def check_for_release_restore(self) -> tuple[bool, UpdateInfo | None, str]:
        """Return the newest stable full patch when leaving the beta channel."""
        if not getattr(sys, "frozen", False):
            return (
                False,
                None,
                "개발 실행 환경에서는 정식 릴리즈 복구를 적용하지 않습니다. 배포된 EXE에서 사용해 주세요.",
            )

        try:
            releases = self._fetch_json(self.RELEASES_API, self._api_headers())
            if not isinstance(releases, list):
                return False, None, "정식 릴리즈 정보를 가져오지 못했습니다."

            candidates = self._release_candidates(
                releases,
                include_prerelease=False,
                include_delta=False,
            )
            if not candidates:
                return False, None, "적용 가능한 정식 릴리즈 패치를 찾지 못했습니다."

            candidate = None
            manifest: dict[str, Any] = {}
            for possible in candidates:
                possible_manifest = self._fetch_manifest(possible.get("manifest_url"))
                try:
                    self._validate_release_manifest(possible, possible_manifest)
                except ValueError as exc:
                    logger.warning(
                        "[%s] invalid stable release %s: %s",
                        self.app_name,
                        possible["version"],
                        exc,
                    )
                    continue
                candidate = possible
                manifest = possible_manifest
                break
            if candidate is None:
                return False, None, "검증 가능한 정식 릴리즈 패치를 찾지 못했습니다."
            release_version = candidate["version"]
            if self._release_restore_not_needed(release_version, manifest):
                return False, None, f"이미 최신 정식 릴리즈(v{release_version}) 상태입니다."

            full_patch = manifest.get("full_patch", {}) if isinstance(manifest, dict) else {}
            full_hash = full_patch.get("sha256") if isinstance(full_patch, dict) else None
            info = UpdateInfo(
                version=release_version,
                changelog=str(candidate["release"].get("body") or "변경사항 없음"),
                download_url=str(candidate["patch_url"]),
                patch_size=int(candidate["patch_size"] or 0),
                manifest=manifest,
                expected_hash=full_hash or manifest.get("zip_hash"),
                patch_mode="full",
                is_release_restore=True,
            )
            return True, info, ""
        except error.URLError as exc:
            logger.warning("Release restore check failed (network): %s", exc)
            return False, None, f"네트워크 오류로 정식 릴리즈를 확인하지 못했습니다: {exc}"
        except Exception as exc:
            logger.warning("Release restore check failed: %s", exc)
            return False, None, f"정식 릴리즈 확인 중 오류가 발생했습니다: {exc}"

    def download_patch(
        self,
        info: UpdateInfo,
        progress_callback: ProgressCallback | None = None,
    ) -> Path:
        """Download and verify the selected full or delta patch."""
        try:
            if not info.expected_hash:
                raise ValueError("패치 SHA-256 정보가 없어 다운로드를 거부했습니다.")
            if self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
            self.temp_dir.mkdir(parents=True, exist_ok=True)

            zip_path = self.temp_dir / "patch.zip"
            download_request = request.Request(
                info.download_url,
                headers={"User-Agent": f"{self.app_name}-Updater"},
            )
            with request.urlopen(download_request, timeout=60) as response:
                total_size = int(response.headers.get("Content-Length", 0) or 0)
                downloaded = 0
                with zip_path.open("wb") as handle:
                    while True:
                        chunk = response.read(256 * 1024)
                        if not chunk:
                            break
                        handle.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size > 0:
                            progress_callback(min(100, int(downloaded / total_size * 100)))

            expected_hash = info.expected_hash or info.manifest.get("zip_hash")
            if expected_hash:
                actual_hash = self._sha256(zip_path)
                if actual_hash.lower() != str(expected_hash).lower():
                    zip_path.unlink(missing_ok=True)
                    raise ValueError(
                        f"해시 불일치: 기대 {str(expected_hash)[:8]}, 실제 {actual_hash[:8]}"
                    )
            return zip_path
        except Exception as exc:
            logger.error("Patch download failed: %s", exc)
            raise RuntimeError(f"다운로드 실패: {exc}") from exc

    def apply_update(
        self,
        zip_path: Path,
        new_version: str | dict[str, Any] | None = None,
        manifest: dict[str, Any] | None = None,
    ) -> bool:
        """Apply a patch after exit and restart the executable.

        ``new_version`` is optional for compatibility with simpler callers. A
        manifest may also be passed as the second positional argument.
        """
        try:
            if not getattr(sys, "frozen", False):
                raise RuntimeError("개발 실행 환경에서는 자동 패치를 적용하지 않습니다.")
            if isinstance(new_version, dict) and manifest is None:
                manifest = new_version
                new_version = str(manifest.get("version", ""))

            extract_dir = self.temp_dir / "extracted"
            if extract_dir.exists():
                shutil.rmtree(extract_dir)
            extract_dir.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(zip_path, "r") as archive:
                self._validate_archive_members(archive)
                archive.extractall(extract_dir)

            manifest_copy_command = ""
            obsolete_delete_commands = ""
            if manifest:
                manifest_target = self._local_manifest_path()
                manifest_temp = self.temp_dir / manifest_target.name
                manifest_temp.write_text(
                    json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
                    encoding="utf-8",
                )
                manifest_copy_command = (
                    f'copy /Y "{manifest_temp}" "{manifest_target}" >nul\r\n'
                )
                obsolete_delete_commands = self._obsolete_delete_commands(manifest)

            executable = Path(sys.executable).resolve()
            batch_path = self.temp_dir / "apply_update.bat"
            batch_path.write_text(
                "@echo off\r\n"
                "ping 127.0.0.1 -n 2 >nul\r\n"
                ":wait_exit\r\n"
                f'tasklist /FI "IMAGENAME eq {executable.name}" | find /I "{executable.name}" >nul\r\n'
                "if not errorlevel 1 (\r\n"
                "    ping 127.0.0.1 -n 2 >nul\r\n"
                "    goto wait_exit\r\n"
                ")\r\n"
                f"{obsolete_delete_commands}"
                f'xcopy /E /Y /I "{extract_dir}\\*" "{self.app_dir}\\" >nul\r\n'
                "if errorlevel 1 exit /b 1\r\n"
                f"{manifest_copy_command}"
                f'start "" "{executable}"\r\n'
                "ping 127.0.0.1 -n 3 >nul\r\n"
                f'rmdir /S /Q "{self.temp_dir}" 2>nul\r\n'
                "exit /b 0\r\n",
                encoding="cp949",
                errors="replace",
            )

            startupinfo = None
            if hasattr(subprocess, "STARTUPINFO"):
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0
            subprocess.Popen(
                ["cmd", "/c", str(batch_path)],
                creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
                startupinfo=startupinfo,
                close_fds=True,
            )
            logger.info("Update batch started for version %s", new_version or "unknown")
            return True
        except Exception as exc:
            logger.warning("Update apply failed: %s", exc)
            return False

    def _release_candidates(
        self,
        releases: list[Any],
        *,
        include_prerelease: bool,
        include_delta: bool,
    ) -> list[dict[str, Any]]:
        candidates: list[dict[str, Any]] = []
        for release in releases:
            if not isinstance(release, dict) or release.get("draft"):
                continue
            if release.get("prerelease") and not include_prerelease:
                continue

            version = str(release.get("tag_name", "")).lstrip("vV")
            if not re.match(r"^\d+(?:\.\d+)+", version):
                continue

            manifest_url = None
            patch_url = None
            patch_size = 0
            delta_patch_url = None
            delta_patch_size = 0
            for asset in release.get("assets", []):
                if not isinstance(asset, dict):
                    continue
                name = asset.get("name", "")
                if name == f"{self.app_prefix}_manifest.json":
                    manifest_url = asset.get("browser_download_url")
                elif name == f"{self.app_prefix}_patch.zip":
                    patch_url = asset.get("browser_download_url")
                    patch_size = int(asset.get("size", 0) or 0)
                elif include_delta and name == f"{self.app_prefix}_delta_patch.zip":
                    delta_patch_url = asset.get("browser_download_url")
                    delta_patch_size = int(asset.get("size", 0) or 0)
            if not patch_url or not manifest_url:
                continue

            candidates.append(
                {
                    "release": release,
                    "version": version,
                    "published_at": str(release.get("published_at", "")),
                    "manifest_url": manifest_url,
                    "patch_url": patch_url,
                    "patch_size": patch_size,
                    "delta_patch_url": delta_patch_url,
                    "delta_patch_size": delta_patch_size,
                }
            )
        candidates.sort(
            key=lambda item: (self._version_key(item["version"]), item["published_at"]),
            reverse=True,
        )
        return candidates

    def _fetch_manifest(self, manifest_url: str | None) -> dict[str, Any]:
        if not manifest_url:
            return {}
        value = self._fetch_json(
            manifest_url,
            {"User-Agent": f"{self.app_name}-Updater", "Cache-Control": "no-cache"},
        )
        return value if isinstance(value, dict) else {}

    def _validate_release_manifest(
        self,
        candidate: dict[str, Any],
        manifest: dict[str, Any],
    ) -> None:
        if not isinstance(manifest, dict) or not isinstance(manifest.get("files"), list):
            raise ValueError("manifest files가 없습니다.")
        if manifest.get("app_name") != self.app_name:
            raise ValueError("manifest 앱 이름이 일치하지 않습니다.")
        if str(manifest.get("version", "")) != str(candidate.get("version", "")):
            raise ValueError("manifest 버전과 릴리즈 태그가 일치하지 않습니다.")

        full_patch = manifest.get("full_patch")
        if not isinstance(full_patch, dict):
            raise ValueError("전체 패치 메타데이터가 없습니다.")
        if full_patch.get("file") != f"{self.app_prefix}_patch.zip":
            raise ValueError("전체 패치 파일명이 올바르지 않습니다.")
        if not re.fullmatch(r"[0-9a-fA-F]{64}", str(full_patch.get("sha256", ""))):
            raise ValueError("전체 패치 SHA-256이 올바르지 않습니다.")

        delta_patch = manifest.get("delta_patch")
        if candidate.get("delta_patch_url") and isinstance(delta_patch, dict):
            if delta_patch.get("file") != f"{self.app_prefix}_delta_patch.zip":
                raise ValueError("델타 패치 파일명이 올바르지 않습니다.")
            if not re.fullmatch(r"[0-9a-fA-F]{64}", str(delta_patch.get("sha256", ""))):
                raise ValueError("델타 패치 SHA-256이 올바르지 않습니다.")

    def _select_patch(
        self,
        candidate: dict[str, Any],
        manifest: dict[str, Any],
    ) -> dict[str, Any]:
        full_patch = manifest.get("full_patch", {}) if isinstance(manifest, dict) else {}
        full_hash = full_patch.get("sha256") if isinstance(full_patch, dict) else None
        selected = {
            "url": candidate["patch_url"],
            "size": candidate["patch_size"],
            "sha256": full_hash or manifest.get("zip_hash"),
            "mode": "full",
        }

        delta_patch = manifest.get("delta_patch") if isinstance(manifest, dict) else None
        if (
            isinstance(delta_patch, dict)
            and candidate.get("delta_patch_url")
            and candidate.get("version") != self._current_version
            and self._can_use_delta_patch(delta_patch)
        ):
            selected = {
                "url": candidate["delta_patch_url"],
                "size": candidate.get("delta_patch_size", 0),
                "sha256": delta_patch.get("sha256"),
                "mode": "delta",
            }
        return selected

    def _can_use_delta_patch(self, delta_patch: dict[str, Any]) -> bool:
        base_version = str(delta_patch.get("base_version", "")).strip()
        return bool(base_version) and base_version == self._current_version

    def _local_manifest_path(self) -> Path:
        if getattr(sys, "frozen", False):
            return self.app_dir / f"{self.app_prefix}_manifest.json"
        return self.app_dir / "release" / f"{self.app_prefix}_manifest.json"

    def _obsolete_delete_commands(self, manifest: dict[str, Any]) -> str:
        """Delete installed distribution files absent from the target manifest.

        User settings and downloads live outside the install directory. The
        NSIS uninstaller and local manifest are preserved explicitly.
        """
        target_paths = {
            Path(str(item.get("path", ""))).as_posix().casefold()
            for item in manifest.get("files", [])
            if isinstance(item, dict) and item.get("path")
        }
        preserved = {
            "uninstall.exe",
            f"{self.app_prefix}_manifest.json".casefold(),
        }
        commands: list[str] = []
        app_root = self.app_dir.resolve()
        for local_path in app_root.rglob("*"):
            if not local_path.is_file():
                continue
            relative = local_path.relative_to(app_root).as_posix()
            folded = relative.casefold()
            if folded in target_paths or folded in preserved:
                continue
            commands.append(f'del /F /Q "{local_path}" >nul 2>&1\r\n')
        return "".join(commands)

    def _load_local_manifest(self) -> dict[str, Any]:
        path = self._local_manifest_path()
        if not path.is_file():
            return {}
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _same_version_update_available(self, manifest: dict[str, Any]) -> bool:
        if not isinstance(manifest, dict) or not manifest.get("files"):
            return False
        local_manifest = self._load_local_manifest()
        if local_manifest:
            return self._manifest_signature(local_manifest) != self._manifest_signature(manifest)
        if not getattr(sys, "frozen", False):
            return False
        return not self._matches_local_installation(manifest)

    def _release_restore_not_needed(
        self,
        release_version: str,
        manifest: dict[str, Any],
    ) -> bool:
        if release_version != self._current_version:
            return False
        if not isinstance(manifest, dict) or not manifest.get("files"):
            return True
        local_manifest = self._load_local_manifest()
        if local_manifest:
            return self._manifest_signature(local_manifest) == self._manifest_signature(manifest)
        return self._matches_local_installation(manifest)

    def _matches_local_installation(self, manifest: dict[str, Any]) -> bool:
        app_root = self.app_dir.resolve()
        for file_info in manifest.get("files", []):
            if not isinstance(file_info, dict):
                continue
            relative_path = file_info.get("path")
            expected_hash = file_info.get("sha256")
            expected_size = file_info.get("size")
            if not relative_path or not expected_hash:
                continue

            local_path = (app_root / Path(str(relative_path))).resolve()
            try:
                local_path.relative_to(app_root)
            except ValueError:
                return False
            if not local_path.is_file():
                return False
            try:
                if expected_size is not None and local_path.stat().st_size != int(expected_size):
                    return False
                if self._sha256(local_path).lower() != str(expected_hash).lower():
                    return False
            except (OSError, TypeError, ValueError):
                return False
        return True

    @staticmethod
    def _manifest_signature(manifest: dict[str, Any]) -> str:
        files = []
        for item in manifest.get("files", []):
            if isinstance(item, dict):
                files.append(
                    {
                        "path": str(item.get("path", "")),
                        "sha256": str(item.get("sha256", "")),
                        "size": int(item.get("size", 0) or 0),
                    }
                )
        files.sort(key=lambda item: item["path"])
        normalized = {"version": str(manifest.get("version", "")), "files": files}
        payload = json.dumps(normalized, sort_keys=True, ensure_ascii=False).encode("utf-8")
        return hashlib.sha256(payload).hexdigest()

    @staticmethod
    def _validate_archive_members(archive: zipfile.ZipFile) -> None:
        for member in archive.infolist():
            path = Path(member.filename.replace("\\", "/"))
            if path.is_absolute() or ".." in path.parts:
                raise ValueError(f"안전하지 않은 패치 경로입니다: {member.filename}")

    @staticmethod
    def _api_headers() -> dict[str, str]:
        return {
            "Accept": "application/vnd.github+json",
            "User-Agent": "UnHelper-Updater",
            "Cache-Control": "no-cache",
            "X-GitHub-Api-Version": "2022-11-28",
        }

    @staticmethod
    def _fetch_json(url: str, headers: dict[str, str] | None = None) -> Any:
        fetch_request = request.Request(url, headers=headers or {})
        with request.urlopen(fetch_request, timeout=15) as response:
            return json.loads(response.read().decode("utf-8"))

    @staticmethod
    def _sha256(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    @staticmethod
    def _version_key(value: str) -> tuple[int, ...]:
        match = re.match(r"^(\d+(?:\.\d+)+)", str(value))
        if not match:
            return (0,)
        return tuple(int(part) for part in match.group(1).split("."))

    @staticmethod
    def _is_newer(latest: str, current: str) -> bool:
        return AutoUpdater._version_key(latest) > AutoUpdater._version_key(current)

    @staticmethod
    def format_size(size_bytes: int) -> str:
        size = float(max(0, size_bytes))
        for unit in ("B", "KB", "MB", "GB"):
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"
