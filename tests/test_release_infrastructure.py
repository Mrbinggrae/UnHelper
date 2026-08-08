from __future__ import annotations

import hashlib
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from urllib import error

import build_release
from build_manifest import build_manifest, build_release_assets
from Modules.Common.AutoUpdater import AutoUpdater
from Modules.Common.GitHubIssueReporter import encode_token_value


def release_payload(
    version: str,
    *,
    prerelease: bool = False,
    include_delta: bool = True,
) -> dict:
    assets = [
        {
            "name": "UnHelper_manifest.json",
            "browser_download_url": f"https://example.test/{version}/manifest",
            "size": 100,
        },
        {
            "name": "UnHelper_patch.zip",
            "browser_download_url": f"https://example.test/{version}/full",
            "size": 1000,
        },
    ]
    if include_delta:
        assets.append(
            {
                "name": "UnHelper_delta_patch.zip",
                "browser_download_url": f"https://example.test/{version}/delta",
                "size": 100,
            }
        )
    return {
        "tag_name": f"v{version}",
        "draft": False,
        "prerelease": prerelease,
        "published_at": f"2026-08-{version[-1:]}T00:00:00Z",
        "body": f"release {version}",
        "assets": assets,
    }


def manifest_payload(version: str, base_version: str = "0.1.0") -> dict:
    file_hash = hashlib.sha256(f"file-{version}".encode()).hexdigest()
    full_hash = hashlib.sha256(f"full-{version}".encode()).hexdigest()
    delta_hash = hashlib.sha256(f"delta-{version}".encode()).hexdigest()
    return {
        "app_name": "UnHelper",
        "version": version,
        "files": [{"path": "UnHelper.exe", "sha256": file_hash, "size": 10}],
        "zip_hash": full_hash,
        "full_patch": {
            "file": "UnHelper_patch.zip",
            "sha256": full_hash,
            "size": 1000,
        },
        "delta_patch": {
            "file": "UnHelper_delta_patch.zip",
            "sha256": delta_hash,
            "size": 100,
            "base_version": base_version,
        },
    }


class ManifestBuilderTests(unittest.TestCase):
    def test_hybrid_build_has_full_patch_and_previous_version_delta(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            previous_dir = root / "previous"
            current_dir = root / "current"
            output_dir = root / "release"
            previous_dir.mkdir()
            current_dir.mkdir()
            (previous_dir / "same.txt").write_text("same", encoding="utf-8")
            (previous_dir / "changed.txt").write_text("before", encoding="utf-8")
            (previous_dir / "removed.txt").write_text("removed", encoding="utf-8")
            (current_dir / "same.txt").write_text("same", encoding="utf-8")
            (current_dir / "changed.txt").write_text("after", encoding="utf-8")
            (current_dir / "new.txt").write_text("new", encoding="utf-8")

            previous = build_manifest(previous_dir)
            previous.update({"app_name": "UnHelper", "version": "0.1.0"})
            manifest = build_release_assets(
                app_name="UnHelper",
                dist_dir=current_dir,
                version="0.2.0",
                output_dir=output_dir,
                previous_manifest=previous,
            )

            self.assertEqual(manifest["patch_mode"], "hybrid")
            self.assertEqual(manifest["delta_patch"]["base_version"], "0.1.0")
            self.assertEqual(manifest["deleted_files"], ["removed.txt"])
            self.assertTrue((output_dir / "UnHelper_patch.zip").is_file())
            with zipfile.ZipFile(output_dir / "UnHelper_delta_patch.zip") as archive:
                self.assertEqual(sorted(archive.namelist()), ["changed.txt", "new.txt"])

    def test_initial_build_removes_stale_delta(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dist_dir = root / "dist"
            output_dir = root / "release"
            dist_dir.mkdir()
            output_dir.mkdir()
            (dist_dir / "UnHelper.exe").write_bytes(b"exe")
            stale_delta = output_dir / "UnHelper_delta_patch.zip"
            stale_delta.write_bytes(b"stale")

            build_release_assets(
                app_name="UnHelper",
                dist_dir=dist_dir,
                version="0.1.0",
                output_dir=output_dir,
            )
            self.assertFalse(stale_delta.exists())


class PreviousManifestSelectionTests(unittest.TestCase):
    def test_local_delta_base_must_be_strictly_older(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            release_dir = root / "release"
            release_dir.mkdir()
            local_manifest = release_dir / "UnHelper_manifest.json"

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(build_release, "fetch_previous_manifest", return_value=None),
            ):
                local_manifest.write_text(
                    json.dumps({"version": "0.2.0", "files": []}),
                    encoding="utf-8",
                )
                self.assertIsNone(build_release.find_previous_manifest("0.2.0"))

                local_manifest.write_text(
                    json.dumps({"version": "0.1.0", "files": []}),
                    encoding="utf-8",
                )
                self.assertEqual(
                    build_release.find_previous_manifest("0.2.0"),
                    local_manifest,
                )


class ReleaseTokenValidationTests(unittest.TestCase):
    @staticmethod
    def _api_response(payload: dict):
        response = mock.MagicMock()
        response.__enter__.return_value = response
        response.read.return_value = json.dumps(payload).encode("utf-8")
        return response

    @staticmethod
    def _api_error(status: int) -> error.HTTPError:
        return error.HTTPError(
            build_release.GITHUB_ISSUES_API,
            status,
            "GitHub API error",
            hdrs=None,
            fp=io.BytesIO(b'{"message":"Validation Failed"}'),
        )

    @staticmethod
    def _write_valid_token(root: Path) -> Path:
        token_path = root / build_release.BUG_REPORT_TOKEN_NAME
        token_path.write_text(
            encode_token_value("github_pat_" + "A" * 40),
            encoding="utf-8",
        )
        return token_path

    def test_release_source_token_must_exist_and_be_nonempty(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            token_path = root / build_release.BUG_REPORT_TOKEN_NAME

            with mock.patch.object(build_release, "PROJECT_ROOT", root):
                with self.assertRaisesRegex(RuntimeError, "토큰 파일이 없습니다"):
                    build_release.require_bug_report_token_source()

                token_path.touch()
                with self.assertRaisesRegex(RuntimeError, "토큰 파일이 비어 있습니다"):
                    build_release.require_bug_report_token_source()

                token_path.write_bytes(b"opaque-token-data")
                self.assertEqual(
                    build_release.require_bug_report_token_source(),
                    token_path,
                )

    def test_valid_token_authenticates_and_accesses_intended_repository(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self._write_valid_token(root)
            responses = (
                self._api_response({"login": "token-owner"}),
                self._api_response({"full_name": "Mrbinggrae/UnHelper"}),
                self._api_error(422),
            )

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(
                    build_release.request,
                    "urlopen",
                    side_effect=responses,
                ) as urlopen,
            ):
                build_release.validate_bug_report_token_preflight()

            self.assertEqual(urlopen.call_count, 3)
            for call in urlopen.call_args_list:
                authorization = call.args[0].headers.get("Authorization", "")
                self.assertTrue(authorization.startswith("Bearer github_pat_"))
            issue_probe = urlopen.call_args_list[-1].args[0]
            self.assertEqual(issue_probe.get_method(), "POST")
            self.assertEqual(issue_probe.data, b"{}")

    def test_missing_issues_write_permission_fails_preflight(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self._write_valid_token(root)
            responses = (
                self._api_response({"login": "token-owner"}),
                self._api_response({"full_name": "Mrbinggrae/UnHelper"}),
                self._api_error(403),
            )

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(
                    build_release.request,
                    "urlopen",
                    side_effect=responses,
                ),
                self.assertRaisesRegex(RuntimeError, "Issues 쓰기 권한"),
            ):
                build_release.validate_bug_report_token_preflight()

    def test_malformed_and_non_fine_grained_tokens_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            token_path = root / build_release.BUG_REPORT_TOKEN_NAME
            cases = (
                "v1:not-valid-base64%%%",
                encode_token_value("ghp_not-a-fine-grained-token"),
            )

            for value in cases:
                with self.subTest(value_kind=value[:3]):
                    token_path.write_text(value, encoding="utf-8")
                    with (
                        mock.patch.object(build_release, "PROJECT_ROOT", root),
                        mock.patch.object(build_release.request, "urlopen") as urlopen,
                        self.assertRaisesRegex(RuntimeError, "토큰"),
                    ):
                        build_release.validate_bug_report_token_preflight()
                    urlopen.assert_not_called()

    def test_http_401_fails_token_preflight(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self._write_valid_token(root)
            unauthorized = error.HTTPError(
                build_release.GITHUB_USER_API,
                401,
                "Unauthorized",
                hdrs=None,
                fp=None,
            )

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(
                    build_release.request,
                    "urlopen",
                    side_effect=unauthorized,
                ),
                self.assertRaisesRegex(RuntimeError, "HTTP 401"),
            ):
                build_release.validate_bug_report_token_preflight()

    def test_wrong_repository_and_network_failure_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self._write_valid_token(root)

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(
                    build_release.request,
                    "urlopen",
                    side_effect=(
                        self._api_response({"login": "token-owner"}),
                        self._api_response({"full_name": "Mrbinggrae/Other"}),
                    ),
                ),
                self.assertRaisesRegex(RuntimeError, "저장소가 UnHelper와 일치하지"),
            ):
                build_release.validate_bug_report_token_preflight()

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(
                    build_release.request,
                    "urlopen",
                    side_effect=error.URLError("offline"),
                ),
                self.assertRaisesRegex(RuntimeError, "네트워크"),
            ):
                build_release.validate_bug_report_token_preflight()

    def test_packaged_token_must_exist_and_be_nonempty(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            dist_dir = Path(temp) / "dist" / "UnHelper"
            packaged_path = dist_dir / build_release.PACKAGED_BUG_REPORT_TOKEN_PATH

            with mock.patch.object(build_release, "DIST_DIR", dist_dir):
                with self.assertRaisesRegex(RuntimeError, "산출물에 오류 신고 토큰"):
                    build_release.require_packaged_bug_report_token()

                packaged_path.parent.mkdir(parents=True)
                packaged_path.touch()
                with self.assertRaisesRegex(RuntimeError, "산출물에 오류 신고 토큰"):
                    build_release.require_packaged_bug_report_token()

                packaged_path.write_bytes(b"opaque-token-data")
                self.assertEqual(
                    build_release.require_packaged_bug_report_token(),
                    packaged_path,
                )

    def test_build_stops_before_manifest_when_pyinstaller_omits_token(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dist_dir = root / "dist" / "UnHelper"
            (root / build_release.BUG_REPORT_TOKEN_NAME).write_bytes(
                b"opaque-token-data"
            )

            with (
                mock.patch.object(build_release, "PROJECT_ROOT", root),
                mock.patch.object(build_release, "DIST_DIR", dist_dir),
                mock.patch.object(build_release, "validate_bug_report_token_preflight"),
                mock.patch.object(build_release, "copy_and_verify_chromedriver"),
                mock.patch.object(build_release, "find_previous_manifest", return_value=None),
                mock.patch.object(build_release, "project_python", return_value=Path("python")),
                mock.patch.object(build_release, "run") as run_command,
                self.assertRaisesRegex(RuntimeError, "산출물에 오류 신고 토큰"),
            ):
                build_release.build_unhelper("0.1.4", with_installer=False)

            run_command.assert_called_once()
            self.assertIn("PyInstaller", run_command.call_args.args[1])


class UpdaterSelectionTests(unittest.TestCase):
    def _wire_api(self, updater: AutoUpdater, releases: list[dict], manifests: dict[str, dict]) -> None:
        updater._fetch_json = lambda url, headers=None: (  # type: ignore[method-assign]
            releases if url == updater.RELEASES_API else manifests[url]
        )

    def test_stable_channel_skips_prerelease_and_uses_delta(self) -> None:
        updater = AutoUpdater("UnHelper", use_prerelease=False)
        updater._current_version = "0.1.0"
        releases = [
            release_payload("0.3.0", prerelease=True),
            release_payload("0.2.0", prerelease=False),
        ]
        manifests = {
            "https://example.test/0.2.0/manifest": manifest_payload("0.2.0"),
        }
        self._wire_api(updater, releases, manifests)

        available, info = updater.check_for_update()
        self.assertTrue(available)
        self.assertIsNotNone(info)
        self.assertEqual(info.version, "0.2.0")
        self.assertEqual(info.patch_mode, "delta")
        self.assertEqual(info.download_url, "https://example.test/0.2.0/delta")

    def test_beta_channel_accepts_prerelease(self) -> None:
        updater = AutoUpdater("UnHelper", use_prerelease=True)
        updater._current_version = "0.1.0"
        releases = [
            release_payload("0.3.0", prerelease=True),
            release_payload("0.2.0", prerelease=False),
        ]
        manifests = {
            "https://example.test/0.3.0/manifest": manifest_payload("0.3.0"),
        }
        self._wire_api(updater, releases, manifests)

        available, info = updater.check_for_update()
        self.assertTrue(available)
        self.assertEqual(info.version, "0.3.0")

    def test_same_version_changed_manifest_is_an_update(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            updater = AutoUpdater("UnHelper")
            updater._current_version = "0.1.0"
            updater.app_dir = Path(temp)
            (updater.app_dir / "release").mkdir()
            local_manifest = manifest_payload("0.1.0")
            local_manifest["files"][0]["sha256"] = "old"
            (updater.app_dir / "release" / "UnHelper_manifest.json").write_text(
                json.dumps(local_manifest),
                encoding="utf-8",
            )
            releases = [release_payload("0.1.0")]
            manifests = {
                "https://example.test/0.1.0/manifest": manifest_payload(
                    "0.1.0", base_version="0.1.0"
                )
            }
            self._wire_api(updater, releases, manifests)

            available, info = updater.check_for_update()
            self.assertTrue(available)
            self.assertTrue(info.is_same_version_update)
            self.assertEqual(info.patch_mode, "full")
            self.assertEqual(info.download_url, "https://example.test/0.1.0/full")

    def test_release_restore_always_selects_stable_full_patch(self) -> None:
        updater = AutoUpdater("UnHelper", use_prerelease=True)
        updater._current_version = "0.2.0"
        releases = [
            release_payload("0.2.0", prerelease=True),
            release_payload("0.1.0", prerelease=False),
        ]
        manifests = {
            "https://example.test/0.1.0/manifest": manifest_payload("0.1.0"),
        }
        self._wire_api(updater, releases, manifests)

        with mock.patch.object(sys, "frozen", True, create=True):
            available, info, message = updater.check_for_release_restore()
        self.assertTrue(available)
        self.assertEqual(message, "")
        self.assertTrue(info.is_release_restore)
        self.assertEqual(info.patch_mode, "full")
        self.assertEqual(info.download_url, "https://example.test/0.1.0/full")

    def test_network_failure_is_not_reported_as_latest(self) -> None:
        updater = AutoUpdater("UnHelper")
        updater._fetch_json = mock.Mock(side_effect=error.URLError("offline"))  # type: ignore[method-assign]
        with self.assertRaisesRegex(RuntimeError, "네트워크 오류"):
            updater.check_for_update()

    def test_release_restore_network_failure_preserves_error(self) -> None:
        updater = AutoUpdater("UnHelper")
        updater._fetch_json = mock.Mock(side_effect=error.URLError("offline"))  # type: ignore[method-assign]
        with (
            mock.patch.object(sys, "frozen", True, create=True),
            self.assertRaisesRegex(RuntimeError, "네트워크 오류"),
        ):
            updater.check_for_release_restore()

    def test_update_apply_failure_is_not_silently_reduced_to_false(self) -> None:
        updater = AutoUpdater("UnHelper")
        with self.assertRaisesRegex(RuntimeError, "적용 준비에 실패"):
            updater.apply_update(Path("missing.zip"), "0.1.1", {})


if __name__ == "__main__":
    unittest.main()
