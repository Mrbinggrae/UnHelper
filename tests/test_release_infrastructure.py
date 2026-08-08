from __future__ import annotations

import json
import hashlib
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


if __name__ == "__main__":
    unittest.main()
