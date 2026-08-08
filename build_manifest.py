from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path
from typing import Any


CHUNK_SIZE = 1024 * 1024


def sha256_file(filepath: Path) -> str:
    """Return the SHA-256 digest for *filepath*."""
    digest = hashlib.sha256()
    with filepath.open("rb") as handle:
        for chunk in iter(lambda: handle.read(CHUNK_SIZE), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_manifest(dist_dir: Path) -> dict[str, Any]:
    """Describe every file in a PyInstaller onedir distribution."""
    files: list[dict[str, Any]] = []
    for filepath in sorted(dist_dir.rglob("*")):
        if not filepath.is_file():
            continue
        files.append(
            {
                "path": filepath.relative_to(dist_dir).as_posix(),
                "sha256": sha256_file(filepath),
                "size": filepath.stat().st_size,
            }
        )
    return {"files": files}


def load_manifest(path: str | Path | None) -> dict[str, Any] | None:
    if not path:
        return None

    manifest_path = Path(path)
    if not manifest_path.is_file():
        return None

    with manifest_path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, dict) else None


def changed_files(current: dict[str, Any], previous: dict[str, Any]) -> list[dict[str, Any]]:
    """Return new and modified files relative to *previous*."""
    previous_hashes = {
        item.get("path"): item.get("sha256")
        for item in previous.get("files", [])
        if isinstance(item, dict) and item.get("path")
    }
    return [
        item
        for item in current.get("files", [])
        if isinstance(item, dict)
        and previous_hashes.get(item.get("path")) != item.get("sha256")
    ]


def create_zip(dist_dir: Path, file_list: list[dict[str, Any]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for item in file_list:
            relative_path = str(item["path"])
            full_path = dist_dir / Path(relative_path)
            if full_path.is_file():
                archive.write(full_path, relative_path)
    print(f"  -> {output_path} ({output_path.stat().st_size / 1024 / 1024:.1f} MB)")


def zip_metadata(path: Path) -> dict[str, Any]:
    return {
        "file": path.name,
        "sha256": sha256_file(path),
        "size": path.stat().st_size,
    }


def build_release_assets(
    *,
    app_name: str,
    dist_dir: Path,
    version: str,
    output_dir: Path,
    previous_manifest: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build a full patch and, when possible, a previous-release delta."""
    if not dist_dir.is_dir():
        raise FileNotFoundError(f"dist folder not found: {dist_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    prefix = app_name.replace(" ", "_")
    manifest_path = output_dir / f"{prefix}_manifest.json"
    full_patch_path = output_dir / f"{prefix}_patch.zip"
    delta_patch_path = output_dir / f"{prefix}_delta_patch.zip"

    print(f"[1/3] manifest: {dist_dir}")
    manifest = build_manifest(dist_dir)
    manifest.update(
        {
            "app_name": app_name,
            "version": version,
            "patch_mode": "hybrid",
        }
    )
    total_size = sum(int(item["size"]) for item in manifest["files"])
    print(f"  -> files {len(manifest['files'])}, total {total_size / 1024 / 1024:.1f} MB")

    print("[2/3] full patch")
    create_zip(dist_dir, manifest["files"], full_patch_path)
    full_patch = zip_metadata(full_patch_path)
    manifest["full_patch"] = full_patch
    # Kept for compatibility with older TC Helper-style clients.
    manifest["zip_hash"] = full_patch["sha256"]

    print("[3/3] previous-release delta")
    if previous_manifest:
        current_paths = {
            str(item.get("path"))
            for item in manifest.get("files", [])
            if isinstance(item, dict) and item.get("path")
        }
        previous_paths = {
            str(item.get("path"))
            for item in previous_manifest.get("files", [])
            if isinstance(item, dict) and item.get("path")
        }
        manifest["deleted_files"] = sorted(previous_paths - current_paths)
        delta_items = changed_files(manifest, previous_manifest)
        if delta_items:
            create_zip(dist_dir, delta_items, delta_patch_path)
            delta_patch = zip_metadata(delta_patch_path)
            delta_patch["base_version"] = str(previous_manifest.get("version", ""))
            delta_patch["base_file_count"] = len(previous_manifest.get("files", []))
            manifest["delta_patch"] = delta_patch
            print(
                f"  -> changed files {len(delta_items)} "
                f"(base v{delta_patch['base_version'] or 'unknown'})"
            )
        else:
            delta_patch_path.unlink(missing_ok=True)
            print("  -> no changed files; delta omitted")
    else:
        delta_patch_path.unlink(missing_ok=True)
        print("  -> previous manifest unavailable; delta omitted")

    with manifest_path.open("w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2, ensure_ascii=False)
        handle.write("\n")
    print(f"[OK] saved: {manifest_path}")
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="UnHelper manifest and patch builder")
    parser.add_argument("--app", default="UnHelper")
    parser.add_argument("--dist-dir", default="dist/UnHelper")
    parser.add_argument("--version", required=True)
    parser.add_argument("--prev-manifest")
    parser.add_argument("--output-dir", default="release")
    args = parser.parse_args()

    try:
        build_release_assets(
            app_name=args.app,
            dist_dir=Path(args.dist_dir),
            version=args.version,
            output_dir=Path(args.output_dir),
            previous_manifest=load_manifest(args.prev_manifest),
        )
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
