"""Build the deterministic source-only ZIP; executable downloads happen at first setup."""

from __future__ import annotations

import argparse
import configparser
import hashlib
import json
import re
import subprocess
import tomllib
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def payload():
    paths = {
        "folder-review.bat": ROOT / "folder-review.bat",
        "folder-change-review.bat": ROOT / "folder-change-review.bat",
        "folder-review.ini": ROOT / "folder-review.ini",
        "folder-change-review.ini": ROOT / "folder-change-review.ini",
        "support/folder_change_review.py": ROOT / "folder_change_review.py",
        "support/review_state.py": ROOT / "review_state.py",
        "support/folder_review.py": ROOT / "folder_review.py",
        "support/LICENSE": ROOT / "LICENSE",
        "support/SECURITY.md": ROOT / "SECURITY.md",
    }
    for name in (
        "launcher.ps1",
        "pins.json",
        "requirements.txt",
        "run_review.py",
        "run_change_review.py",
        "runtime_check.py",
        "README.txt",
        "THIRD_PARTY.txt",
    ):
        paths["support/" + name] = ROOT / "portable" / name
    contents = {}
    for name, path in paths.items():
        if path.is_symlink():
            raise ValueError("Linked release source is not permitted")
        data = path.read_bytes().replace(b"\r\n", b"\n")
        if name.endswith(".bat"):
            data = data.replace(b"\n", b"\r\n")
        if name.endswith(".py"):
            compile(data, name, "exec")
        contents[name] = data
    return contents


def build(output, release=False, source_commit=None):
    project = tomllib.loads((ROOT / "pyproject.toml").read_text())["project"]
    lock = tomllib.loads((ROOT / "uv.lock").read_text())
    pins = json.loads((ROOT / "portable/pins.json").read_text())
    assert pins["python"]["version"] == "3.13.16" and pins["uv"]["version"] == "0.13.1"
    packages = {p["name"]: p["version"] for p in lock["package"]}
    assert all(packages[n] == v for n, v in pins["packages"].items())
    exported = subprocess.check_output(
        [
            "uv",
            "--offline",
            "export",
            "--frozen",
            "--no-dev",
            "--no-emit-project",
            "--no-header",
            "--no-annotate",
        ],
        cwd=ROOT,
    )
    assert (
        exported.replace(b"\r\n", b"\n")
        == (ROOT / "portable/requirements.txt").read_bytes()
    ), "Regenerate frozen requirements from uv.lock"
    assert re.fullmatch(r"\d+\.\d+\.\d+", project["version"])
    if source_commit is not None and not re.fullmatch(r"[0-9a-f]{40}", source_commit):
        raise ValueError("Source commit must be a complete lowercase Git SHA-1")
    try:
        dirty = subprocess.check_output(
            ["git", "status", "--porcelain", "--untracked-files=normal"],
            cwd=ROOT,
            stderr=subprocess.DEVNULL,
        )
        head = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip()
        if source_commit is not None and source_commit != head:
            raise ValueError("Provided source commit differs from checkout")
    except (OSError, subprocess.CalledProcessError):
        dirty, head = b"", source_commit
    if release and (dirty or head is None):
        raise ValueError(
            "Use a clean checkout, or specify --source-commit when building a source archive"
        )
    contents = payload()
    settings = configparser.ConfigParser(interpolation=None)
    settings.read_string(contents["folder-review.ini"].decode("utf-8-sig"))
    if settings["review"]["folder"].strip():
        raise ValueError("Release settings must leave the target folder blank")
    comparison = configparser.ConfigParser(interpolation=None)
    comparison.read_string(contents["folder-change-review.ini"].decode("utf-8-sig"))
    if any(comparison["compare"][key].strip() for key in ("before", "after")):
        raise ValueError(
            "Release comparison settings must leave both input paths blank"
        )
    manifest = {
        "format": 1,
        "version": project["version"],
        "source_commit": head if not dirty else None,
        "uv_lock_sha256": hashlib.sha256((ROOT / "uv.lock").read_bytes()).hexdigest(),
        "immutable_sha256": {
            n: hashlib.sha256(b).hexdigest()
            for n, b in contents.items()
            if not n.endswith(".ini")
        },
        "delivery": "Source-only Windows x86-64; fixed verified first-time setup; offline recurring scans",
    }
    contents["support/manifest.json"] = (json.dumps(manifest, indent=2) + "\n").encode()
    output.mkdir(parents=True, exist_ok=True)
    target = output / "folder-review-windows-x64.zip"
    with zipfile.ZipFile(
        target, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as archive:
        for name, data in sorted(contents.items()):
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data, compresslevel=9)
    with zipfile.ZipFile(target) as archive:
        assert archive.testzip() is None
        assert set(archive.namelist()) == set(contents)
    result = {
        "version": project["version"],
        "source_commit": manifest["source_commit"],
        "zip_bytes": target.stat().st_size,
        "zip_sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
        "files": {
            n: hashlib.sha256(b).hexdigest() for n, b in sorted(contents.items())
        },
    }
    (output / "build-manifest.json").write_text(json.dumps(result, indent=2) + "\n")
    (output / "SHA256SUMS.txt").write_text(f"{result['zip_sha256']}  {target.name}\n")
    print(json.dumps({k: v for k, v in result.items() if k != "files"}))
    return target


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "dist")
    parser.add_argument("--release", action="store_true")
    parser.add_argument(
        "--source-commit",
        help="Release commit ID when building a GitHub source archive without .git",
    )
    args = parser.parse_args()
    build(args.output, args.release, args.source_commit)
