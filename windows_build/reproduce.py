"""Require a byte-identical production rebuild from a fresh source directory."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUILD_ENV = {"PYTHONHASHSEED": "0", "SOURCE_DATE_EPOCH": "1767225600"}


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", required=True, type=Path)
    parser.add_argument("--results", required=True, type=Path)
    args = parser.parse_args()
    if os.name != "nt" or sys.flags.optimize:
        parser.error("Use Windows x86-64 and Python optimization level zero")
    for name, value in BUILD_ENV.items():
        if os.environ.get(name) != value:
            raise RuntimeError(f"Required build setting: {name}={value}")
    original = args.exe.resolve()
    inputs = [
        ROOT / "folder_review.py",
        ROOT / "folder_review_ui.py",
        ROOT / "tests/test_folder_review_ui.py",
        ROOT / "BUILD-EXE.bat",
        ROOT / "RUN.bat",
        ROOT / "SETUP.bat",
        ROOT / "README.md",
        ROOT / "LICENSE",
        ROOT / "pyproject.toml",
        ROOT / "uv.lock",
        ROOT / "tests/test_folder_review.py",
        *sorted((ROOT / "tests/fixtures").rglob("*")),
        *sorted((ROOT / "windows_build").glob("*")),
    ]
    inputs = [path for path in inputs if path.is_file()]
    hashes = {path.relative_to(ROOT).as_posix(): sha256(path) for path in inputs}
    with tempfile.TemporaryDirectory(prefix="rebuild-", dir=ROOT / ".build") as name:
        fresh = Path(name)
        for source in inputs:
            destination = fresh / source.relative_to(ROOT)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, destination)
        env = os.environ.copy()
        env.pop("FOLDER_REVIEW_BUILD_TESTS", None)
        subprocess.run(
            [
                sys.executable,
                "-m",
                "PyInstaller",
                "--noconfirm",
                "--clean",
                "--distpath",
                str(fresh / "dist"),
                "--workpath",
                str(fresh / ".build"),
                str(fresh / "windows_build/folder-review.spec"),
            ],
            cwd=fresh,
            env=env,
            check=True,
        )
        rebuilt = fresh / "dist/folder-review.exe"
        expected, actual = sha256(original), sha256(rebuilt)
        if expected != actual:
            raise RuntimeError(f"Clean rebuild differs: {expected} != {actual}")
        report = {
            "byte_identical": True,
            "exe_sha256": expected,
            "rebuild_sha256": actual,
            "fresh_source_directory": True,
            "build_environment": BUILD_ENV,
            "source_input_sha256": hashes,
        }
    args.results.parent.mkdir(parents=True, exist_ok=True)
    args.results.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                key: value
                for key, value in report.items()
                if key != "source_input_sha256"
            }
        )
    )


if __name__ == "__main__":
    main()
