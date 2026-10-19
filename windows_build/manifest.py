"""Record reproducible inputs, executable identity, and third-party licenses."""

import argparse
import hashlib
import io
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import tomllib
import urllib.request
from importlib import metadata
from pathlib import Path

from PyInstaller.archive.readers import CArchiveReader

ROOT = Path(__file__).resolve().parent.parent


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def collect_licenses(destination, distributions, lock):
    destination.mkdir(parents=True, exist_ok=True)
    packages = {p["name"]: p for p in lock["package"]}
    for name, distribution in distributions.items():
        folder = destination / name
        folder.mkdir(exist_ok=True)
        (folder / "METADATA.txt").write_text(
            distribution.read_text("METADATA") or "", encoding="utf-8"
        )
        found = False
        for file in distribution.files or []:
            if (
                "license" not in str(file).lower()
                and "copying" not in str(file).lower()
            ):
                continue
            path = Path(distribution.locate_file(file))
            if path.is_file():
                shutil.copyfile(
                    path, folder / str(file).replace("/", "_").replace("\\", "_")
                )
                found = True
        # A few wheels omit the actual MIT text; copy it from the hash-pinned
        # source archive in uv.lock. Never execute or unpack source code.
        if not found and name in packages:
            source = packages[name]["sdist"]
            with urllib.request.urlopen(source["url"], timeout=30) as response:
                data = response.read()
            assert "sha256:" + hashlib.sha256(data).hexdigest() == source["hash"]
            with tarfile.open(fileobj=io.BytesIO(data), mode="r:*") as archive:
                for member in archive.getmembers():
                    if member.isfile() and (
                        "license" in Path(member.name).name.lower()
                        or "licence" in Path(member.name).name.lower()
                        or "copying" in Path(member.name).name.lower()
                    ):
                        stream = archive.extractfile(member)
                        (folder / member.name.replace("/", "_")).write_bytes(
                            stream.read()
                        )
                        found = True
            assert found, f"No license text found for {name}"
    for filename in ("LICENSE.txt", "LICENSE"):
        source = Path(sys.base_prefix) / filename
        if source.is_file():
            shutil.copyfile(source, destination / "Python-LICENSE.txt")
            break
    else:
        raise FileNotFoundError("Python runtime license text not found")


def main():
    if sys.flags.optimize:
        raise RuntimeError("Release preparation requires Python optimization level 0")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    lock = tomllib.loads((ROOT / "uv.lock").read_text(encoding="utf-8"))
    distributions = {}
    for name in [p["name"] for p in lock["package"] if p["name"] != "folder-review"] + [
        "pyinstaller",
        "pyinstaller-hooks-contrib",
    ]:
        try:
            distributions[name] = metadata.distribution(name)
        except metadata.PackageNotFoundError:
            continue  # uv.lock also describes other Python implementations/OSes.
    collect_licenses(args.output.parent / "third-party-licenses", distributions, lock)
    tkinter_license = Path(sys.base_prefix) / "tcl/tk8.6/license.terms"
    shutil.copyfile(
        tkinter_license, args.output.parent / "third-party-licenses/Tcl-Tk-LICENSE.txt"
    )
    archive = CArchiveReader(str(args.exe.resolve()))
    try:
        commit = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True, stderr=subprocess.DEVNULL
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        commit = os.environ.get("FOLDER_REVIEW_SOURCE_COMMIT")
        if commit is not None and re.fullmatch(r"[0-9a-f]{40}", commit) is None:
            raise ValueError("Source commit must be a full lowercase Git SHA-1")
    inputs = [
        ROOT / "folder_review.py",
        ROOT / "folder_review_ui.py",
        ROOT / "tests/test_folder_review_ui.py",
        ROOT / "BUILD-EXE.bat",
        ROOT / ".gitattributes",
        ROOT / "RUN.bat",
        ROOT / "SETUP.bat",
        ROOT / "README.md",
        ROOT / "LICENSE",
        ROOT / "SECURITY.md",
        *sorted((ROOT / "third_party").rglob("*")),
        ROOT / "pyproject.toml",
        ROOT / "uv.lock",
        ROOT / "tests/test_folder_review.py",
        *sorted((ROOT / "tests/fixtures").rglob("*")),
        *sorted((ROOT / "windows_build").glob("*")),
    ]
    python_files = (
        "folder_review.py",
        "folder_review_ui.py",
        "pyproject.toml",
        "uv.lock",
        "RUN.bat",
        "SETUP.bat",
        "README.md",
        "LICENSE",
    )
    python_files += tuple(
        path.relative_to(ROOT).as_posix()
        for path in sorted((ROOT / "third_party").rglob("*"))
        if path.is_file()
    )
    python_delivery = args.output.parent / "python"
    for name in python_files:
        assert (python_delivery / name).read_bytes() == (ROOT / name).read_bytes(), (
            "Python delivery differs from validated build input",
            name,
        )
    reproducibility = json.loads(
        (args.output.parent / "reproducibility.json").read_text()
    )
    assert reproducibility["byte_identical"]
    assert (
        reproducibility["exe_sha256"]
        == reproducibility["rebuild_sha256"]
        == sha256(args.exe)
    )
    runtime_files = {
        p.relative_to(args.exe.parent).as_posix(): sha256(p)
        for p in sorted((args.exe.parent / "_internal").rglob("*"))
        if p.is_file()
    }
    assert reproducibility["runtime_tree_byte_identical"]
    assert runtime_files == reproducibility["runtime_file_sha256"], (
        "Published runtime differs from clean rebuild"
    )
    manifest = {
        "source_version": tomllib.loads((ROOT / "pyproject.toml").read_text())[
            "project"
        ]["version"],
        "source_commit": commit,
        "python_delivery_sha256": {
            name: sha256(python_delivery / name) for name in python_files
        },
        "input_sha256": {
            p.relative_to(ROOT).as_posix(): sha256(p) for p in inputs if p.is_file()
        },
        "input_sha256_lf": {
            p.relative_to(ROOT).as_posix(): hashlib.sha256(
                p.read_bytes().replace(b"\r\n", b"\n")
                if p.suffix
                in {
                    ".py",
                    ".toml",
                    ".lock",
                    ".md",
                    ".txt",
                    ".json",
                    ".ps1",
                    ".bat",
                    ".spec",
                }
                else p.read_bytes()
            ).hexdigest()
            for p in inputs
            if p.is_file()
        },
        "python": platform.python_version(),
        "build_platform": platform.platform(),
        "package_versions": {name: d.version for name, d in distributions.items()},
        "python_runtime_pin": json.loads(
            (ROOT / "windows_build/python-pins.json").read_text()
        ),
        "component_inventory_sha256": sha256(args.output.parent / "components.json"),
        "corresponding_source_bundle_sha256": sha256(
            args.output.parent / "third-party-sources.zip"
        ),
        "report_writer": {"library": "XlsxWriter", "constant_memory": True},
        "content_identity": {
            "algorithm": "BLAKE3-256",
            "display": "Consecutive report-local content group IDs starting at 1",
            "last": "Always 1; filter virtual rows and states explicitly",
        },
        "test_reader_only_packages": ["openpyxl", "et-xmlfile"],
        "sevenzip": json.loads((ROOT / "windows_build/7zip-pins.json").read_text()),
        "executable": {
            "name": args.exe.name,
            "bytes": args.exe.stat().st_size,
            "sha256": sha256(args.exe),
        },
        "distribution": "Portable PyInstaller onedir; keep EXE and _internal together",
        "runtime_file_sha256": runtime_files,
        "compression": "PyInstaller standard zlib, level 9; no UPX, strip or bytecode optimization",
        "excluded_build_only_modules": ["setuptools", "pkg_resources", "packaging"],
        "payload_entries": len(archive.toc),
        "payload_sha256": {
            name: hashlib.sha256(data).hexdigest()
            for name in sorted(archive.toc)
            if (data := archive.extract(name)) is not None
        },
        "scanner_logic": "Unmodified folder_review.py; embedded bytecode checked against optimize=0 source compilation",
        "packaging_adaptations": [
            "Default output beside executable",
            "Pinned bundled RAR backend",
            "Unprintable diagnostic characters escaped on redirected legacy-codepage streams",
        ],
        "validation": {
            "source_suite": json.loads(
                (args.output.parent / "source-tests.json").read_text()
            ),
            "frozen_suite": json.loads(
                (args.output.parent / "frozen-tests.json").read_text()
            ),
            "executable_differential_suite": "verification.json",
            "byte_identical_clean_rebuild": "reproducibility.json",
        },
        "validation_limits": [
            "Synthetic corpus cannot prove every possible input",
            "Archive time limits retain the original cooperative behavior",
            "Windows 11 tested; Windows 10 remains untested",
            "Access time and generated run/workbook timestamps differ legitimately between runs",
        ],
    }
    args.output.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
