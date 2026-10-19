"""Inventory build/runtime components and reject native payloads of unknown origin."""

import argparse
import base64
import hashlib
import json
import platform
import ssl
import tarfile
import tomllib
from importlib import metadata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUILD_ONLY = {
    "altgraph",
    "packaging",
    "pefile",
    "pyinstaller",
    "pyinstaller-hooks-contrib",
    "pywin32-ctypes",
    "setuptools",
}
TEST_ONLY = {"openpyxl", "et-xmlfile"}


def canonical(name):
    return name.lower().replace("_", "-").replace(".", "-")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--exe", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    lock = tomllib.loads((ROOT / "uv.lock").read_text())
    packages = {p["name"]: p for p in lock["package"]}
    runtime = json.loads((ROOT / "windows_build/python-pins.json").read_text())
    origins = {}
    with tarfile.open(ROOT / ".build/downloads" / runtime["archive"]) as archive:
        for item in archive:
            if item.isfile() and item.name.lower().endswith((".exe", ".dll", ".pyd")):
                digest = hashlib.sha256(archive.extractfile(item).read()).hexdigest()
                origins[digest] = {
                    "component": "CPython runtime",
                    "file": item.name,
                    "archive_sha256": runtime["sha256"],
                }
    components = []
    for distribution in sorted(
        metadata.distributions(), key=lambda d: canonical(d.metadata["Name"])
    ):
        name = canonical(distribution.metadata["Name"])
        role = (
            "build"
            if name in BUILD_ONLY
            else "test reader"
            if name in TEST_ONLY
            else "runtime"
        )
        files_checked = 0
        for file in distribution.files or []:
            if str(file).lower().endswith((".py", ".pyd", ".dll", ".exe")):
                if ".." in Path(str(file)).parts:
                    # Installer-generated entry points live outside site-packages;
                    # builds use python -m, and these wrappers are not bundled.
                    continue
                path = Path(distribution.locate_file(file))
                raw = path.read_bytes()
                if file.hash is None or file.hash.mode != "sha256":
                    raise RuntimeError(f"Missing code integrity record: {name}/{file}")
                digest = hashlib.sha256(raw).digest()
                if (
                    base64.urlsafe_b64encode(digest).decode().rstrip("=")
                    != file.hash.value
                ):
                    raise RuntimeError(
                        f"Installed wheel code differs from RECORD: {name}/{file}"
                    )
                origins[digest.hex()] = {
                    "component": name,
                    "version": distribution.version,
                    "file": str(file).replace("\\", "/"),
                }
                files_checked += 1
        package = packages.get(name, {})
        components.append(
            {
                "name": name,
                "version": distribution.version,
                "role": role,
                "purl": f"pkg:pypi/{name}@{distribution.version}",
                "license": distribution.metadata.get("License-Expression")
                or distribution.metadata.get("License"),
                "project_urls": distribution.metadata.get_all("Project-URL") or [],
                "code_files_record_verified": files_checked,
                "source": package.get("sdist"),
                "artifact_lock": "uv.lock"
                if package
                else "windows_build/requirements-build.txt",
            }
        )
    sevenzip = json.loads((ROOT / "windows_build/7zip-pins.json").read_text())
    for name in ("7z.exe", "7z.dll"):
        origins[sevenzip["sha256"][name]] = {
            "component": "7-Zip",
            "version": sevenzip["version"],
            "file": name,
            "installer_sha256": sevenzip["installer_sha256"],
        }
    native = []
    for path in sorted((args.exe.parent / "_internal").rglob("*")):
        name = path.relative_to(args.exe.parent).as_posix()
        if path.is_file() and name.lower().endswith((".dll", ".pyd", ".exe")):
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            if digest not in origins:
                raise RuntimeError(
                    f"Native payload has no verified upstream origin: {name}"
                )
            native.append(
                {
                    "path": name.replace("\\", "/"),
                    "sha256": digest,
                    "origin": origins[digest],
                }
            )
    report = {
        "format": "Folder Review component inventory 1",
        "source_version": tomllib.loads((ROOT / "pyproject.toml").read_text())[
            "project"
        ]["version"],
        "executable_sha256": hashlib.sha256(args.exe.read_bytes()).hexdigest(),
        "python_runtime": runtime,
        "openssl": ssl.OPENSSL_VERSION,
        "python": platform.python_version(),
        "components": components,
        "native_payloads": native,
        "unknown_native_payloads": 0,
        "bootloader": "PyInstaller wheel, fixed PE resources plus appended application archive; complete EXE verified by clean rebuild",
        "limits": "Provenance and RECORD checks do not prove absence of vulnerabilities or malicious upstream code. Windows system libraries are supplied by the OS.",
    }
    args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {
                "components": len(components),
                "native_payloads": len(native),
                "unknown_native_payloads": 0,
            }
        )
    )


if __name__ == "__main__":
    main()
