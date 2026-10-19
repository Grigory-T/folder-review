"""Package a validated dist directory without requiring private release tooling."""

import argparse
import hashlib
import json
import zipfile
from pathlib import Path


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", type=Path, default=Path("dist"))
    args = parser.parse_args()
    root = args.directory
    manifest = json.loads((root / "build-manifest.json").read_text())
    if sha256(root / "folder-review.exe") != manifest["executable"]["sha256"]:
        raise RuntimeError("EXE identity differs from build manifest")
    for name, digest in manifest["runtime_file_sha256"].items():
        if sha256(root / name) != digest:
            raise RuntimeError(f"Portable runtime mismatch: {name}")
    for name, digest in manifest["python_delivery_sha256"].items():
        if sha256(root / "python" / name) != digest:
            raise RuntimeError(f"Python delivery mismatch: {name}")
    if not json.loads((root / "reproducibility.json").read_text())["byte_identical"]:
        raise RuntimeError("Clean rebuild did not pass")
    target = root / "folder-review-windows-x64.zip"
    with zipfile.ZipFile(
        target, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as archive:
        for path in sorted(root.rglob("*")):
            if "folder-review-data" in path.relative_to(root).parts:
                continue
            if path.is_file() and path not in {target, root / "SHA256SUMS.txt"}:
                entry = zipfile.ZipInfo(
                    path.relative_to(root).as_posix(), date_time=(1980, 1, 1, 0, 0, 0)
                )
                entry.compress_type = zipfile.ZIP_DEFLATED
                entry.external_attr = 0o100644 << 16
                archive.writestr(entry, path.read_bytes(), compresslevel=9)
    files = sorted(
        path
        for path in root.iterdir()
        if path.is_file() and path.name != "SHA256SUMS.txt"
    )
    (root / "SHA256SUMS.txt").write_text(
        "".join(f"{sha256(path)}  {path.name}\n" for path in files), encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "exe_bytes": (root / "folder-review.exe").stat().st_size,
                "exe_sha256": sha256(root / "folder-review.exe"),
                "zip_bytes": target.stat().st_size,
            }
        )
    )


if __name__ == "__main__":
    main()
