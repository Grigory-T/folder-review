"""Assemble notices and corresponding LGPL sources; build-time network only."""

import argparse
import hashlib
import json
import shutil
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--destination", type=Path, required=True)
    args = parser.parse_args()
    destination = args.destination
    shutil.copytree(
        ROOT / "third_party/licenses",
        destination / "third-party-licenses",
        dirs_exist_ok=True,
    )
    pins = json.loads((ROOT / "third_party/sources.json").read_text())
    cache = ROOT / ".build/downloads"
    cache.mkdir(parents=True, exist_ok=True)
    files = {}
    for item in pins["archives"]:
        path = cache / item["name"]
        if not path.exists():
            with urllib.request.urlopen(item["url"], timeout=60) as response:
                path.write_bytes(response.read())
        if hashlib.sha256(path.read_bytes()).hexdigest() != item["sha256"]:
            raise RuntimeError(f"Corresponding source hash mismatch: {item['name']}")
        files[item["name"]] = path.read_bytes()
    for name in ("README.md", "sources.json"):
        files[name] = (ROOT / "third_party" / name).read_bytes()
    files["LGPL-2.1.txt"] = (ROOT / "third_party/licenses/LGPL-2.1.txt").read_bytes()
    with zipfile.ZipFile(
        destination / "third-party-sources.zip",
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as archive:
        for name, data in sorted(files.items()):
            entry = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr = 0o100644 << 16
            archive.writestr(entry, data, compresslevel=9)
    print(
        json.dumps(
            {
                "corresponding_source_archives": len(pins["archives"]),
                "sha256_verified": True,
            }
        )
    )


if __name__ == "__main__":
    main()
