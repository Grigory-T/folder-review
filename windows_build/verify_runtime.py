"""Verify the complete installed runtime against its hash-pinned archive."""

import hashlib
import json
import sys
import tarfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def verify_runtime():
    pins = json.loads((ROOT / "windows_build/python-pins.json").read_text())
    archive = ROOT / ".build/downloads" / pins["archive"]
    if hashlib.sha256(archive.read_bytes()).hexdigest() != pins["sha256"]:
        raise RuntimeError("Python archive hash mismatch")
    base = Path(sys.base_prefix)
    checked = 0
    with tarfile.open(archive) as source:
        for member in source:
            if not member.isfile():
                continue
            relative = Path(member.name)
            if relative.parts[0] != "python" or ".." in relative.parts:
                raise RuntimeError("Unexpected runtime archive member")
            installed = base.joinpath(*relative.parts[1:])
            expected = hashlib.sha256(source.extractfile(member).read()).digest()
            if hashlib.sha256(installed.read_bytes()).digest() != expected:
                raise RuntimeError(f"Runtime file mismatch: {relative.as_posix()}")
            checked += 1
    return {
        "version": pins["version"],
        "archive_sha256": pins["sha256"],
        "files_checked": checked,
        "all_files_match": True,
    }


if __name__ == "__main__":
    print(json.dumps(verify_runtime()))
