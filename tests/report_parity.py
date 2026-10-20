"""Compare the new settings runner with reviewed v0.18 source on synthetic inputs."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

import py7zr
from report_compare import normalized_workbook

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--tool", required=True, type=Path)
parser.add_argument("--baseline", required=True, type=Path)
parser.add_argument("--results", required=True, type=Path)
args = parser.parse_args()
args.results.parent.mkdir(parents=True, exist_ok=True)
tool = args.tool.resolve()
python = tool / "support/state/venv/Scripts/python.exe"
fixtures = Path(__file__).parent / "fixtures/windows"
settings = tool / "settings.ini"
original = settings.read_bytes()
results = []
try:
    with tempfile.TemporaryDirectory(
        prefix="report-parity-", dir=args.results.parent
    ) as temp:
        work = Path(temp)
        baseline = work / "baseline"
        baseline.mkdir()
        shutil.copyfile(args.baseline, baseline / "folder_review.py")
        wrapper = baseline / "run.py"
        wrapper.write_text("""import sys
from pathlib import Path
import folder_review as core
backend=sys.argv.pop(1)
def configure():
    core.rarfile.SEVENZIP_TOOL=backend
    core.rarfile.tool_setup(force=True,unrar=False,unar=False,bsdtar=False,sevenzip=True,sevenzip2=False)
core.configure_rar_backend=configure
raise SystemExit(core.main())
""")
        complete = work / "complete"
        complete.mkdir()
        (complete / "empty/deep").mkdir(parents=True)
        payload = b"content duplicate" * 200
        for name in ("one.txt", "two.bin", "=literal.txt"):
            (complete / name).write_bytes(payload)
        (complete / "unique.txt").write_bytes(b"unique length")
        (complete / "empty.bin").write_bytes(b"")
        (complete / "unicode-\u044f.txt").write_bytes(b"\xef\xbb\xbfUnicode BOM")
        with zipfile.ZipFile(complete / "diagram.vsdx", "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in {
                "[Content_Types].xml": b"<Types/>",
                "_rels/.rels": b"<Relationships/>",
                "visio/document.xml": b"<VisioDocument/>",
            }.items():
                z.writestr(name, data)
        inner = work / "inner.7z"
        with py7zr.SevenZipFile(inner, "w") as z:
            z.writestr(payload, "inside.txt")
        with zipfile.ZipFile(complete / "nested.zip", "w", zipfile.ZIP_DEFLATED) as z:
            z.write(inner, "inner.7z")
            z.writestr("empty/deep/", b"")
        shutil.copyfile(fixtures / "rar5-solid.rar", complete / "solid.rar")
        partial = work / "partial"
        partial.mkdir()
        for source, name in [
            ("rar5-encrypted-headers.rar", "encrypted.rar"),
            ("rar5-corrupt.rar", "corrupt.rar"),
        ]:
            shutil.copyfile(fixtures / source, partial / name)
        with zipfile.ZipFile(partial / "unusable.zip", "w") as z:
            z.writestr(".", b"cannot disappear")
        before = {
            p.relative_to(work).as_posix(): (
                hashlib.sha256(p.read_bytes()).hexdigest(),
                p.stat().st_mtime_ns,
            )
            for folder in (complete, partial)
            for p in folder.rglob("*")
            if p.is_file()
        }
        cases = [
            ("filtered", complete, {}, 0),
            ("quick", complete, {"duplicates": "off"}, 0),
            ("full", complete, {"duplicates": "all"}, 0),
            ("advanced", complete, {"full_metadata": "yes"}, 0),
            ("partial", partial, {}, 2),
            ("row-limit", complete, {"max_rows": "2"}, 2),
            ("byte-limit", complete, {"max_expanded_gib": "0.000000001"}, 2),
            ("time-limit", complete, {"max_archive_seconds": "0.000000001"}, 2),
        ]
        for name, source, options, expected in cases:
            report_name = "parity-" + work.name + ".xlsx"
            configuration = {
                "folder": str(source),
                "workers": "1",
                "report": report_name,
                **options,
            }
            settings.write_text(
                "[review]\n"
                + "".join(f"{k} = {v}\n" for k, v in configuration.items()),
                encoding="utf-8",
            )
            new = tool / report_name
            old = baseline / "report.xlsx"
            flags = []
            for key, value in options.items():
                flags.append("--" + key.replace("_", "-"))
                if key != "full_metadata":
                    flags.append(value)
            for command, report in [
                (
                    [
                        str(python),
                        str(wrapper),
                        str(tool / "support/state/backend/7z.exe"),
                        "--root",
                        str(source),
                        "--workers",
                        "1",
                        "--output",
                        str(old),
                        *flags,
                    ],
                    old,
                ),
                ([str(python), str(tool / "support/run_review.py")], new),
            ]:
                run = subprocess.run(
                    command, capture_output=True, text=True, check=False, timeout=60
                )
                assert run.returncode == expected, (
                    name,
                    run.returncode,
                    run.stdout,
                    run.stderr,
                )
                assert report.exists()
            a, status, rows = normalized_workbook(old)
            b, new_status, new_rows = normalized_workbook(new)
            different = [part for part in a if a[part] != b.get(part)]
            assert a.keys() == b.keys() and not different, (name, different)
            stable_status = {
                k: v
                for k, v in status.items()
                if k not in {"Started at", "Completed at"}
            }
            stable_new_status = {
                k: v
                for k, v in new_status.items()
                if k not in {"Started at", "Completed at"}
            }
            assert stable_status == stable_new_status and rows == new_rows, name
            new.unlink()
            results.append(
                {
                    "case": name,
                    "exit_code": expected,
                    "rows": rows,
                    "all_workbook_members_match": True,
                }
            )
        after = {
            p.relative_to(work).as_posix(): (
                hashlib.sha256(p.read_bytes()).hexdigest(),
                p.stat().st_mtime_ns,
            )
            for folder in (complete, partial)
            for p in folder.rglob("*")
            if p.is_file()
        }
        assert before == after, "Synthetic inputs were modified"
finally:
    settings.write_bytes(original)
receipt = {
    "baseline": "v0.18.0",
    "cases": results,
    "input_bytes_and_modification_times_unchanged": True,
    "volatile_fields": [
        "access time",
        "scan start/end",
        "workbook creation/modification time",
    ],
}
args.results.write_text(json.dumps(receipt, indent=2) + "\n")
print(json.dumps(receipt))
