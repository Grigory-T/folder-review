"""Versioned physical-file states embedded in ordinary Folder Review workbooks.

The hidden worksheet contains compressed, checksummed JSON, never executable
objects. Precise integers and original names survive Excel cell conversions.
"""

from __future__ import annotations

import base64
import hashlib
import json
import ntpath
import posixpath
import re
import time
import xml.etree.ElementTree as ET
import zipfile
import zlib
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from pathlib import Path
from types import SimpleNamespace

import folder_change_review as changes
import folder_review as core

SHEET = "_folder_review_state"
FORMAT = "folder-review.comparison-state"
MARKER = "folder-review-state-v1"
POLICY = "physical-v1"
CHUNK = 30000
MAX_JSON = 256 * 1024 * 1024
MAX_PACKED = 64 * 1024 * 1024
MAX_XML = 96 * 1024 * 1024
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def path_module(root):
    return (
        ntpath if ntpath.splitdrive(root)[0] or root.startswith("\\\\") else posixpath
    )


def capture_state(root, physical_files, stats, mode, workers, output):
    """Reuse enumeration metadata; content reads are explicitly opt-in."""
    if mode not in changes.REVIEW_MODES:
        raise ValueError("state_mode must be quick, sampled, or full.")
    started = time.monotonic()
    root = core.absolute_path(root)
    module = path_module(root)
    base = core.application_directory()
    scopes = []
    outputs = [
        Path(output),
        base / changes.OUTPUT_FILE,
        *core.inventory_report_files(Path(output)),
    ]
    try:
        values = core.read_ini(
            base / "folder-change-review.txt",
            "compare",
            {"before", "after", "report", "diagnostics"},
        )["compare"]
        outputs.append(
            base / core.report_filename(values.get("report", changes.OUTPUT_FILE))
        )
    except (OSError, ValueError):
        pass
    for path in [
        *core.support_directories(),
        *outputs,
        *[
            base / name
            for name in (
                "folder-review.bat",
                "folder-review.txt",
                "folder-change-review.bat",
                "folder-change-review.txt",
            )
        ],
    ]:
        try:
            relative = module.relpath(str(path), root)
        except ValueError:
            continue
        if relative != ".." and not relative.startswith(".." + module.sep):
            scopes.append(relative)

    def excluded(path, parts, kind="file"):
        if ".states" in (p.casefold() for p in parts):
            return True
        if kind == "file" and parts and changes.ignored(parts[-1]):
            return True
        candidate = module.normcase(module.normpath(path))
        return any(
            candidate == (scope := module.normcase(module.join(root, item)))
            or candidate.startswith(scope + module.sep)
            for item in scopes
        )

    def issue(path, kind, message):
        relative = module.relpath(path, root)
        parts = [] if relative == "." else relative.split(module.sep)
        return {
            "path": path,
            "parts": parts,
            "kind": kind,
            "message": message,
            "errno": None,
            "winerror": None,
        }

    errors = []
    for diagnostic in stats.diagnostics:
        kind = "directory" if diagnostic.stage == "folder enumeration" else "unknown"
        item = issue(diagnostic.path, kind, diagnostic.error)
        if not excluded(item["path"], item["parts"], kind):
            errors.append(item)
    files = []
    jobs = []
    for item in physical_files:
        if item.virtual_folder or excluded(item.path, item.levels):
            continue
        if item.item_type in {"Symbolic link", "Directory junction", "Reparse point"}:
            continue
        if not item.inspect_content or item.properties is None:
            errors.append(issue(item.path, "unknown", "Unsupported filesystem item."))
            continue
        jobs.append(item)

    def inspect(item):
        stamp = item.properties.modified_ns
        if type(stamp) is not int:
            raise OSError("Modification time is unavailable for comparison.")
        if mode == "quick":
            return None, item.size, stamp, 0
        if mode == "sampled":
            digest, size, read = changes.sample_file(
                item.path, SimpleNamespace(st_size=item.size, st_mtime_ns=stamp)
            )
        else:
            digest, size = changes.hash_file(item.path)
            read = size
        return digest, size, stamp, read

    read_bytes = 0

    def finish(item, get_result):
        nonlocal read_bytes
        try:
            digest, size, stamp, read = get_result()
        except OSError as error:
            error_item = issue(
                item.path, "file", str(getattr(error, "strerror", None) or error)
            )
            errors.append(error_item)
            stats.diagnostics.append(
                core.Diagnostic(item.path, "comparison state", error_item["message"])
            )
        else:
            files.append(
                {
                    "path": item.path,
                    "parts": list(item.levels),
                    "size": size,
                    "mtime_ns": stamp,
                    "sha256": digest,
                }
            )
            read_bytes += read
            core.diagnostic_count("state_files")
            core.diagnostic_count("state_bytes_read", read)

    if mode == "quick":
        for item in jobs:
            finish(item, lambda item=item: inspect(item))
    else:
        # Match the original bounded content queue; no unbounded file futures.
        iterator = iter(jobs)
        with ThreadPoolExecutor(max_workers=workers) as pool:
            pending = {}
            exhausted = False
            while pending or not exhausted:
                while not exhausted and len(pending) < workers * 2:
                    item = next(iterator, None)
                    if item is None:
                        exhausted = True
                    else:
                        pending[pool.submit(inspect, item)] = item
                if pending:
                    done, _ = wait(pending, timeout=1, return_when=FIRST_COMPLETED)
                    for future in done:
                        finish(pending.pop(future), future.result)
    files.sort(key=lambda record: record["path"])
    errors.sort(key=lambda record: (record["path"], record["kind"]))
    result = {
        "format": FORMAT,
        "version": 1,
        "policy": POLICY,
        "root": root,
        "review_mode": mode,
        "hash_algorithm": {
            "quick": "none",
            "sampled": "sha256-sample-v1",
            "full": "sha256",
        }[mode],
        "files": files,
        "file_count": len(files),
        "read_errors": errors,
        "complete": not errors,
        "excluded_tool_paths": sorted(set(scopes)),
        "statistics": {
            "content_bytes_read": read_bytes,
            "capture_seconds": time.monotonic() - started,
        },
    }
    validate_snapshot(result)
    return result


def validate_snapshot(snapshot):
    """Reject unsupported, ambiguous, escaped or incomplete state schemas."""
    if (
        not isinstance(snapshot, dict)
        or snapshot.get("format") != FORMAT
        or snapshot.get("version") != 1
        or snapshot.get("policy") != POLICY
    ):
        raise ValueError(
            "Unsupported comparison state. Generate a new state with Folder Review."
        )
    root = snapshot.get("root")
    if not isinstance(root, str) or not root or not path_module(root).isabs(root):
        raise ValueError("Invalid state root.")
    module = path_module(root)

    def valid_parts(parts, allow_root=False):
        return (
            isinstance(parts, list)
            and (bool(parts) or allow_root)
            and all(
                isinstance(part, str)
                and part not in ("", ".", "..")
                and not any(c in part for c in (module.sep, module.altsep) if c)
                and not module.splitdrive(part)[0]
                for part in parts
            )
        )

    mode = snapshot.get("review_mode")
    if (
        mode not in changes.REVIEW_MODES
        or snapshot.get("hash_algorithm")
        != {"quick": "none", "sampled": "sha256-sample-v1", "full": "sha256"}[mode]
    ):
        raise ValueError("Invalid state review mode or algorithm.")
    scopes = snapshot.get("excluded_tool_paths")
    if (
        not isinstance(scopes, list)
        or len(scopes) > core.EXCEL_MAX_DATA_ROWS
        or any(
            not isinstance(scope, str)
            or not valid_parts(scope.replace("\\", module.sep).split(module.sep))
            or module.isabs(scope)
            for scope in scopes
        )
    ):
        raise ValueError("Invalid state exclusion scope.")
    files = snapshot.get("files")
    errors = snapshot.get("read_errors")
    if (
        not isinstance(files, list)
        or len(files) > core.EXCEL_MAX_DATA_ROWS
        or type(snapshot.get("file_count")) is not int
        or snapshot["file_count"] != len(files)
        or not isinstance(errors, list)
        or len(errors) > core.EXCEL_MAX_DATA_ROWS
        or type(snapshot.get("complete")) is not bool
        or snapshot["complete"] != (not errors)
    ):
        raise ValueError("Invalid state counts or coverage metadata.")
    seen = set()
    for record in files:
        if not isinstance(record, dict) or not valid_parts(record.get("parts")):
            raise ValueError("Invalid state file path.")
        path = module.join(root, *record["parts"])
        digest = record.get("sha256")
        if (
            record.get("path") != path
            or path in seen
            or type(record.get("size")) is not int
            or record["size"] < 0
            or type(record.get("mtime_ns")) is not int
            or (
                digest is not None
                if mode == "quick"
                else not isinstance(digest, str)
                or not re.fullmatch("[a-f0-9]{64}", digest)
            )
        ):
            raise ValueError("Invalid or duplicate state file record.")
        seen.add(path)
    for error in errors:
        if (
            not isinstance(error, dict)
            or error.get("kind") not in ("file", "directory", "unknown")
            or not valid_parts(error.get("parts"), error.get("kind") != "file")
            or error.get("path") != module.join(root, *error["parts"])
            or not isinstance(error.get("message"), str)
            or any(
                error.get(key) is not None and type(error[key]) is not int
                for key in ("errno", "winerror")
            )
        ):
            raise ValueError("Invalid state read issue.")
    return snapshot


def write_state(workbook, snapshot):
    validate_snapshot(snapshot)
    compact = {key: value for key, value in snapshot.items() if key != "files"}
    compact["files"] = [
        [r["parts"], r["size"], r["sha256"], r["mtime_ns"]] for r in snapshot["files"]
    ]
    data = json.dumps(compact, ensure_ascii=True, separators=(",", ":")).encode("ascii")
    if len(data) > MAX_JSON:
        raise ValueError("Comparison state exceeds the 256 MiB decoded limit.")
    packed = zlib.compress(data, 6)
    if len(packed) > MAX_PACKED:
        raise ValueError("Comparison state exceeds the 64 MiB compressed limit.")
    encoded = base64.b64encode(packed).decode("ascii")
    chunks = [encoded[i : i + CHUNK] for i in range(0, len(encoded), CHUNK)]
    sheet = workbook.add_worksheet(SHEET)
    sheet.hide()
    for row, value in enumerate(
        [
            MARKER,
            hashlib.sha256(data).hexdigest(),
            str(len(data)),
            str(len(chunks)),
            *chunks,
        ]
    ):
        if sheet.write_string(row, 0, value):
            raise ValueError("Comparison state cannot fit in the workbook.")


def read_xml(archive, name, limit):
    if archive.getinfo(name).file_size > limit:
        raise ValueError("Workbook XML exceeds the supported state size limit.")
    data = archive.read(name)
    text = data.decode("utf-8-sig")
    if "\x00" in text or "<!DOCTYPE" in text.upper() or "<!ENTITY" in text.upper():
        raise ValueError("XML declarations/entities are not allowed in a state.")
    return ET.fromstring(text)


def load_snapshot(path):
    """Read only XLSX package parts; never access its recorded project root."""
    try:
        with zipfile.ZipFile(core.extended_path(path)) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("Duplicate workbook package members.")
            book = read_xml(archive, "xl/workbook.xml", 4 * 1024 * 1024)
            sheets = [
                s
                for s in book.findall(NS + "sheets/" + NS + "sheet")
                if s.get("name") == SHEET
            ]
            if len(sheets) != 1:
                raise ValueError(
                    "This workbook has no comparison state. Generate new XLSX states with Folder Review 0.20 or later."
                )
            rid = sheets[0].get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            rels = read_xml(archive, "xl/_rels/workbook.xml.rels", 4 * 1024 * 1024)
            relation = [r for r in rels if r.get("Id") == rid]
            if len(relation) != 1 or relation[0].get("TargetMode") == "External":
                raise ValueError("Invalid comparison-state relationship.")
            target = relation[0].get("Target", "")
            name = posixpath.normpath(
                target.lstrip("/") if target.startswith("/") else "xl/" + target
            )
            if not name.startswith("xl/worksheets/") or ".." in name.split("/"):
                raise ValueError("Invalid comparison-state worksheet target.")
            sheet = read_xml(archive, name, MAX_XML)
            cells = list(sheet.iter(NS + "c"))
            values = []
            shared = set()
            for index, cell in enumerate(cells):
                if cell.get("r") != f"A{index + 1}" or cell.find(NS + "f") is not None:
                    raise ValueError("Invalid comparison-state cells.")
                kind = cell.get("t")
                if kind == "inlineStr":
                    values.append("".join(cell.itertext()))
                elif kind == "s":
                    value = int(cell.findtext(NS + "v", "-1"))
                    if value < 0:
                        raise ValueError("Invalid shared string index.")
                    shared.add(value)
                    values.append(value)
                else:
                    raise ValueError(
                        "Comparison-state cells must contain literal text."
                    )
            if shared:
                strings = read_xml(archive, "xl/sharedStrings.xml", 128 * 1024 * 1024)
                strings = ["".join(item.itertext()) for item in strings]
                values = [
                    strings[value] if type(value) is int else value for value in values
                ]
            if (
                len(values) < 5
                or values[0] != MARKER
                or not re.fullmatch("[a-f0-9]{64}", values[1])
            ):
                raise ValueError("Invalid comparison-state header.")
            length, count = int(values[2]), int(values[3])
            if (
                not 0 < length <= MAX_JSON
                or count != len(values) - 4
                or count > MAX_PACKED * 4 // 3 // CHUNK + 2
                or any(len(chunk) > CHUNK for chunk in values[4:])
            ):
                raise ValueError("Invalid comparison-state length.")
            packed = base64.b64decode("".join(values[4:]), validate=True)
            if len(packed) > MAX_PACKED:
                raise ValueError("Comparison state is too large.")
            decoder = zlib.decompressobj()
            data = decoder.decompress(packed, length + 1)
            if (
                len(data) != length
                or not decoder.eof
                or decoder.unused_data
                or decoder.unconsumed_tail
                or hashlib.sha256(data).hexdigest() != values[1]
            ):
                raise ValueError("Comparison state is truncated, changed or corrupt.")
            snapshot = json.loads(data)
            module = path_module(snapshot["root"])
            snapshot["files"] = [
                {
                    "parts": r[0],
                    "size": r[1],
                    "sha256": r[2],
                    "mtime_ns": r[3],
                    "path": module.join(snapshot["root"], *r[0]),
                }
                for r in snapshot["files"]
            ]
            return validate_snapshot(snapshot)
    except (
        KeyError,
        TypeError,
        IndexError,
        ET.ParseError,
        UnicodeError,
        zipfile.BadZipFile,
        zlib.error,
    ) as error:
        raise ValueError("Unreadable or unsupported XLSX comparison state.") from error
