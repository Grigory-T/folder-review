"""Change comparison/report logic migrated from folder-change-review (MIT).

The comparison and formatting rules are retained. Snapshot capture belongs to
Folder Review; this module reads its embedded states, never the live project.
"""

from __future__ import annotations

import errno
import hashlib
import ntpath
import os
import re
import stat
import time
from collections import Counter
from pathlib import Path

import folder_review as core

absolute = core.absolute_path
extended = core.extended_path

OUTPUT_FILE = "folder-change-review.xlsx"


MAX_DATA_ROWS = 1_048_570


HASH_BUFFER_SIZE = 256 * 1024


SAMPLE_SIZE = 64 * 1024


REVIEW_MODES = {
    "quick": "Quick review",
    "sampled": "Sampled contents",
    "full": "Full contents",
}


MODE_DETAILS = {
    "quick": "Size and modification time only; no file contents are read. Preserved timestamps can hide edits.",
    "sampled": "Size and first, middle and last 64 KiB; edits outside these samples can be missed.",
    "full": "SHA-256 of complete contents; reads every byte and can be slow over SMB.",
}


def review_mode(snapshot):
    return snapshot.get("review_mode", "full")


def comparison_mode(before, after):
    mode = review_mode(before)
    if mode != review_mode(after):
        raise ValueError(
            "These states use different review modes. Choose two states created with the same mode."
        )
    return mode


def report_error_message(error):
    if isinstance(error, PermissionError) or getattr(error, "winerror", None) in (
        5,
        32,
        33,
    ):
        path = getattr(error, "filename2", None) or getattr(error, "filename", None)
        location = f"\n\nPath: {absolute(path)}" if path else ""
        return (
            "The report could not be replaced. Close it in Excel and retry. "
            "Also check write permission and the SMB connection for the report folder."
            + location
            + "\n\nThe previous report is unchanged."
        )
    return str(error)


def ignored(name: str) -> bool:
    name = name.lower()
    return (
        name.startswith("~")
        or name in ("thumbs.db", "folder-change-review.exe", OUTPUT_FILE)
        or name.endswith((".tmp", ".temp", ".fcrstate"))
        or re.fullmatch(r"folder-changes-\d{8}-\d{6}-[0-9a-f]{6}\.xlsx", name)
        is not None
    )


def hash_file(path: str) -> tuple[str, int]:
    error = f"File changed while being hashed. Retry the snapshot: {path}"
    with open(extended(path), "rb") as stream:
        before = os.fstat(stream.fileno())
        digest = hashlib.sha256()
        remaining = before.st_size
        buffer = bytearray(min(HASH_BUFFER_SIZE, remaining))
        view = memoryview(buffer)
        # Never chase a growing file to EOF. Read its initial size, then probe
        # one byte for growth. A short positive read is valid, even over SMB.
        while remaining:
            count = stream.readinto(view[: min(len(view), remaining)])
            if not count:
                raise OSError(error)
            digest.update(view[:count])
            remaining -= count
        grew = bool(stream.read(1))
        after = os.fstat(stream.fileno())
    current = os.stat(extended(path), follow_symlinks=False)

    def identity(info):
        return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns

    metadata_changed = before.st_ctime_ns != after.st_ctime_ns
    # On Windows, fstat() can expose change time while stat() exposes creation
    # time as ctime. Compare ctime only through the same API there.
    if os.name != "nt":
        metadata_changed |= after.st_ctime_ns != current.st_ctime_ns
    if (
        grew
        or metadata_changed
        or identity(before) != identity(after)
        or identity(after) != identity(current)
    ):
        raise OSError(error)
    return digest.hexdigest(), after.st_size


def read_issue(path, parts, kind, error):
    message = str(getattr(error, "strerror", None) or error)
    if (
        os.name == "nt"
        and kind == "file"
        and getattr(error, "errno", None) == errno.EACCES
        and getattr(error, "winerror", None) is None
    ):
        # Python's CRT-backed open() may map both ACL denials and sharing
        # violations to errno 13 without exposing the original Win32 code.
        message = "Access denied or file locked by another application."
    return {
        "path": path,
        "parts": list(parts),
        "kind": kind,
        "errno": getattr(error, "errno", None),
        "winerror": getattr(error, "winerror", None),
        "message": message,
    }


def sample_file(path, initial):
    """Bounded content sampling; never represented as a full-file fingerprint."""
    size = initial.st_size
    offsets = sorted({0, max(0, (size - SAMPLE_SIZE) // 2), max(0, size - SAMPLE_SIZE)})
    # Merge overlapping ranges; small files are read completely, once.
    ranges = []
    for offset in offsets:
        end = min(size, offset + SAMPLE_SIZE)
        if ranges and offset <= ranges[-1][1]:
            ranges[-1] = (ranges[-1][0], max(end, ranges[-1][1]))
        else:
            ranges.append((offset, end))
    digest = hashlib.sha256(
        b"folder-change-review.sample.v1\0" + str(size).encode() + b"\0"
    )
    total = 0
    with open(extended(path), "rb", buffering=0) as stream:
        for start, end in ranges:
            digest.update(start.to_bytes(8, "little"))
            stream.seek(start)
            remaining = end - start
            while remaining:
                data = stream.read(remaining)
                if not data:
                    raise OSError("File changed while being sampled. Retry the state.")
                digest.update(data)
                total += len(data)
                remaining -= len(data)
        # Do not use fstat() twice per sampled file: only the metadata required
        # by this approximate mode is checked. Full mode keeps its stronger checks.
        current = os.stat(extended(path), follow_symlinks=False)
        if (current.st_size, current.st_mtime_ns) != (size, initial.st_mtime_ns):
            raise OSError("File changed while being sampled. Retry the state.")
        if getattr(current, "st_file_attributes", 0) & 0x400 or stat.S_ISLNK(
            current.st_mode
        ):
            raise OSError("File became a link while being sampled. Retry the state.")
    return digest.hexdigest(), size, total


def compare_snapshots(before: dict, after: dict) -> list[dict]:
    mode = comparison_mode(before, after)
    excluded_scopes = []
    for snapshot in (before, after):
        root = snapshot.get("root", "")
        module = (
            ntpath if ntpath.splitdrive(root)[0] or root.startswith("\\\\") else os.path
        )
        for path in snapshot.get("excluded_tool_paths", []):
            excluded_scopes.append(
                (module, module.normcase(module.normpath(module.join(root, path))))
            )

    def tool_artifact(path):
        for module, scope in excluded_scopes:
            candidate = module.normcase(module.normpath(path))
            if candidate == scope or candidate.startswith(scope + module.sep):
                return True
        return False

    # Coverage lookup is O(path depth), not O(files * read errors).
    def inventory(snapshot):
        # Apply reserved-artifact exclusions to old inventories too, so a policy
        # change does not report the tool's former outputs as project deletions.
        records = {
            record["path"]: record
            for record in snapshot["files"]
            if not tool_artifact(record["path"])
            and not ignored(record["parts"][-1])
            and ".states" not in (part.casefold() for part in record["parts"])
        }
        scopes = set()
        root = snapshot.get("root", "")
        module = (
            ntpath if ntpath.splitdrive(root)[0] or root.startswith("\\\\") else os.path
        )
        for issue in snapshot.get("read_errors", []):
            if tool_artifact(issue["path"]):
                continue
            if issue["parts"] and (
                (issue["kind"] == "file" and ignored(issue["parts"][-1]))
                or ".states" in (part.casefold() for part in issue["parts"])
            ):
                continue
            if issue["kind"] == "file":
                records.setdefault(
                    issue["path"],
                    {
                        "path": issue["path"],
                        "parts": issue["parts"],
                        "size": None,
                        "sha256": None,
                        "unavailable": True,
                    },
                )
            else:
                scopes.add(module.normpath(issue["path"]))

        def unavailable(path, record):
            if record is not None:
                return record.get("unavailable", False)
            current = module.normpath(path)
            while True:
                if current in scopes:
                    return True
                parent = module.dirname(current)
                if parent == current:
                    return False
                current = parent

        return records, unavailable

    old, old_unavailable = inventory(before)
    new, new_unavailable = inventory(after)
    rows = []
    for path in sorted(old.keys() | new.keys()):
        a, b = old.get(path), new.get(path)
        if old_unavailable(path, a) or new_unavailable(path, b):
            status = "Unknown"
        elif a is None:
            status = "Added"
        elif b is None:
            status = "Deleted"
        elif a["size"] != b["size"] or (
            (a.get("mtime_ns") != b.get("mtime_ns"))
            if mode == "quick"
            else a["sha256"] != b["sha256"]
        ):
            status = "Changed"
        else:
            status = "Unchanged"
        record = b if b is not None else a
        filename = record["parts"][-1]
        rows.append(
            {
                "parts": record["parts"],
                "FullPath": path,
                "FileName": filename,
                "Extension": os.path.splitext(filename)[1].lstrip(".").lower(),
                "Status": status,
                "Added": 1 if status == "Added" else None,
                "Deleted": 1 if status == "Deleted" else None,
                "Changed": 1 if status == "Changed" else None,
                "SizeA": None if a is None else a["size"],
                "SizeB": None if b is None else b["size"],
                "SizeDelta": None
                if status == "Unknown"
                else (b["size"] if b else 0) - (a["size"] if a else 0),
            }
        )
    return rows


def report_hierarchy(
    rows: list[dict], root_hint: str = ""
) -> tuple[str, list[list[str]]]:
    """LCA of union file parents; single-file LCA is its folder, never the file."""
    if not rows:
        return root_hint, []
    paths = [row["FullPath"] for row in rows]
    path_module = (
        ntpath
        if ntpath.splitdrive(paths[0])[0] or paths[0].startswith("\\\\")
        else os.path
    )
    try:
        ancestor = path_module.commonpath([path_module.dirname(path) for path in paths])
    except ValueError:
        ancestor = ""  # roots are intentionally not validated
    if ancestor:
        parts = [
            path_module.relpath(path, ancestor).split(path_module.sep) for path in paths
        ]
    else:
        parts = [list(row["parts"]) for row in rows]
    return ancestor, parts


def write_report(
    rows: list[dict],
    output: str | Path,
    root_hint: str = "",
    read_errors=(),
    issue_roots=None,
    mode=None,
) -> None:
    if len(rows) > MAX_DATA_ROWS:
        raise ValueError(
            f"Union has {len(rows):,} files; one worksheet allows {MAX_DATA_ROWS:,} data rows."
        )
    ancestor, paths = report_hierarchy(rows, root_hint)
    depth = max((len(parts) for parts in paths), default=1)
    fields = [
        "Leaf depth",
        "Name without extension",
        "File name",
        "Extension",
        "Size A (bytes)",
        "Size A (MiB)",
        "Size B (bytes)",
        "Size B (MiB)",
        "Size delta (bytes)",
        "Size delta (MiB)",
        "Status",
        "Added",
        "Deleted",
        "Changed",
    ]
    headers = [
        "Row number",
        *[f"Level {index}" for index in range(1, depth + 1)],
        *fields,
        "last",
    ]
    if len(headers) + 1 > 16_384:  # column A is blank
        raise ValueError("Folder hierarchy exceeds Excel's column limit.")
    with core._report_workbook(Path(output)) as workbook:
        workbook.set_calc_mode("auto", 191029)
        workbook.set_size(1484, 843)
        workbook.set_tab_ratio(92.8)
        sheet = workbook.add_worksheet("Changes")
        sheet.hide_gridlines(2)
        sheet.set_zoom(70)
        sheet.set_tab_color("#002060")
        sheet.outline_settings(True, False, False, False)
        sheet.set_default_row(14.5)
        sheet.set_portrait()
        sheet.set_paper(9)
        sheet.set_margins(0.7, 0.75, 0.75, 0.75)
        sheet.set_header("", {"margin": 0.3})
        sheet.set_footer("", {"margin": 0.3})
        font = {"font_name": "Calibri", "font_size": 11, "font_color": "#000000"}
        border = {**font, "border": 1, "border_color": "#000000"}
        body = workbook.add_format(border)
        centered = workbook.add_format(
            {**border, "align": "center", "valign": "vcenter"}
        )
        size_format = workbook.add_format({**border, "num_format": "#,##0.000"})
        bytes_format = workbook.add_format({**border, "num_format": "#,##0"})
        title = workbook.add_format({**font, "bold": True})
        header_format = {
            **border,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
        header = workbook.add_format({**header_format, "bg_color": "#00B0F0"})
        boundary = workbook.add_format(
            {**header_format, "bg_color": "#000000", "font_color": "#FFFFFF"}
        )
        sheet.set_column(0, len(headers), 5)
        sheet.freeze_panes("A7")
        # Match the accepted Folder Review layout, including its taller header.
        sheet.set_row(5, 72)
        sheet.write_string(
            1, 1, f"Table: File changes in folder {ancestor or root_hint}", title
        )
        if read_errors:
            sheet.write_string(
                2,
                1,
                "WARNING: Incomplete coverage. Unknown is not Added, Deleted or Changed. See Read issues.",
                title,
            )
        for column, text in enumerate(headers, 1):
            sheet.write_number(3, column, column, centered)
        if mode is not None:
            sheet.write_string(
                4, 1, f"Review mode: {REVIEW_MODES[mode]}. {MODE_DETAILS[mode]}", title
            )
        for column, text in enumerate(headers, 1):
            sheet.write_string(
                5, column, text, boundary if column in (1, len(headers)) else header
            )
        for index, row in enumerate(rows, 1):
            parts = paths[index - 1]
            values = {
                "Row number": index,
                "Leaf depth": len(parts),
                "Name without extension": ntpath.splitext(row["FileName"])[0],
                "File name": row["FileName"],
                "Extension": row["Extension"],
                "Status": row["Status"],
                "Added": row["Added"],
                "Deleted": row["Deleted"],
                "Changed": row["Changed"],
                "last": 1,
                **{f"Level {i}": part for i, part in enumerate(parts, 1)},
            }
            for label, key in (("A", "SizeA"), ("B", "SizeB"), ("delta", "SizeDelta")):
                amount = row.get(key)
                values[f"Size {label} (bytes)"] = amount
                values[f"Size {label} (MiB)"] = (
                    None if amount is None else amount / 1_048_576
                )
            for column, label in enumerate(headers, 1):
                value = values.get(label)
                fmt = (
                    size_format
                    if label.endswith("(MiB)")
                    else bytes_format
                    if label.endswith("(bytes)")
                    else centered
                    if label
                    in (
                        "Row number",
                        "last",
                        "Leaf depth",
                        "Extension",
                        "Status",
                        "Added",
                        "Deleted",
                        "Changed",
                    )
                    else body
                )
                if isinstance(value, str):
                    result = sheet.write_string(index + 5, column, value, fmt)
                    if result != 0:
                        raise ValueError(
                            f"Text cannot fit in an Excel cell: {row['FullPath']}"
                        )
                elif value is not None:
                    sheet.write_number(index + 5, column, value, fmt)
                else:
                    sheet.write_blank(index + 5, column, None, fmt)
        sheet.autofilter(5, 1, 5 + len(rows), len(headers))
        sheet.set_selection("B6")
        if read_errors:
            issues_sheet = workbook.add_worksheet("Read issues")
            issues_sheet.set_column(0, 0, 10)
            issues_sheet.set_column(1, 1, 15)
            issues_sheet.set_column(2, 2, 80)
            issues_sheet.set_column(3, 3, 12)
            issues_sheet.set_column(4, 4, 65)
            roots = issue_roots or {}
            issues_sheet.write_string(0, 0, "Read issues", title)
            issues_sheet.write_string(1, 0, "State A root", body)
            issues_sheet.write_string(1, 1, roots.get("A", root_hint), body)
            issues_sheet.write_string(2, 0, "State B root", body)
            issues_sheet.write_string(2, 1, roots.get("B", root_hint), body)
            issues_sheet.write_row(
                4,
                0,
                ["State", "Item type", "Relative path", "OS error", "Reason"],
                header,
            )
            for index, (label, issue) in enumerate(read_errors, 5):
                root = roots.get(label, root_hint)
                module = ntpath if ntpath.splitdrive(root)[0] else os.path
                try:
                    relative = (
                        module.relpath(issue["path"], root) if root else issue["path"]
                    )
                except ValueError:
                    relative = issue["path"]
                values = [
                    label,
                    issue["kind"],
                    relative,
                    str(
                        issue["winerror"]
                        if issue["winerror"] is not None
                        else issue["errno"] or ""
                    ),
                    issue["message"],
                ]
                for column, value in enumerate(values):
                    if issues_sheet.write_string(index, column, value, body) != 0:
                        raise ValueError("Read issue text cannot fit in an Excel cell.")
            issues_sheet.autofilter(4, 0, 4 + len(read_errors), 4)
            issues_sheet.freeze_panes(5, 0)


def compare_files(before_path, after_path, output, progress=None, diagnostics=True):
    """Compare saved workbooks offline; the recorded project paths stay lexical."""
    from review_state import load_snapshot

    emit = progress or print
    output = Path(core.absolute_path(output))
    inputs = [
        core.absolute_path(core.pasted_path(str(path)))
        for path in (before_path, after_path)
    ]
    if output.suffix.lower() != ".xlsx" or core.resolved_path(
        output.parent
    ) != core.resolved_path(core.application_directory()):
        raise ValueError(
            "The comparison report must be an .xlsx file beside the application."
        )
    for path in inputs:
        core.validate_storage_transport(path)
        if core.normalized_key(path) == core.normalized_key(output) or (
            os.path.exists(core.extended_path(output))
            and os.path.samefile(core.extended_path(path), core.extended_path(output))
        ):
            raise ValueError("The comparison output must not replace an input state.")
    args = core.parse_args([])
    args.workers = 1
    args.duplicates = "off"
    args.no_log = not diagnostics
    started = time.monotonic()
    with (
        core.safe_runtime(str(core.application_directory()), output),
        core.performance_log(args, str(core.application_directory()), progress) as log,
    ):
        if log is not None:
            log.set_phase("state_load")
        before, after = [load_snapshot(path) for path in inputs]
        if log is not None:
            log.set_phase("comparison")
        rows = compare_snapshots(before, after)
        errors = [
            (label, issue)
            for label, state in (("A", before), ("B", after))
            for issue in state["read_errors"]
        ]
        if len(errors) > 1_048_571:
            raise ValueError("Read issues exceed Excel's worksheet row limit.")
        if log is not None:
            log.set_phase("report")
        emit("Writing the comparison workbook...")
        write_report(
            rows,
            output,
            after["root"],
            errors,
            {"A": before["root"], "B": after["root"]},
            comparison_mode(before, after),
        )
        counts = Counter(row["Status"] for row in rows)
        counts.partial = bool(errors)
        code = 2 if errors else 0
        if log is not None:
            for status in ("Added", "Deleted", "Changed", "Unchanged", "Unknown"):
                log.count(status.lower(), counts[status])
            log.exit_code = code
        emit(
            " | ".join(
                f"{status}: {counts[status]:,}"
                for status in ("Added", "Deleted", "Changed", "Unchanged", "Unknown")
            )
        )
        emit(
            f"Comparison status: {'PARTIAL' if code else 'COMPLETE'} | {time.monotonic() - started:.1f}s"
        )
        emit(f"Result: {output}")
        return counts


def run_settings(path, progress=None):
    parser = core.read_ini(
        Path(path), "compare", {"before", "after", "report", "diagnostics"}
    )
    values = parser["compare"]
    inputs = [core.pasted_path(values.get(key, "")) for key in ("before", "after")]
    if not all(inputs) or any(not os.path.isabs(p) for p in inputs):
        raise ValueError(
            "Set before and after to two full XLSX state paths in folder-change-review.ini."
        )
    name = core.report_filename(values.get("report", OUTPUT_FILE))
    counts = compare_files(
        *inputs,
        core.application_directory() / name,
        progress,
        values.getboolean("diagnostics", True),
    )
    return 2 if counts.partial else 0


def main():
    try:
        core.validate_platform()
        if not core.local_storage(Path(__file__).parent):
            raise ValueError("Copy the tool to a local folder before running it.")
        import sys

        if sys.argv[1:]:
            raise ValueError(
                "Edit folder-change-review.ini; this entry point accepts no arguments."
            )
        return run_settings(core.application_directory() / "folder-change-review.ini")
    except Exception as error:  # noqa: BLE001 -- console boundary preserves existing reports
        print("ERROR: " + report_error_message(error), flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
