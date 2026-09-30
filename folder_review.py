"""Create one flat-table XLSX inventory for a manually selected folder."""

# ruff: noqa: BLE001, S110 -- archive libraries expose broad failure modes;
# each failure is converted into an inventory error instead of stopping the scan.

from __future__ import annotations

import argparse
import ntpath
import os
import re
import shutil
import stat
import sys
import tempfile
import time
import zipfile
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from threading import Semaphore

import olefile
import py7zr
import rarfile
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from py7zr import Py7zIO, WriterFactory

# User configuration ---------------------------------------------------------

ROOT_FOLDER = r"PASTE_FOLDER_PATH_HERE"

OUTPUT_FILE = "folder-review.xlsx"


# Runtime limits -------------------------------------------------------------

DEFAULT_WORKERS = 16
MAX_CONCURRENT_ARCHIVES = 4

MAGIC_READ_SIZE = 8 * 1024
MAX_ARCHIVE_DEPTH = 20
MAX_NESTED_ARCHIVE_SIZE = 50 * 1024 * 1024 * 1024
SEVEN_ZIP_INSPECT_MAX_MEMBERS = 25_000
EXCEL_MAX_DATA_ROWS = 1_048_570  # rows 7..1,048,576
HEAVY_ARCHIVE_SLOTS = Semaphore(MAX_CONCURRENT_ARCHIVES)
ARCHIVE_TYPES = {"ZIP archive", "7z archive", "RAR archive"}
STREAM_ARCHIVE_TYPES = {
    "ZIP container": "ZIP archive",
    "7z archive": "7z archive",
    "RAR archive": "RAR archive",
}


MAGIC_SIGNATURES = (
    (b"\x37\x7a\xbc\xaf\x27\x1c", "7z archive"),
    (b"Rar!\x1a\x07\x01\x00", "RAR archive"),
    (b"Rar!\x1a\x07\x00", "RAR archive"),
    (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "OLE2 compound file"),
    (b"SQLite format 3\x00", "SQLite DB"),
    (b"\x00\x01\x00\x00Standard Jet DB", "Microsoft Access / Jet DB"),
    (b"\x00\x01\x00\x08\x00\x01\x00\x01\x01", "Microsoft Access MDB"),
    (b"%PDF-", "PDF"),
    (b"{\\rtf", "RTF document"),
    (b"\x89PNG\r\n\x1a\n", "PNG image"),
    (b"\xff\xd8\xff", "JPEG image"),
    (b"GIF87a", "GIF image"),
    (b"GIF89a", "GIF image"),
    (b"II*\x00", "TIFF image"),
    (b"MM\x00*", "TIFF image"),
    (b"BM", "BMP image"),
    (b"\x1f\x8b", "GZIP archive"),
    (b"BZh", "BZIP2 archive"),
    (b"\xfd7zXZ\x00", "XZ archive"),
    (b"\x28\xb5\x2f\xfd", "Zstandard compressed file"),
    (b"\x7fELF", "ELF executable"),
    (b"MZ", "Windows executable / DLL"),
    (b"ID3", "MP3 audio"),
    (b"OggS", "Ogg media"),
    (b"fLaC", "FLAC audio"),
    (b"PAR1", "Apache Parquet"),
    (b"-----BEGIN PGP SIGNATURE-----", "PGP signature (ASCII armored)"),
    (b"-----BEGIN PKCS7-----", "PKCS#7 signature (PEM)"),
    (b"-----BEGIN CMS-----", "CMS signature (PEM)"),
)

ZIP_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")

BOM_SIGNATURES = (
    (b"\xef\xbb\xbf", "UTF-8"),
    (b"\xff\xfe\x00\x00", "UTF-32 LE"),
    (b"\x00\x00\xfe\xff", "UTF-32 BE"),
    (b"\xff\xfe", "UTF-16 LE"),
    (b"\xfe\xff", "UTF-16 BE"),
)
BOM_NOT_PRESENT = "нет"


# Data models ----------------------------------------------------------------


@dataclass(slots=True, frozen=True)
class FileProperties:
    hidden: str
    mode: int
    inode: str
    device: str
    links: int
    user_id: str | None
    group_id: str | None
    accessed_at: datetime | None
    modified_at: datetime | None
    created_at: datetime | None


@dataclass(slots=True, frozen=True)
class PhysicalFile:
    path: str
    levels: tuple[str, ...]
    size: int
    properties: FileProperties | None = None


@dataclass(slots=True, frozen=True)
class ArchiveNode:
    source_path: str
    logical_path: str
    levels: tuple[str, ...]
    size: int
    depth: int
    properties: FileProperties | None = None


@dataclass(slots=True, frozen=True)
class ResultRow:
    levels: tuple[str, ...]
    full_path: str
    leaf_name: str
    extension: str
    size: int
    content_type: str
    in_archive: str
    bom: str | None = None
    properties: FileProperties | None = None
    error: str | None = None


@dataclass(slots=True)
class ScanStats:
    folders: int = 0
    skipped_folders: int = 0
    skipped_entries: int = 0


# Path and row helpers --------------------------------------------------------


def extended_path(path: str) -> str:
    """Return a Windows extended path while leaving other platforms unchanged."""
    if os.name != "nt" or path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + os.path.abspath(path)


def normalized_key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def split_archive_path(name: str) -> tuple[str, ...]:
    return tuple(part for part in re.split(r"[\\/]+", name) if part not in ("", "."))


def extension_of(name: str) -> str:
    return ntpath.splitext(name)[1].lstrip(".").lower()


def timestamp_value(value: float | None) -> datetime | None:
    if value is None:
        return None
    try:
        return (
            datetime.fromtimestamp(value, tz=UTC)
            .astimezone()
            .replace(tzinfo=None, microsecond=0)
        )
    except (OSError, OverflowError, ValueError):
        return None


def file_properties(name: str, result: os.stat_result) -> FileProperties:
    if os.name == "nt":
        hidden_mask = getattr(stat, "FILE_ATTRIBUTE_HIDDEN", 0x2)
        hidden = bool(getattr(result, "st_file_attributes", 0) & hidden_mask)
    else:
        hidden = name.startswith(".")

    created_timestamp = getattr(result, "st_birthtime", None)
    if created_timestamp is None and os.name == "nt":
        created_timestamp = result.st_ctime

    user_id = getattr(result, "st_uid", None)
    group_id = getattr(result, "st_gid", None)
    return FileProperties(
        hidden="да" if hidden else "нет",
        mode=result.st_mode,
        inode=str(result.st_ino),
        device=str(result.st_dev),
        links=result.st_nlink,
        user_id=None if user_id is None else str(user_id),
        group_id=None if group_id is None else str(group_id),
        accessed_at=timestamp_value(result.st_atime),
        modified_at=timestamp_value(result.st_mtime),
        created_at=timestamp_value(created_timestamp),
    )


def make_row(
    full_path: str,
    levels: tuple[str, ...],
    size: int,
    content_type: str,
    in_archive: bool,
    error: str | None = None,
    *,
    bom: str | None = None,
    properties: FileProperties | None = None,
) -> ResultRow:
    leaf_name = levels[-1]
    return ResultRow(
        levels=levels,
        full_path=full_path,
        leaf_name=leaf_name,
        extension=extension_of(leaf_name),
        size=size,
        content_type=content_type,
        in_archive="да" if in_archive else "нет",
        bom=bom,
        properties=properties,
        error=error,
    )


# Content detection ----------------------------------------------------------


def is_zip_signature(data: bytes) -> bool:
    return any(data.startswith(signature) for signature in ZIP_SIGNATURES)


def detect_bom(data: bytes) -> str | None:
    for signature, name in BOM_SIGNATURES:
        if data.startswith(signature):
            return name
    return None


def detect_magic(data: bytes) -> str:
    if not data:
        return "Empty file"

    bom = detect_bom(data)
    if bom is not None:
        return f"Text ({bom} BOM)"

    if is_zip_signature(data):
        return "ZIP container"

    for signature, file_type in MAGIC_SIGNATURES:
        if data.startswith(signature):
            return file_type

    if data.startswith(b"RIFF") and len(data) >= 12:
        return {
            b"WAVE": "WAV audio",
            b"AVI ": "AVI video",
            b"WEBP": "WebP image",
        }.get(data[8:12], "RIFF container")

    if len(data) >= 12 and data[4:8] == b"ftyp":
        return "MP4 / QuickTime container"

    if len(data) >= 262 and data[257:262] == b"ustar":
        return "TAR archive"

    sample = data[:8192]
    if b"\x00" not in sample:
        control_bytes = sum(
            byte < 32 and byte not in (8, 9, 10, 12, 13) for byte in sample
        )
        if control_bytes / max(len(sample), 1) < 0.01:
            try:
                sample.decode("utf-8")
                return "Text (UTF-8 / ASCII)"
            except UnicodeDecodeError:
                return "Text (unknown single-byte encoding)"

    return "Unknown binary"


def classify_zip_names(names: list[str]) -> str:
    normalized = [name.replace("\\", "/").lower() for name in names]

    if any(name.startswith("word/") for name in normalized):
        return (
            "Word document (DOCM)"
            if "word/vbaproject.bin" in normalized
            else "Word document (DOCX)"
        )

    if any(name.startswith("xl/") for name in normalized):
        if "xl/workbook.bin" in normalized:
            return "Excel binary workbook (XLSB)"
        if "xl/vbaproject.bin" in normalized:
            return "Excel workbook (XLSM)"
        return "Excel workbook (XLSX)"

    if any(name.startswith("ppt/") for name in normalized):
        return (
            "PowerPoint presentation (PPTM)"
            if "ppt/vbaproject.bin" in normalized
            else "PowerPoint presentation (PPTX)"
        )

    if "meta-inf/manifest.mf" in normalized:
        return "JAR archive"
    if "meta-inf/container.xml" in normalized and "mimetype" in normalized:
        return "EPUB document"
    return "ZIP archive"


def classify_zip_path(file_path: str) -> str:
    try:
        with zipfile.ZipFile(extended_path(file_path)) as archive:
            return classify_zip_names(archive.namelist())
    except (OSError, zipfile.BadZipFile, RuntimeError):
        return "ZIP container (unreadable)"


def classify_ole_path(file_path: str) -> str:
    try:
        with olefile.OleFileIO(extended_path(file_path)) as ole:
            streams = {
                "/".join(parts).lower()
                for parts in ole.listdir(streams=True, storages=False)
            }
    except Exception:
        return "OLE2 compound file"

    if "encryptedpackage" in streams and "encryptioninfo" in streams:
        return "Encrypted Microsoft Office OOXML package"
    if "worddocument" in streams:
        return "Word document (DOC)"
    if "workbook" in streams or "book" in streams:
        return "Excel workbook (XLS)"
    if "powerpoint document" in streams:
        return "PowerPoint presentation (PPT)"
    if "__properties_version1.0" in streams:
        return "Outlook message (MSG)"
    return "OLE2 compound file"


def inspect_file(file_path: str) -> tuple[str, str | None]:
    try:
        with open(extended_path(file_path), "rb", buffering=0) as file:
            head = file.read(MAGIC_READ_SIZE)
    except OSError:
        return "Unreadable", None

    file_type = detect_magic(head)
    if file_type == "ZIP container":
        file_type = classify_zip_path(file_path)
    elif file_type == "OLE2 compound file":
        file_type = classify_ole_path(file_path)
    return file_type, detect_bom(head) or BOM_NOT_PRESENT


def detect_file_type(file_path: str) -> str:
    return inspect_file(file_path)[0]


# Physical folder scan -------------------------------------------------------


def scan_folder(
    folder_path: str,
    levels: tuple[str, ...],
    excluded_paths: set[str],
) -> tuple[list[tuple[str, tuple[str, ...]]], list[PhysicalFile], int, bool]:
    subfolders: list[tuple[str, tuple[str, ...]]] = []
    files: list[PhysicalFile] = []
    skipped_entries = 0

    try:
        with os.scandir(extended_path(folder_path)) as entries:
            for entry in entries:
                try:
                    logical_path = os.path.join(folder_path, entry.name)
                    if entry.is_dir(follow_symlinks=False):
                        subfolders.append((logical_path, levels + (entry.name,)))
                    elif entry.is_file(follow_symlinks=False):
                        if normalized_key(logical_path) in excluded_paths:
                            continue
                        stat_result = entry.stat(follow_symlinks=False)
                        files.append(
                            PhysicalFile(
                                logical_path,
                                levels + (entry.name,),
                                stat_result.st_size,
                                file_properties(entry.name, stat_result),
                            )
                        )
                except OSError:
                    skipped_entries += 1
    except OSError:
        return [], [], skipped_entries, True

    return subfolders, files, skipped_entries, False


def scan_tree(
    root_folder: str,
    excluded_paths: set[str],
    workers: int = DEFAULT_WORKERS,
) -> tuple[list[PhysicalFile], ScanStats]:
    all_files: list[PhysicalFile] = []
    stats = ScanStats()
    last_print = time.monotonic()

    with ThreadPoolExecutor(max_workers=workers) as executor:
        pending = {executor.submit(scan_folder, root_folder, (), excluded_paths)}

        while pending:
            done, pending = wait(pending, return_when=FIRST_COMPLETED)
            for future in done:
                subfolders, files, skipped_entries, skipped_folder = future.result()
                stats.folders += 1
                stats.skipped_folders += int(skipped_folder)
                stats.skipped_entries += skipped_entries
                all_files.extend(files)
                for subfolder_path, levels in subfolders:
                    pending.add(
                        executor.submit(
                            scan_folder, subfolder_path, levels, excluded_paths
                        )
                    )

            now = time.monotonic()
            if now - last_print >= 10:
                print(
                    f"Folders: {stats.folders:,} | files: {len(all_files):,} | "
                    f"skipped folders: {stats.skipped_folders:,}"
                )
                last_print = now

    all_files.sort(key=lambda item: tuple(part.casefold() for part in item.levels))
    return all_files, stats


# Recursive archive inspection ----------------------------------------------


def archive_row(
    node: ArchiveNode,
    file_type: str,
    error: str | None = None,
    bom: str | None = None,
) -> ResultRow:
    return make_row(
        node.logical_path,
        node.levels,
        node.size,
        file_type,
        node.depth > 0,
        error,
        bom=bom,
        properties=node.properties,
    )


def make_temp_path(temp_dir: str, logical_name: str) -> str:
    suffix = ntpath.splitext(logical_name)[1]
    if not suffix or len(suffix) > 16:
        suffix = ".bin"
    descriptor, path = tempfile.mkstemp(prefix="member-", suffix=suffix, dir=temp_dir)
    os.close(descriptor)
    return path


def raw_archive_type(data: bytes) -> str | None:
    return STREAM_ARCHIVE_TYPES.get(detect_magic(data))


def inspect_member_stream(
    stream,
    *,
    size: int,
    logical_path: str,
    levels: tuple[str, ...],
    parent_depth: int,
    temp_dir: str,
) -> list[ResultRow]:
    try:
        head = stream.read(MAGIC_READ_SIZE)
    except Exception as exc:
        return [
            make_row(
                logical_path,
                levels,
                size,
                "Unreadable archive member",
                True,
                type(exc).__name__,
            )
        ]

    archive_type = raw_archive_type(head)
    if archive_type is None:
        return [
            make_row(
                logical_path,
                levels,
                size,
                detect_magic(head),
                True,
                bom=detect_bom(head) or BOM_NOT_PRESENT,
            )
        ]

    depth = parent_depth + 1
    if depth >= MAX_ARCHIVE_DEPTH:
        return [
            make_row(
                logical_path,
                levels,
                size,
                archive_type,
                True,
                f"Archive nesting limit reached ({MAX_ARCHIVE_DEPTH})",
            )
        ]
    if size > MAX_NESTED_ARCHIVE_SIZE:
        return [
            make_row(
                logical_path,
                levels,
                size,
                archive_type,
                True,
                "Nested archive exceeds 50 GiB safety limit",
            )
        ]

    temp_path = make_temp_path(temp_dir, levels[-1])
    try:
        with open(temp_path, "wb") as output:
            output.write(head)
            shutil.copyfileobj(stream, output, length=1024 * 1024)
        node = ArchiveNode(temp_path, logical_path, levels, size, depth)
        return process_archive_node(node, temp_dir)
    except Exception as exc:
        return [
            make_row(
                logical_path,
                levels,
                size,
                archive_type,
                True,
                type(exc).__name__,
            )
        ]
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass


def expand_zip(node: ArchiveNode, temp_dir: str) -> list[ResultRow]:
    rows: list[ResultRow] = []
    try:
        with zipfile.ZipFile(extended_path(node.source_path)) as archive:
            infos = [info for info in archive.infolist() if not info.is_dir()]
            if not infos:
                return [archive_row(node, "ZIP archive (empty)")]

            for info in infos:
                parts = split_archive_path(info.filename)
                if not parts:
                    continue
                levels = node.levels + parts
                logical_path = node.logical_path + "::" + "\\".join(parts)
                try:
                    with archive.open(info) as member:
                        rows.extend(
                            inspect_member_stream(
                                member,
                                size=info.file_size,
                                logical_path=logical_path,
                                levels=levels,
                                parent_depth=node.depth,
                                temp_dir=temp_dir,
                            )
                        )
                except Exception as exc:
                    rows.append(
                        make_row(
                            logical_path,
                            levels,
                            info.file_size,
                            "Unreadable archive member",
                            True,
                            type(exc).__name__,
                        )
                    )
    except Exception as exc:
        return [archive_row(node, "ZIP archive", type(exc).__name__)]
    return rows


class MemberCaptureIO(Py7zIO):
    def __init__(
        self,
        expected_size: int,
        filename: str,
        temp_dir: str,
    ):
        self.expected_size = expected_size or 0
        self.filename = filename
        self.temp_dir = temp_dir
        self.head = bytearray()
        self.total_written = 0
        self.mode: str | None = None
        self.archive_type: str | None = None
        self.temp_path: str | None = None
        self.output = None
        self.oversized = False

    def write(self, chunk):
        raw = bytes(chunk)
        if self.mode == "archive":
            self.output.write(raw)
            self.total_written += len(raw)
            return len(raw)
        if self.mode == "head":
            self.total_written += len(raw)
            return len(raw)

        take = min(len(raw), max(MAGIC_READ_SIZE - len(self.head), 0))
        self.head.extend(raw[:take])
        remainder = raw[take:]
        total_after = self.total_written + len(raw)
        ready = (
            len(self.head) >= 8
            or len(self.head) >= MAGIC_READ_SIZE
            or total_after >= self.expected_size
        )

        if ready:
            self.archive_type = raw_archive_type(bytes(self.head))
            if self.archive_type and self.expected_size <= MAX_NESTED_ARCHIVE_SIZE:
                self.temp_path = make_temp_path(self.temp_dir, self.filename)
                self.output = open(self.temp_path, "wb")  # noqa: SIM115
                self.output.write(self.head)
                self.output.write(remainder)
                self.mode = "archive"
            else:
                self.mode = "head"
                self.oversized = bool(
                    self.archive_type and self.expected_size > MAX_NESTED_ARCHIVE_SIZE
                )

        self.total_written = total_after
        return len(raw)

    def close_capture(self) -> None:
        if self.output is not None:
            self.output.close()
            self.output = None

    def read(self, size=None):
        return bytes(self.head if size is None else self.head[:size])

    def seek(self, offset, whence=0):
        return 0

    def flush(self):
        if self.output is not None:
            self.output.flush()

    def size(self):
        return self.total_written


class MemberCaptureFactory(WriterFactory):
    def __init__(self, sizes: dict[str, int], temp_dir: str):
        self.sizes = sizes
        self.temp_dir = temp_dir
        self.products: dict[str, MemberCaptureIO] = {}

    def create(self, filename):
        product = MemberCaptureIO(
            self.sizes.get(filename, 0),
            filename,
            self.temp_dir,
        )
        self.products[filename] = product
        return product

    def close_all(self) -> None:
        for product in self.products.values():
            product.close_capture()


def seven_zip_limit_rows(
    node: ArchiveNode,
    infos,
    reason: str,
) -> list[ResultRow]:
    rows = []
    for info in infos:
        parts = split_archive_path(info.filename)
        if not parts:
            continue
        rows.append(
            make_row(
                node.logical_path + "::" + "\\".join(parts),
                node.levels + parts,
                info.uncompressed or 0,
                "Not inspected",
                True,
                reason,
            )
        )
    return rows


def expand_7z(node: ArchiveNode, temp_dir: str) -> list[ResultRow]:
    factory = None
    try:
        with py7zr.SevenZipFile(extended_path(node.source_path), mode="r") as archive:
            infos = [info for info in archive.list() if not info.is_directory]
            if not infos:
                return [archive_row(node, "7z archive (empty)")]
            if len(infos) > SEVEN_ZIP_INSPECT_MAX_MEMBERS:
                return seven_zip_limit_rows(
                    node,
                    infos,
                    f"Content not read: more than {SEVEN_ZIP_INSPECT_MAX_MEMBERS:,} 7z members",
                )

            sizes = {info.filename: info.uncompressed or 0 for info in infos}
            factory = MemberCaptureFactory(sizes, temp_dir)
            archive.extractall(factory=factory)
            factory.close_all()
    except Exception as exc:
        if factory is not None:
            factory.close_all()
        return [archive_row(node, "7z archive", type(exc).__name__)]

    rows: list[ResultRow] = []
    for info in infos:
        parts = split_archive_path(info.filename)
        if not parts:
            continue
        levels = node.levels + parts
        logical_path = node.logical_path + "::" + "\\".join(parts)
        size = info.uncompressed or 0
        product = factory.products.get(info.filename)

        if size == 0:
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    0,
                    "Empty file",
                    True,
                    bom=BOM_NOT_PRESENT,
                )
            )
        elif product is None or not product.head:
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    size,
                    "Not inspected",
                    True,
                    "No captured member data",
                )
            )
        elif product.oversized:
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    size,
                    product.archive_type or "Archive",
                    True,
                    "Nested archive exceeds 50 GiB safety limit",
                )
            )
        elif product.temp_path:
            child = ArchiveNode(
                product.temp_path,
                logical_path,
                levels,
                size,
                node.depth + 1,
            )
            try:
                rows.extend(process_archive_node(child, temp_dir))
            finally:
                try:
                    os.unlink(product.temp_path)
                except OSError:
                    pass
        else:
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    size,
                    detect_magic(bytes(product.head)),
                    True,
                    bom=detect_bom(bytes(product.head)) or BOM_NOT_PRESENT,
                )
            )
    return rows


def configure_rar_backend() -> None:
    if os.name != "nt":
        return
    candidates = (
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "7-Zip" / "7z.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
        / "7-Zip"
        / "7z.exe",
    )
    for candidate in candidates:
        if candidate.is_file():
            rarfile.SEVENZIP_TOOL = str(candidate)
            try:
                rarfile.tool_setup(force=True)
            except Exception:
                pass
            return


def expand_rar(node: ArchiveNode, temp_dir: str) -> list[ResultRow]:
    rows: list[ResultRow] = []
    try:
        with rarfile.RarFile(extended_path(node.source_path)) as archive:
            infos = [info for info in archive.infolist() if not info.is_dir()]
            if not infos:
                return [archive_row(node, "RAR archive (empty)")]

            for info in infos:
                parts = split_archive_path(info.filename)
                if not parts:
                    continue
                levels = node.levels + parts
                logical_path = node.logical_path + "::" + "\\".join(parts)
                try:
                    with archive.open(info) as member:
                        rows.extend(
                            inspect_member_stream(
                                member,
                                size=info.file_size,
                                logical_path=logical_path,
                                levels=levels,
                                parent_depth=node.depth,
                                temp_dir=temp_dir,
                            )
                        )
                except Exception as exc:
                    rows.append(
                        make_row(
                            logical_path,
                            levels,
                            info.file_size,
                            "Unreadable archive member",
                            True,
                            type(exc).__name__,
                        )
                    )
    except Exception as exc:
        return [archive_row(node, "RAR archive", type(exc).__name__)]
    return rows


def process_archive_node(
    node: ArchiveNode,
    temp_dir: str,
    known_file_type: str | None = None,
    known_bom: str | None = None,
) -> list[ResultRow]:
    if known_file_type is None:
        file_type, bom = inspect_file(node.source_path)
    else:
        file_type, bom = known_file_type, known_bom
    if file_type not in ARCHIVE_TYPES:
        error = None
        if file_type == "Unreadable":
            error = "File could not be read"
        elif file_type == "ZIP container (unreadable)":
            error = "Invalid or unsupported ZIP container"
        return [archive_row(node, file_type, error, bom)]

    if node.depth >= MAX_ARCHIVE_DEPTH:
        return [
            archive_row(
                node,
                file_type,
                f"Archive nesting limit reached ({MAX_ARCHIVE_DEPTH})",
            )
        ]

    if file_type == "ZIP archive":
        return expand_zip(node, temp_dir)
    if file_type == "7z archive":
        return expand_7z(node, temp_dir)
    return expand_rar(node, temp_dir)


def process_file(item: PhysicalFile) -> list[ResultRow]:
    file_type, bom = inspect_file(item.path)
    if file_type in ARCHIVE_TYPES:
        with (
            HEAVY_ARCHIVE_SLOTS,
            tempfile.TemporaryDirectory(prefix="folder-review-") as temp_dir,
        ):
            node = ArchiveNode(
                item.path,
                item.path,
                item.levels,
                item.size,
                0,
                item.properties,
            )
            return process_archive_node(node, temp_dir, file_type, bom)

    error = None
    if file_type == "Unreadable":
        error = "File could not be read"
    elif file_type == "ZIP container (unreadable)":
        error = "Invalid or unsupported ZIP container"
    return [
        make_row(
            item.path,
            item.levels,
            item.size,
            file_type,
            False,
            error,
            bom=bom,
            properties=item.properties,
        )
    ]


def inspect_files(
    physical_files: list[PhysicalFile],
    workers: int = DEFAULT_WORKERS,
) -> list[ResultRow]:
    rows: list[ResultRow] = []
    completed = 0
    last_print = time.monotonic()

    items = iter(physical_files)
    with ThreadPoolExecutor(max_workers=workers) as executor:
        pending = set()
        for _ in range(workers * 2):
            try:
                pending.add(executor.submit(process_file, next(items)))
            except StopIteration:
                break

        while pending:
            done, pending = wait(pending, return_when=FIRST_COMPLETED)
            for future in done:
                rows.extend(future.result())
                completed += 1
                try:
                    pending.add(executor.submit(process_file, next(items)))
                except StopIteration:
                    pass

            now = time.monotonic()
            if now - last_print >= 10:
                print(
                    f"Inspected physical files: {completed:,} / {len(physical_files):,} | "
                    f"logical rows: {len(rows):,}"
                )
                last_print = now

    rows.sort(key=lambda row: row.full_path.casefold())
    return rows


# Excel output ---------------------------------------------------------------


def excel_text(value: str | None) -> str | None:
    if value is None:
        return None
    text = ILLEGAL_CHARACTERS_RE.sub("�", str(value))
    if len(text) > 32_767:
        text = text[:32_766] + "…"
    return text


def styled_cell(
    ws,
    value,
    *,
    font,
    border=None,
    fill=None,
    alignment=None,
    number_format=None,
    force_text=False,
):
    cell = WriteOnlyCell(ws, value=value)
    if force_text and isinstance(value, str):
        cell.data_type = "s"
    cell.font = font
    if border is not None:
        cell.border = border
    if fill is not None:
        cell.fill = fill
    if alignment is not None:
        cell.alignment = alignment
    if number_format is not None:
        cell.number_format = number_format
    return cell


def write_workbook(rows: list[ResultRow], output_path: Path, root_folder: str) -> None:
    if len(rows) > EXCEL_MAX_DATA_ROWS:
        raise ValueError(
            f"The result has {len(rows):,} rows; one Excel sheet supports at most "
            f"{EXCEL_MAX_DATA_ROWS:,} data rows in this layout."
        )

    max_level = max((len(row.levels) for row in rows), default=0)
    level_headers = [f"Уровень {index}" for index in range(1, max_level + 1)]
    data_headers = level_headers + [
        "Уровень листа",
        "Полный путь к файлу",
        "Имя файла",
        "Расширение",
        "Полное имя файла",
        "Скрытый",
        "st_mode",
        "st_ino",
        "st_dev",
        "st_nlink",
        "st_uid (User ID of the owner)",
        "st_gid (Group ID of the owner)",
        "Размер файла в мб",
        "Размер файла, байт",
        "Время последнего обращения к файлу",
        "Время последнего изменения файла",
        "Время создания файла",
        "BOM",
        "Тип по содержимому",
        "В архиве",
        "Ошибка",
    ]
    headers = ["№пп", *data_headers, "last"]

    wb = Workbook(write_only=True)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.calculation.calcId = 191029
    wb.views[0].windowWidth = 22260
    wb.views[0].windowHeight = 12650
    wb.views[0].tabRatio = 928

    ws = wb.create_sheet("folder_review")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_view.zoomScaleNormal = 85
    ws.sheet_properties.tabColor = "FF002060"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.sheet_format.defaultRowHeight = 14.5
    ws.freeze_panes = None
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = "9"  # A4
    ws.page_margins.left = 0.7
    ws.page_margins.right = 0.7
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    thin = Side(style="thin", color=Color(indexed=64))
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    normal_font = Font(name="Calibri", size=11, color=Color(theme=1))
    title_font = Font(
        name="Calibri", size=11, bold=True, color=Color(theme=1), charset=204
    )
    header_font = Font(
        name="Calibri", size=11, bold=True, color=Color(theme=1), charset=204
    )
    boundary_font = Font(
        name="Calibri", size=11, bold=True, color=Color(theme=0), charset=204
    )
    blue_fill = PatternFill("solid", fgColor="FF00B0F0")
    black_fill = PatternFill("solid", fgColor=Color(theme=1))
    centered = Alignment(horizontal="center", vertical="center")
    wrapped_centered = Alignment(horizontal="center", vertical="center", wrap_text=True)

    first_table_column = 2
    last_table_column = first_table_column + len(headers) - 1
    ws.column_dimensions["A"].width = 5.81640625
    wide_columns = {
        "Полный путь к файлу": 42,
        "Имя файла": 20,
        "Полное имя файла": 24,
        "st_uid (User ID of the owner)": 17,
        "st_gid (Group ID of the owner)": 17,
        "Время последнего обращения к файлу": 20,
        "Время последнего изменения файла": 20,
        "Время создания файла": 20,
        "Тип по содержимому": 24,
        "Ошибка": 20,
    }
    for offset, column in enumerate(range(first_table_column, last_table_column + 1)):
        letter = get_column_letter(column)
        boundary = column in (first_table_column, last_table_column)
        header = headers[offset]
        ws.column_dimensions[letter].width = (
            5.81640625 if boundary else wide_columns.get(header, 12.54296875)
        )
    ws.row_dimensions[6].height = 72

    ws.append([None])
    root_name = os.path.basename(root_folder.rstrip("\\/")) or root_folder
    title = f"Таблица: Обзор файлов папки {root_name}"
    ws.append(
        [None, styled_cell(ws, excel_text(title), font=title_font, force_text=True)]
    )
    ws.append([None])

    number_row = [None]
    for _ in headers:
        number_row.append(
            styled_cell(
                ws,
                "=COLUMN()-COLUMN($A$4)",
                font=normal_font,
                border=border,
                alignment=centered,
            )
        )
    ws.append(number_row)
    ws.append([None])

    header_row = [None]
    for index, header in enumerate(headers):
        boundary = index in (0, len(headers) - 1)
        header_row.append(
            styled_cell(
                ws,
                header,
                font=boundary_font if boundary else header_font,
                border=border,
                fill=black_fill if boundary else blue_fill,
                alignment=wrapped_centered,
            )
        )
    ws.append(header_row)

    centered_headers = {
        "Уровень листа",
        "Расширение",
        "Скрытый",
        "st_mode",
        "st_ino",
        "st_dev",
        "st_nlink",
        "st_uid (User ID of the owner)",
        "st_gid (Group ID of the owner)",
        "BOM",
        "В архиве",
    }
    number_formats = {
        "Размер файла в мб": "#,##0.000",
        "Размер файла, байт": "#,##0",
        "Время последнего обращения к файлу": "yyyy-mm-dd hh:mm:ss",
        "Время последнего изменения файла": "yyyy-mm-dd hh:mm:ss",
        "Время создания файла": "yyyy-mm-dd hh:mm:ss",
    }
    for row in rows:
        properties = row.properties
        stem_name = ntpath.splitext(row.leaf_name)[0]
        values = [
            *(
                row.levels[index] if index < len(row.levels) else None
                for index in range(max_level)
            ),
            len(row.levels),
            row.full_path,
            stem_name,
            row.extension,
            row.leaf_name,
            properties.hidden if properties else None,
            properties.mode if properties else None,
            properties.inode if properties else None,
            properties.device if properties else None,
            properties.links if properties else None,
            properties.user_id if properties else None,
            properties.group_id if properties else None,
            row.size / (1024 * 1024),
            row.size,
            properties.accessed_at if properties else None,
            properties.modified_at if properties else None,
            properties.created_at if properties else None,
            row.bom,
            row.content_type,
            row.in_archive,
            row.error,
        ]

        output_row = [None]
        output_row.append(
            styled_cell(
                ws,
                "=ROW()-ROW($B$6)",
                font=normal_font,
                border=border,
                alignment=centered,
            )
        )

        for header, value in zip(data_headers, values):
            is_text = isinstance(value, str)
            if is_text:
                value = excel_text(value)
            output_row.append(
                styled_cell(
                    ws,
                    value,
                    font=normal_font,
                    border=border,
                    alignment=centered if header in centered_headers else None,
                    number_format=number_formats.get(header),
                    force_text=is_text,
                )
            )

        output_row.append(
            styled_cell(ws, 1, font=normal_font, border=border, alignment=centered)
        )
        ws.append(output_row)

    last_row = max(6, 6 + len(rows))
    ws.auto_filter.ref = f"B6:{get_column_letter(last_table_column)}{last_row}"
    ws.sheet_view.selection[0].activeCell = "B6"
    ws.sheet_view.selection[0].sqref = "B6"

    temporary = output_path.with_name(f".{output_path.stem}.tmp.xlsx")
    try:
        wb.save(temporary)
        os.replace(temporary, output_path)
    finally:
        if temporary.exists():
            temporary.unlink()


# Command-line entry point ---------------------------------------------------


def validate_platform() -> None:
    if sys.platform == "darwin":
        raise RuntimeError("macOS is not supported. Use Windows or Linux.")
    if os.name != "nt" and not sys.platform.startswith("linux"):
        raise RuntimeError("This operating system is not supported.")


def worker_count(value: str) -> int:
    workers = int(value)
    if not 1 <= workers <= 64:
        raise argparse.ArgumentTypeError("workers must be between 1 and 64")
    return workers


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", help="Override ROOT_FOLDER for this run")
    parser.add_argument("--output", help="Override the local XLSX output path")
    parser.add_argument(
        "--workers",
        type=worker_count,
        default=DEFAULT_WORKERS,
        help=f"Concurrent file workers (default: {DEFAULT_WORKERS})",
    )
    return parser.parse_args()


def run() -> int:
    validate_platform()
    args = parse_args()
    script_dir = Path(__file__).resolve().parent

    root_folder = os.path.normpath((args.root or ROOT_FOLDER).strip().strip('"'))
    if (
        not args.root
        and root_folder == ROOT_FOLDER
        and ROOT_FOLDER == r"PASTE_FOLDER_PATH_HERE"
    ):
        raise ValueError(
            "Edit ROOT_FOLDER near the top of folder_review.py before running."
        )
    if not os.path.isdir(extended_path(root_folder)):
        raise ValueError(f"Folder is unavailable: {root_folder}")

    output_path = (
        Path(args.output).resolve() if args.output else script_dir / OUTPUT_FILE
    )
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Output file must use the .xlsx extension.")

    configure_rar_backend()
    started = time.monotonic()
    print(f"Scanning: {root_folder}")
    print(f"Workers: {args.workers}")
    physical_files, stats = scan_tree(
        root_folder,
        {normalized_key(str(output_path))},
        args.workers,
    )
    print(f"Physical files found: {len(physical_files):,}")

    rows = inspect_files(physical_files, args.workers)
    print(f"Logical file rows: {len(rows):,}")
    print("Writing Excel workbook...")
    write_workbook(rows, output_path, root_folder)

    errors = sum(row.error is not None for row in rows)
    elapsed = time.monotonic() - started
    print(
        f"Done in {elapsed:.1f}s | folders: {stats.folders:,} | "
        f"skipped folders: {stats.skipped_folders:,} | skipped entries: {stats.skipped_entries:,} | "
        f"rows with errors: {errors:,}"
    )
    print(f"Result: {output_path}")
    return 0


def main() -> int:
    try:
        return run()
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
