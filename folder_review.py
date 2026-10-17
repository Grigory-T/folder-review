"""Create one flat-table XLSX inventory for a manually selected folder."""

# ruff: noqa: BLE001, S110 -- archive libraries expose broad failure modes;
# each failure is converted into an inventory error instead of stopping the scan.

from __future__ import annotations

import argparse
import errno
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
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from datetime import UTC, datetime
from itertools import chain, islice
from pathlib import Path
from threading import Lock, Semaphore

# A normal source launch must not create caches beside inspected source files.
sys.dont_write_bytecode = True

import olefile
import py7zr
import rarfile
import xlsxwriter
from blake3 import blake3
from py7zr import Py7zIO, WriterFactory
from xlsxwriter.color import Color
from xlsxwriter.utility import xl_col_to_name

# User configuration ---------------------------------------------------------

ROOT_FOLDER = r"PASTE_FOLDER_PATH_HERE"

OUTPUT_FILE = "folder-review.xlsx"

# Diagnostic metadata remains available through --full-metadata. The default
# inventory keeps fields useful for an ordinary content/folder review.
ADVANCED_METADATA_HEADERS = frozenset(
    {
        "Content hash (BLAKE3)",
        "Raw suffix",
        "Link type",
        "Link target",
        "Permissions",
        "Read-only",
        "Executable",
        "System (Windows)",
        "Archive attribute (Windows)",
        "Temporary (Windows)",
        "Offline (Windows)",
        "Compressed (Windows)",
        "Encrypted (Windows)",
        "Sparse (Windows)",
        "Content indexing disabled (Windows)",
        "File attributes (Windows)",
        "File attributes (st_file_attributes)",
        "Reparse tag (st_reparse_tag)",
        "Mode (st_mode)",
        "File ID (st_ino)",
        "Device ID (st_dev)",
        "Device type (st_rdev)",
        "Link count (st_nlink)",
        "Owner ID (st_uid)",
        "Group ID (st_gid)",
        "Block size (st_blksize)",
        "Allocated blocks (st_blocks)",
    }
)


DAX_MEASURE = """=VAR Source =
    SELECTCOLUMNS ( 'Data', "Value", 'Data'[TextColumn] )

VAR Counts =
    GROUPBY (
        FILTER ( Source, LEN ( TRIM ( [Value] ) ) > 0 ),
        [Value],
        "Count", COUNTX ( CURRENTGROUP (), 1 )
    )

RETURN
    CONCATENATEX (
        Counts,
        [Value] & " (" & [Count] & ")",
        " · ",
        [Count], DESC,
        [Value], ASC
    )"""


# Runtime limits -------------------------------------------------------------

DEFAULT_WORKERS = 16
MAX_CONCURRENT_ARCHIVES = 4

MAGIC_READ_SIZE = 8 * 1024
MAX_ARCHIVE_DEPTH = 20
MAX_NESTED_ARCHIVE_SIZE = 50 * 1024 * 1024 * 1024
DEFAULT_MAX_EXPANDED_BYTES = 100 * 1024 * 1024 * 1024
DEFAULT_MAX_ARCHIVE_SECONDS = 6 * 60 * 60
SEVEN_ZIP_INSPECT_MAX_MEMBERS = 25_000
EXCEL_MAX_DATA_ROWS = 1_048_570  # rows 7..1,048,576
EXCEL_MAX_TEXT_LENGTH = 32_767
HEAVY_ARCHIVE_SLOTS = Semaphore(MAX_CONCURRENT_ARCHIVES)
ARCHIVE_TYPES = {"ZIP archive", "7z archive", "RAR archive"}
ARCHIVE_EXTENSIONS = frozenset({"zip", "7z", "rar"})
# These names improve leaf labels only. Expansion uses the positive archive
# extension policy, so an unknown ZIP-based application format is a leaf too.
ZIP_LEAF_FORMATS = {
    extension: f"{label} ({extension.upper()})"
    for extensions, label in (
        (("docx", "docm"), "Word document"),
        (("dotx", "dotm"), "Word template"),
        (("xlsx", "xlsm"), "Excel workbook"),
        (("xlsb",), "Excel binary workbook"),
        (("xltx", "xltm"), "Excel template"),
        (("xlam",), "Excel add-in"),
        (("pptx", "pptm"), "PowerPoint presentation"),
        (("potx", "potm"), "PowerPoint template"),
        (("ppsx", "ppsm"), "PowerPoint slideshow"),
        (("sldx", "sldm"), "PowerPoint slide"),
        (("ppam",), "PowerPoint add-in"),
        (("thmx",), "Office theme"),
        (("vsdx", "vsdm"), "Visio drawing"),
        (("vstx", "vstm"), "Visio template"),
        (("vssx", "vssm"), "Visio stencil"),
        (("odt", "ott", "odm"), "OpenDocument text"),
        (("ods", "ots"), "OpenDocument spreadsheet"),
        (("odp", "otp"), "OpenDocument presentation"),
        (("odg", "otg"), "OpenDocument drawing"),
        (("odf", "odb"), "OpenDocument file"),
        (("epub",), "EPUB document"),
        (("jar",), "JAR archive"),
        (("war", "ear"), "Java application package"),
        (("apk", "aab", "apks", "xapk"), "Android application package"),
        (("ipa",), "iOS application package"),
        (("appx", "appxbundle", "msix", "msixbundle"), "Windows application package"),
        (("xpi",), "Browser extension package"),
        (("nupkg", "snupkg"), "NuGet package"),
        (("whl",), "Python wheel"),
        (("kmz",), "Google Earth document"),
        (("xps", "oxps", "dwfx"), "Fixed-layout document"),
        (("3mf", "usdz"), "3D model package"),
        (("pages", "numbers", "key"), "iWork document"),
        (("kra", "ora", "sketch"), "Graphics document"),
        (("sb2", "sb3", "ggb"), "Application document"),
        (("idml",), "InDesign markup document"),
        (("xmind",), "Mind-map document"),
        (("mxl",), "Compressed MusicXML document"),
        (("cbz", "cbr", "cb7"), "Comic-book document"),
    )
    for extension in extensions
}
# Keep the existing labels for these two recognized package formats.
ZIP_LEAF_FORMATS.update({"epub": "EPUB document", "jar": "JAR archive"})
STREAM_ARCHIVE_TYPES = {
    "ZIP container": "ZIP archive",
    "7z archive": "7z archive",
    "RAR archive": "RAR archive",
}

OFFICE_DOCUMENT_EXTENSIONS = frozenset(
    {
        "doc",
        "docm",
        "docx",
        "dot",
        "dotm",
        "dotx",
        "ods",
        "odp",
        "odt",
        "pot",
        "potm",
        "potx",
        "pps",
        "ppsm",
        "ppsx",
        "ppt",
        "pptm",
        "pptx",
        "rtf",
        "vsd",
        "vsdm",
        "vsdx",
        "xls",
        "xlsb",
        "xlsm",
        "xlsx",
        "xlt",
        "xltm",
        "xltx",
    }
)


MAGIC_SIGNATURES = (
    (b"\x37\x7a\xbc\xaf\x27\x1c", "7z archive"),
    (b"Rar!\x1a\x07\x01\x00", "RAR archive"),
    (b"Rar!\x1a\x07\x00", "RAR archive"),
    (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "OLE2 compound file"),
    (
        (
            b"\x4c\x00\x00\x00\x01\x14\x02\x00\x00\x00"
            b"\x00\x00\xc0\x00\x00\x00\x00\x00\x00\x46"
        ),
        "Windows shortcut (LNK)",
    ),
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
INTEGRITY_VERIFIED = "проверена"
INTEGRITY_NOT_VERIFIED = "не проверена"
INTEGRITY_ERROR = "ошибка"

YES = "да"
NO = "нет"

WINDOWS_FILE_ATTRIBUTES = (
    (0x00000001, "READONLY"),
    (0x00000002, "HIDDEN"),
    (0x00000004, "SYSTEM"),
    (0x00000010, "DIRECTORY"),
    (0x00000020, "ARCHIVE"),
    (0x00000040, "DEVICE"),
    (0x00000080, "NORMAL"),
    (0x00000100, "TEMPORARY"),
    (0x00000200, "SPARSE_FILE"),
    (0x00000400, "REPARSE_POINT"),
    (0x00000800, "COMPRESSED"),
    (0x00001000, "OFFLINE"),
    (0x00002000, "NOT_CONTENT_INDEXED"),
    (0x00004000, "ENCRYPTED"),
    (0x00008000, "INTEGRITY_STREAM"),
    (0x00020000, "NO_SCRUB_DATA"),
    (0x00040000, "RECALL_ON_OPEN"),
    (0x00080000, "PINNED"),
    (0x00100000, "UNPINNED"),
    (0x00400000, "RECALL_ON_DATA_ACCESS"),
)


# Data models ----------------------------------------------------------------


@dataclass(slots=True, frozen=True)
class FileProperties:
    permissions: str
    hidden: str
    read_only: str
    executable: str
    system: str | None
    archive_attribute: str | None
    temporary: str | None
    offline: str | None
    compressed: str | None
    encrypted: str | None
    sparse: str | None
    not_content_indexed: str | None
    file_attributes: str | None
    file_attributes_value: int | None
    reparse_tag: str | None
    mode: int
    inode: str
    device: str
    device_type: str | None
    links: int
    user_id: str | None
    group_id: str | None
    block_size: int | None
    blocks: int | None
    accessed_at: datetime | None
    modified_at: datetime | None
    metadata_changed_at: datetime | None
    created_at: datetime | None


@dataclass(slots=True, frozen=True)
class PhysicalFile:
    path: str
    levels: tuple[str, ...]
    size: int
    properties: FileProperties | None = None
    item_type: str = "Файл"
    link_type: str | None = None
    link_target: str | None = None
    inspect_content: bool = True
    virtual_folder: bool = False
    folder_unreadable: bool = False
    folder_error: str | None = None


@dataclass(slots=True, frozen=True)
class ArchiveNode:
    source_path: str
    logical_path: str
    levels: tuple[str, ...]
    size: int
    depth: int
    properties: FileProperties | None = None
    content_hash: bytes | None = None


@dataclass(slots=True, frozen=True)
class ResultRow:
    levels: tuple[str, ...]
    full_path: str
    leaf_name: str
    extension: str
    size: int
    item_type: str
    recognized_role: str | None
    link_type: str | None
    link_target: str | None
    content_type: str
    in_archive: str
    bom: str | None = None
    integrity_status: str | None = None
    properties: FileProperties | None = None
    error: str | None = None
    content_hash: bytes | None = None
    virtual_folder: bool = False
    virtual_reason: str | None = None

    @property
    def is_virtual(self) -> bool:
        return self.virtual_folder or self.virtual_reason is not None


@dataclass(slots=True, frozen=True)
class Diagnostic:
    path: str
    stage: str
    error: str


@dataclass(slots=True)
class ScanStats:
    folders: int = 0
    skipped_folders: int = 0
    skipped_entries: int = 0
    diagnostics: list[Diagnostic] = field(default_factory=list)


class BudgetExceeded(RuntimeError):
    pass


@dataclass(slots=True)
class WorkBudget:
    max_expanded_bytes: int | None = DEFAULT_MAX_EXPANDED_BYTES
    max_rows: int = EXCEL_MAX_DATA_ROWS
    deadline: float | None = None
    expanded_bytes: int = 0
    rows_claimed: int = 0
    _lock: Lock = field(default_factory=Lock, repr=False)

    def check_time(self) -> None:
        if self.deadline is not None and time.monotonic() > self.deadline:
            raise BudgetExceeded("Archive inspection time budget exceeded")

    def consume_expanded(self, amount: int) -> None:
        if amount <= 0:
            self.check_time()
            return
        with self._lock:
            self.check_time()
            total = self.expanded_bytes + amount
            if self.max_expanded_bytes is not None and total > self.max_expanded_bytes:
                raise BudgetExceeded(
                    f"Expanded-data budget exceeded ({self.max_expanded_bytes:,} bytes)"
                )
            self.expanded_bytes = total

    def ensure_declared_size(self, amount: int) -> None:
        with self._lock:
            self.check_time()
            if (
                self.max_expanded_bytes is not None
                and self.expanded_bytes + amount > self.max_expanded_bytes
            ):
                raise BudgetExceeded(
                    "Declared archive content exceeds remaining expanded-data budget"
                )

    def expand_one_row(self, member_count: int) -> None:
        additional = max(member_count - 1, 0)
        with self._lock:
            self.check_time()
            total = self.rows_claimed + additional
            if total > self.max_rows:
                raise BudgetExceeded(
                    f"Logical-row budget exceeded ({self.max_rows:,} rows)"
                )
            self.rows_claimed = total

    def initialize_rows(self, row_count: int) -> None:
        with self._lock:
            if row_count > self.max_rows:
                raise BudgetExceeded(
                    f"Logical-row budget exceeded ({self.max_rows:,} rows)"
                )
            self.rows_claimed = row_count


@dataclass(slots=True, frozen=True)
class RunReport:
    status: str
    started_at: datetime
    completed_at: datetime
    workers: int
    physical_items: int
    logical_rows: int
    row_errors: int
    stats: ScanStats
    expanded_bytes: int
    max_expanded_bytes: int | None
    max_archive_seconds: float | None
    max_rows: int


# Path and row helpers --------------------------------------------------------


def regular_path(path: str | os.PathLike[str]) -> str:
    """Remove a Windows extended-path prefix for display and path comparison."""
    value = os.fspath(path)
    if os.name != "nt":
        return value
    folded = value.casefold()
    if folded.startswith("\\\\?\\unc\\"):
        return "\\\\" + value[8:]
    if folded.startswith("\\\\?\\"):
        return value[4:]
    return value


def absolute_path(path: str | os.PathLike[str]) -> str:
    """Return an absolute path without exposing a Windows device prefix."""
    return regular_path(os.path.abspath(os.path.expanduser(regular_path(path))))


def extended_path(path: str | os.PathLike[str]) -> str:
    """Return a Windows extended-length path for every physical filesystem call."""
    value = os.fspath(path)
    if os.name != "nt" or value.startswith(("\\\\?\\", "\\\\.\\")):
        return value
    value = os.path.abspath(value)
    if value.startswith("\\\\"):
        return "\\\\?\\UNC\\" + value[2:]
    return "\\\\?\\" + value


def normalized_key(path: str | os.PathLike[str]) -> str:
    return os.path.normcase(absolute_path(path))


# Runtime safety -------------------------------------------------------------


def resolved_path(path: str | os.PathLike[str]) -> str:
    return os.path.normcase(regular_path(os.path.realpath(extended_path(path))))


def path_within(path: str | os.PathLike[str], root: str) -> bool:
    """Include junction/symlink and Windows mapped/UNC aliases of the root."""
    candidate = resolved_path(path)
    source = resolved_path(root)
    try:
        if os.path.commonpath((candidate, source)) == source:
            return True
    except ValueError:  # different Windows drives
        pass
    # A subst or mapped-drive alias can retain a different textual spelling.
    while True:
        try:
            if os.path.samefile(extended_path(candidate), extended_path(source)):
                return True
        except OSError:
            pass
        parent = os.path.dirname(candidate)
        if parent == candidate:
            return False
        candidate = parent


def local_storage(path: str | os.PathLike[str]) -> bool:
    """Reject remote working/output storage; remote input reads stay allowed."""
    candidate = resolved_path(path)
    if os.name == "nt":
        import ctypes

        drive, _ = os.path.splitdrive(candidate)
        if drive.startswith("\\\\"):
            return False
        get_drive_type = ctypes.windll.kernel32.GetDriveTypeW
        get_drive_type.argtypes = [ctypes.c_wchar_p]
        get_drive_type.restype = ctypes.c_uint
        return get_drive_type(drive + "\\") in (2, 3, 6)  # removable, fixed, RAM
    # Linux mount metadata is local; this does not contact a server.
    network_types = {"cifs", "smb3", "nfs", "nfs4", "fuse.sshfs", "9p", "afs"}
    best_mount = ""
    filesystem = ""
    with open("/proc/self/mountinfo", encoding="utf-8") as mounts:
        for line in mounts:
            fields, details = line.rstrip().split(" - ", 1)
            mount = re.sub(
                r"\\([0-7]{3})", lambda match: chr(int(match[1], 8)), fields.split()[4]
            )
            if (
                candidate == mount or candidate.startswith(mount.rstrip("/") + "/")
            ) and len(mount) > len(best_mount):
                best_mount, filesystem = mount, details.split()[0]
    return filesystem not in network_types


def validate_output_path(output: Path, root: str) -> None:
    """Reports may be inside input/on SMB; never follow a report symlink."""
    if os.path.islink(extended_path(output)):
        raise ValueError("Report output must not be a symbolic link.")


_protected_root: str | None = None
_report_output_path: str | None = None
_report_write_paths: set[str] = set()


def runtime_audit(event: str, arguments: tuple) -> None:
    """Defence against accidental Python writes/network calls, not a sandbox."""
    root = _protected_root
    if root is None:
        return
    if event.startswith("socket.") or event in {
        "urllib.Request",
        "http.client.connect",
        "webbrowser.open",
        "winreg.ConnectRegistry",
    }:
        raise PermissionError("Network APIs are disabled during folder review.")

    def protect(value, dir_fd=None):
        if isinstance(value, int):
            if os.name == "nt":
                import ctypes
                import msvcrt

                buffer = ctypes.create_unicode_buffer(32_768)
                get_name = ctypes.windll.kernel32.GetFinalPathNameByHandleW
                get_name.argtypes = [
                    ctypes.c_void_p,
                    ctypes.c_wchar_p,
                    ctypes.c_uint,
                    ctypes.c_uint,
                ]
                get_name.restype = ctypes.c_uint
                length = get_name(msvcrt.get_osfhandle(value), buffer, len(buffer), 0)
                if not 0 < length < len(buffer):
                    raise PermissionError(
                        "Unresolved file descriptor mutation is disabled."
                    )
                value = buffer.value
            else:
                value = os.readlink(f"/proc/self/fd/{value}")
        if not isinstance(value, (str, bytes, os.PathLike)):
            return
        value = os.fsdecode(value)
        if dir_fd is not None and dir_fd != -1 and not os.path.isabs(value):
            value = os.path.join(os.readlink(f"/proc/self/fd/{dir_fd}"), value)
        if normalized_key(value) in _report_write_paths:
            if os.path.islink(extended_path(value)):
                raise PermissionError("Report writes must not follow symbolic links.")
            return
        if path_within(value, root):
            raise PermissionError(
                "The inspected folder is read-only for folder review."
            )
        if not local_storage(value):
            raise PermissionError("Working-file writes must stay on local storage.")

    if event == "open":
        filename, _mode, flags = arguments
        # subprocess opens the local null device read/write for disconnected
        # extractor stdin. It is a sink, not network or persistent storage.
        if (
            isinstance(filename, (str, bytes))
            and os.fsdecode(filename).casefold() == os.devnull.casefold()
        ):
            return
        if flags & (os.O_WRONLY | os.O_RDWR | os.O_CREAT | os.O_TRUNC | os.O_APPEND):
            protect(filename)
    elif event in {"os.remove", "os.rmdir"}:
        protect(arguments[0], arguments[1])
    elif event == "os.mkdir":
        protect(arguments[0], arguments[2])
    elif event in {"os.rename", "os.link"}:
        protect(arguments[0], arguments[2])
        protect(arguments[1], arguments[3])
    elif event == "os.symlink":
        protect(arguments[1], arguments[2])
    elif event in {"os.chmod", "os.chown", "os.utime"}:
        protect(arguments[0], arguments[-1])
    elif event in {"os.truncate", "os.setxattr", "os.removexattr", "shutil.rmtree"}:
        protect(arguments[0])
    elif event == "subprocess.Popen":
        executable, command, _, _ = arguments
        # Windows audits the serialized command line and may leave executable
        # unset even when Popen was given an argv list. Decode it with the OS
        # parser, preserving quotes/spaces in extractor and member paths.
        if os.name == "nt" and isinstance(command, str):
            import ctypes

            count = ctypes.c_int()
            parse_command = ctypes.windll.shell32.CommandLineToArgvW
            parse_command.argtypes = [ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_int)]
            parse_command.restype = ctypes.POINTER(ctypes.c_wchar_p)
            parsed = parse_command(command, ctypes.byref(count))
            if not parsed:
                raise PermissionError("Unable to validate archive backend command.")
            try:
                command = [parsed[index] for index in range(count.value)]
            finally:
                free = ctypes.windll.kernel32.LocalFree
                free.argtypes = [ctypes.c_void_p]
                free.restype = ctypes.c_void_p
                free(parsed)
        if not isinstance(command, (list, tuple)) or not command:
            raise PermissionError(
                "Only read-only archive backend commands are allowed."
            )
        executable = executable or command[0]
        allowed = {
            "7z": (("i",), ("e", "-so", "-bb0")),
            "7za": (("i",), ("e", "-so", "-bb0")),
            "7zz": (("i",), ("e", "-so", "-bb0")),
            "unrar": (("-inul", "-?"), ("p", "-inul")),
            "unar": (("-version",), ("-q", "-o", "-")),
            "bsdtar": (("--version",), ("-x", "--to-stdout", "-f")),
        }
        name = os.path.basename(os.fsdecode(executable)).lower().removesuffix(".exe")
        prefixes = allowed.get(name, ())
        if not any(
            tuple(command[1 : 1 + len(prefix)]) == prefix for prefix in prefixes
        ):
            raise PermissionError(
                "Only read-only archive backend commands are allowed."
            )
        executable = shutil.which(executable)
        if executable is None:
            # Preserve rarfile's normal missing-tool detection and fallback.
            raise FileNotFoundError(
                errno.ENOENT, "The archive backend is unavailable.", command[0]
            )
        protect(executable)
    elif event in {
        "os.system",
        "os.exec",
        "os.posix_spawn",
        "os.spawn",
        "os.startfile",
        "os.startfile/2",
    }:
        raise PermissionError(
            "Executing inspected content or shell commands is disabled."
        )


sys.addaudithook(runtime_audit)


@contextmanager
def safe_runtime(root: str, output: Path):
    """Keep workspaces outside input; only the selected report may be saved in it."""
    global _protected_root, _report_output_path
    if _protected_root is not None:
        raise RuntimeError("Concurrent scans in one Python process are not supported.")
    validate_output_path(output, root)
    if getattr(sys, "frozen", False) and path_within(sys._MEIPASS, root):
        raise ValueError(
            "The inspected folder contains the running temporary runtime. Select a narrower input folder."
        )
    previous_tempdir = tempfile.tempdir
    _protected_root = root
    _report_output_path = normalized_key(output)
    try:
        # gettempdir() may probe candidate folders by creating a file. Guard
        # those probes too, including when TEMP/TMP points into the input.
        try:
            workspace_parent = tempfile.gettempdir()
        except FileNotFoundError:
            workspace_parent = None
        candidates = [
            workspace_parent,
            str(output.parent),
            os.path.dirname(absolute_path(root)),
            str(Path.home()),
        ]
        workspace_parent = next(
            (
                candidate
                for candidate in candidates
                if candidate
                and not path_within(candidate, root)
                and local_storage(candidate)
                and os.path.isdir(extended_path(candidate))
            ),
            None,
        )
        if workspace_parent is None:
            raise ValueError(
                "No local temporary folder outside the inspected folder is available."
            )
        with tempfile.TemporaryDirectory(
            prefix="folder-review-work-", dir=extended_path(workspace_parent)
        ) as workspace:
            tempfile.tempdir = workspace
            try:
                yield
            finally:
                tempfile.tempdir = previous_tempdir
    finally:
        tempfile.tempdir = previous_tempdir
        _protected_root = None
        _report_output_path = None
        _report_write_paths.clear()


def split_archive_path(name: str) -> tuple[str, ...]:
    return tuple(part for part in re.split(r"[\\/]+", name) if part not in ("", "."))


def extension_of(name: str) -> str:
    """Conservative display suffix; preserve unknown and long real extensions."""
    suffix = raw_extension_of(name)
    if len(suffix) > 64 or re.fullmatch(r"\d{1,3}_\d{15,20}", suffix):
        return ""
    # Package-cache version/ABI names are extensionless, unlike wheel files.
    if re.fullmatch(
        r"(?:\d+-)?(?:cp\d+|py\d+|abi\d+|windows|linux|macosx|manylinux|musllinux)"
        r"[-_][a-z0-9_.-]+",
        suffix,
    ):
        return ""
    if re.fullmatch(r"\d+(?:\.\d+)+(?:[-+][a-z0-9_-]+)?", name.casefold()):
        return ""
    return suffix


def raw_extension_of(name: str) -> str:
    return ntpath.splitext(name)[1].lstrip(".").lower()


VIRTUAL_MARKERS = {
    "empty": ("dummy.dummy", "Нет файловых строк"),
    "denied": ("deny.deny", "Нет доступа"),
    "password": ("password.password", "Требуется пароль"),
    "corrupt": ("corrupt.corrupt", "Ошибка целостности"),
    "limit": ("limit.limit", "Лимит инспекции"),
    "error": ("error.error", "Ошибка инспекции"),
    "uninspected": ("error.error", "Не прочитано"),
}


def problem_kind(error: str) -> str:
    # Quoted filenames such as password.txt are not diagnoses.
    text = re.sub(r"(['\"]).*?\1", "", error.casefold())
    if any(
        part in text
        for part in ("permissionerror:", "access is denied", "permission denied")
    ) or re.search(r"\bwinerror\s+5\b", text):
        return "denied"
    if any(
        part in text
        for part in (
            "crcerror:",
            "bad crc",
            "crc failed",
            "crc mismatch",
            "checksum mismatch",
            "cyclic redundancy",
            "corrupt archive",
        )
    ):
        return "corrupt"
    if re.search(
        r"password(?:required|wrong|incorrect)?\s*:|password (?:is )?required|(?:without a|wrong|incorrect|invalid) password|encrypted (?:7z|rar) headers",
        text,
    ):
        return "password"
    if any(
        part in text
        for part in (
            "budgetexceeded:",
            "archive nesting limit",
            "nested archive exceeds",
            "content not read: more than",
        )
    ):
        return "limit"
    return "error"


def display_row(row: ResultRow) -> ResultRow:
    """Keep failed locations in hierarchy with a clearly virtual marker leaf."""
    if row.is_virtual or not row.error or row.content_hash is not None:
        return row
    reason = problem_kind(row.error)
    name = VIRTUAL_MARKERS[reason][0]
    return replace(
        row,
        levels=row.levels + (name,),
        leaf_name=name,
        extension="",
        content_hash=None,
        virtual_reason=reason,
        item_type="Виртуальная строка",
    )


# Practical MIME-style families, subdividing application/* into useful groups.
# Fixed lists keep Windows/Python/EXE classification independent of OS registry.
_EXTENSION_GROUPS = (
    (
        "Документы",
        "Текстовые документы",
        "doc docx docm dot dotx dotm odt ott rtf pages wpd chm",
    ),
    ("Документы", "Таблицы", "xls xlsx xlsm xlsb xlt xltx xltm xlc ods ots numbers"),
    ("Документы", "Презентации", "ppt pptx pptm potx potm pps ppsx ppsm odp otp"),
    ("Документы", "PDF и печатные документы", "pdf xps oxps ps eps"),
    ("Документы", "Электронные книги", "epub mobi azw azw3 fb2 djvu djv"),
    (
        "Документы",
        "Обычный текст и разметка",
        "txt text md markdown rst adoc tex nfo man 1 2 3 4 5 6 7 8 9",
    ),
    ("Документы", "Почта и календарь", "eml msg mbox pst ost ics vcf oft"),
    (
        "Изображения",
        "Растровые изображения",
        "jpg jpeg jpe png gif bmp dib tif tiff webp heic heif avif ico cur psd xcf xbm xpm ppm pgm pbm pnm icns",
    ),
    ("Изображения", "Векторные изображения", "svg svgz ai emf wmf cdr"),
    ("Изображения", "Фотографии RAW", "raw cr2 cr3 nef nrw arw dng raf orf rw2 pef"),
    (
        "Аудио",
        "Аудиозаписи",
        "mp3 wav flac aac m4a ogg oga opus wma aiff aif ape alac amr",
    ),
    ("Аудио", "Музыка и списки воспроизведения", "mid midi kar m3u m3u8 pls cue"),
    (
        "Видео",
        "Видеозаписи",
        "mp4 m4v mkv avi mov qt webm wmv mpg mpeg m2v m2ts vob flv 3gp",
    ),
    ("Видео", "Субтитры", "srt ass ssa vtt sub"),
    ("Архивы и образы", "Архивы", "zip rar 7z tar cab cpio ar lzh lha ace"),
    ("Архивы и образы", "Сжатые файлы", "gz bz2 xz zst zstd lz lzma br z tgz tbz2 txz"),
    (
        "Архивы и образы",
        "Образы дисков",
        "iso img vhd vhdx vmdk vdi qcow qcow2 dmg wim esd",
    ),
    (
        "Программы и код",
        "Программы и библиотеки",
        "exe dll sys ocx cpl scr com so dylib pyd wasm lib a node class pyc pyo",
    ),
    (
        "Программы и код",
        "Пакеты и установщики",
        "msi msp msix appx msixbundle appxbundle whl egg deb rpm apk aab ipa nupkg vsix crx xpi asar jar war ear",
    ),
    (
        "Программы и код",
        "Исходный код и сценарии",
        "py pyw js mjs cjs ts tsx jsx c cc cpp cxx h hh hpp cs java go rs rb php pl pm sh bash zsh fish bat cmd ps1 psm1 psd1 vbs vbe bas vb sql lua r swift kt kts dart scala groovy f f90 pas asm s cts tcl tk tm csh tcsh nu xsh pyx pxd pxi pyf f95 f03 f08 for f77 coffee idl sct pys",
    ),
    (
        "Программы и код",
        "Сборка и метаданные пакетов",
        "make cmake sln csproj vbproj vcxproj gradle nuspec nuspectemplate ninja nmake gyp gypi pc def nsi nsh wxs wixproj vc rc res",
    ),
    (
        "Программы и код",
        "Веб-страницы и стили",
        "html htm xhtml css scss sass less asp aspx jsp vue svelte jspf jspx tag tld",
    ),
    (
        "Данные и настройки",
        "Табличные данные",
        "csv tsv parquet arrow feather avro orc tab",
    ),
    (
        "Данные и настройки",
        "Структурированные данные",
        "json jsonl ndjson xml rdf ttl yaml yml pb binarypb msgpack cbor bson pkl pickle dtd webmanifest",
    ),
    (
        "Данные и настройки",
        "Базы данных",
        "db sqlite sqlite3 db3 mdb accdb dbf odb db-journal db-wal db-shm ldb",
    ),
    (
        "Данные и настройки",
        "Настройки",
        "ini cfg conf config cnf properties env reg desktop service socket target policy plist manifest jupyterlab-workspace toml mplstyle defaults opts",
    ),
    (
        "Проектирование",
        "Схемы и диаграммы",
        "vsd vsdx vsdm vss vssx vssm vst vstx vstm drawio dia plantuml puml mmd",
    ),
    (
        "Проектирование",
        "Чертежи CAD",
        "dwg dxf dgn rvt rfa rft ifc fcstd sldprt sldasm slddrw ipt iam idw step stp iges igs",
    ),
    ("Проектирование", "3D-модели", "stl fbx blend gltf glb 3ds dae skp max"),
    ("Шрифты", "Файлы шрифтов", "ttf otf ttc woff woff2 eot pfb pfm afm bdf pcf cff"),
    (
        "Системные и служебные",
        "Сертификаты и ключи",
        "pem crt cer der pfx p12 p7b keytab pub cert certs keychain",
    ),
    (
        "Системные и служебные",
        "Кэш и временные файлы",
        "tmp temp cache dmp dump swp swo part crdownload",
    ),
    ("Системные и служебные", "Ярлыки и ссылки", "lnk url webloc"),
    ("Системные и служебные", "Резервные копии", "bak backup old orig sav"),
    (
        "Системные и служебные",
        "Файлы блокировки и состояния",
        "lock pid state stamp tsbuildinfo fcrstate",
    ),
    ("Документы", "Лицензии и уведомления", "license licence"),
    ("Системные и служебные", "Журналы", "log journal"),
    (
        "Системные и служебные",
        "Контрольные суммы и подписи",
        "sha256 sha256sum sha512 sha512sum md5 md5sum sig asc",
    ),
    ("Программы и код", "Интерактивные блокноты", "ipynb"),
    ("Программы и код", "Объявления типов", "pyi"),
    ("Программы и код", "Шаблоны", "j2 jinja jinja2 tpl tmpl template in"),
    ("Программы и код", "Локализация и переводы", "po mo xliff xlf"),
    (
        "Программы и код",
        "Схемы и описания интерфейсов",
        "xsd xsl xslt proto pdl bnf ebnf lark",
    ),
    (
        "Данные и настройки",
        "Научные массивы и наборы данных",
        "npy npz mat h5 hdf5 nc fits fit fts",
    ),
    ("Данные и настройки", "Модели машинного обучения", "tflite onnx safetensors"),
    ("Данные и настройки", "Захваты сетевого трафика", "pcap pcapng cap"),
    ("Данные и настройки", "Ресурсы приложений", "pak"),
    (
        "Данные и настройки",
        "Данные неопределенного формата",
        "dat bin enc idx store pack model",
    ),
)
EXTENSION_GROUPS = {
    extension: (family, subtype)
    for family, subtype, extensions in _EXTENSION_GROUPS
    for extension in extensions.split()
}


def file_type_groups(row: ResultRow) -> tuple[str, str]:
    """Deterministic display groups; never decide whether an archive expands."""
    if row.is_virtual:
        return "Служебные строки", VIRTUAL_MARKERS[row.virtual_reason or "empty"][1]
    name = row.leaf_name.casefold()
    extension = row.extension
    content = row.content_type
    # Exact conventions precede broad suffix rules (CMakeLists.txt is not prose).
    if name in {
        "dockerfile",
        "containerfile",
        "makefile",
        "gnumakefile",
        "cmakelists.txt",
        "meson.build",
        "meson_options.txt",
        "manifest.in",
        "setup.py",
        "setup.cfg",
        "pyproject.toml",
        "package.json",
        "package-lock.json",
        "npm-shrinkwrap.json",
        "cargo.toml",
        "cargo.lock",
        "go.mod",
        "go.sum",
        "pom.xml",
        "build.gradle",
        "settings.gradle",
        "requirements.txt",
        "environment.yml",
    }:
        return "Программы и код", "Сборка и метаданные пакетов"
    if name in {"sha256sums", "sha512sums", "md5sums", "checksums"}:
        return "Системные и служебные", "Контрольные суммы и подписи"
    if name in {"thumbs.db", "ehthumbs.db", ".ds_store"} or re.fullmatch(
        r"(?:thumbcache|iconcache)_[^.]+\.db", name
    ):
        return "Системные и служебные", "Кэш и временные файлы"
    if name == "py.typed" or name.endswith((".d.ts", ".d.cts", ".d.mts")):
        return "Программы и код", "Объявления типов"
    if name in {
        ".gitignore",
        ".gitattributes",
        ".gitmodules",
        ".dockerignore",
        ".env",
        ".editorconfig",
        ".npmrc",
        ".yarnrc",
    }:
        return "Данные и настройки", "Настройки"
    if name in {".lock", ".pid"}:
        return "Системные и служебные", "Файлы блокировки и состояния"
    if name.endswith(
        (
            ".js.map",
            ".mjs.map",
            ".cjs.map",
            ".css.map",
            ".ts.map",
            ".tsx.map",
            ".jsx.map",
        )
    ):
        return "Программы и код", "Отладка и карты исходного кода"
    if re.match(r"^(?:license|licence|copying|notice)(?:[._-]|$)", name):
        return "Документы", "Лицензии и уведомления"
    if re.match(r"^readme(?:[._-]|$)", name):
        return "Документы", "Обычный текст и разметка"
    if content in {"PEM private key", "PEM public key", "PEM certificate"}:
        return "Системные и служебные", "Сертификаты и ключи"
    if content == "Gettext translation catalog":
        return "Программы и код", "Локализация и переводы"
    if content in {
        "COFF object file",
        "ELF executable",
        "Windows executable / DLL",
        "Mach-O binary",
    }:
        return "Программы и код", "Программы и библиотеки"
    if content == "Program database (PDB)":
        return "Программы и код", "Отладка и карты исходного кода"
    if content == "Wavefront OBJ model":
        return "Проектирование", "3D-модели"
    if content == "MPEG transport stream":
        return "Видео", "Видеозаписи"
    if extension == "mts" and content.startswith("Text ("):
        return "Программы и код", "Исходный код и сценарии"
    if extension == "pot" and content.startswith("PowerPoint "):
        return "Документы", "Презентации"
    if extension == "key" and content.startswith("ZIP "):
        return "Документы", "Презентации"
    # These extensions genuinely have unrelated uses; no single blind mapping.
    if extension in {"key", "pot", "obj", "mts", "pdb", "map"}:
        return "Прочее", "Неоднозначный формат"
    known = EXTENSION_GROUPS.get(extension)
    if known is not None:
        return known
    return "Прочее", "Неизвестное расширение" if extension else "Без расширения"


def recognized_file_role(name: str, content_type: str | None = None) -> str | None:
    """Return a high-confidence role without replacing content classification."""
    lowered = name.casefold()
    extension = extension_of(name)

    if lowered == "thumbs.db":
        return "Windows thumbnail cache (Thumbs.db)"
    if lowered == "ehthumbs.db":
        return "Windows Media Center thumbnail cache (ehthumbs.db)"
    if re.fullmatch(r"thumbcache(?:_[^.]+)?\.db", lowered):
        return "Windows thumbnail cache database"
    if re.fullmatch(r"iconcache(?:_[^.]+)?\.db", lowered):
        return "Windows icon cache database"
    if lowered == "desktop.ini":
        return "Windows folder settings (desktop.ini)"
    if lowered == "autorun.inf":
        return "Windows AutoRun configuration (autorun.inf)"

    if lowered.startswith("~$") and extension in OFFICE_DOCUMENT_EXTENSIONS:
        return "Microsoft Office owner/lock file"
    if lowered.startswith("~") and extension in OFFICE_DOCUMENT_EXTENSIONS:
        return "Probable Microsoft Office temporary file"

    if extension == "lnk" or content_type == "Windows shortcut (LNK)":
        return "Windows shortcut"
    if extension == "url":
        return "Windows Internet shortcut"
    if extension == "tmp":
        return "Temporary file"
    if extension == "asd":
        return "Microsoft Word AutoRecover file"
    if extension == "wbk":
        return "Microsoft Word backup file"
    if extension == "xlk":
        return "Microsoft Excel backup file"
    return None


def refine_named_content_type(name: str, content_type: str) -> str:
    """Add a specific label only when filename and content evidence agree."""
    lowered = name.casefold()
    if content_type == "OLE2 compound file":
        if lowered == "thumbs.db":
            return "Windows thumbnail cache (Thumbs.db / OLE2)"
        if lowered == "ehthumbs.db":
            return "Windows Media Center thumbnail cache (ehthumbs.db / OLE2)"
    return content_type


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


def local_now() -> datetime:
    return datetime.now(tz=UTC).astimezone().replace(tzinfo=None, microsecond=0)


def yes_no(value: bool) -> str:
    return YES if value else NO


def attribute_flag(attributes: int | None, mask: int) -> str | None:
    if attributes is None:
        return None
    return yes_no(bool(attributes & mask))


def windows_attribute_names(attributes: int | None) -> str | None:
    if attributes is None:
        return None
    names = [name for mask, name in WINDOWS_FILE_ATTRIBUTES if attributes & mask]
    known_mask = sum(mask for mask, _ in WINDOWS_FILE_ATTRIBUTES)
    unknown = attributes & ~known_mask
    if unknown:
        names.append(f"UNKNOWN_0x{unknown:08X}")
    return " | ".join(names) if names else "0"


def exception_text(exc: BaseException) -> str:
    name = type(exc).__name__
    if name == "PasswordRequired":
        return "PasswordRequired: password is required"
    message = str(exc).strip()
    if len(message) > 500:
        message = message[:499] + "…"
    return f"{name}: {message}" if message else name


def item_type_from_mode(mode: int) -> str:
    if stat.S_ISREG(mode):
        return "Файл"
    if stat.S_ISDIR(mode):
        return "Папка"
    if stat.S_ISLNK(mode):
        return "Символическая ссылка"
    if stat.S_ISFIFO(mode):
        return "Именованный канал (FIFO)"
    if stat.S_ISSOCK(mode):
        return "Сокет"
    if stat.S_ISCHR(mode):
        return "Символьное устройство"
    if stat.S_ISBLK(mode):
        return "Блочное устройство"
    return "Другой объект файловой системы"


def file_properties(name: str, result: os.stat_result) -> FileProperties:
    attributes = getattr(result, "st_file_attributes", None)
    if attributes is not None:
        hidden = bool(attributes & 0x00000002)
        read_only = bool(attributes & 0x00000001)
    else:
        hidden = name.startswith(".")
        write_bits = stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH
        read_only = not bool(result.st_mode & write_bits)

    created_timestamp = getattr(result, "st_birthtime", None)
    if created_timestamp is None and os.name == "nt":
        created_timestamp = result.st_ctime

    user_id = getattr(result, "st_uid", None)
    group_id = getattr(result, "st_gid", None)
    device_type = getattr(result, "st_rdev", None)
    block_size = getattr(result, "st_blksize", None)
    blocks = getattr(result, "st_blocks", None)
    reparse_tag = getattr(result, "st_reparse_tag", None)
    execute_bits = stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH
    return FileProperties(
        permissions=stat.filemode(result.st_mode),
        hidden=yes_no(hidden),
        read_only=yes_no(read_only),
        executable=yes_no(bool(result.st_mode & execute_bits)),
        system=attribute_flag(attributes, 0x00000004),
        archive_attribute=attribute_flag(attributes, 0x00000020),
        temporary=attribute_flag(attributes, 0x00000100),
        offline=attribute_flag(attributes, 0x00001000),
        compressed=attribute_flag(attributes, 0x00000800),
        encrypted=attribute_flag(attributes, 0x00004000),
        sparse=attribute_flag(attributes, 0x00000200),
        not_content_indexed=attribute_flag(attributes, 0x00002000),
        file_attributes=windows_attribute_names(attributes),
        file_attributes_value=attributes,
        reparse_tag=None if reparse_tag is None else f"0x{reparse_tag:08X}",
        mode=result.st_mode,
        inode=str(result.st_ino),
        device=str(result.st_dev),
        device_type=None if device_type is None else str(device_type),
        links=result.st_nlink,
        user_id=None if user_id is None else str(user_id),
        group_id=None if group_id is None else str(group_id),
        block_size=block_size,
        blocks=blocks,
        accessed_at=timestamp_value(result.st_atime),
        modified_at=timestamp_value(result.st_mtime),
        metadata_changed_at=(
            None if os.name == "nt" else timestamp_value(result.st_ctime)
        ),
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
    integrity_status: str | None = None,
    properties: FileProperties | None = None,
    item_type: str | None = None,
    link_type: str | None = None,
    link_target: str | None = None,
    content_hash: bytes | None = None,
    virtual_folder: bool = False,
    virtual_reason: str | None = None,
) -> ResultRow:
    leaf_name = levels[-1]
    content_type = refine_named_content_type(leaf_name, content_type)
    return ResultRow(
        levels=levels,
        full_path=full_path,
        leaf_name=leaf_name,
        extension="" if virtual_folder or virtual_reason else extension_of(leaf_name),
        size=size,
        item_type=item_type or ("Файл" if properties else "Файл внутри архива"),
        recognized_role=recognized_file_role(leaf_name, content_type),
        link_type=link_type,
        link_target=link_target,
        content_type=content_type,
        in_archive="да" if in_archive else "нет",
        bom=bom,
        integrity_status=integrity_status,
        properties=properties,
        error=error,
        content_hash=content_hash,
        virtual_folder=virtual_folder,
        virtual_reason=virtual_reason,
    )


def folder_placeholder(
    path: str,
    levels: tuple[str, ...],
    *,
    in_archive: bool,
    unreadable: bool = False,
    folder_error: str | None = None,
) -> ResultRow:
    """A logical marker only: never create a file or hash imaginary content."""
    reason = (
        problem_kind(folder_error)
        if folder_error
        else ("uninspected" if unreadable else "empty")
    )
    return make_row(
        path,
        levels + (VIRTUAL_MARKERS[reason][0],),
        0,
        "Folder contents not inspected"
        if unreadable
        else "Folder with no inventoried leaf files",
        in_archive,
        item_type="Папка (виртуальная строка)",
        virtual_folder=True,
        virtual_reason=reason,
    )


def file_digest(path: str, expected_size: int | None = None) -> bytes:
    """Hash complete bytes, rejecting observable changes during the read."""
    digest = blake3()
    total = 0
    with open(extended_path(path), "rb") as stream:
        before = os.fstat(stream.fileno())
        if expected_size is not None and before.st_size != expected_size:
            raise OSError(
                "File changed during inspection; Content fingerprint was not retained"
            )
        remaining = before.st_size
        buffer = bytearray(min(1024 * 1024, max(remaining, 1)))
        view = memoryview(buffer)
        while remaining:
            count = stream.readinto(view[: min(len(buffer), remaining)])
            if not count:
                break
            digest.update(view[:count])
            total += count
            remaining -= count
        grew = bool(stream.read(1))
        after = os.fstat(stream.fileno())
    fields = ("st_size", "st_mtime_ns", "st_ctime_ns", "st_ino", "st_dev")
    if (
        grew
        or any(getattr(before, key) != getattr(after, key) for key in fields)
        or total != before.st_size
        or (expected_size is not None and total != expected_size)
    ):
        raise OSError(
            "File changed during inspection; Content fingerprint was not retained"
        )
    return digest.digest()


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
    normalized = {name.replace("\\", "/").lower() for name in names}
    package_markers = {"[content_types].xml", "_rels/.rels"}

    if package_markers <= normalized and "word/document.xml" in normalized:
        return (
            "Word document (DOCM)"
            if "word/vbaproject.bin" in normalized
            else "Word document (DOCX)"
        )

    if package_markers <= normalized and (
        "xl/workbook.xml" in normalized or "xl/workbook.bin" in normalized
    ):
        if "xl/workbook.bin" in normalized:
            return "Excel binary workbook (XLSB)"
        if "xl/vbaproject.bin" in normalized:
            return "Excel workbook (XLSM)"
        return "Excel workbook (XLSX)"

    if package_markers <= normalized and "ppt/presentation.xml" in normalized:
        return (
            "PowerPoint presentation (PPTM)"
            if "ppt/vbaproject.bin" in normalized
            else "PowerPoint presentation (PPTX)"
        )

    if package_markers <= normalized and "visio/document.xml" in normalized:
        return "Visio document (ZIP package)"

    if "meta-inf/manifest.mf" in normalized:
        return "JAR archive"
    if "meta-inf/container.xml" in normalized and "mimetype" in normalized:
        return "EPUB document"
    if "meta-inf/manifest.xml" in normalized and "mimetype" in normalized:
        return "OpenDocument package"
    if package_markers <= normalized:
        return "Open Packaging Conventions package"
    return "ZIP archive"


def classify_zip_path(file_path: str) -> str:
    try:
        with zipfile.ZipFile(extended_path(file_path)) as archive:
            file_type = classify_zip_names(archive.namelist())
            if extension_of(file_path) not in ARCHIVE_EXTENSIONS:
                hint = ZIP_LEAF_FORMATS.get(extension_of(file_path))
                # Preserve authoritative package-family detection even when
                # a filename claims another format. Hints refine a recognized
                # family or label otherwise unrecognized ZIP-based leaves.
                if file_type == "ZIP archive":
                    return hint or "ZIP container (leaf file)"
                if hint and (
                    file_type == "Open Packaging Conventions package"
                    or (
                        file_type == "Visio document (ZIP package)"
                        and hint.startswith("Visio ")
                    )
                    or (
                        file_type == "OpenDocument package"
                        and hint.startswith("OpenDocument ")
                    )
                ):
                    return hint
            return file_type
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


def refine_grouping_content(name: str, data: bytes, content: str) -> str:
    """Use bounded, already-read headers to resolve common ambiguous suffixes.

    These are format hints, not validators. Never execute/deserialise input or
    inspect package internals to group a leaf. Preserve archive routing labels.
    """
    extension = extension_of(name)
    sample = data[:MAGIC_READ_SIZE]
    text = content.startswith("Text (")
    if sample.startswith(
        (
            b"-----BEGIN PRIVATE KEY-----",
            b"-----BEGIN RSA PRIVATE KEY-----",
            b"-----BEGIN EC PRIVATE KEY-----",
            b"-----BEGIN DSA PRIVATE KEY-----",
            b"-----BEGIN ENCRYPTED PRIVATE KEY-----",
            b"-----BEGIN OPENSSH PRIVATE KEY-----",
        )
    ):
        return "PEM private key"
    if sample.startswith(
        (b"-----BEGIN PUBLIC KEY-----", b"-----BEGIN RSA PUBLIC KEY-----")
    ):
        return "PEM public key"
    if sample.startswith(b"-----BEGIN CERTIFICATE-----"):
        return "PEM certificate"
    if (
        extension in {"po", "pot"}
        and text
        and re.search(rb'^msgid\s+"', sample, re.MULTILINE)
        and re.search(rb'^msgstr\s+"', sample, re.MULTILINE)
    ):
        return "Gettext translation catalog"
    if extension == "obj":
        if len(sample) >= 20:
            machine = int.from_bytes(sample[:2], "little")
            sections = int.from_bytes(sample[2:4], "little")
            optional_size = int.from_bytes(sample[16:18], "little")
            flags = int.from_bytes(sample[18:20], "little")
            if (
                machine in {0x14C, 0x8664, 0xAA64, 0x1C0, 0x1C4}
                and 0 < sections <= 96
                and optional_size == 0
                and not flags & 0x2002
                and 20 + sections * 40 <= len(sample)
            ):
                return "COFF object file"
        if text and re.search(
            rb"^v[ \t]+[-+.0-9]+[ \t]+[-+.0-9]+[ \t]+[-+.0-9]+", sample, re.MULTILINE
        ):
            return "Wavefront OBJ model"
    if extension in {"mts", "ts", "m2ts"}:
        for offset, stride in ((0, 188), (4, 192)):
            if len(sample) > offset + 2 * stride and all(
                sample[offset + i * stride] == 0x47 for i in range(3)
            ):
                return "MPEG transport stream"
    if extension == "pdb" and sample.startswith(
        b"Microsoft C/C++ MSF 7.00\r\n\x1aDS\x00\x00\x00"
    ):
        return "Program database (PDB)"
    if content == "Unknown binary" and sample[:4] in {
        b"\xfe\xed\xfa\xce",
        b"\xce\xfa\xed\xfe",
        b"\xfe\xed\xfa\xcf",
        b"\xcf\xfa\xed\xfe",
    }:
        return "Mach-O binary"
    return content


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
    return refine_grouping_content(file_path, head, file_type), detect_bom(
        head
    ) or BOM_NOT_PRESENT


def detect_file_type(file_path: str) -> str:
    return inspect_file(file_path)[0]


# Physical folder scan -------------------------------------------------------


def link_target(path: str) -> str | None:
    try:
        return regular_path(os.readlink(extended_path(path)))
    except (OSError, ValueError):
        return None


def scan_folder(
    folder_path: str,
    levels: tuple[str, ...],
    excluded_paths: set[str],
) -> tuple[
    list[tuple[str, tuple[str, ...]]],
    list[PhysicalFile],
    list[Diagnostic],
    bool,
]:
    subfolders: list[tuple[str, tuple[str, ...]]] = []
    files: list[PhysicalFile] = []
    diagnostics: list[Diagnostic] = []
    skipped_folder = False

    try:
        with os.scandir(extended_path(folder_path)) as entries:
            for entry in entries:
                logical_path = os.path.join(folder_path, entry.name)
                try:
                    if normalized_key(logical_path) in excluded_paths:
                        continue

                    if os.name == "nt":
                        stat_result = os.stat(
                            extended_path(logical_path),
                            follow_symlinks=False,
                        )
                    else:
                        stat_result = entry.stat(follow_symlinks=False)
                    mode = stat_result.st_mode
                    properties = file_properties(entry.name, stat_result)
                    attributes = getattr(stat_result, "st_file_attributes", None)
                    is_reparse = bool(
                        attributes is not None and attributes & 0x00000400
                    )
                    is_junction_method = getattr(entry, "is_junction", None)
                    is_junction = bool(
                        is_junction_method is not None and is_junction_method()
                    )
                    is_symbolic_link = entry.is_symlink() or stat.S_ISLNK(mode)

                    if is_junction or is_symbolic_link or is_reparse:
                        if is_junction:
                            item_type = "Соединение каталогов (junction)"
                            link_type_value = "Junction"
                        elif is_symbolic_link:
                            item_type = "Символическая ссылка"
                            link_type_value = "Символическая ссылка"
                        else:
                            item_type = "Точка повторной обработки"
                            link_type_value = "Reparse point"
                        files.append(
                            PhysicalFile(
                                logical_path,
                                levels + (entry.name,),
                                stat_result.st_size,
                                properties,
                                item_type,
                                link_type_value,
                                link_target(logical_path),
                                False,
                            )
                        )
                    elif stat.S_ISDIR(mode):
                        subfolders.append((logical_path, levels + (entry.name,)))
                    else:
                        item_type = item_type_from_mode(mode)
                        is_regular = stat.S_ISREG(mode)
                        files.append(
                            PhysicalFile(
                                logical_path,
                                levels + (entry.name,),
                                stat_result.st_size,
                                properties,
                                item_type,
                                (
                                    "Жесткая ссылка"
                                    if is_regular and stat_result.st_nlink > 1
                                    else None
                                ),
                                None,
                                is_regular,
                            )
                        )
                except OSError as exc:
                    diagnostics.append(
                        Diagnostic(
                            logical_path,
                            "entry metadata",
                            exception_text(exc),
                        )
                    )
    except OSError as exc:
        diagnostics.append(
            Diagnostic(folder_path, "folder enumeration", exception_text(exc))
        )
        skipped_folder = True

    if skipped_folder or (not subfolders and not files):
        folder_diagnostic = next(
            (
                item
                for item in reversed(diagnostics)
                if item.stage == "folder enumeration"
            ),
            diagnostics[0] if diagnostics else None,
        )
        files.append(
            PhysicalFile(
                folder_path,
                levels,
                0,
                inspect_content=False,
                virtual_folder=True,
                folder_unreadable=bool(diagnostics),
                folder_error=folder_diagnostic.error if folder_diagnostic else None,
            )
        )
    return subfolders, files, diagnostics, skipped_folder


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
                subfolders, files, diagnostics, skipped_folder = future.result()
                stats.folders += 1
                stats.skipped_folders += int(skipped_folder)
                stats.skipped_entries += sum(
                    diagnostic.stage == "entry metadata" for diagnostic in diagnostics
                )
                stats.diagnostics.extend(diagnostics)
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
    integrity_status: str | None = None,
) -> ResultRow:
    digest = None
    if not error and integrity_status == INTEGRITY_VERIFIED:
        try:
            digest = node.content_hash or file_digest(node.source_path, node.size)
        except OSError as exc:
            error = exception_text(exc)
    return make_row(
        node.logical_path,
        node.levels,
        node.size,
        file_type,
        node.depth > 0,
        error,
        bom=bom,
        integrity_status=(
            integrity_status or (INTEGRITY_ERROR if error else INTEGRITY_NOT_VERIFIED)
        ),
        properties=node.properties,
        content_hash=digest,
    )


def make_temp_path(temp_dir: str, logical_name: str) -> str:
    suffix = ntpath.splitext(logical_name)[1]
    if not re.fullmatch(r"\.[A-Za-z0-9]{1,15}", suffix):
        suffix = ".bin"
    descriptor, path = tempfile.mkstemp(prefix="member-", suffix=suffix, dir=temp_dir)
    os.close(descriptor)
    return path


def archive_temp_directory():
    """Create a short-name archive workspace in the OS temp directory."""
    return tempfile.TemporaryDirectory(
        prefix="folder-review-",
        dir=extended_path(tempfile.gettempdir()),
    )


def raw_archive_type(data: bytes) -> str | None:
    return STREAM_ARCHIVE_TYPES.get(detect_magic(data))


def expandable_archive_type(data: bytes, name: str) -> str | None:
    if raw_extension_of(name) not in ARCHIVE_EXTENSIONS:
        return None
    return raw_archive_type(data)


def should_expand_archive(file_type: str, name: str) -> bool:
    return file_type in ARCHIVE_TYPES and raw_extension_of(name) in ARCHIVE_EXTENSIONS


def member_content_type(data: bytes, name: str) -> str:
    """Label leaf members without staging or opening their internal package."""
    file_type = detect_magic(data)
    if file_type == "ZIP container" and extension_of(name) not in ARCHIVE_EXTENSIONS:
        return ZIP_LEAF_FORMATS.get(extension_of(name), "ZIP container (leaf file)")
    return refine_grouping_content(name, data, file_type)


def read_counted(stream, size: int, budget: WorkBudget) -> bytes:
    budget.check_time()
    data = stream.read(size)
    budget.consume_expanded(len(data))
    return data


def copy_counted(stream, output, budget: WorkBudget) -> None:
    while True:
        chunk = read_counted(stream, 1024 * 1024, budget)
        if not chunk:
            return
        output.write(chunk)


def drain_counted(stream, budget: WorkBudget) -> None:
    while read_counted(stream, 1024 * 1024, budget):
        pass


def inspect_member_stream(
    stream,
    *,
    size: int,
    logical_path: str,
    levels: tuple[str, ...],
    parent_depth: int,
    temp_dir: str,
    budget: WorkBudget,
) -> list[ResultRow]:
    try:
        head = read_counted(stream, MAGIC_READ_SIZE, budget)
    except Exception as exc:
        return [
            make_row(
                logical_path,
                levels,
                size,
                "Unreadable archive member",
                True,
                exception_text(exc),
                integrity_status=INTEGRITY_ERROR,
            )
        ]

    archive_type = expandable_archive_type(head, levels[-1])
    if archive_type is None:
        error = None
        digest = blake3(head)
        total = len(head)
        integrity_status = INTEGRITY_VERIFIED
        try:
            while chunk := read_counted(stream, 1024 * 1024, budget):
                digest.update(chunk)
                total += len(chunk)
            if total != size:
                raise OSError("Archive member size differs from fully read content")
        except Exception as exc:
            error = exception_text(exc)
            integrity_status = INTEGRITY_ERROR
        return [
            make_row(
                logical_path,
                levels,
                size,
                member_content_type(head, levels[-1]),
                True,
                error,
                bom=detect_bom(head) or BOM_NOT_PRESENT,
                integrity_status=integrity_status,
                content_hash=digest.digest() if error is None else None,
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
                integrity_status=INTEGRITY_NOT_VERIFIED,
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
                integrity_status=INTEGRITY_NOT_VERIFIED,
            )
        ]

    temp_path = make_temp_path(temp_dir, levels[-1])
    try:
        with open(temp_path, "wb") as output:
            output.write(head)
            copy_counted(stream, output, budget)
        node = ArchiveNode(temp_path, logical_path, levels, size, depth)
        return process_archive_node(node, temp_dir, budget=budget)
    except Exception as exc:
        return [
            make_row(
                logical_path,
                levels,
                size,
                archive_type,
                True,
                exception_text(exc),
                integrity_status=INTEGRITY_ERROR,
            )
        ]
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass


def archive_info_size(info) -> int:
    return int(
        getattr(info, "file_size", None) or getattr(info, "uncompressed", None) or 0
    )


def archive_member_error_rows(
    node: ArchiveNode,
    infos,
    reason: str,
) -> list[ResultRow]:
    rows = []
    for info in infos:
        location = archive_member_location(node, info.filename)
        if location is None:
            rows.append(
                unusable_archive_member_path_row(
                    node,
                    info.filename,
                    archive_info_size(info),
                    reason,
                )
            )
            continue
        levels, logical_path = location
        rows.append(
            make_row(
                logical_path,
                levels,
                archive_info_size(info),
                "Not inspected",
                True,
                reason,
                integrity_status=INTEGRITY_NOT_VERIFIED,
            )
        )
    return rows


def archive_member_location(
    node: ArchiveNode,
    filename: str,
) -> tuple[tuple[str, ...], str] | None:
    parts = split_archive_path(filename)
    if not parts:
        return None
    return node.levels + parts, node.logical_path + "::" + "\\".join(parts)


def unusable_archive_member_path_row(
    node: ArchiveNode,
    filename: str,
    size: int,
    preceding_error: str | None = None,
) -> ResultRow:
    displayed_name = filename or "<empty member name>"
    path_error = f"Archive member path has no usable components: {filename!r}"
    error = f"{preceding_error}; {path_error}" if preceding_error else path_error
    return make_row(
        node.logical_path + "::" + displayed_name,
        node.levels + ("<unusable archive member path>",),
        size,
        "Not inspected",
        True,
        error,
        integrity_status=INTEGRITY_ERROR,
    )


def prepare_archive_members(
    node: ArchiveNode,
    infos,
    budget: WorkBudget,
    archive_type: str,
) -> list[ResultRow] | None:
    try:
        budget.expand_one_row(len(infos))
    except BudgetExceeded as exc:
        return [archive_row(node, archive_type, exception_text(exc))]

    try:
        budget.ensure_declared_size(sum(archive_info_size(info) for info in infos))
    except BudgetExceeded as exc:
        return archive_member_error_rows(node, infos, exception_text(exc))
    return None


def retain_archive_folders(node, directories, rows, budget) -> list[ResultRow]:
    """Keep terminal listed folders that no existing leaf row represents."""
    paths = {split_archive_path(name) for name in directories}
    paths.discard(())
    ancestors = {parts[:index] for parts in paths for index in range(1, len(parts))}
    represented = {
        row.levels[:index] for row in rows for index in range(1, len(row.levels))
    }
    missing = [
        parts
        for parts in sorted(paths)
        if parts not in ancestors and node.levels + parts not in represented
    ]
    if not missing:
        return rows
    try:
        budget.expand_one_row(len(missing) + 1)
    except BudgetExceeded as exc:
        reason = "Folder placeholders omitted: " + exception_text(exc)
        if rows:
            rows[-1] = replace(rows[-1], error=join_errors(rows[-1].error, reason))
            return rows
        return [archive_row(node, "Archive", reason)]
    return rows + [
        folder_placeholder(
            node.logical_path + "::" + "\\".join(parts),
            node.levels + parts,
            in_archive=True,
        )
        for parts in missing
    ]


def join_errors(first: str | None, second: str) -> str:
    return f"{first}; {second}" if first else second


def expand_zip(
    node: ArchiveNode,
    temp_dir: str,
    budget: WorkBudget,
) -> list[ResultRow]:
    rows: list[ResultRow] = []
    directories = []
    try:
        with zipfile.ZipFile(extended_path(node.source_path)) as archive:
            entries = archive.infolist()
            directories = [info.filename for info in entries if info.is_dir()]
            infos = [info for info in entries if not info.is_dir()]
            if not infos:
                return retain_archive_folders(
                    node,
                    directories,
                    [
                        archive_row(
                            node,
                            "ZIP archive (empty)",
                            integrity_status=INTEGRITY_VERIFIED,
                        )
                    ],
                    budget,
                )
            limit_rows = prepare_archive_members(node, infos, budget, "ZIP archive")
            if limit_rows is not None:
                return limit_rows

            for info in infos:
                location = archive_member_location(node, info.filename)
                if location is None:
                    rows.append(
                        unusable_archive_member_path_row(
                            node,
                            info.filename,
                            info.file_size,
                        )
                    )
                    continue
                levels, logical_path = location
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
                                budget=budget,
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
                            exception_text(exc),
                            integrity_status=INTEGRITY_ERROR,
                        )
                    )
    except Exception as exc:
        return [archive_row(node, "ZIP archive", exception_text(exc))]
    return retain_archive_folders(node, directories, rows, budget)


class MemberCaptureIO(Py7zIO):
    def __init__(
        self,
        expected_size: int,
        filename: str,
        temp_dir: str,
        budget: WorkBudget,
    ):
        self.expected_size = expected_size or 0
        self.filename = filename
        self.temp_dir = temp_dir
        self.budget = budget
        self.head = bytearray()
        self.total_written = 0
        self.mode: str | None = None
        self.archive_type: str | None = None
        self.temp_path: str | None = None
        self.output = None
        self.oversized = False
        self.digest = blake3()

    def write(self, chunk):
        raw = bytes(chunk)
        self.budget.consume_expanded(len(raw))
        self.digest.update(raw)
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
        archive_probe_ready = len(self.head) >= 8 or total_after >= self.expected_size
        if archive_probe_ready:
            self.archive_type = expandable_archive_type(bytes(self.head), self.filename)
            if self.archive_type and self.expected_size <= MAX_NESTED_ARCHIVE_SIZE:
                self.temp_path = make_temp_path(self.temp_dir, self.filename)
                self.output = open(self.temp_path, "wb")  # noqa: SIM115
                self.output.write(self.head)
                self.output.write(remainder)
                self.mode = "archive"
            elif len(self.head) >= MAGIC_READ_SIZE or total_after >= self.expected_size:
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
    def __init__(
        self,
        sizes: dict[str, int],
        temp_dir: str,
        budget: WorkBudget,
    ):
        self.sizes = sizes
        self.temp_dir = temp_dir
        self.budget = budget
        self.products: dict[str, MemberCaptureIO] = {}

    def create(self, filename):
        product = MemberCaptureIO(
            self.sizes.get(filename, 0),
            filename,
            self.temp_dir,
            self.budget,
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
    return archive_member_error_rows(node, infos, reason)


def expand_7z(
    node: ArchiveNode,
    temp_dir: str,
    budget: WorkBudget,
) -> list[ResultRow]:
    infos = []
    directories = []
    factory = None
    extraction_exception: Exception | None = None
    try:
        with py7zr.SevenZipFile(extended_path(node.source_path), mode="r") as archive:
            entries = archive.list()
            directories = [info.filename for info in entries if info.is_directory]
            infos = [info for info in entries if not info.is_directory]
            if not infos:
                if archive.needs_password():
                    return [
                        archive_row(
                            node,
                            "7z archive",
                            "Encrypted 7z headers could not be inspected without a password",
                            integrity_status=INTEGRITY_NOT_VERIFIED,
                        )
                    ]
                return retain_archive_folders(
                    node,
                    directories,
                    [
                        archive_row(
                            node,
                            "7z archive (empty)",
                            integrity_status=INTEGRITY_VERIFIED,
                        )
                    ],
                    budget,
                )
            if len(infos) > SEVEN_ZIP_INSPECT_MAX_MEMBERS:
                return seven_zip_limit_rows(
                    node,
                    infos,
                    f"Content not read: more than {SEVEN_ZIP_INSPECT_MAX_MEMBERS:,} 7z members",
                )
            limit_rows = prepare_archive_members(node, infos, budget, "7z archive")
            if limit_rows is not None:
                return limit_rows

            sizes = {info.filename: info.uncompressed or 0 for info in infos}
            factory = MemberCaptureFactory(sizes, temp_dir, budget)
            try:
                archive.extractall(factory=factory)
            except Exception as exc:
                extraction_exception = exc
            finally:
                factory.close_all()
    except py7zr.exceptions.PasswordRequired:
        if factory is not None:
            factory.close_all()
        return [
            archive_row(
                node,
                "7z archive",
                "Encrypted 7z headers could not be inspected without a password",
                integrity_status=INTEGRITY_NOT_VERIFIED,
            )
        ]
    except Exception as exc:
        if factory is not None:
            factory.close_all()
        if infos:
            return archive_member_error_rows(node, infos, exception_text(exc))
        return [archive_row(node, "7z archive", exception_text(exc))]

    rows: list[ResultRow] = []
    for info in infos:
        location = archive_member_location(node, info.filename)
        if location is None:
            rows.append(
                unusable_archive_member_path_row(
                    node,
                    info.filename,
                    info.uncompressed or 0,
                )
            )
            continue
        levels, logical_path = location
        size = info.uncompressed or 0
        product = factory.products.get(info.filename)
        if extraction_exception is not None:
            extraction_error = exception_text(extraction_exception)
            integrity_status = (
                INTEGRITY_NOT_VERIFIED
                if isinstance(
                    extraction_exception,
                    (BudgetExceeded, py7zr.exceptions.PasswordRequired),
                )
                else INTEGRITY_ERROR
            )
            if product is not None and product.temp_path is not None:
                try:
                    os.unlink(product.temp_path)
                except OSError:
                    pass

            if size == 0:
                content_type = "Empty file"
                bom = BOM_NOT_PRESENT
            elif product is not None and product.head:
                content_type = product.archive_type or member_content_type(
                    bytes(product.head), info.filename
                )
                bom = detect_bom(bytes(product.head)) or BOM_NOT_PRESENT
            else:
                content_type = "Not inspected"
                bom = None
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    size,
                    content_type,
                    True,
                    extraction_error,
                    bom=bom,
                    integrity_status=integrity_status,
                )
            )
        elif size == 0:
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    0,
                    "Empty file",
                    True,
                    bom=BOM_NOT_PRESENT,
                    integrity_status=INTEGRITY_VERIFIED,
                    content_hash=blake3(b"").digest(),
                )
            )
        elif product is None or not product.head or product.total_written != size:
            rows.append(
                make_row(
                    logical_path,
                    levels,
                    size,
                    "Not inspected",
                    True,
                    "Missing or incomplete captured member data",
                    integrity_status=INTEGRITY_ERROR,
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
                    integrity_status=INTEGRITY_NOT_VERIFIED,
                )
            )
        elif product.temp_path:
            child = ArchiveNode(
                product.temp_path,
                logical_path,
                levels,
                size,
                node.depth + 1,
                content_hash=product.digest.digest(),
            )
            try:
                rows.extend(process_archive_node(child, temp_dir, budget=budget))
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
                    member_content_type(bytes(product.head), info.filename),
                    True,
                    bom=detect_bom(bytes(product.head)) or BOM_NOT_PRESENT,
                    integrity_status=INTEGRITY_VERIFIED,
                    content_hash=product.digest.digest(),
                )
            )
    return retain_archive_folders(node, directories, rows, budget)


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


def expand_rar(
    node: ArchiveNode,
    temp_dir: str,
    budget: WorkBudget,
) -> list[ResultRow]:
    rows: list[ResultRow] = []
    directories = []
    try:
        with rarfile.RarFile(extended_path(node.source_path)) as archive:
            entries = archive.infolist()
            directories = [info.filename for info in entries if info.is_dir()]
            infos = [info for info in entries if not info.is_dir()]
            if not infos:
                if archive.needs_password():
                    return [
                        archive_row(
                            node,
                            "RAR archive",
                            "Encrypted RAR headers could not be inspected without a password",
                            integrity_status=INTEGRITY_NOT_VERIFIED,
                        )
                    ]
                return retain_archive_folders(
                    node,
                    directories,
                    [
                        archive_row(
                            node,
                            "RAR archive (empty)",
                            integrity_status=INTEGRITY_VERIFIED,
                        )
                    ],
                    budget,
                )
            limit_rows = prepare_archive_members(node, infos, budget, "RAR archive")
            if limit_rows is not None:
                return limit_rows

            for info in infos:
                location = archive_member_location(node, info.filename)
                if location is None:
                    rows.append(
                        unusable_archive_member_path_row(
                            node,
                            info.filename,
                            info.file_size,
                        )
                    )
                    continue
                levels, logical_path = location
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
                                budget=budget,
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
                            exception_text(exc),
                            integrity_status=INTEGRITY_ERROR,
                        )
                    )
    except Exception as exc:
        return [archive_row(node, "RAR archive", exception_text(exc))]
    return retain_archive_folders(node, directories, rows, budget)


def process_archive_node(
    node: ArchiveNode,
    temp_dir: str,
    known_file_type: str | None = None,
    known_bom: str | None = None,
    *,
    budget: WorkBudget | None = None,
) -> list[ResultRow]:
    if budget is None:
        budget = WorkBudget(rows_claimed=1)
    try:
        budget.check_time()
    except BudgetExceeded as exc:
        return [
            archive_row(
                node,
                known_file_type or "Archive",
                exception_text(exc),
            )
        ]

    if known_file_type is None:
        file_type, bom = inspect_file(node.source_path)
    else:
        file_type, bom = known_file_type, known_bom
    if not should_expand_archive(file_type, node.levels[-1]):
        error = None
        if file_type == "Unreadable":
            error = "File could not be read"
        elif file_type == "ZIP container (unreadable)":
            error = "Invalid or unsupported ZIP container"
        return [
            archive_row(
                node,
                file_type,
                error,
                bom,
                integrity_status=(
                    INTEGRITY_ERROR
                    if error
                    else (
                        INTEGRITY_VERIFIED if node.depth > 0 else INTEGRITY_NOT_VERIFIED
                    )
                ),
            )
        ]

    if node.depth >= MAX_ARCHIVE_DEPTH:
        return [
            archive_row(
                node,
                file_type,
                f"Archive nesting limit reached ({MAX_ARCHIVE_DEPTH})",
            )
        ]

    if file_type == "ZIP archive":
        return expand_zip(node, temp_dir, budget)
    if file_type == "7z archive":
        return expand_7z(node, temp_dir, budget)
    return expand_rar(node, temp_dir, budget)


def process_file(
    item: PhysicalFile,
    budget: WorkBudget | None = None,
) -> list[ResultRow]:
    if budget is None:
        budget = WorkBudget(rows_claimed=1)
    if item.virtual_folder:
        return [
            folder_placeholder(
                item.path,
                item.levels,
                in_archive=False,
                unreadable=item.folder_unreadable,
                folder_error=item.folder_error,
            )
        ]
    if not item.inspect_content:
        if item.link_type is not None:
            content_type = "Filesystem link (not followed)"
        else:
            content_type = "Special filesystem item (not read)"
        return [
            make_row(
                item.path,
                item.levels,
                item.size,
                content_type,
                False,
                properties=item.properties,
                item_type=item.item_type,
                link_type=item.link_type,
                link_target=item.link_target,
            )
        ]

    file_type, bom = inspect_file(item.path)
    if should_expand_archive(file_type, item.levels[-1]):
        with (
            HEAVY_ARCHIVE_SLOTS,
            archive_temp_directory() as temp_dir,
        ):
            node = ArchiveNode(
                item.path,
                item.path,
                item.levels,
                item.size,
                0,
                item.properties,
            )
            return process_archive_node(
                node,
                temp_dir,
                file_type,
                bom,
                budget=budget,
            )

    error = None
    if file_type == "Unreadable":
        error = "File could not be read"
    elif file_type == "ZIP container (unreadable)":
        error = "Invalid or unsupported ZIP container"
    digest = None
    if error is None:
        try:
            digest = file_digest(item.path, item.size)
        except OSError as exc:
            error = exception_text(exc)
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
            item_type=item.item_type,
            link_type=item.link_type,
            link_target=item.link_target,
            content_hash=digest,
        )
    ]


def inspect_files(
    physical_files: list[PhysicalFile],
    workers: int = DEFAULT_WORKERS,
    budget: WorkBudget | None = None,
) -> list[ResultRow]:
    if budget is None:
        budget = WorkBudget()
    budget.initialize_rows(len(physical_files))
    rows: list[ResultRow] = []
    completed = 0
    last_print = time.monotonic()

    items = iter(physical_files)
    with ThreadPoolExecutor(max_workers=workers) as executor:
        pending = set()
        for _ in range(workers * 2):
            try:
                pending.add(executor.submit(process_file, next(items), budget))
            except StopIteration:
                break

        while pending:
            done, pending = wait(pending, return_when=FIRST_COMPLETED)
            for future in done:
                rows.extend(future.result())
                completed += 1
                try:
                    pending.add(executor.submit(process_file, next(items), budget))
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


def apply_physical_row_limit(
    physical_files: list[PhysicalFile],
    max_rows: int,
    stats: ScanStats,
    root_folder: str,
) -> tuple[list[PhysicalFile], int]:
    """Select a deterministic physical subset that can fit the logical-row budget."""
    total = sum(not item.virtual_folder for item in physical_files)
    omitted = max(len(physical_files) - max_rows, 0)
    if omitted == 0:
        return physical_files, total

    stats.skipped_entries += omitted
    omitted_kind = (
        "physical items or folder placeholders"
        if any(item.virtual_folder for item in physical_files[max_rows:])
        else "physical items"
    )
    stats.diagnostics.append(
        Diagnostic(
            root_folder,
            "row limit",
            (
                f"{omitted:,} {omitted_kind} omitted because the logical-row "
                f"budget is {max_rows:,} rows"
            ),
        )
    )
    return physical_files[:max_rows], total


# Excel output ---------------------------------------------------------------


def excel_text(value: str | None) -> str | None:
    if value is None:
        return None
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "�", str(value))
    if len(text) > EXCEL_MAX_TEXT_LENGTH:
        text = text[: EXCEL_MAX_TEXT_LENGTH - 1] + "…"
    return text


def get_column_letter(index: int) -> str:
    return xl_col_to_name(index - 1)


class _ReportSheet:
    """Write report rows sequentially without retaining worksheet cell objects."""

    def __init__(self, workbook, name: str, formats: dict):
        self.workbook = workbook
        self.worksheet = workbook.add_worksheet(name)
        self.formats = formats
        self.row = 0

    def append(self, cells) -> None:
        for column, cell in enumerate(cells):
            if cell is None:
                continue
            value, cell_format = cell
            if value is None:
                result = self.worksheet.write_blank(self.row, column, None, cell_format)
            elif isinstance(value, str):
                result = self.worksheet.write_string(
                    self.row, column, value, cell_format
                )
            elif isinstance(value, datetime):
                result = self.worksheet.write_datetime(
                    self.row, column, value, cell_format
                )
            elif isinstance(value, bool):
                result = self.worksheet.write_boolean(
                    self.row, column, value, cell_format
                )
            else:
                result = self.worksheet.write_number(
                    self.row, column, value, cell_format
                )
            if result:
                raise ValueError(
                    f"Excel cell could not be written: row {self.row + 1}, column {column + 1}, code {result}"
                )
        self.row += 1


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
    # Every string uses write_string(), including formula/URL-like filenames.
    # Cache format objects, not rows or cells. Font/border definitions live for
    # the whole report; alignment is keyed by value to share diagnostic formats.
    alignment = alignment or {}
    key = (
        id(font),
        id(border),
        fill,
        alignment.get("align"),
        alignment.get("valign"),
        alignment.get("text_wrap"),
        number_format,
    )
    cell_format = ws.formats.get(key)
    if cell_format is None:
        options = dict(font)
        options.update(border or {})
        options.update(alignment)
        if fill:
            options.update(pattern=1, fg_color=fill)
        if number_format:
            options["num_format"] = number_format
        cell_format = ws.workbook.add_format(options)
        ws.formats[key] = cell_format
    return value, cell_format


@contextmanager
def _report_workbook(output_path: Path):
    """Close/compress successfully before atomically replacing the report."""
    output_name = absolute_path(output_path)
    output_folder = os.path.dirname(output_name)
    if (
        _protected_root is not None
        and normalized_key(output_name) != _report_output_path
    ):
        raise PermissionError("Only the selected report output may be written.")
    temporary = os.path.join(
        output_folder, f".folder-review-{os.urandom(16).hex()}.tmp.xlsx"
    )
    allowed = {normalized_key(temporary), normalized_key(output_name)}
    _report_write_paths.update(allowed)
    try:
        descriptor = os.open(
            extended_path(temporary), os.O_RDWR | os.O_CREAT | os.O_EXCL, 0o600
        )
        os.close(descriptor)
    except BaseException:
        _report_write_paths.difference_update(allowed)
        raise
    try:
        with xlsxwriter.Workbook(
            extended_path(temporary),
            {
                "constant_memory": True,
                "strings_to_formulas": False,
                "strings_to_urls": False,
                "strings_to_numbers": False,
                "tmpdir": extended_path(tempfile.gettempdir()),
            },
        ) as workbook:
            yield workbook
        os.replace(temporary, extended_path(output_name))
    finally:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        finally:
            _report_write_paths.difference_update(allowed)


def write_workbook(
    rows: list[ResultRow],
    output_path: Path,
    root_folder: str,
    report: RunReport | None = None,
    *,
    full_metadata: bool = False,
) -> None:
    if len(rows) > EXCEL_MAX_DATA_ROWS:
        raise ValueError(
            f"The result has {len(rows):,} rows; one Excel sheet supports at most "
            f"{EXCEL_MAX_DATA_ROWS:,} data rows in this layout."
        )

    rows = [display_row(row) for row in rows]
    content_counts = {}
    for row in rows:
        if row.content_hash is not None and not row.is_virtual:
            key = (row.size, row.content_hash)
            content_counts[key] = content_counts.get(key, 0) + 1
    content_groups = {
        key: index
        for index, key in enumerate(
            (key for key, count in content_counts.items() if count > 1), start=1
        )
    }

    max_level = max((len(row.levels) for row in rows), default=0)
    level_headers = [f"Level {index}" for index in range(1, max_level + 1)]
    all_data_headers = level_headers + [
        "Leaf depth",
        "Full path",
        "Name without extension",
        "File name",
        "Extension",
        "Type group",
        "Type subgroup",
        "Raw suffix",
        "File size (bytes)",
        "File size (MiB)",
        "Duplicate group",
        "Content hash (BLAKE3)",
        "Item type",
        "Virtual row",
        "Inspection state",
        "Inside archive",
        "Hidden",
        "Created at",
        "Modified at",
        "Accessed at",
        "Metadata changed at",
        "Detected role",
        "Detected content type",
        "BOM",
        "Integrity status",
        "Error",
        "Link type",
        "Link target",
        "Permissions",
        "Read-only",
        "Executable",
        "System (Windows)",
        "Archive attribute (Windows)",
        "Temporary (Windows)",
        "Offline (Windows)",
        "Compressed (Windows)",
        "Encrypted (Windows)",
        "Sparse (Windows)",
        "Content indexing disabled (Windows)",
        "File attributes (Windows)",
        "File attributes (st_file_attributes)",
        "Reparse tag (st_reparse_tag)",
        "Mode (st_mode)",
        "File ID (st_ino)",
        "Device ID (st_dev)",
        "Device type (st_rdev)",
        "Link count (st_nlink)",
        "Owner ID (st_uid)",
        "Group ID (st_gid)",
        "Block size (st_blksize)",
        "Allocated blocks (st_blocks)",
    ]
    visible_columns = [
        index
        for index, header in enumerate(all_data_headers)
        if full_metadata or header not in ADVANCED_METADATA_HEADERS
    ]
    data_headers = [all_data_headers[index] for index in visible_columns]
    headers = ["Row number", *data_headers, "last"]

    with _report_workbook(output_path) as wb:
        wb.set_calc_mode("auto", 191029)
        wb.set_size(1484, 843)
        wb.set_tab_ratio(92.8)
        formats = {}
        ws = _ReportSheet(wb, "folder_review", formats)
        sheet = ws.worksheet
        sheet.hide_gridlines(2)
        sheet.set_zoom(85)
        sheet.set_tab_color("#002060")
        sheet.outline_settings(True, False, False, False)
        sheet.set_default_row(14.5)
        sheet.set_portrait()
        sheet.set_paper(9)
        sheet.set_margins(0.7, 0.75, 0.75, 0.75)
        sheet.set_header("", {"margin": 0.3})
        sheet.set_footer("", {"margin": 0.3})

        border = {"border": 1, "border_color": "#000000"}
        normal_font = {
            "font_name": "Calibri",
            "font_size": 11,
            "font_color": Color((1, 0)),
        }
        title_font = {**normal_font, "bold": True, "font_charset": 204}
        header_font = dict(title_font)
        boundary_font = {
            "font_name": "Calibri",
            "font_size": 11,
            "bold": True,
            "font_color": Color((0, 0)),
            "font_charset": 204,
        }
        blue_fill = "#00B0F0"
        black_fill = "#000000"
        green_fill = "#C6E0B4"
        amber_fill = "#FFE699"
        centered = {"align": "center", "valign": "vcenter"}
        wrapped_centered = {**centered, "text_wrap": True}

        first_table_column = 2
        last_table_column = first_table_column + len(headers) - 1
        sheet.set_column(0, last_table_column - 1, 5)
        sheet.freeze_panes("A7")
        sheet.set_row(5, 72)

        ws.append([None])
        root_name = os.path.basename(root_folder.rstrip("\\/")) or root_folder
        title = f"Folder file inventory: {root_name}"
        ws.append(
            [None, styled_cell(ws, excel_text(title), font=title_font, force_text=True)]
        )
        ws.append([None])

        number_row = [None]
        for index, _ in enumerate(headers, start=1):
            number_row.append(
                styled_cell(
                    ws,
                    index,
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
            "Duplicate group",
            "Virtual row",
            "Leaf depth",
            "Extension",
            "Item type",
            "Link type",
            "Permissions",
            "Hidden",
            "Read-only",
            "Executable",
            "System (Windows)",
            "Archive attribute (Windows)",
            "Temporary (Windows)",
            "Offline (Windows)",
            "Compressed (Windows)",
            "Encrypted (Windows)",
            "Sparse (Windows)",
            "Content indexing disabled (Windows)",
            "File attributes (st_file_attributes)",
            "Reparse tag (st_reparse_tag)",
            "Mode (st_mode)",
            "File ID (st_ino)",
            "Device ID (st_dev)",
            "Device type (st_rdev)",
            "Link count (st_nlink)",
            "Owner ID (st_uid)",
            "Group ID (st_gid)",
            "Block size (st_blksize)",
            "Allocated blocks (st_blocks)",
            "BOM",
            "Integrity status",
            "Inside archive",
        }
        number_formats = {
            "File ID (st_ino)": "@",
            "Device ID (st_dev)": "@",
            "Device type (st_rdev)": "@",
            "Owner ID (st_uid)": "@",
            "Group ID (st_gid)": "@",
            "File size (MiB)": "#,##0.000",
            "File size (bytes)": "#,##0",
            "Accessed at": "yyyy-mm-dd hh:mm:ss",
            "Modified at": "yyyy-mm-dd hh:mm:ss",
            "Created at": "yyyy-mm-dd hh:mm:ss",
            "Metadata changed at": "yyyy-mm-dd hh:mm:ss",
        }
        for row_number, row in enumerate(rows, start=1):
            properties = row.properties
            family, subtype = file_type_groups(row)
            state = (
                VIRTUAL_MARKERS[row.virtual_reason or "empty"][1]
                if row.is_virtual
                else VIRTUAL_MARKERS[problem_kind(row.error)][1]
                if row.error
                else "Прочитан"
                if row.content_hash is not None
                else "Ссылка"
                if row.link_type
                else "Не хеширован"
            )
            stem_name = (
                ntpath.splitext(row.leaf_name)[0] if row.extension else row.leaf_name
            )
            values = {
                **{
                    header: row.levels[index] if index < len(row.levels) else None
                    for index, header in enumerate(level_headers)
                },
                "Leaf depth": len(row.levels),
                "Full path": row.full_path,
                "Name without extension": stem_name,
                "Extension": row.extension,
                "Type group": family,
                "Type subgroup": subtype,
                "Raw suffix": raw_extension_of(row.leaf_name)
                if not row.is_virtual
                else None,
                "File name": row.leaf_name,
                "Duplicate group": content_groups.get((row.size, row.content_hash))
                if not row.is_virtual
                else None,
                "Content hash (BLAKE3)": row.content_hash.hex()
                if row.content_hash is not None
                else None,
                "Item type": row.item_type,
                "Virtual row": "да" if row.is_virtual else "нет",
                "Inspection state": state,
                "Detected role": row.recognized_role,
                "Link type": row.link_type,
                "Link target": row.link_target,
                "Permissions": properties.permissions if properties else None,
                "Hidden": properties.hidden if properties else None,
                "Read-only": properties.read_only if properties else None,
                "Executable": properties.executable if properties else None,
                "System (Windows)": properties.system if properties else None,
                "Archive attribute (Windows)": properties.archive_attribute
                if properties
                else None,
                "Temporary (Windows)": properties.temporary if properties else None,
                "Offline (Windows)": properties.offline if properties else None,
                "Compressed (Windows)": properties.compressed if properties else None,
                "Encrypted (Windows)": properties.encrypted if properties else None,
                "Sparse (Windows)": properties.sparse if properties else None,
                "Content indexing disabled (Windows)": properties.not_content_indexed
                if properties
                else None,
                "File attributes (Windows)": properties.file_attributes
                if properties
                else None,
                "File attributes (st_file_attributes)": properties.file_attributes_value
                if properties
                else None,
                "Reparse tag (st_reparse_tag)": properties.reparse_tag
                if properties
                else None,
                "Mode (st_mode)": properties.mode if properties else None,
                "File ID (st_ino)": properties.inode if properties else None,
                "Device ID (st_dev)": properties.device if properties else None,
                "Device type (st_rdev)": properties.device_type if properties else None,
                "Link count (st_nlink)": properties.links if properties else None,
                "Owner ID (st_uid)": properties.user_id if properties else None,
                "Group ID (st_gid)": properties.group_id if properties else None,
                "Block size (st_blksize)": properties.block_size
                if properties
                else None,
                "Allocated blocks (st_blocks)": properties.blocks
                if properties
                else None,
                "File size (MiB)": row.size / (1024 * 1024),
                "File size (bytes)": row.size,
                "Accessed at": properties.accessed_at if properties else None,
                "Modified at": properties.modified_at if properties else None,
                "Created at": properties.created_at if properties else None,
                "Metadata changed at": properties.metadata_changed_at
                if properties
                else None,
                "BOM": row.bom,
                "Integrity status": row.integrity_status,
                "Detected content type": row.content_type,
                "Inside archive": row.in_archive,
                "Error": row.error,
            }

            output_row = [None]
            output_row.append(
                styled_cell(
                    ws,
                    row_number,
                    font=normal_font,
                    border=border,
                    alignment=centered,
                )
            )

            for header, value in zip(
                data_headers, (values[header] for header in data_headers)
            ):
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
                styled_cell(
                    ws,
                    1,
                    font=normal_font,
                    border=border,
                    alignment=centered,
                )
            )
            ws.append(output_row)

        last_row = max(6, 6 + len(rows))
        sheet.autofilter(f"B6:{get_column_letter(last_table_column)}{last_row}")
        sheet.set_selection("B6")

        if report is None:
            now = local_now()
            report = RunReport(
                status="PARTIAL"
                if any(
                    row.error
                    or (row.is_virtual and row.virtual_reason not in (None, "empty"))
                    for row in rows
                )
                else "COMPLETE",
                started_at=now,
                completed_at=now,
                workers=DEFAULT_WORKERS,
                physical_items=sum(not row.is_virtual for row in rows),
                logical_rows=len(rows),
                row_errors=sum(row.error is not None for row in rows),
                stats=ScanStats(),
                expanded_bytes=0,
                max_expanded_bytes=DEFAULT_MAX_EXPANDED_BYTES,
                max_archive_seconds=None,
                max_rows=EXCEL_MAX_DATA_ROWS,
            )

        status_ws = _ReportSheet(wb, "scan_status", formats)
        status = status_ws.worksheet
        status.hide_gridlines(2)
        status.set_tab_color("#70AD47" if report.status == "COMPLETE" else "#FFC000")
        status.freeze_panes(1, 0)
        status.set_column(0, 2, 5)
        status.set_margins(0.75, 0.75, 1, 1)
        status.set_header("", {"margin": 0.5})
        status.set_footer("", {"margin": 0.5})
        status.set_row(16, 32)
        status.set_row(18, 32)

        status_header = [
            styled_cell(
                status_ws,
                "Parameter",
                font=header_font,
                border=border,
                fill=blue_fill,
                alignment=wrapped_centered,
            ),
            styled_cell(
                status_ws,
                "Value",
                font=header_font,
                border=border,
                fill=blue_fill,
                alignment=wrapped_centered,
            ),
        ]
        status_ws.append(status_header)

        max_expanded = (
            "Unlimited"
            if report.max_expanded_bytes is None
            else report.max_expanded_bytes
        )
        max_seconds = (
            "Unlimited"
            if report.max_archive_seconds is None
            else report.max_archive_seconds
        )
        summary_rows = [
            ("Status", report.status),
            ("Root folder", root_folder),
            ("Started at", report.started_at),
            ("Completed at", report.completed_at),
            ("Workers", report.workers),
            ("Physical items", report.physical_items),
            ("Report rows", report.logical_rows),
            ("Folders visited", report.stats.folders),
            ("Skipped folders", report.stats.skipped_folders),
            ("Skipped entries", report.stats.skipped_entries),
            ("Rows with errors", report.row_errors),
            ("Expanded data (bytes)", report.expanded_bytes),
            ("Expanded data limit (bytes)", max_expanded),
            ("Archive time limit (seconds)", max_seconds),
            ("Report row limit", report.max_rows),
            (
                "Inventory notes",
                (
                    "Logical file inventory: folders define the hierarchy; last is always 1. "
                    "Filter Virtual row and Inspection state to select folder/error markers. "
                    "Full path retains the original location. Sizes use bytes and MiB (1 MiB = 1,048,576 bytes); timestamps use local time."
                ),
            ),
            ("Virtual rows", sum(row.is_virtual for row in rows)),
            (
                "Content fingerprints",
                "BLAKE3-256 of complete contents. Only repeated contents receive a Duplicate group ID, consecutive from 1 within this report. Unique, virtual and unhashed rows have blank IDs.",
            ),
            (
                "Rows with fingerprints",
                sum(row.content_hash is not None for row in rows),
            ),
            ("Nonvirtual rows", sum(not row.is_virtual for row in rows)),
            ("Duplicate groups", len(content_groups)),
            ("Duplicate file rows", sum(content_counts[key] for key in content_groups)),
        ]
        for label, value in summary_rows:
            if isinstance(value, datetime):
                number_format = "yyyy-mm-dd hh:mm:ss"
            elif isinstance(value, int):
                number_format = "#,##0"
            elif isinstance(value, float):
                number_format = "#,##0.###"
            else:
                number_format = None
            value_fill = None
            value_font = normal_font
            value_alignment = None
            if label == "Status":
                value_fill = green_fill if value == "COMPLETE" else amber_fill
                value_font = title_font
                value_alignment = centered
            elif label in ("Inventory notes", "Content fingerprints"):
                value_alignment = {"valign": "top", "text_wrap": True}
            status_ws.append(
                [
                    styled_cell(
                        status_ws,
                        label,
                        font=title_font,
                        border=border,
                    ),
                    styled_cell(
                        status_ws,
                        excel_text(value) if isinstance(value, str) else value,
                        font=value_font,
                        border=border,
                        fill=value_fill,
                        alignment=value_alignment,
                        number_format=number_format,
                        force_text=isinstance(value, str),
                    ),
                ]
            )

        status_ws.append([None])
        diagnostic_headers = ("Stage", "Path", "Error")
        status_ws.append(
            [
                styled_cell(
                    status_ws,
                    header,
                    font=header_font,
                    border=border,
                    fill=blue_fill,
                    alignment=wrapped_centered,
                )
                for header in diagnostic_headers
            ]
        )
        diagnostic_total = len(report.stats.diagnostics) + report.row_errors
        diagnostic_iterator = chain(
            report.stats.diagnostics,
            (
                Diagnostic(row.full_path, "archive or content inspection", row.error)
                for row in rows
                if row.error is not None
            ),
        )
        diagnostic_header_row = len(summary_rows) + 3
        max_diagnostic_rows = 1_048_576 - diagnostic_header_row
        displayed_diagnostics = (
            max_diagnostic_rows - 1
            if diagnostic_total > max_diagnostic_rows
            else diagnostic_total
        )

        def append_diagnostic(worksheet_row: int, diagnostic: Diagnostic) -> None:
            status.set_row(worksheet_row - 1, 45)
            status_ws.append(
                [
                    styled_cell(
                        status_ws,
                        excel_text(diagnostic.stage),
                        font=normal_font,
                        border=border,
                        force_text=True,
                    ),
                    styled_cell(
                        status_ws,
                        excel_text(diagnostic.path),
                        font=normal_font,
                        border=border,
                        alignment={"valign": "top", "text_wrap": True},
                        force_text=True,
                    ),
                    styled_cell(
                        status_ws,
                        excel_text(diagnostic.error),
                        font=normal_font,
                        border=border,
                        alignment={"valign": "top", "text_wrap": True},
                        force_text=True,
                    ),
                ]
            )

        if diagnostic_total:
            for worksheet_row, diagnostic in enumerate(
                islice(diagnostic_iterator, displayed_diagnostics),
                start=diagnostic_header_row + 1,
            ):
                append_diagnostic(worksheet_row, diagnostic)
            if diagnostic_total > displayed_diagnostics:
                append_diagnostic(
                    diagnostic_header_row + 1 + displayed_diagnostics,
                    Diagnostic(
                        "",
                        "diagnostics truncated",
                        f"{diagnostic_total - displayed_diagnostics:,} additional diagnostics omitted",
                    ),
                )
        else:
            status_ws.append(
                [
                    styled_cell(
                        status_ws,
                        "No errors",
                        font=normal_font,
                        border=border,
                        force_text=True,
                    ),
                    styled_cell(status_ws, None, font=normal_font, border=border),
                    styled_cell(status_ws, None, font=normal_font, border=border),
                ]
            )

        dax = wb.add_worksheet("dax_reference")
        dax.set_column(0, 0, 5)
        dax.write_string("A1", DAX_MEASURE)
        dax.write_string(
            "A3",
            "Copy the text of A1 from the formula bar into a Power Pivot measure. Edit the first SELECTCOLUMNS line for your table and column. This cell is text, not an Excel formula.",
        )


# Command-line entry point ---------------------------------------------------


def validate_platform() -> None:
    if sys.platform == "darwin":
        raise RuntimeError("macOS is not supported. Use Windows or Linux.")
    if os.name != "nt" and not sys.platform.startswith("linux"):
        raise RuntimeError("This operating system is not supported.")


def worker_count(value: str) -> int | None:
    if value.casefold() == "auto":
        return None
    workers = int(value)
    if not 1 <= workers <= 64:
        raise argparse.ArgumentTypeError("workers must be between 1 and 64")
    return workers


def choose_workers(root: str) -> int:
    """Bounded I/O concurrency, without benchmarking or rereading input files."""
    cpu_count = getattr(os, "process_cpu_count", os.cpu_count)() or 1
    remote = not local_storage(root)
    return min(
        32 if remote else 16, max(8 if remote else 4, cpu_count * (4 if remote else 2))
    )


def pasted_path(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        return value[1:-1]
    return value


@contextmanager
def console_paste_mode():
    """Enable standard console paste/editing while prompting; restore host mode."""
    if os.name != "nt":
        yield
        return
    import ctypes

    api = ctypes.windll.kernel32
    api.GetStdHandle.argtypes = [ctypes.c_ulong]
    api.GetStdHandle.restype = ctypes.c_void_p
    api.GetConsoleMode.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
    api.SetConsoleMode.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
    handle = api.GetStdHandle(-10)
    mode = ctypes.c_ulong()
    changed = bool(api.GetConsoleMode(handle, ctypes.byref(mode)))
    if changed:
        changed = bool(api.SetConsoleMode(handle, mode.value | 0xE7))
    try:
        yield
    finally:
        if changed:
            api.SetConsoleMode(handle, mode.value)


def non_negative_float(value: str) -> float:
    number = float(value)
    if number < 0:
        raise argparse.ArgumentTypeError("value must be zero or greater")
    return number


def row_limit(value: str) -> int:
    limit = int(value)
    if not 1 <= limit <= EXCEL_MAX_DATA_ROWS:
        raise argparse.ArgumentTypeError(
            f"max rows must be between 1 and {EXCEL_MAX_DATA_ROWS:,}"
        )
    return limit


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", help="Override ROOT_FOLDER for this run")
    parser.add_argument("--output", help="Override the XLSX report path (local or SMB)")
    parser.add_argument(
        "--full-metadata",
        action="store_true",
        help="Include diagnostic link, permission, Windows-attribute and st_* columns",
    )
    parser.add_argument(
        "--workers",
        type=worker_count,
        default=None,
        help="Concurrent file workers: auto (default) or 1-64; auto uses CPU capacity and local/network input",
    )
    parser.add_argument(
        "--max-expanded-gib",
        type=non_negative_float,
        default=DEFAULT_MAX_EXPANDED_BYTES / (1024**3),
        help="Maximum total decompressed archive data; 0 disables the limit",
    )
    parser.add_argument(
        "--max-archive-seconds",
        type=non_negative_float,
        default=DEFAULT_MAX_ARCHIVE_SECONDS,
        help="Maximum archive-inspection time; 0 disables the limit",
    )
    parser.add_argument(
        "--max-rows",
        type=row_limit,
        default=EXCEL_MAX_DATA_ROWS,
        help=f"Maximum logical rows (default: {EXCEL_MAX_DATA_ROWS:,})",
    )
    return parser.parse_args()


def run() -> int:
    validate_platform()
    args = parse_args()
    script_dir = Path(os.path.dirname(absolute_path(__file__)))

    configured_root = pasted_path(args.root or ROOT_FOLDER)
    if (
        not args.root
        and configured_root == ROOT_FOLDER
        and ROOT_FOLDER == r"PASTE_FOLDER_PATH_HERE"
    ):
        if not sys.stdin.isatty():
            raise ValueError("Run with --root or start RUN.bat interactively.")
        with console_paste_mode():
            configured_root = pasted_path(
                input(
                    "Folder to inspect (paste with Ctrl+V or Shift+Insert, then Enter): "
                )
            )
        if not configured_root:
            raise ValueError("No folder was entered.")
    root_folder = absolute_path(os.path.normpath(configured_root))
    if not os.path.isdir(extended_path(root_folder)):
        raise ValueError(f"Folder is unavailable: {root_folder}")

    output_path = Path(
        absolute_path(args.output) if args.output else str(script_dir / OUTPUT_FILE)
    )
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Output file must use the .xlsx extension.")
    if not os.path.isdir(extended_path(output_path.parent)):
        raise ValueError(f"Output folder is unavailable: {output_path.parent}")

    if not local_storage(script_dir):
        raise ValueError("Copy the tool to a local folder before running it.")
    validate_output_path(output_path, root_folder)

    if args.workers is None:
        args.workers = choose_workers(root_folder)
    with safe_runtime(root_folder, output_path):
        return scan_and_report(args, root_folder, output_path)


def scan_and_report(
    args: argparse.Namespace, root_folder: str, output_path: Path
) -> int:
    configure_rar_backend()
    started = time.monotonic()
    started_at = local_now()
    print(f"Scanning: {root_folder}")
    print(f"Workers: {args.workers}")
    physical_files, stats = scan_tree(
        root_folder,
        {normalized_key(str(output_path))},
        args.workers,
    )
    print(
        f"Physical items found: {sum(not item.virtual_folder for item in physical_files):,}"
    )
    physical_files, physical_items_found = apply_physical_row_limit(
        physical_files,
        args.max_rows,
        stats,
        root_folder,
    )
    selected_physical = sum(not item.virtual_folder for item in physical_files)
    if selected_physical < physical_items_found:
        print(
            f"Physical items selected for inspection: {selected_physical:,} "
            f"({physical_items_found - selected_physical:,} omitted by row limit)"
        )

    max_expanded_bytes = (
        None if args.max_expanded_gib == 0 else int(args.max_expanded_gib * (1024**3))
    )
    max_archive_seconds = (
        None if args.max_archive_seconds == 0 else args.max_archive_seconds
    )
    budget = WorkBudget(
        max_expanded_bytes=max_expanded_bytes,
        max_rows=args.max_rows,
        deadline=(
            None
            if max_archive_seconds is None
            else time.monotonic() + max_archive_seconds
        ),
    )
    rows = inspect_files(physical_files, args.workers, budget)
    print(f"Logical file rows: {len(rows):,}")

    errors = sum(row.error is not None for row in rows)
    partial = bool(
        stats.skipped_folders or stats.skipped_entries or stats.diagnostics or errors
    )
    report = RunReport(
        status="PARTIAL" if partial else "COMPLETE",
        started_at=started_at,
        completed_at=local_now(),
        workers=args.workers,
        physical_items=physical_items_found,
        logical_rows=len(rows),
        row_errors=errors,
        stats=stats,
        expanded_bytes=budget.expanded_bytes,
        max_expanded_bytes=max_expanded_bytes,
        max_archive_seconds=max_archive_seconds,
        max_rows=args.max_rows,
    )
    print("Writing Excel workbook...")
    write_workbook(
        rows, output_path, root_folder, report, full_metadata=args.full_metadata
    )

    elapsed = time.monotonic() - started
    print(
        f"Done in {elapsed:.1f}s | folders: {stats.folders:,} | "
        f"skipped folders: {stats.skipped_folders:,} | skipped entries: {stats.skipped_entries:,} | "
        f"rows with errors: {errors:,}"
    )
    print(f"Scan status: {report.status}")
    print(f"Result: {output_path}")
    if partial:
        print(
            "WARNING: Partial result. Review the scan_status sheet.",
            flush=True,
        )
        return 2
    return 0


def main() -> int:
    try:
        return run()
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
