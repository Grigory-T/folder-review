# Folder Review

Folder Review creates one Excel inventory for a selected folder. It scans
filesystem items and recursively opens ZIP, 7z, and RAR archives in any nesting
order. Archive names and internal folders remain visible as hierarchy levels.
The workbook contains the flat inventory and an explicit scan-status sheet.

## Requirements

- Windows 10 or 11 (primary platform)
- [uv](https://docs.astral.sh/uv/getting-started/installation/)
- [7-Zip](https://www.7-zip.org/) for RAR extraction

Linux is supported as a secondary platform and needs an available `7z`,
`unrar`, or `unar` command for RAR extraction. macOS is not supported.

Clone the repository into a local folder:

```powershell
git clone https://github.com/Grigory-T/folder-review.git
Set-Location .\folder-review
```

After cloning, only two actions are needed:

1. Run `RUN.bat`.
2. Enter the local, mapped-drive, or UNC folder when prompted.

The result is `folder-review.xlsx` beside the script. `RUN.bat` creates a local
`.venv` with uv and installs the locked dependencies automatically.

## Behavior

- Files are classified from their content, not only their extension.
- ZIP, 7z, and RAR archives can contain folders and more supported archives in
  any order.
- DOCX, XLSX, XLSB, PPTX, EPUB, and JAR ZIP containers remain single logical
  files. Office formats require their real package markers; a ZIP merely
  containing a `word`, `xl`, or `ppt` folder is still expanded as an archive.
- Symbolic links, junctions, reparse points, and special filesystem items are
  reported as rows but are not followed or opened.
- Local folders, mapped network drives, and Windows UNC paths are supported.
- Folder enumeration and file inspection use 16 workers by default. Archive
  extraction is capped separately to avoid excessive memory use.
- Archive members are read to completion. ZIP CRC and backend-provided 7z/RAR
  integrity failures are reported in `Проверка целостности` and `Ошибка`, even
  when a backend reports the failure only after delivering every member byte.
- Empty, encrypted, damaged, unsupported, or safety-limited archive members are
  retained as rows with errors whenever their names are readable.
- Nested archives use temporary local storage and are removed after inspection.

## Completion status

The `scan_status` sheet records `COMPLETE` or `PARTIAL`, run counts, configured
limits, and path-specific folder/archive diagnostics. Exit codes are:

- `0`: complete inventory;
- `2`: usable workbook with partial results or errors;
- `1`: fatal failure.

`RUN.bat` does not print `Completed successfully` for a partial result.

Archive processing is bounded by default to 100 GiB of total decompressed data,
six hours of archive-inspection time, the Excel logical-row limit, 20 nesting
levels, and four concurrently expanded archives. Limits can be changed with
`--max-expanded-gib`, `--max-archive-seconds`, and `--max-rows`; use `0` for an
unlimited byte or time budget. Reaching a limit produces a partial workbook
instead of silent omission. If physical items alone exceed `--max-rows`, the
first path-sorted items that fit are inventoried and the omitted count is
recorded in `scan_status`.

## Long paths and temporary files

On Windows, physical filesystem access uses the extended-length path namespace
for local, mapped-drive, and UNC paths. This applies to folder enumeration,
file metadata, content reads, archive access, workbook staging, replacement,
and cleanup.

Nested archive members are never unpacked into their original hierarchy on
disk. Only a nested archive that must be opened is copied to the local OS temp
directory, using a short random filename; it is deleted automatically after
inspection. The workbook is also written through a short random temporary name
in the output folder and then replaced atomically.

Keep the cloned repository and its `.venv` in a normal short local path. The
folder being inspected may use long local or network paths. Filesystems and
servers can still impose their own limits, and a single Windows path component
normally cannot exceed 255 characters. Excel cells are limited to 32,767
characters, so only a still-longer logical path created by archive nesting may
need to be shortened in the workbook.

Physical items include a readable item type, link type and target, permissions,
hidden/read-only/executable flags, `st_mode`, `st_ino`, `st_dev`, `st_rdev`,
`st_nlink`, `st_uid`, `st_gid`, `st_blksize`, `st_blocks`, size, access time,
modification time, creation time when available, and metadata-change time on
Linux. Regular files with more than one filesystem link are marked as hard
links.

On Windows the workbook also includes the numeric and readable file-attribute
set, reparse tag, and separate system, archive, temporary, offline, compressed,
encrypted, sparse, and content-indexing flags. These fields use metadata already
returned by the operating system; the tool does not enumerate ACLs or alternate
data streams. Archive members are logical files rather than OS filesystem
objects, so their OS-only fields remain blank and their item type is
`Файл внутри архива`. Their size, content type, hierarchy, BOM, integrity
status, and errors are reported.

This is a logical file inventory. Ordinary folders define hierarchy and are not
separate rows. A successfully expanded archive is represented by its final
member rows rather than an additional parent row; do not sum logical member
sizes as physical disk usage.

BOM detection is independent from file-type detection and recognizes UTF-8,
UTF-16 LE/BE, and UTF-32 LE/BE markers.

For automation, pass the folder explicitly. This example uses 32 workers for a
network share:

```powershell
.\RUN.bat --root "\\server\share\folder" --workers 32
```

On Linux, pass `--root`:

```bash
uv sync --frozen --no-dev
uv run python folder_review.py --root /data/folder --workers 16
```

## Development

The implementation intentionally remains in one Python file, separated into
configuration, detection, scanning, archive handling, Excel output, and CLI
sections. Run the regression tests with:

```powershell
uv run python -m unittest discover -s tests -v
```
