# Folder Review

Folder Review creates one flat Excel inventory for a selected folder. It scans
ordinary files and recursively opens ZIP, 7z, and RAR archives in any nesting
order. Archive names and internal folders remain visible as hierarchy levels in
the output.

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

1. Open `folder_review.py` and replace `PASTE_FOLDER_PATH_HERE` in
   `ROOT_FOLDER` with the local, mapped-drive, or UNC folder to inspect, for
   example `r"\\server\share\folder"`.
2. Run `RUN.bat`.

The result is `folder-review.xlsx` beside the script. `RUN.bat` creates a local
`.venv` with uv and installs the locked dependencies automatically.

## Behavior

- Files are classified from their content, not only their extension.
- ZIP, 7z, and RAR archives can contain folders and more supported archives in
  any order.
- DOCX, XLSX, PPTX, EPUB, and JAR ZIP containers remain single logical files.
- Directory links are not followed.
- Local folders, mapped network drives, and Windows UNC paths are supported.
- Folder enumeration and file inspection use 16 workers by default. Archive
  extraction is capped separately to avoid excessive memory use.
- Empty, encrypted, damaged, unsupported, or safety-limited archive members are
  retained as rows with an error instead of stopping the scan.
- Nested archives use temporary local storage and are removed after inspection.

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

Physical files include Windows/Python filesystem properties: hidden status,
`st_mode`, `st_ino`, `st_dev`, `st_nlink`, `st_uid`, `st_gid`, size, access
time, modification time, and creation time when the operating system provides
it. Archive members are logical files rather than OS filesystem objects, so
their OS-only fields remain blank. Their size, type, hierarchy, and BOM are
still reported.

BOM detection is independent from file-type detection and recognizes UTF-8,
UTF-16 LE/BE, and UTF-32 LE/BE markers.

For one-off automation, the configured path can be overridden:

```powershell
.\RUN.bat --root "D:\Data" --output ".\data-review.xlsx" --workers 16
```

On Linux, use the same configured `ROOT_FOLDER` or pass `--root`:

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
