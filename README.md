# Folder Review

Local, read-only folder inventory with Excel output. Windows EXE and Python
use the same scanner. Supports local folders, SMB shares and mapped drives.

## Windows — ready to run

[Download EXE](../../releases/latest/download/folder-review.exe) ·
[Download full package](../../releases/latest/download/folder-review-windows-x64.zip) ·
[Releases and checksums](../../releases/latest)

The x86-64 EXE includes Python and 7-Zip. Copy it to a
local folder and run it, or use:

```powershell
.\folder-review.exe --root "C:\Data" --output "C:\Reports\folder-review.xlsx"
```

The output folder must exist and be outside the scanned tree. Without
`--output`, the report is saved beside the program; unsafe defaults prompt
for another path.

## Python

Requires Python 3.11+, [uv](https://docs.astral.sh/uv/), and a compatible
[7-Zip](https://www.7-zip.org/) extractor for compressed RAR files.

On Windows, run `SETUP.bat` once, then `RUN.bat`. Setup may download dependencies;
normal launches are offline. On Linux:

```bash
uv sync --frozen --no-dev
.venv/bin/python -B folder_review.py --root /data/input --output /data/report.xlsx
```

## Behavior

- Only `.zip`, `.rar` and `.7z` filenames with recognized archive content expand.
  Documents and application packages, including Visio, remain single leaf rows.
  Raw archives renamed `.bin` also remain leaves.
- All headers and tool-authored `scan_status` text are English. File names,
  paths and original diagnostics retain their original text. Columns run from
  hierarchy and adjacent names through types, sizes, duplicates, flags, dates
  and inspection details; advanced filesystem fields come last. File sizes use
  bytes and MiB (1 MiB = 1,048,576 bytes); timestamps use local time.
- Reports include hierarchy, sizes, dates, content types, hidden flags and errors.
  Add `--full-metadata` for advanced filesystem columns.
- `Duplicate group` is filled only for repeated complete contents. Matching
  contents share a consecutive numeric ID from 1; unique files stay blank.
  IDs apply only within one report. For duplicate pivots, exclude blank IDs, group by this column
  and sum `last`. `scan_status` gives duplicate group and duplicate file-row counts.
  Native BLAKE3-256 hashes full contents, reusing a read buffer and existing
  archive decompression streams. Hashing adds disk/SMB traffic. Links,
  virtual/unreadable rows have no ID; observable file changes produce an error.
  This is not a filesystem snapshot. Raw BLAKE3 is available with `--full-metadata`.
- **`last` is always 1**, including virtual rows. Filter `Virtual row`
  and `Inspection state` to select the rows to count. Virtual leaves are `dummy.dummy`
  (no inventoried files), `deny.deny` (access denied), `password.password`,
  `corrupt.corrupt`, `limit.limit` or `error.error`. No marker file is created.
  Failed locations stay in the hierarchy as parents of their marker; original
  full paths, errors and known sizes remain. Virtual hash/extension are blank.
  Errors and unreadable folders still produce PARTIAL.
- `Extension` filters recognizable package-cache version/ABI suffixes,
  generated numeric timestamps and suffixes longer than 64 characters.
  Unknown ordinary suffixes, long legitimate suffixes and `.1` manual pages
  remain. `--full-metadata` includes `Raw suffix`; filenames are unchanged.
  These display heuristics do not infer extensions from content or change
  which archive formats may expand.
- `Type group` and `Type subgroup` add two grouping levels above the
  extension: documents → spreadsheets, images → raster/vector, design → CAD/
  diagrams, software → code/installers, data → tables/databases/settings, etc.
  Fixed MIME-style families and common filename conventions work identically
  across machines. They classify suffixes, not verified content; the existing
  content-type column remains separate. Unknown suffixes and extensionless
  names have explicit fallback groups. Virtual rows have their own family.
- XlsxWriter writes both delivery formats in constant-memory mode. Filenames
  remain text, including formula-like names and URLs. Colors, filters, dates,
  hierarchy and diagnostics retain the existing layout; minor width/row-height
  differences from older reports are expected. Literal `_xNNNN_` names and
  carriage returns are preserved correctly in Excel. Some third-party readers
  may expose OOXML escape sequences in these unusual strings.
- `scan_status` records COMPLETE (exit `0`) or PARTIAL (exit `2`). Fatal errors
  return `1`. Archive integrity failures and inspection limits remain visible.
- No web requests, telemetry or runtime downloads. Input is read-only to the
  application; reports and temporary files use local storage outside it. Links
  are listed without being followed. OS access-time updates/cloud hydration may
  still occur. The one-file runtime unpacks to OS temp before scanning.

## Inspect or rebuild

[Scanner source](folder_review.py) · [PyInstaller build and verification](windows_build/README.md)

Releases include matching Python files, licenses, input/artifact hashes,
source-versus-EXE tests and a byte-identical clean-rebuild check.
Windows 11 is tested; Windows 10 remains unverified. MIT license.
