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

The output folder must exist. Without `--output`, the report is saved beside
the program. The tool and report may sit inside the scanned folder; the selected
report is excluded from the inventory. Local and SMB report destinations are
supported. Existing reports are atomically replaced after a successful save.
Paste a quoted or unquoted folder path at the prompt (Ctrl+V or Shift+Insert in
Windows), then press Enter.

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
  Groups distinguish source/type declarations, compiled code, build metadata,
  debugging/source maps, templates, translations, scientific arrays, ML models,
  network captures, application resources, logs and licenses. Exact names and
  compound suffixes precede broad suffix mappings (`CMakeLists.txt`, `py.typed`,
  `.d.mts`, `.js.map`). Small headers already read by the scanner disambiguate
  PEM keys, gettext templates, COFF/OBJ models, MPEG transport streams and PDBs.
  These are classification hints, not format validation. No extra reads,
  dependencies, OS registry or deserialization are used. The content-type column
  remains separate. Ambiguous formats have an explicit fallback; `.dat`, `.bin`
  and `.enc` receive generic data labels without claiming their exact format.
  Unknown suffixes and extensionless names retain explicit fallbacks. Virtual
  rows have their own family. Grouping never changes archive expansion.
- XlsxWriter writes both delivery formats in constant-memory mode. Filenames
  remain text, including formula-like names and URLs. Colors, filters, dates,
  hierarchy and diagnostics retain the existing layout; minor width/row-height
  differences from older reports are expected. Literal `_xNNNN_` names and
  carriage returns are preserved correctly in Excel. Some third-party readers
  may expose OOXML escape sequences in these unusual strings.
- Column widths use 5 character units throughout. The inventory freezes rows
  1-6 at A7; no columns are frozen. `dax_reference!A1` contains the supplied DAX
  frequency-list measure as literal text. Copy its text from the formula bar and
  edit the first SELECTCOLUMNS line for your model's table and column.
- Worker selection defaults to `auto`: CPU-based I/O concurrency, 4-16 for local
  inputs and 8-32 for network inputs. This is a bounded heuristic, not a universal
  optimum; `--workers 1..64` remains available. No trial scans or extra input reads
  are performed to choose workers.
- `scan_status` records COMPLETE (exit `0`) or PARTIAL (exit `2`). Fatal errors
  return `1`. Archive integrity failures and inspection limits remain visible.
- No web requests, telemetry or runtime downloads. Existing input files remain
  read-only except the explicitly selected report. Its temporary atomic-save file
  may also be created beside it; saving changes the parent directory timestamp.
  The application and archive/XLSX processing workspace remain local and outside
  the inspected tree; only report output may use SMB. Links
  are listed without being followed. OS access-time updates/cloud hydration may
  still occur. The one-file runtime unpacks to OS temp before scanning.

## Inspect or rebuild

[Scanner source](folder_review.py) · [PyInstaller build and verification](windows_build/README.md)

Releases include matching Python files, licenses, input/artifact hashes,
source-versus-EXE tests and a byte-identical clean-rebuild check.
Windows 11 is tested; Windows 10 remains unverified. MIT license.
