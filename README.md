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
- Reports include hierarchy, sizes, dates, content types, hidden flags and errors.
  Add `--full-metadata` for advanced filesystem columns.
- `SHA-256` identifies equal complete file contents, including files inside
  archives. For duplicate pivots, exclude blank hashes, group by `SHA-256`, and
  sum `last`. Hashing reads whole files and increases disk/SMB traffic. Links,
  virtual rows and incompletely read content have no hash. Observable changes
  during hashing produce an error; this is not a filesystem snapshot.
- Terminal folders without inventoried leaves have a virtual `dummy.dummy`
  row, also inside archives. No dummy file is created. Its item type identifies
  a folder, its path points to that folder, its hash/extension are blank and
  `last=0`; ordinary file rows keep `last=1`. Unreadable folders remain unknown
  and produce PARTIAL, even when a virtual row retains their location.
- `Расширение` filters recognizable package-cache version/ABI suffixes,
  generated numeric timestamps and suffixes longer than 64 characters.
  Unknown ordinary suffixes, long legitimate suffixes and `.1` manual pages
  remain. `--full-metadata` includes `Исходный суффикс`; filenames are unchanged.
  These display heuristics do not infer extensions from content or change
  which archive formats may expand.
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
