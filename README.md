# Folder Review

Local, read-only folder inventory with Excel output. Windows EXE and Python
use the same scanner. Supports local folders, SMB shares and mapped drives.

## Windows — ready to run

[Download portable ZIP](../../releases/latest/download/folder-review-windows-x64.zip) ·
[Releases and checksums](../../releases/latest)

Extract the ZIP to a writable local folder. Keep `folder-review.exe` and
`_internal` together; Python and 7-Zip are included. No installation is needed.

```powershell
.\folder-review.exe --root "C:\Data"
```

Double-click `folder-review.exe` to open the desktop window. Browse for a folder
or paste its path and click **Create Excel report**.
The window shows progress and COMPLETE/PARTIAL status, with an action to open the
report folder. Choose **Quick inventory** (no duplicate hashing), **Find duplicates**
(filtered, default), or **Full fingerprints** (fully reads every candidate/leaf,
potentially slow on SMB). All modes keep archive integrity checks enabled.
Worker and advanced-metadata options use the same scanner as CLI.
The desktop window is also the default for Python/RUN.bat; Python needs Tkinter
in its installation. The standalone EXE bundles Tcl/Tk.

Passing command-line options preserves CLI operation. Use `--cli` for the
terminal folder prompt, or `--gui` on its own for the desktop window.

All application-created files stay in this portable folder:

- `folder-review.exe` and `_internal/`: program and bundled runtime; keep together.
- `folder-review.xlsx`: report, atomically replaced after a successful save.
- `folder-review-data/logs/`: neutral diagnostic JSONL files.
- `folder-review-data/work-*/`: per-run temporary archive/XLSX work; removed on normal exit.

The working directory, TEMP/TMP and AppData are not used for application state
or runtime extraction. Move/copy/delete the whole folder as one unit. A crash
can leave a work folder; close all instances before removing leftovers.
Reports always stay beside the program; `--output` may change the filename in
that same folder. Other destinations are rejected. An unwritable program folder
fails clearly instead of falling back elsewhere. The tool may sit inside the
input: its support/runtime folders and selected report are excluded from scans.
The selected input cannot itself be a tool support/runtime folder.

## Python

Requires Python 3.11+, [uv](https://docs.astral.sh/uv/), and a compatible
[7-Zip](https://www.7-zip.org/) extractor for compressed RAR files.

On Windows, run `SETUP.bat` once, then `RUN.bat`. Setup may download dependencies;
normal launches are offline. Setup selects Python 3.13.16. On Linux:

```bash
uv sync --frozen --no-dev
.venv/bin/python -B folder_review.py --root /data/input --output /data/report.xlsx
```

## Behavior

- Only `.zip`, `.rar` and `.7z` filenames with recognized archive content expand.
  Documents and application packages, including Visio, remain single leaf rows.
  Raw archives renamed `.bin` also remain leaves.
- All headers, type groups/subgroups, flags, inspection states and other
  tool-authored report values are English. File names,
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
  Lossless size and already-read prefix filters avoid full reads of ordinary
  files that cannot be duplicates. Archive members participate in size matching;
  their full hashes are collected during unchanged integrity inspection.
  Candidates still receive native BLAKE3-256 hashes of complete contents; a prefix
  never proves equality. This preserves duplicate matching on unchanged inputs
  while reducing disk/SMB traffic. Links,
  virtual/unreadable rows have no ID; observable file changes produce an error.
  This is not a filesystem snapshot. Raw BLAKE3 is available with `--full-metadata`,
  with blanks for skipped files. Use `--hash-all` to fingerprint every readable
  file, including unique files; this can be much slower on large SMB datasets.
  CLI modes are `--duplicates off`, `--duplicates filtered` (default) and
  `--duplicates all`; `--hash-all` is an alias for the last option. Quick mode
  leaves duplicate IDs and fingerprints blank, including archive members, while
  retaining the same archive integrity checks and PARTIAL diagnostics.
- **`last` is always 1**, including virtual rows. Filter `Virtual row`
  and `Inspection state` to select the rows to count. Virtual leaves are `dummy.dummy`
  (no inventoried files), `deny.deny` (access denied), `password.password`,
  `corrupt.corrupt`, `limit.limit` or `error.error`. No marker file is created.
  Failed locations stay in the hierarchy as parents of their marker; original
  relative paths, errors and known sizes remain. Virtual hash/extension are blank.
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
- The full selected folder path appears once above the inventory. `Relative path`
  is relative to that folder; `::` separates archive members. Virtual rows retain
  the original failed location. `scan_status` also records the selected root and
  uses relative diagnostic paths.
- Inventory columns use 5 character units and 70% zoom. `scan_status` uses
  readable widths for parameters, values and errors. The inventory freezes rows
  1-6 at A7; no columns are frozen. `dax_reference!A1` contains the supplied DAX
  frequency-list measure as literal text. Copy its text from the formula bar and
  edit the first SELECTCOLUMNS line for your model's table and column.
- Worker selection defaults to `auto`: CPU-based I/O concurrency, 4-16 for local
  inputs and 8-32 for network inputs. This is a bounded heuristic, not a universal
  optimum; `--workers 1..64` remains available. No trial scans or extra input reads
  are performed to choose workers.
- A neutral JSONL diagnostic log is enabled by default. Its location is shown in
  progress: `folder-review-data/logs/` beside the application. Logs contain
  phase timings, active-operation ages, workers, counts and application bytes read,
  **no input names, paths,
  content, hashes, machine/account names or exception text**. Heartbeats every
  5 seconds also cover stalled operations. Repeated heartbeats stop after 8 MiB;
  the final record is still attempted. Logging failures warn without discarding
  the report. Keep the log after a slow run; it separates listing, classification,
  metadata queries, hashing, archives and writing. Operation times sum across
  workers and can exceed elapsed time; byte counts are not network-traffic measurements.
  `--log` selects a new file in that same logs folder;
  existing files are never overwritten. Use `--no-log` to disable diagnostics.
- `scan_status` records COMPLETE (exit `0`) or PARTIAL (exit `2`). Fatal errors
  return `1`. Archive integrity failures and inspection limits remain visible.
- No web requests, telemetry or runtime downloads. Existing input files remain
  read-only except the explicitly selected report. Its temporary atomic-save file
  may also be created beside it; saving changes the parent directory timestamp.
  Runtime, logs and archive/XLSX processing files stay beside the application.
  Input may use SMB; reports are local beside the program. URLs, WebDAV and known
  other network filesystems are rejected. Links are listed without being followed.
  OS access-time updates/cloud hydration may still occur. Application guards are
  not an OS sandbox; Windows, antivirus and other processes can maintain their own files.

## Inspect or rebuild

[Scanner source](folder_review.py) · [PyInstaller build and verification](windows_build/README.md)

Releases include matching Python files, licenses, input/artifact hashes,
source-versus-EXE tests and a byte-identical clean-rebuild check.
Windows 11 is tested; Windows 10 remains unverified. Folder Review source is MIT;
dependencies retain their licenses. [Security boundaries](SECURITY.md) ·
[Third-party licenses and corresponding sources](third_party/README.md).
Use `--licenses` on its own to print the bundled notices locally.
Releases include `components.json` and `third-party-sources.zip`.
