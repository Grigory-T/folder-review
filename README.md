# Folder Review

Local, read-only folder inventory and saved-state comparison with Excel output. Supports local folders,
SMB shares and mapped drives. Documents remain files; ZIP, RAR and 7z archives
are inspected recursively. XlsxWriter is the only report writer.

[Download the small Windows ZIP](../../releases/latest/download/folder-review-windows-x64.zip) ·
[Releases and checksums](../../releases/latest)

Extract into a writable local folder. Open **folder-review.ini** in Windows Notepad,
set `folder`, save, then double-click **folder-review.bat**. Use that same batch file for
every scan. Use **folder-change-review.bat** with its separate INI to compare two saved states. No GUI, system Python, uv or 7-Zip installation is required.

The first start downloads the fixed, checksum-verified runtime and packages.
Later unchanged starts run offline without invoking uv or checking for updates.
User settings do not change the installed Python, uv or dependency versions.

```text
folder-review.bat          prepare once if needed, then create an inventory/state
folder-review.ini          folder, inventory and state settings
folder-change-review.bat   compare two saved XLSX states; same shared runtime
folder-change-review.ini   before/after full paths and comparison report filename
support/                  code, notices and all generated runtime/state
folder-review.xlsx        inventory/state after a successful scan
folder-change-review.xlsx comparison report
```

All application-created files stay in that folder. Runtime, cache, setup history,
neutral scan logs and temporary work stay under `support/state/`. There is no
AppData/system-temp fallback. Temporary scan work is removed on normal exit;
a crash can leave `work-*` directories. Close all instances before cleaning them.
Keep the whole folder together when copying/moving it: after a move the launcher
recreates its venv from the retained cache **offline**, preserving settings and
reports. If that cache is incomplete, use a fresh extraction with online setup.

Both reports always stay beside the batch files. Its filename can be changed in settings.
An existing report is atomically replaced only after a successful save. The tool
may be inside input: its support folder and selected report are excluded. Other
input files stay read-only. Selecting support itself as input is rejected. SMB
input remains supported; the tool and report stay local. Links are listed without
following their targets. OS access-time updates/cloud hydration may still occur.

## Settings

`folder-review.ini` uses literal paths: paste backslashes as-is, without escaping.
Quoted paths are accepted. Percent signs and Unicode remain literal. UTF-8 and
UTF-16 with BOM are supported. Unknown/duplicate settings and invalid values fail
before scanning. Keep the `[review]` section; edit values only.

| Setting | Meaning / default |
|---|---|
| `folder` | Required folder path; blank in the download |
| `state_mode` | `quick` default; `sampled` content samples; `full` complete SHA-256 |
| `duplicates` | `filtered` default; `off` quick inventory; `all` full fingerprints |
| `workers` | `auto`, or 1–64 |
| `report` | `folder-review.xlsx`; filename only |
| `full_metadata` | `no`; `yes` adds advanced filesystem/raw fingerprint columns |
| `diagnostics` | `yes`; bounded neutral diagnostic logs |
| `max_expanded_gib` | `100`; `0` disables the archive byte limit |
| `max_archive_seconds` | `21600`; `0` disables the cooperative archive time limit |
| `max_rows` | `1048570`; Excel inventory-row limit |

Filtered duplicate detection fully hashes only possible equal-content candidates,
using lossless size/prefix filters. Only repeated complete contents get consecutive
numeric `Duplicate group` IDs. Unique/unhashed/virtual rows have blank IDs. Full
fingerprints read every readable leaf and can be slow over SMB. Quick inventory
skips duplicate hashing. All modes retain archive CRC/integrity checks; archive
members participate in duplicate matching when enabled. Content IDs belong to
one report; concurrent input changes are not a filesystem snapshot.

## Compare two saved states

1. Create the first state with folder-review.bat. Keep that XLSX under a distinct
   filename; the next inventory run replaces its configured report.
2. Later create a second state of the same project with the same `state_mode`.
3. In folder-change-review.ini, set `before` and `after` to the two full XLSX paths.
   Quoted, Unicode and SMB paths are accepted. Run folder-change-review.bat.
4. Open folder-change-review.xlsx. Comparison only reads the saved workbooks;
   the original project can be offline. No live recapture or GUI is involved.

`report` is a filename beside the batch files; `diagnostics=yes/no` controls the
neutral aggregate log. Invalid input, incompatible modes or a failed write keep
an existing comparison report unchanged. Output cannot replace an input state.

New inventories contain a hidden, versioned comparison-state worksheet. It stores
precise integer timestamps, sizes, coverage, exclusions and optional fingerprints
as checksummed compressed JSON text. No additional runtime dependency is needed.
Keep this worksheet intact. Changing visible inventory cells does not change the
recorded state. Workbooks from before 0.20 have no compatible state and are rejected;
create two new states. Old standalone .fcrstate histories are not imported.
States are not encrypted and contain original paths/names; treat them as private.

Original folder-change-review rules are preserved: exact full-path union, no
same-root validation or rename detection; a move is a deletion plus an addition.
ZIP/RAR/7z are opaque physical files in comparison, even while the visible
inventory expands them. Empty-folder/virtual rows and links do not become files.
Temporary files (`~*`, Thumbs.db, *.tmp, *.temp), .states, *.fcrstate and reserved
old comparison artifacts are excluded. Recorded tool-support/output exclusions
apply to both states. The two selected state workbooks are also omitted at matching recorded paths,
including renamed copies. Keep other manually renamed/copied historical workbooks
outside the inspected project so those extra copies do not appear as project files.

| State mode | Same-path change detection | Additional state reads |
|---|---|---:|
| `quick` | Size and nanosecond modification time | None; reuse enumeration metadata |
| `sampled` | Size and SHA-256 of first/middle/last 64 KiB | At most 192 KiB/file |
| `full` | Size and full SHA-256 | Every byte |

These modes are independent of duplicate detection. Quick can miss edits that
preserve size/time. Sampled can miss edits outside its samples; small files are
read completely. Timestamp-only edits count as Changed only in Quick. Full reads
are explicitly opt-in and can add substantial SMB I/O. Live capture is not an
atomic filesystem snapshot. The saved physical state is captured before visible
row/archive limits; archive-inspection errors can make the inventory PARTIAL even
when its opaque physical-file state has complete metadata coverage. Comparison
state coverage is listed separately in scan_status.

The comparison report retains the original **Changes** layout, 1/blank
Added/Deleted/Changed flags, Size A/B/delta in bytes/MiB, `last=1`, width 5,
70% zoom and A7 freeze. The common root appears once. Incomplete state coverage
produces **Unknown**, blank change flags/delta, a warning and **Read issues** sheet;
known changes outside affected scopes are still reported. Even zero Unknown rows
can accompany unreadable scopes. Comparison exit 0 means complete, 2 means usable
partial, 1 means fatal. The union is held in memory and must fit one worksheet.
Decoded state JSON is limited to 256 MiB and its compressed payload to 64 MiB;
oversize or corrupt states fail explicitly, without partial report replacement.

## Report and coverage

- Only ZIP/RAR/7z names with recognized archive content expand. Office, Visio,
  OpenDocument, EPUB and application packages remain leaves, including inside
  archives. A raw archive renamed `.bin` stays a leaf.
- Technical labels/values are English. Original names, relative paths and OS
  diagnostics retain their source text. The full selected root appears once
  above the inventory. `::` separates archive members.
- `last` is always 1. Virtual markers retain empty/inaccessible/problem locations:
  `dummy.dummy`, `deny.deny`, `password.password`, `corrupt.corrupt`, `limit.limit`
  and `error.error`. No physical placeholder file is created.
- Filter `Virtual row` / `Inspection state` when using pivots. Two file-type
  grouping levels support analysis; unknown types remain visible. Hidden is kept;
  advanced low-value filesystem columns are optional.
- Sizes use bytes and MiB. Inventory columns are width 5, zoom 70%, frozen at A7.
  `scan_status` uses readable widths. `dax_reference!A1` contains the frequency-list
  measure as literal text for manual use.
- **0 COMPLETE**; **2 PARTIAL**, usable report with coverage errors; **1 fatal**
  review/configuration error. Setup errors return **10**, launcher errors **30**.
  Read `scan_status` after a PARTIAL run. Archive/header/password/path errors are
  retained, including encrypted-header archives that cannot be inspected.

Workers remain the existing bounded CPU/local/network heuristic, with no trial
scan. Diagnostics contain aggregate timings/counts/bytes and operation ages,
no input names, paths, hashes, contents, machine/account identifiers or exception
text. Scan heartbeats cover stalls; they do not interrupt Windows SMB timeouts.
Setup logs record fixed infrastructure categories/statuses only. App output is
shown in the console rather than captured into the setup journal.

## Fixed setup and audit

This delivery adapts the reviewed
[portable-python application launcher](https://github.com/Grigory-T/portable-python).
The fixed toolchain is **CPython 3.13.16**, **uv 0.13.1**, **7-Zip 26.04 x64**, and
the existing dependency versions in `uv.lock`. Pins are in
[portable/pins.json](portable/pins.json). Exact hashed requirements are exported
from that lock; user machines never resolve a new graph or build source packages.

uv's official archive/executable are verified before execution. Its pinned
catalog selects/checksums the exact official CPython archive; core Python hashes
are checked independently. Wheels come from official PyPI with required hashes,
no source builds and no custom index overrides. Official pinned `7zr.exe` (x86,
setup-only; standard Windows x64 WoW64) extracts just the existing 64-bit 7-Zip
backend and license from the verified installer. The installer is never executed.
Every extracted backend file is independently hash-checked. No system registration
or administrator install is performed. No auto-updater or telemetry is included.

Build the deterministic source-only ZIP with:

```bash
python scripts/build_portable.py --output dist --release
```

The builder requires uv for a frozen export consistency check and a clean source
checkout for `--release` (or `--source-commit` for a GitHub source archive).
It ships only reviewed text files and a source manifest;
no executables, venvs, caches or run data. Historical PyInstaller builds remain
fully available at tag `v0.18.0`; GUI/PyInstaller delivery is no longer maintained.
Nuitka remains a local experiment.

Developer source runs keep the existing CLI: install the frozen dependencies,
then `python -B folder_review.py --root /data/input`. With no arguments it reads
folder-review.ini. Run regression tests with the `test` dependency group; openpyxl is
an independent test reader, never the production writer. Native bootstrap/report
acceptance lives in `tests/windows_portable.ps1`.

[Security boundaries](SECURITY.md) · [Download/license notices](portable/THIRD_PARTY.txt).
Original source is MIT; downloaded components retain their own licenses. Windows
11 x64 is tested; Windows 10 remains unverified. Application protections are not
an OS sandbox. Hashes/provenance checks do not guarantee absence of unknown flaws.
