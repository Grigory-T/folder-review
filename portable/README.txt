FOLDER REVIEW 0.20.0 - WINDOWS X86-64

Extract into a writable LOCAL folder. Two tools share one verified runtime:

1. folder-review.bat: create a folder inventory and embedded comparison state.
   Edit folder-review.ini in Notepad. Set folder; keep state_mode=quick for no
   extra state-content reads. Report: folder-review.xlsx (filename configurable).
2. folder-change-review.bat: compare two SAVED XLSX states.
   Edit folder-change-review.ini. Set before and after to their full file paths.
   Report: folder-change-review.xlsx. The original project can be offline.

Keep the first workbook under a separate name before creating the second.
Both states must use the same state_mode (quick, sampled, full). Keep manually
renamed/copied states OUTSIDE the inspected project. Comparison treats archives
as single physical files and excludes temporary files and virtual folder rows.
Unknown means incomplete coverage; review Read issues. The visible inventory
still expands archives. Workbooks older than 0.20 lack comparison data: create
new states. Keep the hidden comparison-state sheet intact; visible cell edits do
not change it. Reports and embedded states are private, unencrypted project data.

Both INI files support literal paths, quotes, percent signs, Unicode, UTF-8 and
UTF-16 with BOM. No escaping, GUI, Python syntax or normal command arguments.
The first entry point used downloads the fixed pinned Python/uv/packages/7-Zip;
later unchanged starts are offline. No update check or dependency resolution.
No global install or administrator installation is needed. All support files,
runtime, caches, logs and temporary processing stay under support/state/.
Reports stay beside the batch files. No AppData/system-temp fallback.

Keep the WHOLE prepared folder when moving it. A move rebuilds its venv from
its retained cache offline. Incomplete cache fails; use a fresh extraction for
online setup/repair rather than mixing releases. Only one operation per folder.
Temporary work cleans normally; crashes can leave work-* under support/state.
Close all instances before cleaning leftovers. Input files stay read-only.
SMB inputs are supported; execute the tool and keep outputs on a local drive.

Exit codes: 0 complete, 2 usable partial report, 1 review/configuration error,
10 setup/environment failure, 30 batch launcher failure. Failed saves preserve
previous reports. Neutral diagnostics omit input names, paths and contents;
console/report data deliberately contains them. Windows may maintain OS state.

See SECURITY.md, LICENSE and THIRD_PARTY.txt. Source, pins, build process and
validation are published with the matching release. Safeguards are application
checks, not an OS sandbox or a guarantee against unknown upstream flaws.
