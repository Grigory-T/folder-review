FOLDER REVIEW 0.19.1 - WINDOWS X86-64

1. Extract the complete ZIP into a writable local folder.
2. Open settings.ini in Windows Notepad. Set folder to the directory to inspect.
   A local path, UNC SMB path, or mapped SMB drive is supported. Save the file.
3. Double-click SETUP.bat. It prepares the fixed runtime once, then scans.
   Use this same file for every later scan. Later unchanged starts are offline.

No Python, uv, 7-Zip installation or administrator account is required.
The first setup needs Internet access. Setup verifies fixed official downloads;
Python packages are installed from a hash-locked, wheel-only requirements file.
There is no update check or target-side dependency resolution.

The top level contains SETUP.bat, settings.ini, support/ and the generated report.
Report: folder-review.xlsx, or another filename configured in settings.ini.
All code, runtime, caches, setup history and scan diagnostics stay under support/.
Temporary scan work is cleaned on normal completion. A crash can leave work-*.
Close all instances before reviewing/removing leftovers. No AppData/OS-temp fallback.
The support folder is technical: edit only settings.ini. Notepad UTF-8 and UTF-16
with a BOM are supported. Paths are literal, including backslashes and percent signs.
There is no GUI and no folder prompt in normal use.

SETUP.bat exit codes: 0 COMPLETE, 2 PARTIAL report, 1 review/configuration failure,
10 setup/environment failure, 30 launcher failure. A PARTIAL report is usable;
read scan_status for missing coverage and archive errors. Existing reports are
atomically replaced only after a successful save. Only one instance per tool folder.

All previous scanner/report behavior remains: ZIP/RAR/7z inspection, documents as
leaf files, exact duplicate modes, virtual rows, English technical fields and XLSX.
Duplicates: off is quick inventory, filtered is the default lossless candidate
filter, all fully fingerprints readable files. Archive CRC checks stay enabled.
User input remains read-only; the selected report and owned support state are
write exceptions when the tool itself is inside input. Runtime/support/report
are excluded from such self-scans. Do not select support itself as input.

Copy the WHOLE folder when moving it. A changed location rebuilds only the local
venv from the retained cache, offline; settings and reports are preserved.
Do not remove the cache if you need offline relocation. If the copied cache is
incomplete, use a fresh extraction and online setup instead of overlaying versions.
Never mix a prepared state with another tool release. No latest-version selectors.

support/state/setup.jsonl records fixed infrastructure events/exit codes only.
support/state/logs/ contains neutral aggregate scan diagnostics. No input names,
paths, contents, hashes or machine/account names are recorded in these logs.
Console/report data can contain the selected paths and original names; keep it private.
Windows, antivirus and other programs can maintain their own state; this tool is
not an OS sandbox. Checksum/provenance checks do not prove absence of unknown bugs.

Original code: MIT. See LICENSE, THIRD_PARTY.txt and SECURITY.md in support/.
Sources and fixed download hashes are in the matching GitHub release/repository.
For developer use, folder_review.py also supports its existing command-line options.
