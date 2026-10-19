# Security and audit notes

Folder Review processes files locally. It has no updater, telemetry, account,
web service, or automatic runtime installation. Its permanent Python audit hook
rejects sockets/DNS, HTTP requests, browser launching and remote registry calls,
including before and after a scan. SMB uses operating-system file I/O.

Use local storage or SMB UNC/mapped paths for input. Reports stay beside the program. URLs, known
Windows WebDAV paths (including mapped drives), and non-SMB Linux network mounts
are rejected. Run the application itself from a writable local folder.
The standard PyInstaller folder build loads its runtime from `_internal` beside
the EXE; it does not self-extract into AppData or OS temp. Application-owned logs
and per-run processing workspaces stay in `folder-review-data` beside the program,
independently of working directory, TEMP/TMP or profile paths. No external fallback
is used. Python source uses the same layout beside folder_review.py.

The scanner does not edit, delete, rename, chmod or execute inspected files.
During a review, an audit guard rejects Python mutations of the selected tree,
including through known link/mapped/UNC aliases. Write exceptions are the
report beside the application, its exact random atomic-save staging file, and the
application-owned folder-review-data subtree. Runtime/support directories and the
selected report are excluded when the tool folder is within input. Selecting a
support/runtime folder itself as input is rejected. The
report is excluded from scans and replaced only after a successful workbook close.
Existing output is intentionally replaced; all other target files remain read-only.
Links are listed without traversing their targets. Temporary processing files
are cleaned on normal completion/failure; a crash can leave local temporary files.

Neutral performance logs stay in folder-review-data/logs beside the program. They contain only
tool-owned categories, aggregate counts/timings and byte totals, with no input
names/paths/content/fingerprints, host/account identifiers or exception messages.
The log path is displayed in progress, not written inside the log. Logs are not
uploaded. Creation is exclusive; a log never replaces an existing file. Logging
can be disabled with --no-log. Diagnostic heartbeats do not interrupt OS I/O or
change operating-system SMB timeouts.

These are application protections, not an OS sandbox. OS access-time updates,
cloud hydration, antivirus, filesystem redirectors and a compromised host are
outside their control. Saving a report changes its parent directory timestamp.
A concurrently modified tree is not a consistent filesystem snapshot. Archive
budgets are bounded, but time limits are cooperative, not process-level deadlines.
For hostile data or strict enforcement, use a read-only share/account or an
isolated machine with outbound network policy allowing only the required SMB.

# Components, provenance and verification

Runtime dependencies are version/hash locked in uv.lock; build tools are locked
in windows_build/requirements-build.txt. The exact Windows CPython archive and
7-Zip files are hash-pinned separately. Every file in the installed Python runtime
is checked against that archive. Code files in installed wheels are checked
against RECORD, and all bundled native payloads must match a runtime-archive,
wheel, or pinned 7-Zip origin. An unknown native payload stops release preparation.

components.json records versions, roles, native origins and hashes. The build
manifest records source, inputs, component/source bundles and EXE identity.
Source/frozen tests, complete workbook parity and a fresh-directory byte-identical
rebuild must pass before dist/ is published. See windows_build/README.md.

A release security-verification.json records the date, known-vulnerability check
and local antivirus result/availability. These checks can miss unknown flaws and
malicious upstream code. Hashes prove identity, not that code is harmless.
The EXE is unsigned; no verified publisher or malware-free guarantee is claimed.
Windows 11 is tested; Windows 10 remains unverified.

# Reporting a concern

Use the repository Security advisory reporting feature if available, or open an
issue containing a minimal synthetic reproduction. Do not post real file lists,
reports, paths, credentials or other private data in a public issue.
