# Security and audit notes

Folder Review processes files locally. It has no updater, telemetry, account,
web service, or automatic runtime installation. Its permanent Python audit hook
rejects sockets/DNS, HTTP requests, browser launching and remote registry calls,
including before and after a scan. SMB uses operating-system file I/O.

Use local storage or SMB UNC/mapped paths for input and reports. URLs, known
Windows WebDAV paths (including mapped drives), and non-SMB Linux network mounts
are rejected. Run the application itself locally. The EXE unpacks beneath the Windows local
profile temporary directory, independent of TEMP/TMP; a remote unpacked runtime
is rejected. Archive processing and XLSX
workspaces stay on local storage outside the inspected tree.

The scanner does not edit, delete, rename, chmod or execute inspected files.
During a review, an audit guard rejects Python mutations of the selected tree,
including through known link/mapped/UNC aliases. The only write exception is the
explicitly selected report and its exact random atomic-save staging file. The
report is excluded from scans and replaced only after a successful workbook close.
Existing output is intentionally replaced; all other target files remain read-only.
Links are listed without traversing their targets. Temporary processing files
are cleaned on normal completion/failure; a crash can leave local temporary files.

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
