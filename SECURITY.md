# Security and audit notes

Folder Review has two explicit phases. First-time setup uses Internet access to
fetch fixed reviewed components. Subsequent unchanged starts perform local checks
and run the scanner offline; they do not invoke uv, resolve dependencies, install
packages or poll GitHub. There is no updater or telemetry. Python networking is
permanently disabled by the scanner before importing its third-party libraries.
SMB input uses operating-system file I/O rather than Python sockets.

## Setup trust and downloads

The source-only ZIP contains the batch/PowerShell launcher, shared scanner,
settings template, fixed hashes/requirements and notices. Verify its release
checksum and review sources when authenticity matters. The local reviewed
launcher and manifest are trusted inputs; they are not independent signatures
or protection against a local attacker who can replace them together.

The bootstrap is adapted from portable-python's MIT application launcher. It
uses TLS verification, Windows certificate trust and approved official sources.
uv archive/executable hashes are pinned before execution. CPython selection is
fixed in the pinned uv catalog, which verifies the downloaded distribution;
CPython's executable/core DLL hashes are independently checked. Version selectors
are not user settings. Requirements include exact versions/artifact hashes, and
uv installs wheels only with --require-hashes and --no-deps from official PyPI.
Inherited UV_* index/mirror/environment overrides are cleared. Proxy settings
remain supported; TLS verification is never disabled. No origin credentials are
embedded or saved. PowerShell/BITS bootstrap can use the current proxy's Windows
credentials; uv's own client does not guarantee PAC/Negotiate compatibility.

The official 7-Zip installer and setup-only 7zr.exe decoder have pinned release
hashes. The decoder extracts only 7z.exe, 7z.dll and License.txt into local state;
each output is independently verified. The installer is not run. Scan-time RAR
inspection uses the same 64-bit backend as previous releases. Standard Windows
x64 WoW64 is needed for the x86 setup decoder. No global Python/uv/7-Zip, registry
registration or administrator install is used.

Wheels and downloaded interpreters retain their own code-execution capabilities
and licenses. Hashes establish identity, not that upstream code is harmless.
The runtime/package set is fixed rather than guaranteed vulnerability-free forever.
Use a new reviewed release/fresh extraction for upgrades; do not mix prepared state.
The ready marker and version/core checks are not a full installed-file tamper audit.

## Local state and protected input

Run the complete tool on a writable local fixed drive, not a share, junction or
symlink. All state, caches, downloads, optional setup journal/scan logs and workspaces stay
under support/state beside folder-review.bat. There is no AppData/system-temp fallback.
Immutable tool-file hashes are checked on launch; only folder-review.txt and folder-change-review.txt are editable.
A workspace lock prevents concurrent setup/scans in the same folder. A moved
folder rebuilds its venv from retained local cache offline; incomplete cache fails
instead of silently going online. Settings and existing reports are preserved.

Reports stay beside folder-review.bat and are atomically replaced after a successful close.
Only the selected report, its exact random staging file and owned support state
are permitted writes when the tool is inside input; runtime/support/report are
excluded from that scan. Other input mutations/commands are rejected by Python
audit guards, including known mapped/UNC/link aliases. Links are listed without
following targets. Selecting support itself as input is rejected before creating
scan work. Runtime archive commands are allowlisted read-only operations.

Input may be local or SMB. URLs, known WebDAV paths and known non-SMB Linux network
filesystems are rejected. Source/developer APIs use the same scanner and guards.
Normal package settings cannot redirect report/logs/runtime elsewhere.

## Privacy, limitations and reporting

Neutral scan logs record aggregate timings/counts/bytes, not input names, paths,
contents, hashes, machine/account identifiers or exception text. When diagnostics are enabled, the setup journal
records fixed categories and exit/status codes, never command output or exception
messages. App console output and the report deliberately preserve original paths
and names: keep them private. Generated venv/runtime metadata can contain its own
local installation paths; it stays inside the tool folder and is not published.

These are application protections, not an OS sandbox. Windows/antivirus/Explorer
can maintain their own state. Access-time updates, cloud hydration, filesystem
redirectors and a compromised host are outside the tool's control. Concurrently
modified input is not a consistent snapshot. Archive time limits are cooperative.
For strict enforcement, use read-only SMB/filesystem permissions and appropriate
OS outbound policy. Never disable endpoint protections to run the tool.

Windows 11 x64 is tested; Windows 10 is unverified. Source is auditable; no signed
publisher or malware-free guarantee is claimed. Report a concern privately through
GitHub Security advisories if available, or supply a synthetic reproduction in an
issue. Do not publish real reports, file lists, usernames, hosts or credentials.

## Saved-state comparison

Comparison reads only the two explicitly selected XLSX files (local or SMB),
never the recorded project roots or workbook external links. It parses only
required OOXML parts and literal state cells; no formulas/macros are executed.
The versioned state uses checksummed compressed JSON/base64, not pickle or code.
Package/XML/compressed/decoded limits and strict record/coverage validation reject
unsupported, corrupt, ambiguous or oversized states. Checksums detect accidental
changes; they are not signatures against a party able to replace the whole file.
The same offline, local-workspace, atomic-output and read-only guards apply.
Comparison output cannot overwrite either input, including known filesystem aliases.
The two TXT config files are the only editable packaged inputs. Both entry points share
the exclusive workspace lock and exact pinned runtime. No GUI, EXE, alternate
capture engine or user machine history search is included. Historical states must
be kept outside a scanned project if renamed/copied into arbitrary XLSX filenames.

Normal duplicate mode uses bounded samples and labels matches possible; it is not proof of byte equality. Quick duplicate detection and diagnostics-off are defaults. Timestamped reports preserve prior states; collisions receive a numeric suffix. Only the two TXT config files are editable.
