# Security model

Folder Review inventories files and compares saved states. It is a normal user
application for a trusted Windows or Linux host, not a forensic image, antivirus
scanner or security sandbox. Run without administrator/root privileges in a
**private, writable local folder**. Other users must not be able to modify that
folder, its installed runtime or the selected data while a review is running.
Do not use a shared or cloud-synchronized folder for the application or reports.

## Operations and trust

| Phase | Reads and writes | Network |
|---|---|---|
| First setup | Verify packaged code; download fixed uv, CPython, wheels and 7-Zip into support/state; create the local environment | Explicit HTTPS downloads from pinned official sources/PyPI |
| Inventory | Read the selected tree and archive contents; write the report and temporary work in the application folder | No Internet APIs, telemetry, updater or dependency resolution; SMB uses OS file I/O |
| Comparison | Read only the two selected XLSX states; write the comparison report | Same restriction; saved project roots and workbook external links are never opened |
| Relocation | Recreate the environment from the retained local cache | Offline; an incomplete cache fails closed |

The user trusts the downloaded release, launcher, manifest, host, native parsers
and installed runtime. Packaged hashes detect changes against that trusted
manifest. They do not authenticate a manifest an attacker can replace, or defend
against an attacker who can modify the launcher/runtime. Source/developer runs
also require a trusted Python environment and source checkout.

### Fixed setup

The launcher is adapted from portable-python (MIT). It verifies uv's archive and
executable before execution. The pinned uv catalog must select the exact expected
CPython URL. The launcher downloads and checks the **entire CPython archive**
against pins.python.sha256, then tells uv to install from that local file mirror
in offline mode. Core Python executable/DLL hashes are also checked independently.
Version/mirror selectors are not user settings. Inherited UV_* overrides are
cleared; TLS verification is never disabled. Exact hashed wheels are installed
with --require-hashes --no-deps --only-binary :all: from official PyPI.

The official 7-Zip installer and setup-only 7zr.exe have pinned hashes. Only
7z.exe, 7z.dll and License.txt are extracted; their individual hashes are checked.
The installer is never executed. Standard Windows x64 WoW64 supports the x86
setup decoder. There is no global installation, registry registration or updater.
Proxy settings remain supported; bootstrap may use current Windows proxy
credentials. uv does not promise transparent PAC/Negotiate proxy support.

Only the two root TXT configs are editable packaged inputs. Before starting any
Python process, the launcher checks immutable packaged files and rejects extra
files/directories under support, including shadow packages, loose bytecode and
__pycache__. The only generated subtree allowed there is state/. This closes
accidental/unlisted import shadowing; it is **not** a complete installed-file
tamper audit. Ready/version checks do not hash every stdlib, wheel, .pth or venv
file on every start. Keep the whole folder private. Extract a fresh release for
updates/repair rather than mixing versions. Hashes identify upstream artifacts;
they do not establish that upstream code has no vulnerabilities.

### Release authenticity

Checksums on a GitHub release are useful for transfer and reproducibility checks,
but the release channel can replace both a ZIP and its checksums. From v0.21.1,
release tags and SHA256SUMS.txt also have SSH signatures. The public key is in
[security/release-signing-key.pub](https://github.com/Grigory-T/folder-review/blob/main/security/release-signing-key.pub)
and is attached to the release. **Verify its fingerprint through a channel you
already trust or a previously pinned copy.** Fetching a new key from the same
possibly compromised release does not independently establish authenticity.
Signatures do not make malicious or vulnerable code safe.

For users who already trust that key, on a system with OpenSSH and sha256sum:

```sh
printf 'folder-review %s\n' "$(cat release-signing-key.pub)" > allowed_signers
ssh-keygen -Y verify -f allowed_signers -I folder-review -n folder-review-release \
  -s SHA256SUMS.txt.sig < SHA256SUMS.txt
sha256sum -c SHA256SUMS.txt
```

A source rebuild can be compared byte-for-byte with the ZIP. Git tag verification
requires a locally trusted SSH allowed-signers configuration. The batch launcher
uses process-local PowerShell ExecutionPolicy Bypass to run its local script;
it changes no persistent policy and provides no publisher authentication. It is
not Authenticode-signed. Never disable endpoint protection to run the tool.

## Input scope and writes

Reports stay beside the batch files. State, downloads, optional neutral logs and
workspaces stay under support/state; no AppData/system-temp fallback is used.
An exclusive workspace lock serializes both tools. Normal report staging is in
the existing owned work-* workspace, on the same volume, followed by atomic
replacement after a successful close. Failed saves preserve an existing report.
A crash can leave partial work under support/state. After closing all instances,
you may remove leftover work-* directories. Legacy .folder-review-<32 hex>.tmp.xlsx
files at the application's root are excluded from self-scans, not automatically
deleted. Low-level writer calls outside the guarded product flow can still stage
beside their selected output.

The scanner opens source content read-only. Python audit guards reject other
source writes, deletes, renames, permission/time changes and unapproved commands,
including known mapped/UNC aliases. If installed inside the selected input, only
its own support and report writes are allowed and excluded from that inventory.
Selecting support itself as input is rejected. Archive member names are data,
never extraction paths or commands. Archive subprocesses are limited to the
reviewed backend and read-only operations. Formula/URL-like names are literal
Excel text. Invalid XML characters are replaced; unpaired filename surrogates
are visibly escaped as \uXXXX. Hidden state preserves the original names.

Links, junctions and reparse points are listed without following targets. Linux
submounts, including same-device bind mounts, are skipped using a snapshot of
/proc/self/mountinfo taken before traversal. Skipped scope is visible, yields
PARTIAL and propagates Unknown in comparison rather than false deletions.
Select a desired mounted folder as the root to review it directly. Root network
storage supports SMB, not known NFS/FUSE/WebDAV providers or URLs. Mounts/paths
changed concurrently can race these checks; this is not kernel-enforced isolation.
SMB may itself contact a remote server via the OS; there is no promise of zero
network traffic or an allowlist of server addresses. OS redirectors/providers
remain trusted.

## Coverage, private data and resource limits

**COMPLETE means that the requested operations completed without recorded coverage
errors. It does not mean verified contents, a consistent snapshot, or absence of
malware.** Quick state uses size/time and can miss same-size/time-preserving edits.
Sampled state reads at most 192 KiB (three 64 KiB samples). Normal duplicates use
at most 24 KiB and are explicitly possible matches; unsampled differences can be
missed. Full modes read complete content but still cannot freeze a changing tree.
Archive checksums do not validate document semantics or prove safety.

All Windows reparse points, including cloud placeholders, are omitted from the
comparison state under the original rules. They remain visible inventory rows.
A hydration/reparse-status change may therefore appear as an addition/deletion.
The tool neither hydrates them deliberately nor promises complete cloud-file
comparison. OS access-time changes, recall by filesystem providers and antivirus
activity remain outside application control. For strict read-only acquisition,
use filesystem/SMB permissions or a read-only snapshot.

Reports and hidden states are **unencrypted private data**: selected roots,
original names, exact paths/timestamps, sizes and optional fingerprints. Hidden
does not mean secret. Keep reports and crash leftovers in an access-controlled
location. Console output contains paths/names; optional bounded diagnostics use
fixed aggregate counters/timings, not names, paths, hashes, contents or exception
text. Setup journal and diagnostics are off by default. Installed runtime/venv
metadata can contain its own local installation paths; it stays in the tool folder.

Saved-state comparison parses only required literal OOXML parts and checksummed,
compressed JSON/base64, not pickle, formulas, macros or executable serialization.
XML/package/compressed/decoded limits and strict validation reject unsupported,
corrupt, ambiguous or oversized states. Checksums detect accidental corruption,
not forgery by a party able to replace the whole workbook. Comparison cannot
replace either input, including known filesystem aliases.

Archive parsing processes untrusted bytes in-process/native libraries. The
100 GiB expanded-byte and 6-hour cooperative limits preserve normal workload
compatibility; they are **not** CPU/RAM isolation. ZIP/RAR metadata and 7z decoder
dictionaries may allocate memory before those limits apply. A package advisory
check does not audit CPython, 7-Zip, native dependencies, the OS or unknown flaws.
Use lower configured limits and an OS sandbox/VM for hostile archives. This
release does not introduce a new archive sandbox or claim hostile-input hardening.

## Reporting

Windows 11 x64 is tested; Windows 10 is unverified; Linux is a source/developer
platform. A compromised host or writable runtime is outside the threat model.
Report concerns privately through GitHub Security advisories if available, or
provide a synthetic reproduction in an issue. Do not upload real reports, file
lists, private paths, machine/user identifiers or credentials.

## Independent-review disposition (v0.21.1)

| Finding | Assessment / action |
|---|---|
| F1 unusual Unicode filenames | Reproduced and fixed in both visible writers; exact state names retained; invalid XML noncharacters also handled |
| F2 unlisted support imports | Reproduced and rejected before Python starts; a writable host/runtime remains outside the integrity boundary |
| F3 Linux submounts | Confirmed; mount-table boundaries now skipped with PARTIAL/Unknown, including same-device bind mounts |
| F4 release authenticity | Valid channel-trust limitation; signed tags/checksums added, with explicit requirement to trust the key independently |
| F5 unused CPython archive pin | Confirmed verification gap, not evidence of compromise; full pin is now checked before local/offline installation |
| F6 crash staging | Valid privacy/self-scan concern; normal staging moved under owned support work; old root staging excluded and retained for manual review |
| F7 cloud/reparse comparison | Correct description of inherited exclusions; documented without changing comparison rules |
| F8 parser/resource risks | Valid residual risk, not proof of a specific exploit; cooperative limits retained, hostile parsing requires OS isolation |
| F9 COMPLETE / quick modes | Valid interpretation risk; coverage versus verification stated explicitly; defaults and mode semantics unchanged |
| F10 races / installed tampering | Valid trusted-host limits; no claim of OS isolation or full runtime attestation |

Two details in the review need correction: the v0.21.0 ZIP has 18 source payloads
plus its generated manifest (19 entries), and sampled state can read three 64 KiB
ranges, 192 KiB total. Device IDs alone would not detect same-device bind mounts.
The review's absence-of-malicious-code observation is not a proof of safety.
