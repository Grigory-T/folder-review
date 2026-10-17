# Build the Windows executable

The EXE imports [folder_review.py](../folder_review.py) directly. Python and
EXE are two deliveries of the same scanner.

## Build

Use Windows x86-64, [uv](https://docs.astral.sh/uv/) and **7-Zip 26.04 x64**.
Check out the release tag and run from a short local path:

```powershell
.\BUILD-EXE.bat
# Or provide a separate extractor directory:
.\BUILD-EXE.bat -SevenZipDir "C:\BuildTools\7zip-26.04"
```

[7zip-pins.json](7zip-pins.json) identifies the official installer and expected
SHA-256 hashes of `7z.exe`, `7z.dll` and `License.txt`. A mismatch stops the build.
No installed extractor is changed. Setup/building may download dependencies;
the finished application makes no web requests or runtime downloads.

[build.ps1](build.ps1) prepares Python **3.13.14 x64**, installs the hash-locked
[uv.lock](../uv.lock) runtime and [build requirements](requirements-build.txt),
then runs [folder-review.spec](folder-review.spec) with PyInstaller **6.22.0**.
Standard DEFLATE level 9 is used. Only builder-only `setuptools`, `pkg_resources`
and `packaging` are excluded; no UPX, stripping or bytecode optimization.

XlsxWriter is the sole report writer in both deliveries. The `test` dependency
group adds openpyxl solely as an independent verification reader; it is not
bundled in the production EXE. To run the source tests: `uv sync --group test`,
then `uv run --group test python -m unittest discover -s tests`.

## Transparency checks

The build publishes `dist/` only after:

- Source and frozen regression suites pass, with unexpected skips rejected.
- [verify.py](verify.py) compares embedded scanner bytecode with current source,
  checks bundled extractor hashes, and compares complete Python/EXE workbooks.
- [reproduce.py](reproduce.py) rebuilds from a fresh source copy in a different
  directory and requires an identical EXE SHA-256.
- [manifest.py](manifest.py) verifies `dist/python/` matches the validated source
  and records source, toolchain, bundled-payload and artifact hashes.

Fixed `PYTHONHASHSEED` and `SOURCE_DATE_EPOCH` values remove randomized bytecode
ordering and PE timestamps, following [PyInstaller's reproducible-build guidance](https://pyinstaller.org/en/v6.22.0/advanced-topics.html#creating-a-reproducible-build).
The spec also sorts standard-library ZIP entries without changing their contents
or metadata; PyInstaller's module-graph order otherwise varies between processes.
Reproduction requires the same runtime and native dependency inputs; their
payload hashes are recorded for comparison.

`dist/` includes the EXE, matching Python files, test/rebuild receipts and licenses.
Build paths and machine identifiers are not included in public receipts.
When building from a source archive without `.git`, `FOLDER_REVIEW_SOURCE_COMMIT`
may supply the release's full commit ID; input hashes remain independently recorded.

## Packaging behavior

The adapter places default output beside the EXE, selects bundled 7-Zip for RAR,
and escapes unprintable console diagnostics. Scanner logic is shared unchanged.
The runtime unpacks to OS temp and is normally removed on exit. Inputs containing
that runtime are rejected. The EXE and selected report may be inside input.
Reports may use SMB; only the selected report and its atomic-save staging file
are permitted write exceptions. Processing workspaces stay local outside input.
The report is excluded from the inventory. Other input writes remain blocked.

Windows 11 is tested. Synthetic tests cannot cover every possible input;
archive time limits remain cooperative rather than forcibly interrupting every
backend operation.
