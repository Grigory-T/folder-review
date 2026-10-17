# Third-party materials

Folder Review's own code is MIT licensed. Dependencies retain their own licenses.
`licenses/` contains notices for the pinned runtime and archive/report libraries.
The runtime-provider collection includes some optional components; `components.json`
in a release identifies the actual native payloads and their verified origins.
Copyright names in third-party notices must be retained.

`third-party-sources.zip` in each release contains complete, unmodified corresponding
source archives for 7-Zip and the LGPL Python libraries: py7zr, multivolumefile,
inflate64, pybcj and pyppmd. Exact upstream URLs and SHA-256 hashes are in
[sources.json](sources.json). The build downloads and verifies these archives;
normal application execution does not download anything.

You may inspect, modify and rebuild these components under their licenses,
including reverse engineering needed to debug your modifications. The supplied
hash locks protect the published build; they do not prohibit your own changes.
For a modified build, update the relevant source/wheel hash locks and rebuild with
`windows_build/build.ps1`. Source archives contain their upstream build scripts.
The original EXE checksum is expected to change. No signing or encryption prevents
replacing the dependencies in a rebuilt application.

7-Zip's unRAR restriction forbids recreating the proprietary RAR compression
algorithm from that code; retaining RAR decompression does not change Folder
Review's own MIT license. The full LGPL 2.1 text and 7-Zip notices accompany the
release. Python's license also includes conditions for Microsoft runtime code;
those binaries are distributed only for Windows.

Use `folder-review.exe --licenses` (or `python folder_review.py --licenses`) to
print bundled notices locally. The EXE contains the license texts even when
obtained separately from the full package.
