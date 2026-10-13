Neutral synthetic fixtures for Windows EXE compatibility checks.

rar4-stored.rar is a minimal RAR4 stored archive built from documented header
fields and CRCs, with a text member and a nested ZIP. It was checked with
7-Zip and rarfile independently.

RAR5 files were generated with WinRAR 7.23 from repeated ASCII payloads,
a nested ZIP containing a UTF-8 BOM text file, and deterministic pseudorandom
bytes. The password on encrypted fixtures is fixture-password.
rar5-corrupt.rar has one stored payload byte flipped to exercise CRC reporting.
The split archive has four 20 KiB volumes (the final volume is smaller).

The synthetic-* OLE files were generated with Windows structured-storage COM.
They are valid OLE containers containing only the stream names used by the
scanner's classifier. They are deliberately not usable Office documents.
No private documents, malware, macros, or executable payloads are included.

leaf-packages.rar contains 77 synthetic ZIP package shapes under common leaf
suffixes (including all six Visio suffixes), an unknown suffix, and a Visio
package renamed .zip. Its stored RAR5 encoding supports older Linux extractors.
leaf-packages-compressed.rar contains the same shapes with modern WinRAR solid
compression; Windows production-EXE comparisons cover this variant separately.
Linux 7-Zip 23.01 cannot decode that newer compression method and correctly
produces errors; it is not used as a portable Linux unit-test fixture.
The package bodies contain only neutral XML, zero bytes and an intentionally
unreadable inner ZIP to catch accidental traversal. They are classification
fixtures, not complete Office or design documents.
