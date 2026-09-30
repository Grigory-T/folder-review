# Folder Review

Folder Review creates one flat Excel inventory for a selected folder. It scans
ordinary files and recursively opens ZIP, 7z, and RAR archives in any nesting
order. Archive names and internal folders remain visible as hierarchy levels in
the output.

## Requirements

- Windows 10 or 11
- [uv](https://docs.astral.sh/uv/getting-started/installation/)
- [7-Zip](https://www.7-zip.org/) for RAR extraction

Clone the repository into a local folder:

```powershell
git clone https://github.com/Grigory-T/folder-review.git
Set-Location .\folder-review
```

After cloning, only two actions are needed:

1. Open `folder_review.py` and replace `PASTE_FOLDER_PATH_HERE` in
   `ROOT_FOLDER` with the folder to inspect.
2. Run `RUN.bat`.

The result is `folder-review.xlsx` beside the script. `RUN.bat` creates a local
`.venv` with uv and installs the locked dependencies automatically.

## Behavior

- Files are classified from their content, not only their extension.
- ZIP, 7z, and RAR archives can contain folders and more supported archives in
  any order.
- DOCX, XLSX, PPTX, EPUB, and JAR ZIP containers remain single logical files.
- Directory links are not followed.
- Empty, encrypted, damaged, unsupported, or safety-limited archive members are
  retained as rows with an error instead of stopping the scan.
- Nested archives use temporary local storage and are removed after inspection.

For one-off automation, the configured path can be overridden:

```powershell
.\RUN.bat --root "D:\Data" --output ".\data-review.xlsx"
```

## Development

The implementation intentionally remains in one Python file, separated into
configuration, detection, scanning, archive handling, Excel output, and CLI
sections. Run the regression tests with:

```powershell
uv run python -m unittest discover -s tests -v
```
