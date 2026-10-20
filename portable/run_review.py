"""Run the shared scanner from the clean portable delivery layout."""

import sys
from pathlib import Path

import folder_review as core

SUPPORT = Path(__file__).resolve().parent
ROOT = SUPPORT.parent
core.configure_portable_layout(ROOT, SUPPORT)


def configure_backend():
    backend = SUPPORT / "state/backend/7z.exe"
    if not backend.is_file():
        raise FileNotFoundError(
            "The verified archive backend is missing. Extract a fresh copy and run folder-review.bat."
        )
    core.rarfile.SEVENZIP_TOOL = str(backend)
    core.rarfile.tool_setup(
        force=True,
        unrar=False,
        unar=False,
        bsdtar=False,
        sevenzip=True,
        sevenzip2=False,
    )


core.configure_rar_backend = configure_backend
if __name__ == "__main__":
    if sys.argv[1:]:
        raise SystemExit(
            "Edit folder-review.ini; the portable entry point accepts no scan arguments."
        )
    raise SystemExit(core.main())
