"""Adapt only frozen resource/output locations; reuse the original scanner."""

import sys
from pathlib import Path


def configure_frozen_application(app) -> None:
    if not getattr(sys, "frozen", False):
        return

    # A redirected Windows stream can use a legacy code page. A non-ASCII
    # result path must not turn an otherwise completed scan into exit code 1.
    # This affects diagnostics only; workbook text remains exactly unchanged.
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="backslashreplace")

    # run() derives its default output folder from __file__. In a onefile
    # process that would otherwise point into disposable extracted files.
    app.__file__ = sys.executable
    bundled = Path(sys._MEIPASS) / "7z.exe"
    if not bundled.is_file():
        raise FileNotFoundError("The bundled 7-Zip extractor is missing")

    def configure_rar_backend() -> None:
        # Use one tested backend regardless of older tools on the target PC.
        # Only the extractor location/version differs from the source runtime.
        app.rarfile.SEVENZIP_TOOL = str(bundled)
        app.rarfile.tool_setup(
            force=True,
            unrar=False,
            unar=False,
            bsdtar=False,
            sevenzip=True,
            sevenzip2=False,
        )

    app.configure_rar_backend = configure_rar_backend
