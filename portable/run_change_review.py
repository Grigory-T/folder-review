"""Comparison entry point using the same fixed portable environment."""

from pathlib import Path

import folder_change_review as changes
import folder_review as core

SUPPORT = Path(__file__).resolve().parent
core.configure_portable_layout(SUPPORT.parent, SUPPORT)
if __name__ == "__main__":
    raise SystemExit(changes.main())
