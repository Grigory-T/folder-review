"""Check the pinned local runtime/package set without invoking uv or networking."""

import importlib.metadata
import json
import re
import site
import sys
from pathlib import Path

pins = json.loads(Path(sys.argv[1]).read_text())
expected_base = Path(sys.argv[2]).resolve()
if ".".join(map(str, sys.version_info[:3])) != pins["python"]["version"]:
    raise SystemExit("Unexpected Python version")
if site.ENABLE_USER_SITE or Path(sys._base_executable).resolve() != expected_base:
    raise SystemExit("The interpreter is not the expected local environment")
actual = {
    re.sub(r"[-_.]+", "-", d.metadata["Name"]).lower(): d.version
    for d in importlib.metadata.distributions()
}
if actual != pins["packages"]:
    raise SystemExit("Installed package versions differ from this fixed release")
print("Local environment: verified")
