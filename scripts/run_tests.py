"""Run scanner/configuration regressions and emit a neutral receipt."""

import argparse
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--results", type=Path, required=True)
args = parser.parse_args()
result = unittest.TextTestRunner(verbosity=2).run(
    unittest.defaultTestLoader.discover(str(ROOT / "tests"))
)
report = {
    "tests": result.testsRun,
    "passed": result.testsRun
    - len(result.failures)
    - len(result.errors)
    - len(result.skipped),
    "failures": len(result.failures),
    "errors": len(result.errors),
    "skipped": [
        {"test": case.id(), "reason": reason} for case, reason in result.skipped
    ],
}
args.results.parent.mkdir(parents=True, exist_ok=True)
args.results.write_text(json.dumps(report, indent=2) + "\n")
allowed = (
    {"POSIX link and FIFO behavior", "Linux mount metadata"}
    if sys.platform == "win32"
    else {"Windows junction behavior", "Windows drive mappings"}
)
unexpected = any(reason not in allowed for _, reason in result.skipped)
raise SystemExit(0 if result.wasSuccessful() and not unexpected else 1)
