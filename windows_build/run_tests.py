"""Source-runtime release test runner."""

import sys
import unittest
from pathlib import Path

root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root))
from windows_build.suite import run_suite

raise SystemExit(run_suite(unittest.defaultTestLoader.discover(str(root / "tests"))))
