"""Run the original tests and reject unexpected skips during release builds."""

import argparse
import json
import sys
import unittest
from pathlib import Path


def run_suite(suite):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--results", required=True, type=Path)
    args = parser.parse_args()
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    report = {
        "tests": result.testsRun,
        "passed": result.testsRun
        - len(result.skipped)
        - len(result.errors)
        - len(result.failures),
        "skipped": [
            {"test": case.id(), "reason": reason} for case, reason in result.skipped
        ],
        "failures": len(result.failures),
        "errors": len(result.errors),
    }
    args.results.parent.mkdir(parents=True, exist_ok=True)
    args.results.write_text(json.dumps(report, indent=2), encoding="utf-8")
    unexpected_skips = [
        reason
        for _, reason in result.skipped
        if reason not in {"POSIX link and FIFO behavior", "Linux mount metadata"}
    ]
    if unexpected_skips:
        print(
            "Release validation rejected unexpected test skips:",
            unexpected_skips,
            file=sys.stderr,
        )
    return 0 if result.wasSuccessful() and not unexpected_skips else 1
