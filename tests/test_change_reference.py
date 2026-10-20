"""Whole-workbook golden parity with the reviewed original comparison writer."""

import hashlib
import json
import re
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import folder_change_review as changes
import folder_review as core


class OriginalReportParityTests(unittest.TestCase):
    def test_all_cells_styles_types_and_package_parts_match_original_six_cases(self):
        reference = json.loads(
            (Path(__file__).parent / "fixtures/change-reference.json").read_text()
        )
        with (
            tempfile.TemporaryDirectory() as d,
            patch.object(core, "application_directory", return_value=Path(d)),
        ):
            for case in reference["cases"]:
                with self.subTest(case=case["name"]):
                    a, b = case["before"], case["after"]
                    rows = changes.compare_snapshots(a, b)
                    self.assertEqual(rows, case["expected_rows"])
                    errors = [
                        (label, e)
                        for label, s in (("A", a), ("B", b))
                        for e in s["read_errors"]
                    ]
                    output = Path(d) / "report.xlsx"
                    with core.safe_runtime(d, output):
                        changes.write_report(
                            rows,
                            output,
                            b["root"],
                            errors,
                            {"A": a["root"], "B": b["root"]},
                            a["review_mode"],
                        )
                    hashes = {}
                    with zipfile.ZipFile(output) as z:
                        for name in z.namelist():
                            data = z.read(name).replace(b"\r\n", b"\n")
                            if name == "docProps/core.xml":
                                data = re.sub(
                                    rb"(<dcterms:(?:created|modified)\b[^>]*>)[^<]*(</dcterms:(?:created|modified)>)",
                                    rb"\1VOLATILE\2",
                                    data,
                                )
                            hashes[name] = hashlib.sha256(data).hexdigest()
                    self.assertEqual(hashes, case["expected_workbook_sha256"])
