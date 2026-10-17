"""Build-only runner for the original regression suite in a frozen runtime."""

import unittest
from unittest import mock  # noqa: F401 -- include the tests' runtime dependency

import test_folder_review
import test_folder_review_ui

import folder_review
from windows_build.bootstrap import configure_frozen_application
from windows_build.suite import run_suite

configure_frozen_application(folder_review)
suite = unittest.TestSuite(
    unittest.defaultTestLoader.loadTestsFromModule(module)
    for module in (test_folder_review, test_folder_review_ui)
)
raise SystemExit(run_suite(suite))
