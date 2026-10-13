"""Build-only runner for the original regression suite in a frozen runtime."""

import unittest
from unittest import mock  # noqa: F401 -- include the tests' runtime dependency

import test_folder_review

import folder_review
from windows_build.bootstrap import configure_frozen_application
from windows_build.suite import run_suite

configure_frozen_application(folder_review)
suite = unittest.defaultTestLoader.loadTestsFromModule(test_folder_review)
raise SystemExit(run_suite(suite))
