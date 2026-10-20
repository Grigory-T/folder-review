# ruff: noqa: SIM117 -- retain the migrated stream-test structure
"""Original bounded SHA-256 stream/race regressions, migrated under MIT."""

import hashlib
import io
import math
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import folder_change_review as app


class ControlledStream:
    """A real file descriptor with controlled bytes returned to the hasher."""

    def __init__(self, path, content, *, cap=None, after_read=None):
        self.file = open(path, "rb")  # noqa: SIM115 -- closed by this context manager
        self.data = io.BytesIO(content)
        self.cap = cap
        self.after_read = after_read
        self.requests = []
        self.probes = []
        self.delivered = 0

    def __enter__(self):
        return self

    def __exit__(self, *unused):
        self.file.close()

    def fileno(self):
        return self.file.fileno()

    def readable(self):
        return True

    def readinto(self, target):
        self.requests.append(len(target))
        if len(self.requests) > 1000:
            raise AssertionError("Hasher chased an endlessly growing stream")
        count = self.data.readinto(target[: self.cap] if self.cap else target)
        self.delivered += count
        if self.after_read:
            self.after_read(self)
        return count

    def read(self, count):
        self.probes.append(count)
        data = self.data.read(count)
        self.delivered += len(data)
        return data

    def append(self, content):
        position = self.data.tell()
        self.data.seek(0, 2)
        self.data.write(content)
        self.data.seek(position)


class MigratedHashingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.root = self.work / "project"
        self.root.mkdir()
        self.path = self.root / "document.pdf"

    def controlled(self, content, **options):
        self.path.write_bytes(content)
        return ControlledStream(self.path, content, **options)

    def test_stable_sizes_zero_one_and_buffer_boundaries_match_sha256(self):
        for size in (
            0,
            1,
            app.HASH_BUFFER_SIZE - 1,
            app.HASH_BUFFER_SIZE,
            app.HASH_BUFFER_SIZE + 1,
            app.HASH_BUFFER_SIZE * 2 + 17,
        ):
            with self.subTest(size=size):
                content = (b"abc\x00\xff" * math.ceil(size / 5))[:size]
                self.path.write_bytes(content)
                self.assertEqual(
                    app.hash_file(self.path),
                    (hashlib.sha256(content).hexdigest(), size),
                )

    def test_short_positive_reads_are_valid(self):
        content = bytes(range(256)) + b"last"
        stream = self.controlled(content, cap=3)
        with patch.object(app, "open", return_value=stream, create=True):
            self.assertEqual(
                app.hash_file(self.path),
                (hashlib.sha256(content).hexdigest(), len(content)),
            )
        self.assertGreater(len(stream.requests), 1)
        self.assertEqual(stream.probes, [1])
        self.assertEqual(stream.delivered, len(content))

    def test_continuously_growing_stream_is_bounded_even_with_stale_metadata(self):
        content = b"original data"
        stream = self.controlled(
            content, cap=2, after_read=lambda item: item.append(b"growth" * 10)
        )
        # The real descriptor/path metadata stays unchanged. Growth must be
        # detected from the extra byte, not only from a later stat result.
        with patch.object(app, "open", return_value=stream, create=True):
            with self.assertRaisesRegex(OSError, "changed while"):
                app.hash_file(self.path)
        self.assertEqual(stream.delivered, len(content) + 1)
        self.assertEqual(stream.probes, [1])
        self.assertLessEqual(len(stream.requests), math.ceil(len(content) / 2))

    def test_growth_from_empty_file_is_detected(self):
        stream = self.controlled(b"")
        stream.append(b"new byte")
        with patch.object(app, "open", return_value=stream, create=True):
            with self.assertRaisesRegex(OSError, "changed while"):
                app.hash_file(self.path)
        self.assertEqual(stream.requests, [])
        self.assertEqual(stream.probes, [1])
        self.assertEqual(stream.delivered, 1)

    def test_premature_eof_is_rejected_even_with_stale_metadata(self):
        stream = self.controlled(
            b"original data",
            cap=2,
            after_read=lambda item: item.data.truncate(item.data.tell()),
        )
        with patch.object(app, "open", return_value=stream, create=True):
            with self.assertRaisesRegex(OSError, "changed while"):
                app.hash_file(self.path)
        self.assertEqual(stream.delivered, 2)
        self.assertEqual(stream.probes, [])

    def test_same_length_mutation_and_path_replacement_metadata(self):
        self.path.write_bytes(b"original data")
        real = self.path.stat()
        fields = ("st_dev", "st_ino", "st_size", "st_mtime_ns", "st_ctime_ns")
        original = {name: getattr(real, name) for name in fields}
        for field in fields:
            for phase in ("open file", "current path"):
                with self.subTest(field=field, phase=phase):
                    before = SimpleNamespace(**original)
                    changed = SimpleNamespace(
                        **{**original, field: original[field] + 1}
                    )
                    with (
                        patch.object(
                            app.os,
                            "fstat",
                            side_effect=[
                                before,
                                changed if phase == "open file" else before,
                            ],
                        ),
                        patch.object(
                            app.os,
                            "stat",
                            return_value=changed if phase == "current path" else before,
                        ),
                    ):
                        if (
                            os.name == "nt"
                            and field == "st_ctime_ns"
                            and phase == "current path"
                        ):
                            # Stable Windows files can have different ctime
                            # meanings between fstat() and path stat().
                            self.assertEqual(
                                app.hash_file(self.path),
                                (hashlib.sha256(b"original data").hexdigest(), 13),
                            )
                        else:
                            with self.assertRaisesRegex(OSError, "changed while"):
                                app.hash_file(self.path)

    def test_read_error_is_propagated(self):
        stream = self.controlled(b"data")
        with (
            patch.object(app, "open", return_value=stream, create=True),
            patch.object(
                stream, "readinto", side_effect=OSError("disconnected source")
            ),
        ):
            with self.assertRaisesRegex(OSError, "disconnected source"):
                app.hash_file(self.path)
