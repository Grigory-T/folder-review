"""Bounded duplicate samples, shared by ordinary files and archive streams.

Matching samples identify possible duplicates, never proof of byte equality.
"""

from blake3 import blake3

SAMPLE_BYTES = 8192


def sample_ranges(size):
    offsets = sorted(
        {0, max(0, (size - SAMPLE_BYTES) // 2), max(0, size - SAMPLE_BYTES)}
    )
    ranges = []
    for start in offsets:
        end = min(size, start + SAMPLE_BYTES)
        if ranges and start <= ranges[-1][1]:
            ranges[-1] = (ranges[-1][0], max(end, ranges[-1][1]))
        else:
            ranges.append((start, end))
    return ranges


class SampleCapture:
    """Sample an existing integrity stream without retaining file contents."""

    def __init__(self, size):
        self.size = size
        self.ranges = sample_ranges(size)
        self.counts = [0 for _ in self.ranges]
        self.offset = 0
        self.hasher = blake3(
            b"folder-review.possible-duplicates.v1\0"
            + str(size).encode("ascii")
            + b"\0"
        )
        self.result = None
        if size == 0:
            self.hasher.update(bytes(16))
            self.result = self.hasher.digest()
            self.hasher = None

    def update(self, data, offset=None):
        if offset is None:
            offset = self.offset
        end = offset + len(data)
        for index, (start, stop) in enumerate(self.ranges):
            left, right = max(start, offset), min(stop, end)
            if left < right:
                if left != start + self.counts[index]:
                    raise OSError("Incomplete or repeated duplicate sample data")
                if not self.counts[index]:
                    self.hasher.update(start.to_bytes(8, "little"))
                    self.hasher.update((stop - start).to_bytes(8, "little"))
                self.hasher.update(data[left - offset : right - offset])
                self.counts[index] += right - left
        self.offset = end
        if self.result is None and all(
            count == stop - start
            for count, (start, stop) in zip(self.counts, self.ranges)
        ):
            self.result = self.hasher.digest()
            self.hasher = None

    def digest(self):
        if self.result is None:
            raise OSError(
                "File changed during duplicate sampling; sample was not retained"
            )
        return self.result
