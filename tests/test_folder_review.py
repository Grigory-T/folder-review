from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

import py7zr
from openpyxl import load_workbook

import folder_review


class FolderReviewTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory(prefix="folder-review-test-")
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def make_zip(self, path: Path, members: dict[str, bytes | Path]) -> Path:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for archive_name, source in members.items():
                if isinstance(source, Path):
                    archive.write(source, archive_name)
                else:
                    archive.writestr(archive_name, source)
        return path

    def make_7z(self, path: Path, members: dict[str, bytes | Path]) -> Path:
        source_dir = self.root / f"{path.stem}-sources"
        source_dir.mkdir(exist_ok=True)
        with py7zr.SevenZipFile(path, "w") as archive:
            for index, (archive_name, source) in enumerate(members.items()):
                if isinstance(source, bytes):
                    data = source
                    source = source_dir / f"member-{index}.bin"
                    source.write_bytes(data)
                archive.write(source, archive_name)
        return path

    def inspect(self, path: Path) -> list[folder_review.ResultRow]:
        item = folder_review.PhysicalFile(
            path=str(path),
            levels=(path.name,),
            size=path.stat().st_size,
        )
        return folder_review.process_file(item)

    def test_zip_7z_zip_and_office_package(self) -> None:
        office = self.make_zip(
            self.root / "book.xlsx",
            {
                "[Content_Types].xml": b"<Types />",
                "xl/workbook.xml": b"<workbook />",
            },
        )
        inner = self.make_zip(
            self.root / "inner.zip",
            {
                "Deep/final.pdf": b"%PDF-1.4\n%%EOF\n",
                "Deep/book.xlsx": office,
            },
        )
        middle = self.make_7z(
            self.root / "middle.7z",
            {
                "Layer2/inner.zip": inner,
                "Layer2/note.txt": b"terminal text",
            },
        )
        outer = self.make_zip(
            self.root / "outer.zip",
            {
                "Layer1/middle.7z": middle,
                "Direct/picture.bmp": b"BM" + bytes(32),
            },
        )

        rows = self.inspect(outer)
        paths = {row.full_path for row in rows}

        self.assertTrue(
            any(
                path.endswith(
                    r"outer.zip::Layer1\middle.7z::Layer2\inner.zip::Deep\final.pdf"
                )
                for path in paths
            )
        )
        workbook_row = next(row for row in rows if row.leaf_name == "book.xlsx")
        self.assertEqual(workbook_row.content_type, "Excel workbook (XLSX)")
        self.assertFalse(
            any(row.leaf_name in {"middle.7z", "inner.zip"} for row in rows)
        )
        self.assertFalse(any(row.error for row in rows))

    def test_7z_zip_7z_reverse_order(self) -> None:
        inner = self.make_7z(
            self.root / "inner.7z",
            {"Bottom/leaf.txt": b"leaf"},
        )
        middle = self.make_zip(
            self.root / "middle.zip",
            {"Middle/inner.7z": inner},
        )
        outer = self.make_7z(
            self.root / "outer.7z",
            {"Top/middle.zip": middle},
        )

        rows = self.inspect(outer)

        self.assertEqual(len(rows), 1)
        self.assertTrue(
            rows[0].full_path.endswith(
                r"outer.7z::Top\middle.zip::Middle\inner.7z::Bottom\leaf.txt"
            )
        )
        self.assertEqual(rows[0].content_type, "Text (UTF-8 / ASCII)")
        self.assertIsNone(rows[0].error)

    def test_archive_detection_uses_content_not_extension(self) -> None:
        archive = self.make_zip(
            self.root / "archive.bin",
            {"Folder/leaf.txt": b"leaf"},
        )

        rows = self.inspect(archive)

        self.assertEqual([row.leaf_name for row in rows], ["leaf.txt"])
        self.assertTrue(rows[0].full_path.endswith(r"archive.bin::Folder\leaf.txt"))

    def test_corrupt_nested_archive_becomes_error_row(self) -> None:
        outer = self.make_zip(
            self.root / "outer.zip",
            {"Folder/broken.zip": b"PK\x03\x04not-a-valid-zip"},
        )

        rows = self.inspect(outer)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].leaf_name, "broken.zip")
        self.assertEqual(rows[0].content_type, "ZIP container (unreadable)")
        self.assertEqual(rows[0].error, "Invalid or unsupported ZIP container")

    def test_zip_rar_zip_order(self) -> None:
        fixture = Path(__file__).parent / "fixtures" / "rar-with-zip.rar"
        outer = self.make_zip(
            self.root / "outer.zip",
            {"Top/middle.rar": fixture},
        )
        folder_review.configure_rar_backend()

        rows = self.inspect(outer)
        if len(rows) == 1 and rows[0].error == "RarCannotExec":
            self.skipTest("No RAR extraction backend is installed")

        self.assertTrue(
            any(
                row.full_path.endswith(
                    r"outer.zip::Top\middle.rar::Outer\nested.zip::Inner\leaf.txt"
                )
                for row in rows
            )
        )
        self.assertFalse(any(row.error for row in rows))

    def test_workbook_contract(self) -> None:
        rows = [
            folder_review.make_row(
                r"sample.zip::Folder\leaf.pdf",
                ("sample.zip", "Folder", "leaf.pdf"),
                12,
                "PDF",
                True,
            )
        ]
        output = self.root / "folder-review.xlsx"

        folder_review.write_workbook(rows, output, "sample")

        workbook = load_workbook(output, read_only=True, data_only=False)
        try:
            sheet = workbook["folder_review"]
            headers = [
                cell.value for cell in next(sheet.iter_rows(min_row=6, max_row=6))
            ]
            values = next(sheet.iter_rows(min_row=7, max_row=7, values_only=True))
            columns = {header: index for index, header in enumerate(headers) if header}
            self.assertEqual(values[columns["Имя листа"]], "leaf.pdf")
            self.assertEqual(values[columns["Тип по содержимому"]], "PDF")
            self.assertEqual(values[columns["В архиве"]], "да")
        finally:
            workbook.close()

    def test_archive_path_splits_both_separator_styles(self) -> None:
        self.assertEqual(
            folder_review.split_archive_path(r"First/Second\Third/file.txt"),
            ("First", "Second", "Third", "file.txt"),
        )


if __name__ == "__main__":
    unittest.main()
