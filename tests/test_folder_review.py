from __future__ import annotations

import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

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
        stat_result = path.stat()
        item = folder_review.PhysicalFile(
            path=str(path),
            levels=(path.name,),
            size=stat_result.st_size,
            properties=folder_review.file_properties(path.name, stat_result),
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
                "Deep/bom.txt": b"\xff\xfeh\x00i\x00",
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
        bom_row = next(row for row in rows if row.leaf_name == "bom.txt")
        self.assertEqual(workbook_row.content_type, "Excel workbook (XLSX)")
        self.assertEqual(bom_row.bom, "UTF-16 LE")
        self.assertEqual(bom_row.item_type, "Файл внутри архива")
        self.assertIsNone(bom_row.properties)
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
        source = self.root / "leaf.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        properties = folder_review.file_properties(source.name, source.stat())
        rows = [
            folder_review.make_row(
                str(source),
                ("leaf.pdf",),
                source.stat().st_size,
                "PDF",
                False,
                bom=folder_review.BOM_NOT_PRESENT,
                properties=properties,
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
            requested_headers = {
                "Полный путь к файлу",
                "Имя файла",
                "Расширение",
                "Полное имя файла",
                "Тип элемента",
                "Тип ссылки",
                "Цель ссылки",
                "Права доступа",
                "Скрытый",
                "Только чтение",
                "Исполняемый",
                "Системный (Windows)",
                "Архивный (Windows)",
                "Временный (Windows)",
                "Автономный (Windows)",
                "Сжатый (Windows)",
                "Зашифрованный (Windows)",
                "Разреженный (Windows)",
                "Не индексировать содержимое (Windows)",
                "Атрибуты файла (Windows)",
                "st_file_attributes",
                "st_reparse_tag",
                "st_mode",
                "st_ino",
                "st_dev",
                "st_rdev",
                "st_nlink",
                "st_uid (User ID of the owner)",
                "st_gid (Group ID of the owner)",
                "st_blksize",
                "st_blocks",
                "Размер файла в мб",
                "Время последнего обращения к файлу",
                "Время последнего изменения файла",
                "Время создания файла",
                "Время изменения метаданных",
                "BOM",
            }
            self.assertTrue(requested_headers.issubset(columns))
            self.assertEqual(values[columns["Имя файла"]], "leaf")
            self.assertEqual(values[columns["Полное имя файла"]], "leaf.pdf")
            self.assertEqual(values[columns["Тип по содержимому"]], "PDF")
            self.assertEqual(values[columns["В архиве"]], "нет")
            self.assertEqual(values[columns["BOM"]], "нет")
            self.assertEqual(values[columns["Тип элемента"]], "Файл")
            self.assertEqual(values[columns["Скрытый"]], "нет")
            self.assertEqual(
                values[columns["Права доступа"]], properties.permissions
            )
            self.assertEqual(values[columns["st_mode"]], properties.mode)
            self.assertEqual(values[columns["st_ino"]], properties.inode)
            self.assertIsNotNone(values[columns["Время последнего изменения файла"]])
        finally:
            workbook.close()

    def test_physical_properties_and_utf8_bom(self) -> None:
        source = self.root / "bom.txt"
        source.write_bytes(b"\xef\xbb\xbfhello")

        files, stats = folder_review.scan_tree(str(self.root), set(), workers=2)
        rows = folder_review.inspect_files(files, workers=2)

        self.assertEqual(stats.skipped_entries, 0)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].bom, "UTF-8")
        self.assertEqual(rows[0].content_type, "Text (UTF-8 BOM)")
        self.assertIsNotNone(rows[0].properties)
        self.assertEqual(rows[0].item_type, "Файл")
        self.assertIsNone(rows[0].link_type)
        self.assertEqual(rows[0].properties.mode, source.stat().st_mode)
        self.assertTrue(rows[0].properties.permissions.startswith("-"))
        if os.name == "nt":
            self.assertIsNotNone(rows[0].properties.file_attributes_value)
            self.assertIsNone(rows[0].properties.metadata_changed_at)
        else:
            self.assertIsNone(rows[0].properties.file_attributes_value)
            self.assertIsNotNone(rows[0].properties.metadata_changed_at)

    def test_links_and_special_items_are_reported_without_following(self) -> None:
        if os.name == "nt":
            self.skipTest("POSIX link and FIFO behavior")

        target_file = self.root / "target.txt"
        target_file.write_text("target", encoding="utf-8")
        target_folder = self.root / "target-folder"
        target_folder.mkdir()
        (target_folder / "inside.txt").write_text("inside", encoding="utf-8")
        file_link = self.root / "file-link"
        folder_link = self.root / "folder-link"
        file_link.symlink_to(target_file.name)
        folder_link.symlink_to(target_folder.name, target_is_directory=True)
        fifo = self.root / "events.fifo"
        os.mkfifo(fifo)

        files, stats = folder_review.scan_tree(str(self.root), set(), workers=2)
        rows = folder_review.inspect_files(files, workers=2)
        by_name = {row.leaf_name: row for row in rows}

        self.assertEqual(stats.skipped_entries, 0)
        self.assertEqual(by_name["file-link"].item_type, "Символическая ссылка")
        self.assertEqual(by_name["file-link"].link_type, "Символическая ссылка")
        self.assertEqual(by_name["file-link"].link_target, target_file.name)
        self.assertEqual(
            by_name["folder-link"].content_type,
            "Filesystem link (not followed)",
        )
        self.assertEqual(by_name["events.fifo"].item_type, "Именованный канал (FIFO)")
        self.assertEqual(
            by_name["events.fifo"].content_type,
            "Special filesystem item (not read)",
        )
        self.assertEqual(sum(row.leaf_name == "inside.txt" for row in rows), 1)

    def test_hard_links_are_marked(self) -> None:
        source = self.root / "source.txt"
        hard_link = self.root / "hard-link.txt"
        source.write_text("same data", encoding="utf-8")
        try:
            os.link(source, hard_link)
        except OSError as exc:
            self.skipTest(f"Hard links are unavailable: {exc}")

        files, _ = folder_review.scan_tree(str(self.root), set(), workers=2)
        rows = folder_review.inspect_files(files, workers=2)

        self.assertEqual(len(rows), 2)
        self.assertTrue(all(row.link_type == "Жесткая ссылка" for row in rows))
        self.assertTrue(all(row.properties.links == 2 for row in rows))

    def test_windows_attribute_names(self) -> None:
        self.assertEqual(
            folder_review.windows_attribute_names(0x00000001 | 0x00000020),
            "READONLY | ARCHIVE",
        )
        self.assertEqual(
            folder_review.windows_attribute_names(0x01000000),
            "UNKNOWN_0x01000000",
        )
        self.assertIsNone(folder_review.windows_attribute_names(None))

    def test_all_supported_bom_variants(self) -> None:
        cases = {
            b"\xef\xbb\xbftext": "UTF-8",
            b"\xff\xfe\x00\x00text": "UTF-32 LE",
            b"\x00\x00\xfe\xfftext": "UTF-32 BE",
            b"\xff\xfetext": "UTF-16 LE",
            b"\xfe\xfftext": "UTF-16 BE",
            b"plain": None,
        }
        for data, expected in cases.items():
            with self.subTest(expected=expected):
                self.assertEqual(folder_review.detect_bom(data), expected)

    def test_platform_and_worker_defaults(self) -> None:
        self.assertEqual(folder_review.DEFAULT_WORKERS, 16)
        if os.name == "nt":
            self.assertEqual(
                folder_review.regular_path(r"\\?\UNC\server\share\folder"),
                r"\\server\share\folder",
            )
            self.assertEqual(
                folder_review.regular_path(r"\\?\unc\server\share\folder"),
                r"\\server\share\folder",
            )
        with (
            mock.patch.object(folder_review.sys, "platform", "darwin"),
            self.assertRaisesRegex(RuntimeError, "macOS is not supported"),
        ):
            folder_review.validate_platform()

    def test_archive_path_splits_both_separator_styles(self) -> None:
        self.assertEqual(
            folder_review.split_archive_path(r"First/Second\Third/file.txt"),
            ("First", "Second", "Third", "file.txt"),
        )

    def test_long_physical_archive_and_output_paths(self) -> None:
        deep_folder = str(self.root)
        segment = 0
        while len(deep_folder) < 320:
            deep_folder = os.path.join(
                deep_folder,
                f"long-folder-segment-{segment:02d}-abcdefgh",
            )
            segment += 1
        os.makedirs(folder_review.extended_path(deep_folder))

        inner = self.make_zip(
            self.root / "long-inner.zip",
            {"Inside/utf8-bom.txt": b"\xef\xbb\xbflong path"},
        )
        outer = self.make_7z(
            self.root / "long-outer.7z",
            {"Archive folder/long-inner.zip": inner},
        )
        long_archive = os.path.join(deep_folder, "long-outer.7z")
        with (
            open(outer, "rb") as source,
            open(folder_review.extended_path(long_archive), "wb") as destination,
        ):
            destination.write(source.read())

        files, stats = folder_review.scan_tree(deep_folder, set(), workers=2)
        rows = folder_review.inspect_files(files, workers=2)

        self.assertEqual(stats.skipped_folders, 0)
        self.assertEqual(stats.skipped_entries, 0)
        self.assertEqual(len(files), 1)
        self.assertEqual(len(rows), 1)
        self.assertGreater(len(rows[0].full_path), 260)
        self.assertEqual(rows[0].leaf_name, "utf8-bom.txt")
        self.assertEqual(rows[0].bom, "UTF-8")
        self.assertIsNone(rows[0].error)

        output_folder = os.path.join(deep_folder, "workbook-output")
        os.mkdir(folder_review.extended_path(output_folder))
        output = Path(output_folder) / "folder-review.xlsx"
        folder_review.write_workbook(rows, output, deep_folder)

        self.assertTrue(os.path.isfile(folder_review.extended_path(output)))
        self.assertFalse(
            any(
                name.startswith(".folder-review-")
                for name in os.listdir(folder_review.extended_path(output_folder))
            )
        )
        workbook = load_workbook(
            folder_review.extended_path(output),
            read_only=True,
            data_only=False,
        )
        try:
            sheet = workbook["folder_review"]
            values = next(sheet.iter_rows(min_row=7, max_row=7, values_only=True))
            self.assertIn("UTF-8", values)
            self.assertTrue(
                any(isinstance(value, str) and len(value) > 260 for value in values)
            )
        finally:
            workbook.close()

    def test_temp_member_path_uses_safe_short_random_name(self) -> None:
        with tempfile.TemporaryDirectory(prefix="folder-review-temp-test-") as root:
            path = folder_review.make_temp_path(
                root,
                "member-with-invalid-extension." + "x" * 300 + "?zip",
            )
            try:
                self.assertLess(len(os.path.basename(path)), 80)
                self.assertTrue(path.endswith(".bin"))
            finally:
                os.unlink(path)


if __name__ == "__main__":
    unittest.main()
