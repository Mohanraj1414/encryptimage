from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw

import app
from grid_records import MI_FIELDS, MI_HEADERS, canonical_id, export_grid, match_ids
from kyc import FIELDS


class GridExportTests(unittest.TestCase):
    def setUp(self):
        workdir = TemporaryDirectory()
        self.addCleanup(workdir.cleanup)
        self.root = Path(workdir.name)
        self.app_dir = self.root / "app"
        self.app_dir.mkdir()
        self.scans = self.root / "batch" / "images"
        self.scans.mkdir(parents=True)
        self.scan = self.scans / "HnT_DuP-AkyZ (9001).jpg"
        app_file = patch.object(app, "__file__", str(self.app_dir / "app.py"))
        settings = patch.object(app, "mi_settings_path", return_value=self.root / "config.json")
        legacy_config = patch.dict(app.os.environ, {"LOCALAPPDATA": str(self.root)})
        for patcher in (app_file, settings, legacy_config):
            patcher.start()
            self.addCleanup(patcher.stop)

    def make_source(self, path):
        workbook = Workbook()
        identifier = "HnT_DuP-AkyZ-2001"
        record = {"Record No": identifier, "Type Of Account": "Test Account", "Full Name": "SYNTHETIC PERSON"}
        workbook.active.append(MI_HEADERS)
        workbook.active.append([identifier, *(record.get(field) for field in MI_FIELDS)])
        workbook.save(path)
        workbook.close()
        return path

    def test_finds_documented_source_beside_app_from_another_folder(self):
        source = self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        self.assertEqual(app.find_mi_workbook([self.scan], None), source)

    def test_explicit_choices_keep_priority_over_app_source(self):
        self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        imported = self.make_source(self.root / "imported.xlsx")
        saved = self.make_source(self.root / "saved.xlsx")
        session = self.make_source(self.root / "session.xlsx")
        self.assertEqual(app.find_mi_workbook([self.scan], imported), imported)
        app.remember_mi_path(saved)
        self.assertEqual(app.find_mi_workbook([self.scan], imported), saved)
        self.assertEqual(app.find_mi_workbook([self.scan], imported, session), session)

    def test_invalid_app_source_falls_back_to_unique_batch_workbook(self):
        (self.app_dir / "KYC_MI_Source.xlsx").write_text("not Excel", encoding="utf-8")
        source = self.make_source(self.scans.parent / "KYC MI batch.xlsx")
        self.assertEqual(app.find_mi_workbook([self.scan], None), source)
        self.make_source(self.scans / "KYC MI other.xlsx")
        self.assertIsNone(app.find_mi_workbook([self.scan], None))

    def test_grid_export_uses_recovery_without_opening_picker(self):
        workbench = Mock(images=[self.scan], imported_workbook_path=None, source_session_choice=None)
        app.Workbench.save_grid(workbench)
        workbench.choose_mi.assert_not_called()
        workbench.export_scan_batch.assert_called_once_with(None, recover=True)

    def test_grid_export_uses_discovered_source_and_ocr_review_is_explicit(self):
        source = self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        workbench = Mock(images=[self.scan], imported_workbook_path=None, source_session_choice=None)
        app.Workbench.save_grid(workbench)
        workbench.choose_mi.assert_not_called()
        workbench.export_scan_batch.assert_called_once_with(source)

        workbench.export_scan_batch.reset_mock()
        app.Workbench.export_ocr_review(workbench)
        workbench.export_scan_batch.assert_called_once_with(None)

    def test_picker_rejects_generated_output_and_remembers_valid_source(self):
        output = self.root / "generated.xlsx"
        app.export(output, [{"Record No": self.scan.stem, "Image Name": "1"}])
        source = self.make_source(self.root / "chosen.xlsx")
        workbench = Mock()
        with patch.object(app.filedialog, "askopenfilename", side_effect=[str(output), str(source)]):
            with patch.object(app.messagebox, "showerror") as showerror:
                self.assertEqual(app.Workbench.choose_mi(workbench), source)
        showerror.assert_called_once()
        self.assertEqual(app.read_mi_path(), source)

    def test_export_matches_printed_ids_and_reports_unmatched_pages(self):
        source = self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        unmatched_scan = self.scans / "HnT_DuP-AkyZ (9002).jpg"
        output = self.root / "output.xlsx"
        progress = Mock()
        payloads = {self.scan: b"HnT_DuP-AkyZ-2001", unmatched_scan: b"HnT_DuP-AkyZ-9999"}
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", side_effect=lambda path, key: payloads[path]),
            patch("grid_records.ocr_text", side_effect=lambda engine, payload, mode: payload.decode()),
        ):
            result = export_grid(output, [(self.scan, None), (unmatched_scan, None)], source, progress)
        self.assertEqual(result, (1, 1))
        self.assertEqual([call.args for call in progress.call_args_list], [(1,), (2,)])
        workbook = load_workbook(output, read_only=True)
        self.addCleanup(workbook.close)
        headers, values = workbook["KYC Data"].values
        self.assertEqual(headers, FIELDS)
        self.assertEqual(values[:5], ("HnT_DuP-AkyZ-2001", "Test Account", "N.A", "1", "Synthetic Person"))
        _, first_page, second_page = workbook["Scan review"].values
        self.assertEqual((first_page[2], second_page[2], second_page[3]), (1, 0, "HnT_DuP-AkyZ-9999"))

    def test_recovers_exact_scan_and_id_with_current_page_number(self):
        archived = self.root / "KYC_images_20261123_131254_062641.xlsx"
        output = self.root / "KYC_Recovered.xlsx"
        identifier = "HnT_DuP-AkyZ-2001"
        record = {"Record No": identifier, "Image Name": "15", "Full Name": "SYNTHETIC PERSON"}
        app.export(archived, [record], [[15, f"001_images/{self.scan.name}", 1, "", "Review"]])
        other = self.scans / "other.jpg"
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", return_value=b""),
            patch("grid_records.ocr_text", return_value=identifier),
        ):
            self.assertEqual(export_grid(output, [(other, None), (self.scan, None)], None, Mock(), recover=True), (1, 1))
        book = load_workbook(output, read_only=True)
        self.addCleanup(book.close)
        _, pending, recovered = book["KYC Data"].values
        self.assertEqual(pending[:5], (identifier, None, None, "1", None))
        self.assertEqual(pending[-1], "REVIEW: No source data for this record ID")
        self.assertEqual(recovered[:5], (identifier, "N.A", "N.A", "2", "Synthetic Person"))
        _, first, second = book["Scan review"].values
        self.assertEqual((first[2], first[3], second[2]), (0, identifier, 1))

    def test_recovery_without_prior_export_keeps_full_ids_and_reviews_unreadable_pages(self):
        output = self.root / "KYC_Recovered.xlsx"
        unreadable = self.scans / "unreadable.jpg"
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", side_effect=lambda path, key: path.name.encode()),
            patch("grid_records.ocr_text", side_effect=lambda engine, name, mode: (
                "HnT_DuP-AkyZ-2001" if name.decode() == self.scan.name else "DuP-Akyz-2002"
            )),
        ):
            self.assertEqual(
                export_grid(output, [(self.scan, None), (unreadable, None)], None, Mock(), recover=True), (0, 2)
            )
        book = load_workbook(output, read_only=True)
        self.addCleanup(book.close)
        _, pending = book["KYC Data"].values
        self.assertEqual((pending[0], pending[3]), ("HnT_DuP-AkyZ-2001", "1"))
        _, first, second = book["Scan review"].values
        self.assertEqual((first[0], first[2], second[0], second[3]), (1, 0, 2, "DuP-Akyz-2002"))

    def test_page_ids_are_sorted_and_fragmentary_ids_and_gaps_are_reported(self):
        ids = ["HnT_DuP-AkyZ-3968", "HnT_DuP-AkyZ-3970", "HnT_DuP-AkyZ-3969"]
        records = {canonical_id(identifier): identifier for identifier in ids}
        text = (
            "HnT_DuP-AkyZ-3968 HnT_DuP-AkyZ-3970 "
            "garbled _ DuP-Akyz-3969 HTIT Dup-Akyz-3973 HnT_DuP-AkyZ-3973"
        )
        matched, unmatched, gaps = match_ids(text, records)
        self.assertEqual(matched, sorted(ids))
        self.assertEqual(unmatched, ["HnT_DuP-AkyZ-3973"])
        self.assertEqual(gaps, ["3971-3972"])

    def test_fragmentary_ids_without_source_are_unmatched(self):
        matched, unmatched, gaps = match_ids("garbled _ DuP-Akyz-3971", {})
        self.assertEqual(matched, [])
        self.assertEqual(unmatched, ["DuP-Akyz-3971"])
        self.assertEqual(gaps, [])

    def test_export_interleaves_numbered_grid_columns_and_keeps_image_number(self):
        output = self.root / "KYC_Recovered.xlsx"
        ids = [f"HnT_DuP-AkyZ-{number}" for number in range(7700, 7706)]
        with BytesIO() as file, Image.new("L", (400, 500), 255) as grid:
            draw = ImageDraw.Draw(grid)
            draw.line((8, 12, 8, 480), fill=0)
            draw.line((200, 12, 200, 480), fill=0)
            for y in range(12, 480, 78):
                draw.line((8, y, 392, y), fill=0)
            grid.save(file, format="PNG")
            payload = file.getvalue()
        # Whole-page OCR misses the left column. The cell-header OCR supplies it.
        def recognize(engine, image, mode):
            if mode == 11:
                self.assertEqual(image, payload)
                return " ".join(ids[1::2])
            self.assertNotEqual(image, payload)
            return " ".join(ids[::2])

        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", return_value=payload),
            patch("grid_records.ocr_text", side_effect=recognize),
        ):
            export_grid(output, [(self.scan, None)], None, Mock(), recover=True)
        book = load_workbook(output, read_only=True)
        self.addCleanup(book.close)
        values = iter(book["KYC Data"].values)
        self.assertEqual(next(values), FIELDS)
        self.assertEqual([(row[0], row[3]) for row in values], [(identifier, "1") for identifier in ids])


if __name__ == "__main__":
    unittest.main()
