from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw

import app
from center_codes import restore_center_codes
from grid_records import (
    MI_FIELDS, MI_HEADERS, canonical_id, consecutive_scans, detect_ids, export_grid, match_ids, read_prior_export,
    read_grid_cells, select_batch_indexes,
)
from kyc import FIELDS, normalize
from scan_records import OCR_REMARK, parse_scan_record


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

    def make_source(self, path, identifier="HnT_DuP-AkyZ-2001", name="SYNTHETIC PERSON"):
        workbook = Workbook()
        record = {"Record No": identifier, "Type Of Account": "Test Account", "Full Name": name}
        workbook.active.append(MI_HEADERS)
        workbook.active.append([identifier, *(record.get(field) for field in MI_FIELDS)])
        workbook.save(path)
        workbook.close()
        return path

    def test_finds_documented_source_beside_app_from_another_folder(self):
        source = self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        source.with_name("~$" + source.name).write_bytes(b"Excel owner file")
        self.assertEqual(app.find_mi_workbook([self.scan], None), source)

    def test_packaged_app_finds_source_next_to_executable(self):
        source = self.make_source(self.root / "KYC_MI_Source.xlsx")
        self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        with (
            patch.object(app.sys, "frozen", True, create=True),
            patch.object(app.sys, "executable", str(self.root / "KYC_Excel_Automator.exe")),
        ):
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

    def test_unedited_scan_forms_go_through_grid_export(self):
        workbench = Mock(
            images=[self.scan],
            records=[{"Record No": self.scan.stem, "Image Name": "1"}],
        )
        with patch.object(app.filedialog, "asksaveasfilename", return_value="") as picker:
            app.Workbench.save_reviewed_rows(workbench)
        workbench.capture.assert_called_once()
        workbench.save_grid.assert_called_once()
        picker.assert_not_called()

    def test_manually_entered_scan_form_can_still_be_exported(self):
        workbench = Mock(
            images=[self.scan],
            records=[{"Record No": self.scan.stem, "Image Name": "1", "Full Name": "Synthetic Person"}],
        )
        with patch.object(app.filedialog, "asksaveasfilename", return_value="") as picker:
            app.Workbench.save_reviewed_rows(workbench)
        workbench.save_grid.assert_not_called()
        picker.assert_called_once()

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
        headers, values, review = workbook["KYC Data"].values
        self.assertEqual(headers, FIELDS)
        self.assertEqual(values[:5], ("HnT_DuP-AkyZ-2001", "Test Account", "N.A", "1", "Synthetic Person"))
        self.assertEqual(review[:5], ("HnT_DuP-AkyZ-9999", None, None, "2", None))
        self.assertEqual(review[-1], "REVIEW: No source data for this record ID")
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

    def test_detects_new_id_families_without_a_prefix_list(self):
        identifiers = ["FbK_BrT-NaY-7730", "QzP_AxW-TeZ-12345"]
        matched, unmatched, gaps = match_ids(" ".join(identifiers), {})
        self.assertEqual((matched, unmatched, gaps), ([], identifiers, []))
        self.assertEqual(
            detect_ids("QzP AxW TeZ 12345", "QzP_AxW-TeZ (22908)"),
            {"qzpaxwtez12345": identifiers[1]},
        )

    def test_batch_selects_matching_mi_sources_and_exports_populated_rows(self):
        old = self.make_source(self.root / "old.xlsx")
        first_id = "FbK_BrT-NaY-7730"
        second_id = "QzP_AxW-TeZ-12345"
        batch_one = self.make_source(self.scans / "Records.xlsx", first_id, "FIRST PERSON")
        batch_two = self.make_source(self.scans / "Second.xlsx", second_id, "SECOND PERSON")
        next_scan = self.root / "second" / "scan.jpg"
        output = self.root / "new.xlsx"
        text = {self.scan: first_id, next_scan: second_id}
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", side_effect=lambda path, key: text[path].encode()),
            patch("grid_records.ocr_text", side_effect=lambda engine, raw, mode: raw.decode()),
        ):
            self.assertEqual(
                export_grid(output, [(self.scan, None), (next_scan, None)], old, Mock(),
                            alternatives=[batch_one, batch_two], require_matches=True),
                (2, 0),
            )
        workbook = load_workbook(output, read_only=True)
        self.addCleanup(workbook.close)
        _, first, second = workbook["KYC Data"].values
        self.assertEqual((first[:5], second[:5]), (
            (first_id, "Test Account", "N.A", "1", "First Person"),
            (second_id, "Test Account", "N.A", "2", "Second Person"),
        ))

    def test_missing_matching_mi_stops_without_saving_empty_excel(self):
        source = self.make_source(self.root / "old.xlsx")
        output = self.root / "new.xlsx"
        scan = self.scans / "FbK_BrT-NaY (22908).jpg"
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", return_value=b""),
            patch("grid_records.ocr_text", return_value="FOR BiT-NaY-7781 FbK_BrT-NaY-7732 FbK_BrT-NaY-7730"),
        ):
            with self.assertRaisesRegex(ValueError, "No matching MI rows or readable record fields.*FbK_BrT-NaY"):
                export_grid(output, [(scan, None)], source, Mock(), require_matches=True)
        self.assertFalse(output.exists())

    def test_missing_mi_message_distinguishes_a_garbled_prefix(self):
        scan = self.scans / "FbK_BrT-NaY (22908).jpg"
        source = self.make_source(self.root / "old.xlsx")
        output = self.root / "result.xlsx"
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.scan_bytes", return_value=b""),
            patch("grid_records.ocr_text", return_value="FOR BiT-NaY-7781"),
        ):
            with self.assertRaisesRegex(ValueError, "No matching MI rows or readable record fields.*FbK_BrT-NaY"):
                export_grid(output, [(scan, None)], source, Mock(), require_matches=True)
        self.assertFalse(output.exists())

    def test_export_reads_a_new_batch_from_scan_cells_when_the_saved_source_does_not_match(self):
        source = self.make_source(self.root / "old.xlsx")
        scan = self.scans / "FbK_BrT-NaY (22908).jpg"
        output = self.root / "new.xlsx"
        with Image.new("L", (1060, 550), 255) as page:
            grid = ImageDraw.Draw(page)
            for x in (20, 525, 1030):
                grid.line((x, 20, x, 520), fill=0)
            for y in (20, 145, 270, 395, 520):
                grid.line((20, y, 1030, y), fill=0)
            page.save(scan)

        def read_text(engine, payload, mode, *, preserve_spacing=False):
            if not preserve_spacing:
                return "FOR BiT-NaY-7781"
            return ("FbK_BrT-Ma\u00a5-7730 Example Bank (SBS) 432 Example Lane     Sam\n"
                    "Doe     Robin Doe     Thursday,01221959     FEMALE\n"
                    "UNITED STATES     117 River Road     EXAMPLE     MN 56560\n"
                    "490-675-0932    sam@example.org    443-730-4904   419-620-1668 VERIFIED YES Example\n"
                    "ABcdEF12mnOP34qrST56uvWXyZ\u00a578")

        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.ocr_text", side_effect=read_text),
        ):
            result = export_grid(output, [(scan, None)], source, Mock(), require_matches=True)
        self.assertEqual(result, (1, 1))
        book = load_workbook(output, read_only=True)
        self.addCleanup(book.close)
        headers, row = book["KYC Data"].values
        self.assertEqual(headers, FIELDS)
        record = dict(zip(headers, row))
        self.assertEqual(record["Record No"], "FbK_BrT-NaY-7730")
        self.assertEqual(record["Image Name"], "1")
        self.assertEqual(record["Full Name"], "Sam Doe")
        self.assertEqual(record["Date Of Birth"], "Thursday, 01/22/1959")
        self.assertEqual(record["Email"], "sam@example.org")
        self.assertEqual(record["Center Code"], "ABcdEF12mnOP34qrST56uvWXyZY78")
        self.assertIn(OCR_REMARK, record["Remark"])
        _, review = book["Scan review"].values
        self.assertEqual(review[2], 0)
        self.assertIn("1 rows read by OCR", review[-1])

    def test_ocr_rows_are_not_reused_as_mi_rows_on_a_page_with_verified_rows(self):
        output = self.root / "KYC_images_20261126_100000_000001.xlsx"
        verified = {"Record No": "HnT_DuP-AkyZ-2001", "Image Name": "1", "Full Name": "Verified"}
        ocr = {**verified, "Record No": "HnT_DuP-AkyZ-2002", "Remark": OCR_REMARK}
        app.export(output, [verified, ocr], [[1, f"001_images/{self.scan.name}", 1, "", "Review"]])
        entries = read_prior_export(output, {("images", self.scan.name.casefold())})
        self.assertEqual([key for _, key, _ in entries], ["hntdupakyz2001"])

    def test_scan_parser_leaves_unseparated_fields_blank(self):
        info = parse_scan_record(
            "FbK_BrT-NaY-7730",
            "Example Bank (SBS) 432 Example Lane Sam Doe Robin Doe Tuesday,01-22-1959"
            " MALE AMERICAN sam@example.org 443-730-4904 VERIFIED YES Example\nNot A Center Code",
        )
        self.assertEqual(info["Email"], "sam@example.org")
        for field in ("Full Name", "Father/Husband Name", "Contact 1", "Contact 2", "Center Code"):
            self.assertNotIn(field, info)

    def test_scan_parser_recovers_wrapped_fields_and_keeps_invalid_printed_email(self):
        info = parse_scan_record(
            "FbK_BrT-NaY-7730",
            "Credit Union (Sbs} 999 Example Road\n"
            "samantha mock Robin SampleTUESDAY,09-22-\n"
            "1959Male Unitid\nStates 456 Sample Lane    Sampleton    CA = 12345\n"
            "500-123-4567 sam. mock#example. org 500-321-7654 500-789-0123\n"
            "VERIFIED YES Example  aB34Yz78 ZnmQWvRSTu56x¥ABC",
        )
        self.assertEqual(info["Rag Address"], "999 Example Road")
        self.assertEqual(info["Full Name"], "samantha mock")
        self.assertEqual(info["Father/Husband Name"], "Robin Sample")
        self.assertEqual(info["Date Of Birth"], "TUESDAY,09-22-1959")
        self.assertEqual(info["Address"], "456 Sample Lane")
        self.assertEqual(info["City/District/Town"], "Sampleton")
        self.assertEqual(info["Email"], "sam.mock#example.org")
        self.assertEqual(info["Center Code"], "aB34Yz78ZnmQWvRSTu56xYABC")

    def test_scan_phones_keep_printed_digits_and_stop_at_field_boundaries(self):
        record = parse_scan_record(
            "FbK_BrT-NaY-7730",
            "Example Bank (SBS) 42 Main Street\nSarah Doe    Robin Doe Friday03-04-1960 Female. USA "
            "741 West Road   Sample   VA 12345 1-519-\n445-5264 7389 "
            "sample@example.org.1-506-593-5269 736.310.8894 58040 VERIFIED YES Example",
        )
        row = dict(zip(FIELDS, normalize({**record, "Image Name": "1"}, scan_export=True)))
        self.assertEqual(record["Mobile Number"], "1-519-445-5264 7389")
        self.assertEqual(record["Email"], "sample@example.org")
        self.assertEqual(record["Contact 1"], "1-506-593-5269")
        self.assertEqual(record["Contact 2"], "736.310.8894 58040")
        self.assertEqual(record["Gender"], "Female.")
        self.assertEqual(row["Date Of Birth"], "Friday, 03/04/1960")
        self.assertEqual(row["Contact 1"], "(150) 659-35269")
        self.assertIn("MOBILE NUMBER IS INVALID", row["Remark"])
        self.assertIn("CONTACT NO 2 IS INVALID", row["Remark"])

    def test_scan_separates_joined_contacts_after_email(self):
        record = parse_scan_record("FbK_BrT-NaY-7730",
                                   "sample@example.org 725-227-6030547.536.1659 381 VERIFIED YES Example")
        self.assertEqual(record["Contact 1"], "725-227-6030")
        self.assertEqual(record["Contact 2"], "547.536.1659 381")

    def test_scans_only_uses_the_production_excel_export(self):
        workbench = Mock(images=[self.scan])
        app.Workbench.export_scans_only(workbench)
        workbench.export_scan_batch.assert_called_once_with(None, scan_only=True)

    def test_grid_continuations_require_consecutive_filenames_in_the_same_folder(self):
        first = self.scans / "FbK_BrT-NaY (22910).jpg"
        second = self.scans / "FbK_BrT-NaY (22911).jpg"
        self.assertTrue(consecutive_scans(first, second))
        self.assertFalse(consecutive_scans(first, second.with_name("FbK_BrT-NaY (22912).jpg")))
        self.assertFalse(consecutive_scans(first, second.with_name("HnT_DuP-AkyZ (22911).jpg")))
        self.assertFalse(consecutive_scans(first, self.root / second.name))

    def test_grid_chooses_id_agreed_by_multiple_readings(self):
        record = " Example Bank (SBS) 42 Main Street"
        with (
            Image.new("L", (200, 100), 255) as gray,
            patch("grid_records.ruled_grid", return_value=None) as grid,
            patch("grid_records.cell_readings", return_value=[
                "FbK_BrT-NaY-17345" + record,
                "FbK_BrT-NaY-17346" + record,
                "FbK_BrT-MaY-17346" + record,
            ]),
        ):
            grid.return_value = (gray, (4, 100), [(4, 40)])
            rows = list(read_grid_cells(b"", "FbK_BrT-NaY (23488)", "ocr"))
        self.assertEqual([row["Record No"] for row in rows], ["FbK_BrT-NaY-17346"] * 2)

    def test_conflicting_nearby_mi_sources_require_an_explicit_choice(self):
        identifier = "FbK_BrT-NaY-7730"
        one = self.make_source(self.scans / "First.xlsx", identifier, "FIRST PERSON")
        two = self.make_source(self.scans / "Second.xlsx", identifier, "SECOND PERSON")
        with self.assertRaisesRegex(ValueError, "Several MI workbooks match"):
            select_batch_indexes({}, [one, two], {self.scans: {canonical_id(identifier)}})

    def test_export_keeps_dob_and_email_in_their_cells_and_preserves_center_code(self):
        output = self.root / "fields.xlsx"
        record = {
            "Date Of Birth": "Tuesday, 08/15/1939",
            "Gender": "Male",
            "Email": "synthetic.person@example.org",
            "Contact 1": "(130) 959-7595",
            "Center Name": "SYNTHETIC CENTER",
            "Center Code": "AB Y CD ef yZ",
        }
        app.export(output, [record, {**record, "Center Code": "N.A"}], normalized=True)
        book = load_workbook(output)
        self.addCleanup(book.close)
        sheet = book["KYC Data"]
        self.assertEqual((sheet["G2"].value, sheet["H2"].value), (record["Date Of Birth"], "Male"))
        self.assertEqual((sheet["O2"].value, sheet["P2"].value), (record["Email"], record["Contact 1"]))
        self.assertTrue(sheet["G2"].alignment.wrap_text and sheet["O2"].alignment.wrap_text)
        self.assertGreaterEqual(sheet.column_dimensions["G"].width, 26)
        self.assertGreaterEqual(sheet.column_dimensions["O"].width, 42)
        self.assertEqual((sheet["T2"].value, sheet["U2"].value, sheet["U3"].value),
                         ("SYNTHETIC CENTER", "AB Y CD ef yZ", "N.A"))

    def test_center_code_requires_agreement_on_position_and_case(self):
        record = {"Record No": "HnT_DuP-AkyZ-2001", "Center Code": "abcD1234Efgh56jk MN78opQR90st"}
        lower = "abcD1234Efgh56jkyMN78opQR90st"
        upper = "abcD1234Efgh56jkYMN78opQR90st"
        for readings in ([lower], [lower, upper], [lower, lower + " Y", upper]):
            self.assertEqual(restore_center_codes([record], readings), [record])
        other_character = lower.replace("1234", "1238")
        self.assertEqual(restore_center_codes([record], [other_character] * 3), [record])
        self.assertEqual(
            restore_center_codes([record], ["VERIFIED YES " + lower, "abcD1234Efgh56jky MN78opQR90st"]),
            [{**record, "Center Code": lower}],
        )
        standalone = record["Center Code"].replace(" ", " Y ")
        self.assertEqual(restore_center_codes([record], [standalone] * 2), [record])

    def test_grid_export_restores_confirmed_y_and_preserves_source_fields(self):
        source = self.make_source(self.app_dir / "KYC_MI_Source.xlsx")
        workbook = load_workbook(source)
        workbook.active["U2"] = "abcD1234Efgh56jk MN78opQR90st"
        workbook.save(source)
        workbook.close()
        output = self.root / "center.xlsx"
        reading = "abcD1234Efgh56jkYMN78opQR90st"
        with (
            patch("grid_records.tesseract_path", return_value="ocr"),
            patch("grid_records.windows_ocr_path", return_value=None),
            patch("grid_records.scan_bytes", return_value=b"page"),
            patch("grid_records.grid_id_images", return_value=[]),
            patch("grid_records.grid_center_images", return_value=[b"center-a", b"center-b"]),
            patch("grid_records.ocr_text", side_effect=["HnT_DuP-AkyZ-2001", reading, reading]),
        ):
            self.assertEqual(export_grid(output, [(self.scan, None)], source, Mock()), (1, 0))
        result = load_workbook(output, read_only=True)
        self.addCleanup(result.close)
        _, row = result["KYC Data"].values
        self.assertEqual((row[0], row[1], row[4], row[20]),
                         ("HnT_DuP-AkyZ-2001", "Test Account", "Synthetic Person", reading))

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
