"""KYC field normalization and Excel interchange."""
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

FIELDS = (
    "Record No", "Type Of Account", "Rag Address", "Image Name", "Full Name",
    "Father/Husband Name", "Date Of Birth", "Gender", "Nationality", "Address",
    "City/District/Town", "State", "Pin Code", "Mobile Number", "Email",
    "Contact 1", "Contact 2", "Address Verification", "ID Verification",
    "Center Name", "Center Code", "Remark",
)
DAYS = "Monday Tuesday Wednesday Thursday Friday Saturday Sunday".split()
MONTHS = "January February March April May June July August September October November December".split()
MISSING = {"", "na", "n.a", "n.a.", "n/a", "not available", "not apply", "not applicable"}
PHONE_FIELDS = {13: "MOBILE NUMBER", 15: "CONTACT NO_1", 16: "CONTACT NO_2"}


def normalize_date(value):
    prefix = ""
    valid_day = True
    parsed_day = re.fullmatch(r"([A-Za-z]+)([., ]+)(.+)", value)
    if parsed_day:
        day, punctuation, remainder = parsed_day.groups()
        if day.title() not in MONTHS:
            # A weekday label is retained as printed, including an evident typo.
            if day.title() in DAYS or len(re.findall(r"\d+", remainder)) >= 3 or re.match(r"[A-Za-z]+[ ,.-]+\d", remainder):
                valid_day = day.title() in DAYS
                prefix = day.title() + (punctuation.strip()[:1] or ",") + " "
                value = remainder
    month_date = re.fullmatch(r"([A-Za-z]+)[ ,./-]+(\d{1,2})[ ,./-]+(\d+)", value)
    if month_date:
        month, day, year = month_date.groups()
        if month.title() not in MONTHS:
            return prefix + month.title() + value[len(month):], False
        month = str(MONTHS.index(month.title()) + 1)
    else:
        parts = re.fullmatch(r"(\d{1,2})[ ,./-]+(\d{1,2})[ ,./-]+(\d[\d-]*)", value)
        if parts:
            month, day, year = parts.groups()
            year = year.replace("-", "")
        else:
            compact = re.sub(r"[ ,./-]", "", value)
            if not re.fullmatch(r"\d{5,9}", compact):
                return prefix + value, False
            month, day, year = compact[:2], compact[2:4], compact[4:]
    normalized = f"{int(month):02}/{int(day):02}/{year}"
    try:
        date(int(year), int(month), int(day))
    except ValueError:
        return prefix + normalized, False
    return prefix + normalized, valid_day and len(year) == 4


def normalize(record):
    result = []
    issues = []
    for index, field in enumerate(FIELDS[:-1]):
        raw = record.get(field, "")
        if isinstance(raw, (date, datetime)):
            raw = raw.strftime("%m/%d/%Y")
        value = " ".join(str(raw if raw is not None else "").split())
        if value.casefold() in MISSING:
            value = "N.A"
            issues.append(field.upper() + " IS MISSING")
        elif index in (4, 5):
            value = value.replace("ª", "").title()
        elif index == 6:
            value, valid = normalize_date(value)
            if not valid:
                issues.append("DATE OF BIRTH IS INVALID")
        elif index == 7 and value.casefold() not in {"female", "male"}:
            issues.append("GENDER IS INVALID")
        elif index == 8 and value.casefold() not in {"usa", "united states"}:
            issues.append("NATIONALITY IS INVALID")
        elif index == 14:
            value = value.lower()
            if not re.fullmatch(r"[a-z0-9_+%-]+(?:\.[a-z0-9_+%-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,}", value) or value.endswith(".co.in"):
                issues.append("EMAIL IS INVALID")
        elif index in PHONE_FIELDS:
            digits = re.sub(r"[^0-9]", "", value)
            if 8 <= len(digits) <= 12 and re.fullmatch(r"[0-9 ()+./-]+", value):
                value = f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
            else:
                issues.append(PHONE_FIELDS[index] + " IS INVALID")
        result.append(value)
    invalid_contacts = [name for index, name in PHONE_FIELDS.items()
                        if index != 13 and name + " IS INVALID" in issues]
    if len(invalid_contacts) > 1:
        position = issues.index(invalid_contacts[0] + " IS INVALID")
        for name in invalid_contacts:
            issues.remove(name + " IS INVALID")
        issues.insert(position, " , ".join(invalid_contacts) + " ARE INVALID")
    note = " ".join(str(record.get("Remark") or "").upper().split())
    if note.casefold() not in MISSING:
        issues.append(note)
    return result + [" ".join(dict.fromkeys(issues)) or "N.A."]


def read_sheet(path, sheet_name, mapping, header_row=1):
    if not mapping or len(set(mapping.values())) != len(mapping):
        raise ValueError("Map at least one field and use each source column once.")
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name]
        if sheet.max_column is None:
            sheet.calculate_dimension(force=True)
        if any(index < 0 or index >= sheet.max_column for index in mapping.values()):
            raise ValueError("A mapped column is outside the sheet.")
        records = []
        for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            if not any(value is not None for value in row):
                continue
            record = {field: row[index] for field, index in mapping.items()}
            if not record.get("Image Name"):
                record["Image Name"] = str(len(records) + 1)
            records.append(record)
        return records
    finally:
        workbook.close()


def export(path, records, scan_report=None, duplicates=(), *, scan_ocr=None, normalized=False):
    workbook = Workbook(write_only=True)
    from openpyxl.cell import WriteOnlyCell

    if records is not None:
        sheet = workbook.create_sheet("KYC Data")
        sheet.freeze_panes = "F2"

        for index, field in enumerate(FIELDS, 1):
            sheet.column_dimensions[get_column_letter(index)].width = max(17, len(field) + 3)
        sheet.column_dimensions["V"].width = 55
        titles = []
        for field in FIELDS:
            cell = WriteOnlyCell(sheet, value=field)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="18334B")
            titles.append(cell)
        sheet.append(titles)
        for record in records:
            row = []
            values = (str(record.get(field) or "") for field in FIELDS) if normalized else normalize(record)
            for value in values:
                cell = WriteOnlyCell(sheet, value=value)
                cell.data_type = "s"
                row.append(cell)
            sheet.append(row)
    if scan_report is not None:
        report = workbook.create_sheet("Scan review")
        report.freeze_panes = "A2"
        for column, width in {"A": 15, "B": 44, "C": 18, "D": 55, "E": 45}.items():
            report.column_dimensions[column].width = width
        report.append(["Image Name", "Source image", "Matched rows", "Unmatched IDs", "Review"])
        for entry in scan_report:
            report.append(entry)
        if duplicates:
            repeated = workbook.create_sheet("Repeated IDs")
            repeated.append(["Record IDs detected on more than one page"])
            for identifier in duplicates:
                cell = WriteOnlyCell(repeated, value=identifier)
                cell.data_type = "s"
                repeated.append([cell])
    if scan_ocr is not None:
        review = workbook.create_sheet("Scan review")
        review.freeze_panes = "C2"
        review.column_dimensions["B"].width = 45
        review.column_dimensions["D"].width = 100
        review.column_dimensions["E"].width = 58
        review.append(["Image number", "Source image", "Detected IDs", "OCR text", "Review"])
        detected = workbook.create_sheet("Detected IDs")
        detected.freeze_panes = "B2"
        detected.column_dimensions["B"].width = 45
        detected.column_dimensions["C"].width = 32
        detected.append(["Image number", "Source image", "OCR record ID"])

        def text_cell(sheet, value):
            # OCR and filenames are untrusted text. Keep them out of Excel formulas.
            from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
            cell = WriteOnlyCell(sheet, value=ILLEGAL_CHARACTERS_RE.sub("", value))
            cell.data_type = "s"
            return cell

        for number, source, ids, text, error in scan_ocr:
            status = error or ("Review OCR against the image" if ids else "No record IDs detected")
            review.append([number, text_cell(review, source), len(ids),
                           text_cell(review, text), text_cell(review, status)])
            for identifier in ids:
                detected.append([number, text_cell(detected, source), text_cell(detected, identifier)])
    path = Path(path)
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        workbook.save(temporary_path)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)
