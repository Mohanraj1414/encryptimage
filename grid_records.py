"""Match record IDs printed on decrypted grid pages to Client 2 MI data."""
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import PurePosixPath
import re
from subprocess import TimeoutExpired
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from image_reader import scan_bytes
from kyc import FIELDS, MISSING, export
from scan_ocr import ocr_text, tesseract_path

MI_HEADERS = ("ID", *range(1, 17), 18, 19, 20, 21)
MI_FIELDS = FIELDS[:3] + FIELDS[4:-1]
MATCHED_EXPORT = re.compile(r"KYC_.+_\d{8}_\d{6}_\d{6}\.xlsx", re.IGNORECASE)
PRINTED_ID = re.compile(
    r"(?<![a-z0-9])(?:fbt[\s_-]*xhu[\s_-]*ent|hnt[\s_-]*dup[\s_-]*akyz)"
    r"[\s_-]+[0-9]+(?![a-z0-9])", re.IGNORECASE,
)
# OCR can lose the first segment in a two-column grid. The remaining
# two segments still distinguish the supported record families.
PARTIAL_ID = re.compile(
    r"(?<![a-z0-9])(?P<family>dup[\s_-]*akyz|xhu[\s_-]*ent)[\s_-]+"
    r"(?P<number>[0-9]+)(?![a-z0-9])", re.IGNORECASE,
)
ID_PREFIX = {"dupakyz": "hnt", "xhuent": "fbt"}


def canonical_id(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def load_mi(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = book.worksheets[0].values
        headers = tuple(int(v) if str(v).isdigit() else v for v in next(rows, ()))
        if headers[:len(MI_HEADERS)] != MI_HEADERS:
            raise ValueError("Choose the Client 2 KYC MI workbook with ID and numbered column headers.")
        index = {}
        for row in rows:
            if not row[0]:
                continue
            record = dict(zip(MI_FIELDS, row[1:21]))
            identifier = canonical_id(row[0])
            if canonical_id(record["Record No"]) != identifier:
                raise ValueError("The KYC MI workbook contains mismatched ID and Record No values.")
            if identifier in index and index[identifier] != record:
                raise ValueError("The KYC MI workbook contains conflicting rows for one record ID.")
            index[identifier] = record
        if not index:
            raise ValueError("The KYC MI workbook contains no records.")
        return index
    finally:
        book.close()


def record_order(key):
    family = key.rstrip("0123456789")
    return family, int(key[len(family):])


def match_ids(text, index):
    found = {}
    text = " ".join(text.split())
    for match in PRINTED_ID.finditer(text):
        printed = match.group()
        found.setdefault(canonical_id(printed), printed)
    for match in PARTIAL_ID.finditer(PRINTED_ID.sub(" ", text)):
        family = canonical_id(match["family"])
        key = ID_PREFIX[family] + family + match["number"]
        found.setdefault(key, match.group())

    ordered = sorted(found, key=record_order)
    matched = [index[key] for key in ordered if key in index]
    unmatched = [found[key] for key in ordered if key not in index]
    gaps = []
    for left, right in zip(ordered, ordered[1:]):
        family, low = record_order(left)
        next_family, high = record_order(right)
        if family == next_family and high > low + 1:
            gaps.append(str(low + 1) if high == low + 2 else f"{low + 1}-{high - 1}")
    return matched, unmatched, gaps


def read_prior_export(path, scan_names):
    with path.open("rb") as source:
        book = load_workbook(source, read_only=True, data_only=True)
        try:
            if not {"KYC Data", "Scan review"}.issubset(book.sheetnames):
                return []
            records = iter(book["KYC Data"].values)
            report = iter(book["Scan review"].values)
            if next(records, ()) != FIELDS or next(report, ())[:3] != (
                "Image Name", "Source image", "Matched rows"
            ):
                return []
            pages = {}
            for page, name, count, *_ in report:
                if not count or not isinstance(name, str):
                    continue
                source_name = PurePosixPath(name)
                folder = re.sub(r"^\d{3,}_", "", source_name.parent.name)
                scan = (folder.casefold(), source_name.name.casefold())
                if scan in scan_names:
                    pages[str(page)] = scan
            if not pages:
                return []
            return [
                (pages[str(row[3])], canonical_id(row[0]), dict(zip(FIELDS, row)))
                for row in records
                if str(row[3]) in pages
                and isinstance(row[0], str)
                and PRINTED_ID.fullmatch(row[0])
                and any(str(value or "").strip().casefold() not in MISSING for value in row[4:-1])
            ]
        finally:
            book.close()


def load_prior_exports(folder, scans):
    scan_names = {(image.parent.name.casefold(), image.name.casefold()) for image, _ in scans}
    index = {}
    conflicts = set()
    for path in folder.glob("KYC_*.xlsx"):
        if not MATCHED_EXPORT.fullmatch(path.name) or path.name.casefold().startswith(("kyc_ocr_", "kyc_recovered_")):
            continue
        try:
            entries = read_prior_export(path, scan_names)
        except (OSError, ValueError, BadZipFile, InvalidFileException, KeyError, SyntaxError):
            continue
        for scan, identifier, record in entries:
            key = (scan, identifier)
            record.pop("Image Name")
            previous = index.setdefault(scan, {}).setdefault(identifier, record)
            if previous != record:
                conflicts.add(key)
    for scan, identifier in conflicts:
        del index[scan][identifier]
    return index


def export_grid(path, scans, mi_path, progress, *, recover=False):
    index = load_mi(mi_path) if mi_path is not None else {}
    prior = load_prior_exports(path.parent, scans) if recover else {}
    ocr_only = mi_path is None and not recover
    executable = tesseract_path()

    def recognize(scan):
        image, key = scan
        try:
            if isinstance(key, ValueError):
                raise key
            text = ocr_text(executable, scan_bytes(image, key), 11)
            page_index = prior.get((image.parent.name.casefold(), image.name.casefold()), {}) if recover else index
            matched, unmatched, gaps = match_ids(text, page_index)
            return matched, unmatched, gaps, text, ""
        except (OSError, ValueError, TimeoutExpired) as error:
            return [], [], [], "", str(error)

    records = []
    report = []
    scan_ocr = []
    matched_count = 0
    flagged = 0
    folder_numbers = {folder: number for number, folder in enumerate(
        dict.fromkeys(image.parent for image, _ in scans), 1
    )}
    workers = min(4, max(1, len(scans)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for number, (matched, unmatched, gaps, text, error) in enumerate(pool.map(recognize, scans), 1):
            image = scans[number - 1][0]
            matched_count += len(matched)
            page_records = matched
            if recover:
                review_rows = [
                    {"Record No": identifier, "Remark": "REVIEW: No source data for this record ID"}
                    for identifier in unmatched if PRINTED_ID.fullmatch(identifier)
                ]
                page_records = sorted(
                    [*matched, *review_rows], key=lambda row: record_order(canonical_id(row["Record No"]))
                )
            records.extend({**row, "Image Name": str(number)} for row in page_records)
            source = f"{folder_numbers[image.parent]:03d}_{image.parent.name}/{image.name}"
            if ocr_only:
                scan_ocr.append((number, source, unmatched, text, error))
            else:
                status = error or ("No record IDs matched" if not matched else "Check against the page")
                if gaps:
                    status += ". Check ID gaps: " + ", ".join(gaps)
                report.append([number, source, len(matched), ", ".join(unmatched), status])
                flagged += bool(error or not matched or unmatched or gaps)
            progress(number)
    if ocr_only:
        export(path, None, scan_ocr=scan_ocr)
        return sum(len(row[2]) for row in scan_ocr), sum(
            not row[2] or bool(row[4]) for row in scan_ocr
        )
    counts = Counter(record["Record No"] for record in records)
    duplicates = {key for key, count in counts.items() if count > 1}
    export(path, records, report, sorted(duplicates), normalized=recover)
    return matched_count, flagged
