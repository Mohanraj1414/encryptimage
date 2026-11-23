"""Match record IDs printed on decrypted grid pages to Client 2 MI data."""
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import re
from subprocess import TimeoutExpired

from openpyxl import load_workbook

from image_reader import scan_bytes
from kyc import FIELDS, export
from scan_ocr import ocr_text, tesseract_path

MI_HEADERS = ("ID", *range(1, 17), 18, 19, 20, 21)
MI_FIELDS = FIELDS[:3] + FIELDS[4:-1]
PRINTED_ID = re.compile(
    r"(?<![a-z0-9])(?:fbt[\s_-]*xhu[\s_-]*ent|hnt[\s_-]*dup[\s_-]*akyz)"
    r"[\s_-]+[0-9]+(?![a-z0-9])", re.IGNORECASE,
)


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


def match_ids(text, index):
    found = {}
    for match in PRINTED_ID.finditer(text.replace("\n", " ")):
        printed = match.group()
        found.setdefault(canonical_id(printed), printed)
    matched = [index[key] for key in found if key in index]
    unmatched = [printed for key, printed in found.items() if key not in index]
    return matched, unmatched


def export_grid(path, scans, mi_path, progress):
    index = load_mi(mi_path)
    executable = tesseract_path()

    def recognize(scan):
        image, key = scan
        try:
            if isinstance(key, ValueError):
                raise key
            text = ocr_text(executable, scan_bytes(image, key), 11)
            matched, unmatched = match_ids(text, index)
            return matched, unmatched, ""
        except (OSError, ValueError, TimeoutExpired) as error:
            return [], [], str(error)

    records = []
    report = []
    folder_numbers = {folder: number for number, folder in enumerate(
        dict.fromkeys(image.parent for image, _ in scans), 1
    )}
    workers = min(4, max(1, len(scans)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for number, (matched, unmatched, error) in enumerate(pool.map(recognize, scans), 1):
            image = scans[number - 1][0]
            records.extend({**row, "Image Name": str(number)} for row in matched)
            report.append([number, f"{folder_numbers[image.parent]:03d}_{image.parent.name}/{image.name}", len(matched),
                           ", ".join(unmatched), error or (
                               "No record IDs matched" if not matched else "Check against the page"
                           )])
            progress(number)
    counts = Counter(record["Record No"] for record in records)
    duplicates = {key for key, count in counts.items() if count > 1}
    export(path, records, report, sorted(duplicates))
    return len(records), sum(not row[2] or bool(row[3]) for row in report)
