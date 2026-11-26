"""Match record IDs printed on decrypted grid pages to Client 2 MI data."""
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from pathlib import PurePosixPath
import re
from subprocess import TimeoutExpired
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from PIL import Image, ImageOps

from center_codes import is_center_code, restore_center_codes
from image_reader import open_image, scan_bytes
from kyc import FIELDS, MISSING, export, normalize
from scan_ocr import ocr_text, tesseract_path, windows_ocr_path

MI_HEADERS = ("ID", *range(1, 17), 18, 19, 20, 21)
MI_FIELDS = FIELDS[:3] + FIELDS[4:-1]
MATCHED_EXPORT = re.compile(r"KYC_.+_\d{8}_\d{6}_\d{6}\.xlsx", re.IGNORECASE)
PRINTED_ID = re.compile(
    r"(?<![a-z0-9])(?:fbt[\s_-]*xhu[\s_-]*ent|hnt[\s_-]*dup[\s_-]*akyz|"
    r"[a-z]{3}\s*[_-][\s_-]*[a-z]{3}[\s_-]+[a-z]{3,4}|"
    r"[a-z]{3}\s+[a-z]{3}\s*[_-][\s_-]*[a-z]{3,4})"
    r"[\s_-]+[0-9]+(?![a-z0-9])", re.IGNORECASE,
)
# OCR can lose the first segment in a two-column grid. The remaining
# two segments still distinguish the supported record families.
PARTIAL_ID = re.compile(
    r"(?<![a-z0-9])(?P<family>dup[\s_-]*akyz|xhu[\s_-]*ent|brt[\s_-]*nay)[\s_-]+"
    r"(?P<number>[0-9]+)(?![a-z0-9])", re.IGNORECASE,
)
ID_PREFIX = {"dupakyz": "hnt", "xhuent": "fbt", "brtnay": "fbk"}
SCAN_FAMILY = re.compile(r"([a-z]{3})[_-]([a-z]{3})[_-]([a-z]{3,4})\s*\(\d+\)", re.IGNORECASE)


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


def ruled_grid(payload):
    """Return the grayscale image, column starts, and row boundaries."""
    try:
        with open_image(BytesIO(payload)) as image:
            gray = image.convert("L")
    except (OSError, Image.DecompressionBombError):
        return None
    width, height = gray.size
    mask = gray.point(lambda shade: 255 if shade < 155 else 0)
    vertical = mask.resize((width, 1), Image.Resampling.BOX)
    left = [x for x in range(width // 8) if vertical.getpixel((x, 0)) >= 163]
    middle = [x for x in range(width * 2 // 5, width * 3 // 5) if vertical.getpixel((x, 0)) >= 163]
    if not left or not middle:
        return None
    horizontal = mask.resize((1, height), Image.Resampling.BOX)
    rules = []
    for y in range(height):
        if horizontal.getpixel((0, y)) >= 163:
            if rules and y <= rules[-1] + 2:
                rules[-1] = y
            else:
                rules.append(y)
    if not 5 <= len(rules) <= 24:
        return None
    return gray, (left[-1] + 2, middle[-1] + 2), list(zip(rules, rules[1:]))


def ocr_png(sheet, scale):
    with BytesIO() as buffer:
        scale = min(scale, 2400 / max(sheet.size))
        size = tuple(round(length * scale) for length in sheet.size)
        with sheet.resize(size, Image.Resampling.LANCZOS) as expanded:
            expanded.save(buffer, format="PNG")
        return buffer.getvalue()


def grid_id_images(payload):
    """Build two OCR views of cell headers on a ruled, two-column scan."""
    grid = ruled_grid(payload)
    if grid is None:
        return ()
    gray, columns, starts = grid
    width = gray.width
    views = []
    for width_ratio, height_ratio in ((.135, .0154), (.147, .0188)):
        strip_width = max(1, round(width * width_ratio))
        strip_height = max(1, round(width * height_ratio))
        step = max(26, strip_height + 6)
        with Image.new("L", (strip_width + 35, step * len(starts) * 2 + 12), 255) as sheet:
            for row, (top, bottom) in enumerate(starts):
                for column, x in enumerate(columns):
                    start = top + 2
                    end = min(bottom, top + strip_height)
                    if end > start:
                        with gray.crop((x, start, x + strip_width, end)) as strip:
                            sheet.paste(strip, (12, 6 + (2 * row + column) * step))
            views.append(ocr_png(sheet, 2))
    return views


def grid_center_images(payload):
    """Build OCR views of the bottom lines of each two-column record cell."""
    grid = ruled_grid(payload)
    if grid is None:
        return
    gray, columns, bands = grid
    width = columns[1] - columns[0] - 16
    for nominal_depth, scale, contrast in ((24, 2, False), (22, 3, True), (35, 2, False)):
        depth = max(10, round(nominal_depth * gray.width / 1170))
        step = depth + 8
        with Image.new("L", (width + 28, (len(bands) + 1) * 2 * step), 255) as sheet:
            for row, (top, bottom) in enumerate(bands):
                start = max(top + 2, bottom - depth)
                if start >= bottom - 1:
                    continue
                for column, x in enumerate(columns):
                    with gray.crop((x, start, x + width, bottom - 1)) as strip:
                        sheet.paste(ImageOps.autocontrast(strip) if contrast else strip,
                                    (10, (row * 2 + column) * step))
            yield ocr_png(sheet, scale)


def detect_ids(text, scan_stem=""):
    found = {}
    text = " ".join(text.split())
    for match in PRINTED_ID.finditer(text):
        printed = match.group()
        found.setdefault(canonical_id(printed), printed)
    family = SCAN_FAMILY.fullmatch(scan_stem)
    if family:
        # A scan filename supplies the letter pattern only. Its file number
        # has no relation to the individual record numbers printed on the page.
        pattern = r"(?<![a-z0-9])" + r"[\s_-]*".join(family.groups()) + r"[\s_-]+([0-9]+)(?![a-z0-9])"
        for match in re.finditer(pattern, text, re.IGNORECASE):
            printed = f"{family[1]}_{family[2]}-{family[3]}-{match[1]}"
            found.setdefault(canonical_id(printed), printed)
    for match in PARTIAL_ID.finditer(PRINTED_ID.sub(" ", text)):
        family = canonical_id(match["family"])
        key = ID_PREFIX[family] + family + match["number"]
        found.setdefault(key, match.group())

    return found


def match_ids(text, index):
    return match_detected_ids(detect_ids(text), index)


def match_detected_ids(found, index):
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


def select_batch_indexes(preferred, alternatives, detected):
    """Use the chosen MI source where it matches, then compare nearby sources by ID."""
    selected = {folder: preferred for folder, identifiers in detected.items() if preferred.keys() & identifiers}
    remaining = detected.keys() - selected.keys()
    if not remaining:
        return selected

    needed = set().union(*(detected[folder] for folder in remaining))
    if not needed:
        return selected
    indexes = []
    for path in alternatives:
        try:
            index = load_mi(path)
        except (OSError, ValueError, BadZipFile, InvalidFileException, KeyError, SyntaxError):
            continue
        indexes.append({key: index[key] for key in index.keys() & needed})
    for folder in remaining:
        candidates = [{key: index[key] for key in detected[folder] & index.keys()} for index in indexes]
        largest = max(map(len, candidates), default=0)
        if not largest:
            continue
        best = [candidate for candidate in candidates if len(candidate) == largest]
        if any(candidate != best[0] for candidate in best[1:]):
            raise ValueError(
                f"Several MI workbooks match {folder.name} but contain different records. "
                "Choose the correct one under File > Change source workbook."
            )
        selected[folder] = best[0]
    return selected


def missing_mi_reason(folder, scans, recognized):
    pages = [(image, found, error) for (image, _), (found, _, error) in zip(scans, recognized)
             if image.parent == folder]
    families = {}
    for image, _, _ in pages:
        match = SCAN_FAMILY.fullmatch(image.stem)
        if match:
            families["".join(match.groups()).casefold()] = f"{match[1]}_{match[2]}-{match[3]}"

    sample = None
    for _, found, _ in pages:
        candidates = [key for key, printed in found.items() if
                      (key.rstrip("0123456789") in families if families else PRINTED_ID.fullmatch(printed))]
        if candidates:
            sample = found[min(candidates, key=record_order)]
            break
    if sample:
        return (
            f"No MI source matched the printed ID {sample} in {folder.name}. "
            "Choose a KYC MI workbook for this batch under File > Change source workbook."
        )
    if families:
        expected = ", ".join(families.values())
        reason = f"No record IDs with the expected prefix {expected} could be read in {folder.name}."
    else:
        reason = f"No printed record IDs could be read in {folder.name}."
    first_error = next((error for _, _, error in pages if error), "")
    return f"{reason} {first_error or 'Check the source scans.'}"


def export_grid(path, scans, mi_path, progress, *, recover=False, alternatives=(), require_matches=False):
    index = load_mi(mi_path) if mi_path is not None else {}
    ocr_only = mi_path is None and not recover
    executable = tesseract_path()
    center_engine = windows_ocr_path() or executable

    def recognize_ids(scan):
        image, key = scan
        try:
            if isinstance(key, ValueError):
                raise key
            payload = scan_bytes(image, key)
            text = ocr_text(executable, payload, 11)
            id_text = [text]
            for view in grid_id_images(payload):
                try:
                    id_text.append(ocr_text(executable, view, 6))
                except (OSError, ValueError, TimeoutExpired):
                    # Preserve results from the original scan if a supplemental pass fails.
                    continue
            return detect_ids("\n".join(id_text), image.stem), text, ""
        except (OSError, ValueError, TimeoutExpired) as error:
            return {}, "", str(error)

    def prepare_page(item):
        (image, key), (found, text, error) = item
        from_prior = image.parent not in selected
        page_index = (
            prior.get((image.parent.name.casefold(), image.name.casefold()), {})
            if from_prior else selected[image.parent]
        )
        matched, unmatched, gaps = match_detected_ids(found, page_index)
        if any(is_center_code(record.get("Center Code")) for record in matched):
            readings = []
            try:
                payload = scan_bytes(image, key)
            except (OSError, ValueError):
                return matched, unmatched, gaps, text, error, from_prior
            for view in grid_center_images(payload):
                try:
                    readings.append(ocr_text(center_engine, view, 6))
                except (OSError, ValueError, TimeoutExpired):
                    if center_engine != executable:
                        try:
                            readings.append(ocr_text(executable, view, 6))
                        except (OSError, ValueError, TimeoutExpired):
                            continue
            matched = restore_center_codes(matched, readings)
        return matched, unmatched, gaps, text, error, from_prior

    records = []
    report = []
    scan_ocr = []
    matched_count = 0
    flagged = 0
    folder_numbers = {folder: number for number, folder in enumerate(
        dict.fromkeys(image.parent for image, _ in scans), 1
    )}
    import os as _os
    cpu = _os.cpu_count() or 4
    workers = min(max(cpu, 8), max(1, len(scans)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        recognized = []
        detected = {folder: set() for folder in folder_numbers}
        for number, reading in enumerate(pool.map(recognize_ids, scans), 1):
            recognized.append(reading)
            detected[scans[number - 1][0].parent].update(reading[0])
            progress(number)
        selected = select_batch_indexes(index, alternatives, detected) if not ocr_only else {}
        prior = load_prior_exports(path.parent, scans) if not ocr_only and detected.keys() - selected.keys() else {}
        if require_matches:
            supported = set(selected)
            for (image, _), (found, _, _) in zip(scans, recognized):
                if prior.get((image.parent.name.casefold(), image.name.casefold()), {}).keys() & found.keys():
                    supported.add(image.parent)
            for folder in detected.keys() - supported:
                raise ValueError(missing_mi_reason(folder, scans, recognized))
        results = pool.map(prepare_page, zip(scans, recognized))
        for number, (matched, unmatched, gaps, text, error, from_prior) in enumerate(results, 1):
            image = scans[number - 1][0]
            matched_count += len(matched)
            page_records = [{**row, "Image Name": str(number)} for row in matched]
            if not from_prior:
                page_records = [dict(zip(FIELDS, normalize(row))) for row in page_records]
            if not ocr_only and (from_prior or not matched):
                review_rows = [
                    {"Record No": identifier, "Image Name": str(number),
                     "Remark": "REVIEW: No source data for this record ID"}
                    for identifier in unmatched if PRINTED_ID.fullmatch(identifier)
                ]
                page_records.extend(review_rows)
                page_records.sort(key=lambda row: record_order(canonical_id(row["Record No"])))
            records.extend(page_records)
            source = f"{folder_numbers[image.parent]:03d}_{image.parent.name}/{image.name}"
            if ocr_only:
                scan_ocr.append((number, source, [str(u) for u in unmatched if u is not None], text, error))
            else:
                status = error or (
                    "Check against the page" if matched else "No source match. Review the detected IDs."
                )
                if gaps:
                    status += ". Check ID gaps: " + ", ".join(str(g) for g in gaps if g is not None)
                report.append([number, source, len(matched), ", ".join(str(u) for u in unmatched if u is not None), status])
                flagged += bool(error or not matched or unmatched or gaps)
    if ocr_only:
        export(path, None, scan_ocr=scan_ocr)
        return sum(len(row[2]) for row in scan_ocr), sum(
            not row[2] or bool(row[4]) for row in scan_ocr
        )
    counts = Counter(record["Record No"] for record in records)
    duplicates = {key for key, count in counts.items() if count > 1}
    export(path, records, report, sorted(duplicates), normalized=True)
    return matched_count, flagged
