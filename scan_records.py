"""Extract the printed fields from a single OCR-read KYC grid cell."""

import re

from center_codes import is_center_code


OCR_REMARK = "REVIEW: OCR FROM SCAN"
FIELD_GAP = re.compile(r"[ \t]{3,}")
ACCOUNT = re.compile(r"^(.+?\([sn][b8s]s[}\]]?\)?)\s*", re.IGNORECASE)
STREET_END = re.compile(
    r"\b(?:alley|avenue|ave|branch|bridge|boulevard|blvd|circle|court|creek|drive|dr|"
    r"falls|fork|forge|freeway|grove|haven|highway|hill|hills|junction|landing|lane|"
    r"ln|loop|meadow|park|parkway|pass|path|pike|plain|port|road|rd|river|route|run|"
    r"shores?|squares?|street|st|summit|terrace|trail|track|view|walk|way)\b\.?",
    re.IGNORECASE,
)
WEEKDAY = r"(?:monday|tuesday|teusday|wednesday|wenesday|thursday|thrusday|friday|saturday|satruday|sunday)"
DATE_DIGITS = r"(?:\d{8}|\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{2,4})(?!\d)"
WEEKDAY_DOB = re.compile(rf"{WEEKDAY}\s*[,.-]?\s*{DATE_DIGITS}", re.IGNORECASE)
DOB = re.compile(rf"(?<!\w)(?:[a-z]{{3,12}}day\s*[,.-]?\s*)?{DATE_DIGITS}", re.IGNORECASE)
GENDER = re.compile(r"^[\s_]*(female|fimale|firmale|firnale|faimale|femail|male)\b[. ,]*", re.IGNORECASE)
NATIONALITY = re.compile(
    r"^\s*((?:united|unitid|unlted|unitited)\s*states(?:\s+of\s+america)?|"
    r"american|u\.?s\.?a\.?)\b[. ,]*", re.IGNORECASE,
)
STATE_ZIP = re.compile(
    r"(AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|"
    r"MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)"
    r"\s+(?:[=–—]\s*)?([0-9$]{5})(?:-\d{4})?(?!\d)", re.IGNORECASE,
)
EMAIL = re.compile(
    r"(?<![\w.-])[a-z0-9_%+-]+(?:\.[ \t]*[a-z0-9_%+-]+)*"
    r"[ \t]*[@#][ \t]*[a-z0-9-]+(?:[ \t]*\.[ \t]*[a-z0-9-]+)*", re.IGNORECASE,
)
PHONE = re.compile(r"(?<!\d)(?:1[.-])?\(?\d{3}\)?[. -]+\d{3}[. -]+\s*\d{4}"
                   r"(?!\d)(?:[ \t]+\d{3,5}(?!\d|[.-]\d))?")
VERIFICATION = re.compile(r"\b(NOT\s+VERIFIED|VERIFIED|VERIFED|UNVERIFIED)\b", re.IGNORECASE)
ID_VERIFICATION = re.compile(r"^\s*(YES|YESS|YAS|YYS|NO|N\s*/\s*A)\b", re.IGNORECASE)
NAME_WORD = re.compile(r"[^\W\d_]+(?:[’'.-][^\W\d_]+)*\.?", re.UNICODE)
STREET_START = re.compile(r"^[\d$]+[ \t]*[a-z]", re.IGNORECASE)


def clean_boundary(value):
    return value.strip(" \t\n\r=–—_")


def name_words(value):
    words = value.split()
    return 2 <= len(words) <= 8 and all(NAME_WORD.fullmatch(word) for word in words)


def split_names(text):
    text = clean_boundary(text)
    lines = [clean_boundary(line) for line in text.splitlines() if clean_boundary(line)]
    joined = re.sub(r"[ \t]+[=–—][ \t]+", "   ", " ".join(lines))
    groups = [part for part in FIELD_GAP.split(joined) if part]
    if len(groups) == 2 and all(map(name_words, groups)):
        return groups
    if len(lines) == 2 and all(map(name_words, lines)):
        return lines
    # The scan often uses a different casing for the two names, with only one
    # OCR space between them. Accept a unique transition supported by both sides.
    words = joined.split()
    candidates = []
    for index in range(2, len(words) - 1):
        left, right = words[:index], words[index:]
        pair = [" ".join(left), " ".join(right)]
        if all(word.islower() for word in left) and right[0][0].isupper() and not right[0].isupper():
            candidates.append(pair)
        elif all(word.isupper() for word in left) and all(word.islower() for word in right):
            candidates.append(pair)
    if len(candidates) == 1 and all(map(name_words, candidates[0])):
        return candidates[0]
    return ()


def read_account_and_names(record, header):
    account = ACCOUNT.match(header)
    if account:
        record["Type Of Account"] = account[1]
        rest = header[account.end():].lstrip()
    else:
        first = header.split("\n", 1)[0]
        address_at = re.search(r"(?<!\w)\d{2,6}(?=\s*[a-z])", first, re.IGNORECASE)
        if not address_at or not address_at.start():
            return
        record["Type Of Account"] = clean_boundary(header[:address_at.start()])
        rest = header[address_at.start():]
    if not STREET_START.match(rest):
        return
    first, _, tail = rest.partition("\n")
    gap = FIELD_GAP.search(first)
    title = re.search(r"(?<=[a-z])(?=(?:Mr|Mrs|Ms|Dr)\.\s+[A-Z])", first)
    street = list(STREET_END.finditer(first))
    if gap and STREET_START.match(first[:gap.start()]):
        boundary = gap.start()
    elif title:
        boundary = title.start()
    elif street:
        boundary = street[-1].end()
    elif tail and split_names(tail):
        boundary = len(first)
    else:
        return
    record["Rag Address"] = clean_boundary(first[:boundary])
    names = split_names(first[boundary:].strip() + "\n" + tail)
    if names:
        record["Full Name"], record["Father/Husband Name"] = names


def read_location(record, body):
    gender = GENDER.match(body)
    if gender:
        record["Gender"] = gender[1]
        body = body[gender.end():]
    nationality = NATIONALITY.match(body)
    if nationality:
        record["Nationality"] = nationality[1]
        body = body[nationality.end():]
    email = EMAIL.search(body)
    location = body[:email.start()] if email else body
    state_zip = STATE_ZIP.search(location)
    if not state_zip:
        return body
    record["State"] = state_zip[1]
    record["Pin Code"] = state_zip[2]
    address = clean_boundary(body[:state_zip.start()])
    if STREET_START.match(address):
        chunks = FIELD_GAP.split(address, maxsplit=1)
        street = list(STREET_END.finditer(address))
        if len(chunks) == 2:
            record["Address"], record["City/District/Town"] = map(clean_boundary, chunks)
        elif street and clean_boundary(address[street[-1].end():]):
            record["Address"] = address[:street[-1].end()]
            record["City/District/Town"] = clean_boundary(address[street[-1].end():])
    return body[state_zip.end():]


def center_code(text):
    # Tesseract can render the printed uppercase Y as a yen sign.
    code = "".join(clean_boundary(text).strip(".,:").split()).replace("\u00a5", "Y")
    if is_center_code(code) and re.search(r"[A-Z]", code) and re.search(r"[a-z0-9]", code):
        return code
    return ""


def read_center(record, text):
    confirmation = ID_VERIFICATION.match(text)
    if not confirmation:
        return
    record["ID Verification"] = confirmation[1]
    value = clean_boundary(text[confirmation.end():])
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if not lines:
        return
    # The code can follow the center on the same line or wrap onto its own.
    chunks = re.split(r"[ \t]{3,}|[ \t]+[=–—][ \t]+", lines[0], maxsplit=1)
    if len(lines) == 1 and len(chunks) == 1:
        name_and_code = lines[0].split(maxsplit=1)
        if len(name_and_code) == 2 and center_code(name_and_code[1]):
            chunks = name_and_code
    if re.fullmatch(r"[\w .'-]+", chunks[0]):
        record["Center Name"] = chunks[0]
    possible_code = (" ".join(chunks[1:]) + " " + " ".join(lines[1:])).strip()
    code = center_code(possible_code)
    if code:
        record["Center Code"] = code


def parse_scan_record(identifier, body):
    body = "\n".join(line.strip() for line in body.splitlines() if line.strip())
    record = {"Record No": identifier, "Remark": OCR_REMARK}
    dob = WEEKDAY_DOB.search(body) or DOB.search(body)
    if dob:
        read_account_and_names(record, body[:dob.start()])
        record["Date Of Birth"] = re.sub(r"(?<=\d)\s*([-/.])\s*(?=\d)", r"\1", " ".join(dob[0].split()))
        body = read_location(record, body[dob.end():])
    else:
        read_account_and_names(record, body.split("\n", 1)[0])

    email = EMAIL.search(body)
    if email:
        record["Email"] = "".join(email[0].split())
        if "Pin Code" in record:
            mobile = PHONE.findall(" ".join(body[:email.start()].splitlines()))
            if len(mobile) == 1:
                record["Mobile Number"] = mobile[0]
    verification = VERIFICATION.search(body, email.end() if email else 0)
    if verification:
        record["Address Verification"] = verification[0]
        if email:
            contacts = PHONE.findall(" ".join(body[email.end():verification.start()].splitlines()))
            if len(contacts) == 2:
                record["Contact 1"], record["Contact 2"] = contacts
        read_center(record, body[verification.end():])
    if "Center Code" not in record and "\n" in body:
        last_line = body.rsplit("\n", 1)[-1]
        if re.search(r"[A-Za-z0-9\u00a5]{12,}", last_line):
            code = center_code(last_line)
            if code:
                record["Center Code"] = code
    if len(record) > 2:
        return record
