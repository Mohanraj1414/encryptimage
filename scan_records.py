"""Extract fields from the OCR text for one cell of a KYC grid."""

import re

from center_codes import is_center_code


OCR_REMARK = "REVIEW: OCR FROM SCAN"
FIELD_GAP = re.compile(r"[ \t]{3,}")
ACCOUNT = re.compile(r"^(.+?\([sn]bs\))\s*", re.IGNORECASE)
DOB = re.compile(
    r"(?<!\w)(?:[a-z]{3,12}day[ \t]*[,.-]?[ \t]*)?"
    r"(?:\d{8}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})(?!\d)", re.IGNORECASE
)
GENDER = re.compile(r"^\s*(female|fimale|femail|male)\b[. ]*", re.IGNORECASE)
NATIONALITY = re.compile(
    r"^\s*(united\s*states(?:\s+of\s+america)?|unitid\s*states|american|u\.?s\.?a\.?)\b",
    re.IGNORECASE,
)
STATE_ZIP = re.compile(
    r"(AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|"
    r"MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)"
    r"\s+(\d{5})(?:-\d{4})?(?!\d)", re.IGNORECASE
)
EMAIL = re.compile(r"(?<![\w.-])[a-z0-9_%+-]+(?:\.[ \t]*[a-z0-9_%+-]+)*"
                   r"[ \t]*@[ \t]*[a-z0-9-]+(?:\.[a-z0-9-]+)*\.[a-z]{2,}\b", re.IGNORECASE)
PHONE = re.compile(r"(?<!\d)(?:1[.-])?\(?\d{3}\)?[. -]+\d{3}[. -]+\s*\d{4}"
                   r"(?!\d)(?:[ \t]+\d{3,5}(?!\d|[.-]\d))?")
VERIFICATION = re.compile(r"\b(NOT\s+VERIFIED|VERIFIED|VERIFED|UNVERIFIED)\b", re.IGNORECASE)
ID_VERIFICATION = re.compile(r"^\s*(YES|YESS|YAS|NO|N\s*/\s*A)\b", re.IGNORECASE)


def parse_scan_record(identifier, body):
    record = {"Record No": identifier, "Remark": OCR_REMARK}
    dob = DOB.search(body)
    before_dob = body[:dob.start()] if dob else body
    account = ACCOUNT.match(before_dob)
    if account:
        record["Type Of Account"] = account[1]
        if dob:
            # With a field gap, a name can wrap to the next line. Without a
            # boundary, leave the affected fields for review.
            parts = FIELD_GAP.split(before_dob[account.end():].replace("\n", " ").strip())
            if len(parts) >= 2 and re.match(r"^\d+\s*\D", parts[0]):
                record["Rag Address"] = parts[0]
                if len(parts) == 3:
                    record["Full Name"], record["Father/Husband Name"] = parts[1:]
    if dob:
        record["Date Of Birth"] = dob[0]
        location = body[dob.end():]
        gender = GENDER.match(location)
        if gender:
            record["Gender"] = gender[1]
            location = location[gender.end():]
        nationality = NATIONALITY.match(location)
        if nationality:
            record["Nationality"] = nationality[1]
            location = location[nationality.end():].lstrip()
            state_zip = STATE_ZIP.search(location)
            if state_zip:
                record["State"], record["Pin Code"] = state_zip.groups()
                address = location[:state_zip.start()].rstrip()
                parts = re.split(r"(?:[ \t]{3,}|\n)\s*", address)
                if len(parts) >= 2:
                    record["Address"] = " ".join(parts[:-1])
                    record["City/District/Town"] = parts[-1]
                body = location[state_zip.end():]

    email = EMAIL.search(body)
    if email:
        record["Email"] = "".join(email[0].split())
        if "Pin Code" in record:
            mobile = list(PHONE.finditer(" ".join(body[:email.start()].splitlines())))
            if len(mobile) == 1:
                record["Mobile Number"] = mobile[0][0]
    verification = VERIFICATION.search(body, email.end() if email else 0)
    if verification:
        record["Address Verification"] = verification[0]
        if email:
            contacts = PHONE.findall(" ".join(body[email.end():verification.start()].splitlines()))
            if len(contacts) == 2:
                record["Contact 1"], record["Contact 2"] = contacts
        id_verification = ID_VERIFICATION.match(body[verification.end():])
        if id_verification:
            record["ID Verification"] = id_verification[1]
            center = body[verification.end() + id_verification.end():].strip(" \t\ufffd")
            lines = center.splitlines()
            if lines and re.fullmatch(r"[\w .'-]+", lines[0]):
                record["Center Name"] = lines[0].strip()
            if len(lines) == 2:
                # Tesseract can read a printed uppercase Y as a yen sign.
                code = lines[1].strip().replace("\u00a5", "Y")
                if is_center_code(code) and code.isalnum():
                    record["Center Code"] = code
    if len(record) > 2:
        return record
