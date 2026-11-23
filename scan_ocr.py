"""Extract explicitly labeled fields from a scan using local Tesseract."""
import os
from pathlib import Path
import re
import shutil
import subprocess

from image_reader import scan_bytes
from kyc import FIELDS


def tesseract_path():
    executable = shutil.which("tesseract")
    if executable:
        return executable
    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        directory = os.environ.get(variable)
        if directory:
            candidate = Path(directory) / "Tesseract-OCR" / "tesseract.exe"
            if candidate.is_file():
                return str(candidate)
    raise ValueError("Install local Tesseract OCR to read record IDs from the decrypted images.")


def ocr_text(executable, payload, page_mode):
    result = subprocess.run(
        [executable, "stdin", "stdout", "--psm", str(page_mode)],
        input=payload, capture_output=True, timeout=60,
        env={**os.environ, "OMP_THREAD_LIMIT": "1"},
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if result.returncode:
        raise ValueError("OCR could not read this image.")
    return result.stdout.decode("utf-8", errors="replace")


def read_fields(path, key):
    return parse_fields(ocr_text(tesseract_path(), scan_bytes(path, key), 6))


def parse_fields(text):
    labels = {field.casefold(): field for field in FIELDS if field not in {"Image Name", "Remark"}}
    labels.update({"record no.": "Record No", "registration address": "Rag Address",
                   "father name": "Father/Husband Name", "husband name": "Father/Husband Name",
                   "dob": "Date Of Birth", "city": "City/District/Town"})
    fields = {}
    for line in text.splitlines():
        match = re.fullmatch(r"\s*([^:]+):\s*(\S.*?)\s*", line)
        if match:
            label, value = match.groups()
            field = labels.get(label.strip().casefold())
            if field:
                fields.setdefault(field, value)
    return fields
