"""Read plain scans and legacy AES-encrypted scans in memory."""
import os
import re
from io import BytesIO
from pathlib import Path

from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from PIL import Image, ImageOps


def legacy_key(text):
    key = text.encode("utf-16-le")
    if len(key) != 16:
        raise ValueError("Enter the complete legacy key (eight UTF-16 code units).")
    return key


def configured_key(source_folder=None):
    text = os.environ.get("KYC_IMAGE_KEY")
    if text:
        return legacy_key(text)
    app_folder = Path(__file__).resolve().parent
    folders = [Path(source_folder)] if source_folder is not None else []
    folders.extend((app_folder, app_folder.parent, Path.home() / "Downloads"))
    for folder in folders:
        prompt = folder / "MASTER_DECRYPT_PROMPT.md"
        if prompt.is_file():
            contents = prompt.read_text(encoding="utf-8-sig")
            matches = re.findall(r"^\s*KEY STRING:[ \t]*(\S+)[ \t]*$", contents, re.MULTILINE)
            if len(matches) != 1:
                raise ValueError("MASTER_DECRYPT_PROMPT.md needs exactly one KEY STRING line.")
            return legacy_key(matches[0])
    return legacy_key("kyc1Goki")


def scan_bytes(path, key=None):
    """Return the encoded image payload, preserving the original bytes."""
    with Path(path).open("rb") as source:
        signature = source.read(8)
        if signature.startswith((b"\xff\xd8\xff", b"\x89PNG\r\n\x1a\n")):
            source.seek(0)
            return source.read()
        if key is None:
            raise ValueError("The file is not a readable JPG or PNG. For encrypted scans, place MASTER_DECRYPT_PROMPT.md in this image folder.")
        source.seek(0)
        ciphertext = source.read()
    if not ciphertext or len(ciphertext) % 16:
        raise ValueError("The file length does not match the legacy encrypted image format.")
    try:
        decryptor = Cipher(algorithms.AES(key), modes.CBC(key)).decryptor()
        padded = decryptor.update(ciphertext) + decryptor.finalize()
        unpadder = padding.PKCS7(128).unpadder()
        return unpadder.update(padded) + unpadder.finalize()
    except ValueError as error:
        raise ValueError("Decryption did not produce a readable JPG or PNG. Check the complete key and source file.") from error


def load_scan(path, key=None):
    with BytesIO(scan_bytes(path, key)) as source:
        try:
            return open_image(source)
        except OSError as error:
            raise ValueError("The result is not a readable JPG or PNG. Check the key and source file.") from error


def open_image(source):
    with Image.open(source, formats=("JPEG", "PNG")) as image:
        image.load()
        return ImageOps.exif_transpose(image)
