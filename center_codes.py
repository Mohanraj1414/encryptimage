"""Restore printed Y letters only when independent OCR views agree."""

import re
from collections import Counter


CODE = re.compile(r"[A-Za-z0-9\s]+")
# OCR often renders the same printed stroke as a 1, an I, or a lowercase l.
# Use this equivalence for matching only. Write the original character back.
STROKES = str.maketrans({"I": "1", "l": "1"})


def is_center_code(value):
    return isinstance(value, str) and CODE.fullmatch(value) and 24 <= len("".join(value.split())) <= 40


def code_tokens(text):
    tokens = set()
    for line in text.splitlines():
        for block in re.findall(r"[A-Za-z0-9]+(?:[ \t]+[A-Za-z0-9]+)*", line):
            words = block.split()
            for start in range(len(words)):
                candidate = ""
                for word in words[start:]:
                    if word in {"Y", "y"}:
                        break
                    candidate += word
                    if len(candidate) > 43:
                        break
                    if len(candidate) >= 25:
                        tokens.add(candidate)
    return tokens


def insert_printed_y(source, reading):
    if not 1 <= len(reading) - len(source) <= 3:
        return None
    result = []
    position = 0
    for glyph in reading:
        current = source[position:position + 1]
        if current and (
            current == glyph
            or (current not in "Yy" and glyph not in "Yy"
                and current.translate(STROKES).casefold() == glyph.translate(STROKES).casefold())
        ):
            result.append(current)
            position += 1
        elif glyph in "Yy":
            result.append(glyph)
        else:
            return None
    if position == len(source):
        return "".join(result)
    return None


def restore_center_codes(records, readings):
    tokens = [code_tokens(reading) for reading in readings]
    corrected = []
    for record in records:
        value = record.get("Center Code")
        if is_center_code(value):
            code = "".join(value.split())
            candidates = Counter(
                correction
                for view in tokens
                for correction in {insert_printed_y(code, token) for token in view}
                if correction is not None
            )
            if len(candidates) == 1:
                correction, count = next(iter(candidates.items()))
                if count >= 2:
                    record = {**record, "Center Code": correction}
        corrected.append(record)
    return corrected
