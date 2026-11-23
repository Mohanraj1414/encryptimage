"""Convert a folder of legacy encrypted JPG files into verified JPEGs."""
import argparse
from io import BytesIO
import os
from pathlib import Path
import sys
from tempfile import NamedTemporaryFile

from PIL import Image

from image_reader import configured_key, scan_bytes


def save_new_file(target, payload):
    with NamedTemporaryFile(dir=target.parent, delete=False) as temporary:
        staged = Path(temporary.name)
        try:
            temporary.write(payload)
            temporary.close()
            os.link(staged, target)
        finally:
            temporary.close()
            staged.unlink(missing_ok=True)



def save_jpeg(source, target, key):
    if target.exists():
        raise FileExistsError(target)
    payload = scan_bytes(source, key)
    with Image.open(BytesIO(payload), formats=("JPEG",)) as image:
        image.load()
    save_new_file(target, payload)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="Folder containing encrypted JPGs")
    parser.add_argument("output", type=Path, help="Separate folder for verified JPEGs")
    args = parser.parse_args()
    try:
        key = configured_key(args.source)
        if args.source.resolve() == args.output.resolve():
            parser.error("Choose a different output folder.")
        sources = sorted(p for p in args.source.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg"))
        if not sources:
            parser.error("No JPG or JPEG files found in the source folder.")
        args.output.mkdir(parents=True, exist_ok=True)
    except (OSError, ValueError) as error:
        parser.error(str(error))

    saved = skipped = failed = 0
    for source in sources:
        target = args.output / source.name
        try:
            save_jpeg(source, target, key)
            saved += 1
        except FileExistsError:
            print(f"Skipped {source.name}: output already exists.")
            skipped += 1
        except (OSError, ValueError, Image.DecompressionBombError) as error:
            print(f"Failed {source.name}: {error}", file=sys.stderr)
            failed += 1
    print(f"Saved {saved}, skipped {skipped}, failed {failed}.")
    return int(failed > 0)


if __name__ == "__main__":
    sys.exit(main())
