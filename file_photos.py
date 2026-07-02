#!/usr/bin/env python3
"""
file_photos.py — rename newly-added camera photos to the dated convention.

Usage:
    python file_photos.py 2026-07-01

Renames every image in photos/ that is NOT already named <date>_N.jpg
(e.g. camera originals like IMG_8859.jpeg) to <date>_1.jpg, <date>_2.jpg, ...
in filename order, continuing the numbering if some already exist for that date.
Prints the old -> new mapping. Runs the same way every day, so one allow-rule
covers it.
"""

import datetime
import os
import re
import sys

PHOTODIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "photos")


def main():
    if len(sys.argv) == 1:
        date = datetime.date.today().isoformat()  # default: today (keeps the command identical daily)
    elif len(sys.argv) == 2:
        date = sys.argv[1]
    else:
        sys.exit("usage: python file_photos.py [YYYY-MM-DD]  (defaults to today)")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date):
        sys.exit("date must be YYYY-MM-DD")

    files = sorted(os.listdir(PHOTODIR))
    already = re.compile(rf"^{re.escape(date)}_(\d+)\.jpg$", re.I)
    dated_any = re.compile(r"^\d{4}-\d{2}-\d{2}_\d+\.jpg$", re.I)
    is_img = re.compile(r"\.(jpe?g|png|heic|webp)$", re.I)

    existing = [int(m.group(1)) for f in files if (m := already.match(f))]
    n = max(existing) + 1 if existing else 1

    renamed = []
    for f in files:
        if dated_any.match(f) or not is_img.search(f):
            continue
        new = f"{date}_{n}.jpg"
        os.rename(os.path.join(PHOTODIR, f), os.path.join(PHOTODIR, new))
        renamed.append((f, new))
        n += 1

    if not renamed:
        print(f"no new photos to rename in {PHOTODIR}")
    for old, new in renamed:
        print(f"{old} -> {new}")


if __name__ == "__main__":
    main()
