#!/usr/bin/env python3
"""
file_photos.py — rename newly-added camera photos to the dated convention.

Usage:
    python file_photos.py 2026-07-01
    python file_photos.py                # date defaults to today
    python file_photos.py 2026-07-01 --yes   # override the bulk-rename guard

Renames every image in photos/ that is NOT already named <date>_N.<ext>
(e.g. camera originals like IMG_8859.jpeg) to <date>_1.jpg, <date>_2.jpg, ...
in filename order, continuing the numbering if some already exist for that date.
Prints the old -> new mapping. Runs the same way every day, so one allow-rule
covers it.

Safety (added 2026-07-18 after an incident): the "already dated" guard must match
EVERY image extension, not just .jpg -- it previously only matched .jpg, so the
season's `.jpeg` photos weren't recognised as already-dated and got re-dated en
masse. There is also a MAX_AUTO ceiling: a normal day is a handful of photos, so
a huge batch means something is wrong. It aborts with a preview rather than
renaming, unless you pass --yes.
"""

import datetime
import os
import re
import sys

PHOTODIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "photos")
IMG_EXT = r"(?:jpe?g|png|heic|webp)"
MAX_AUTO = 25  # refuse silent bulk renames above this; see docstring


def main():
    argv = sys.argv[1:]
    force = "--yes" in argv
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        date = datetime.date.today().isoformat()
    elif len(args) == 1:
        date = args[0]
    else:
        sys.exit("usage: python file_photos.py [YYYY-MM-DD] [--yes]  (date defaults to today)")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date):
        sys.exit("date must be YYYY-MM-DD")

    files = sorted(os.listdir(PHOTODIR))
    already = re.compile(rf"^{re.escape(date)}_(\d+)\.{IMG_EXT}$", re.I)
    dated_any = re.compile(rf"^\d{{4}}-\d{{2}}-\d{{2}}_\d+\.{IMG_EXT}$", re.I)
    is_img = re.compile(rf"\.{IMG_EXT}$", re.I)

    existing = [int(m.group(1)) for f in files if (m := already.match(f))]
    n = max(existing) + 1 if existing else 1

    todo = [f for f in files if is_img.search(f) and not dated_any.match(f)]

    if not todo:
        print(f"no new photos to rename in {PHOTODIR}")
        return

    if len(todo) > MAX_AUTO and not force:
        preview = "\n  ".join(todo[:10]) + ("\n  ..." if len(todo) > 10 else "")
        sys.exit(
            f"REFUSING to rename {len(todo)} files at once (limit {MAX_AUTO}).\n"
            f"That usually means already-dated photos aren't being recognised -- "
            f"check the extensions before proceeding.\n"
            f"Review this list, then re-run with --yes if it really is intended:\n  {preview}"
        )

    renamed = []
    for f in todo:
        new = f"{date}_{n}.jpg"
        os.rename(os.path.join(PHOTODIR, f), os.path.join(PHOTODIR, new))
        renamed.append((f, new))
        n += 1

    for old, new in renamed:
        print(f"{old} -> {new}")
    print(f"renamed {len(renamed)} file(s)")


if __name__ == "__main__":
    main()
