#!/usr/bin/env python3
"""One-off: transcode/resize Night 2 concert media (icloud/night2 raw -> concert2 web-ready).
Same treatment as Night 1 (see prep_concert.py / POOL.md):
  Photos -> resized JPEG (exif-corrected, max 1600px, q82).
  Videos -> H.264 720p, CRF 23... 26, +faststart (web streaming), AAC 128k.
Originals were already H.264 720p; re-encode is for faststart + size."""
import os, subprocess, sys
from PIL import Image, ImageOps

HERE = r"C:\dev\pool"
SRC = os.path.join(HERE, "icloud", "night2")   # raw originals (gitignored)
DST = os.path.join(HERE, "concert2")           # web-ready, committed
os.makedirs(DST, exist_ok=True)

FFMPEG = None
for cand in ["ffmpeg",
             r"C:\Users\airpe\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.1.2-full_build\bin\ffmpeg.exe"]:
    try:
        subprocess.run([cand, "-version"], capture_output=True, check=True)
        FFMPEG = cand; break
    except Exception:
        continue
if not FFMPEG:
    sys.exit("ffmpeg not found")
print("ffmpeg:", FFMPEG, flush=True)

MAXDIM, Q = 1600, 82
files = sorted(os.listdir(SRC))
photos = [f for f in files if f.lower().endswith((".jpeg", ".jpg", ".png"))]
videos = [f for f in files if f.lower().endswith((".mp4", ".mov"))]

for f in photos:
    base = os.path.splitext(f)[0]
    dst = os.path.join(DST, base + ".jpg")
    with Image.open(os.path.join(SRC, f)) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        w, h = im.size
        if max(w, h) > MAXDIM:
            s = MAXDIM / max(w, h)
            im = im.resize((round(w * s), round(h * s)), Image.LANCZOS)
        im.save(dst, "JPEG", quality=Q, optimize=True)
    print(f"photo {base}.jpg  {os.path.getsize(dst)//1024} KB", flush=True)

for f in videos:
    base = os.path.splitext(f)[0]
    dst = os.path.join(DST, base + ".mp4")
    print(f"transcoding {f} ...", flush=True)
    cmd = [FFMPEG, "-i", os.path.join(SRC, f), "-c:v", "libx264", "-crf", "26",
           "-preset", "medium", "-movflags", "+faststart", "-c:a", "aac",
           "-b:a", "128k", "-pix_fmt", "yuv420p", "-y", dst]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print("FFMPEG ERROR", f, "\n", r.stderr[-1200:], flush=True); continue
    print(f"video {base}.mp4  {os.path.getsize(dst)//1024//1024} MB", flush=True)

tot = sum(os.path.getsize(os.path.join(DST, x)) for x in os.listdir(DST)) // 1024 // 1024
print(f"DONE. {len(photos)} photos + {len(videos)} videos, concert2/ total {tot} MB", flush=True)
