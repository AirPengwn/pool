# Pool Assistant — Session Handoff
_Written 2026-07-16 to carry full context into a fresh session. **Read this AND `POOL.md` before doing anything.** This file = the current live state + active issues; `POOL.md` = the deep source-of-truth history._

---

## 1. What this is
Ongoing pool-maintenance assistant for **John** (Guilford, CT). **34,400-gal inground vinyl-liner** pool, 2026 season. You help him keep chlorine dialed in, log every reading into `Pool_Log.xlsx`, analyze/photograph the pool, and maintain a public dashboard website. `POOL.md` is the source of truth (specs, chemistry approach, website + deploy docs, full changelog).

## 2. How to work — standing rules (non-negotiable)
- **Honest & math-grounded. NEVER round or invent a reading.** If a number looks off, say so and reason it out (this has caught several real errors this season).
- **Twice-daily cadence:** John reports a **noon TEST** and an **evening DOSE**. On EVERY report-in: build the row(s) as JSON → `python append_log.py --json <file>` → `python format_log.py` → `python build_site.py` → `git add Pool_Log.xlsx` (+ any new `photos/`) → commit → **push (durable standing authorization — never ask first).**
- **Photos:** John drops camera originals (`IMG_*.JPEG`) into `photos/`. Rename to `YYYY-MM-DD_N.jpeg` (chronological), put the filenames in that day's log-row Photo column (`;`-separated), and actually LOOK at them (clarity, level, stains, anything off). `.MOV`/`.mov` gitignored; transcode phone video to small `.mp4` before committing.
- **Dosing:** `dose.py` gives a CYA-scaled band + an evening 24h-loss ("L24") projection. Recommend an amount, John pours, you log the actual DOSE with `pred_fc` = next-noon projection.
- **Tone:** warm, concise, decisive — give a recommendation, not a survey.

## 3. Current chemistry snapshot (2026-07-16 noon)
| | Value | Notes |
|---|---|---|
| FC | **~14.5 (corrected est.)** | at aim; see Issue #1 — reagent unreliable |
| CC | 0.4 | low; slight bump from wildfire-smoke ash |
| pH | 7.7 | upper-ideal, fine |
| TA | 150 | high due to CYA contribution (~true 100); fine |
| CH | ~200 | fine (barely matters on vinyl) |
| CYA | **~144** | re-anchored 7/12; **band = floor 10 / aim 14** |
| Water temp | ~80°F | |
| Level | **12.2 cm, dropping ~0.4/day** | dry stretch; season range 11.65–15.5; **hose top-up may be due if it keeps falling** |

Water crystal clear all season, no algae. **Season chlorine total: 28.5 gal.** Last dose: **1.5 gal, 7/15 8pm.** No dose 7/16 (at aim).

## 4. 🔴 ACTIVE ISSUE #1 — the reagent saga (MOST IMPORTANT — read carefully)
The FAS-DPD **FC titrant reagent is bad: it over-reads FC by ~1.5×.** BOTH the "new" bottle (opened 7/11) and the season-start "very old" bottle over-read the same amount — likely aged FAS + a high-CYA endpoint-fade effect. **So the FC *number* is untrustworthy** (a reading of ~21 / 105 drops is really ~14). pH/TA/CH/CYA/CC/temp/level are all still fine (cross-validated by test strips).

**How FC is being handled:**
- **The FC column is left BLANK** on TEST rows since 7/11 (don't log the over-read value; don't invent one). Put the raw drop count + corrected estimate in the **notes**.
- **We track FC by the drop-count TREND, not the absolute.** Bias is ~constant, so day-to-day *change* is real. **Rule of thumb: true FC ≈ (drops × 0.2) ÷ 1.5.** **Dose when the count falls to ~75** (≈ true FC 10 = floor).
- The trend has worked cleanly: 106→96→86→75 drops (~10/day), dosed 1.5 gal 7/15, back up to 109 (~14.5).

**➡️ THE FIX IS IMMINENT — a fresh TF-100 kit is arriving. USPS ETA Friday 7/17 by 9pm.**
**FRIDAY ACTION:** when it lands, have John run a real FC. **Recalibrate the ~1.5× over-read factor** against the true number, note in the log where the estimate era ends and measured FC resumes, then **go back to logging the real FC in the column + precise dosing.** Saga closed.
Endpoint tip for the new kit (high CYA): take the FIRST clear + hold ~15s; a faint re-pink is the CYA effect, not more chlorine. CC precision doesn't matter (it's been low all season).

## 5. 🟡 ACTIVE ISSUE #2 — deploy via manual gh-pages push
A **GitHub Actions incident ("Delays starting Actions runs")** started **2026-07-09**; the auto-deploy Action (`.github/workflows/deploy.yml` — builds `_site/`, pushes to `gh-pages`) has been stuck/failing. **Workaround in use:** after `build_site.py`, manually publish `_site/` to `gh-pages` from a temp dir (do NOT use the session scratchpad path from the old session — make your own):
```bash
git clone --depth 1 --branch gh-pages --single-branch https://github.com/AirPengwn/pool.git /tmp/ghp
find /tmp/ghp -mindepth 1 -maxdepth 1 ! -name '.git' -exec rm -rf {} +
cp -r /c/dev/pool/_site/. /tmp/ghp/ && touch /tmp/ghp/.nojekyll
git -C /tmp/ghp add -A
git -C /tmp/ghp -c user.name="AirPenguin23" -c user.email="johnlorinevans@gmail.com" commit -m "Deploy YYYY-MM-DD (manual)"
git -C /tmp/ghp push origin gh-pages
```
**➡️ FIRST, CHECK whether the incident has cleared** (githubstatus.com, or `gh run list`). If the auto-deploy works again, DROP the manual push — just `git push origin main` and the Action deploys. Live site: **https://airpengwn.github.io/pool/**. (GitHub user = **AirPengwn**, git author = **AirPenguin23** — the odd spelling/caps are real, not typos.)

## 6. The website (generated by build_site.py → `_site/`, gitignored, rebuilt each run)
- **Dashboard** (`index.html`): FC gauge, chemistry chart + bands, ~19 charts in 4 collapsible sections, prediction-accuracy, cost calc, light/dark toggle, "sun-warmed logbook" style.
- **Data** (`log.html`), **Photos** (`photos.html`, day-by-day incl. videos).
- **Concerts** — small "concerts" link at the very bottom of the dashboard → **two-night Noah Kahan @ Fenway** gallery: Night 1 (Jul 8, `concert.html`) + Night 2 (Jul 11, `concert2.html`), pill sub-nav, per-night **setlists** (acts/stages + song notes, sourced from setlist.fm), and an **orientation-aware video layout** (portrait clips solo+centered, landscapes 2-up — uses a pure-python mp4 dimension reader in build_site.py). Web-ready media committed in `concert/` + `concert2/`; raw originals gitignored under `icloud/`. Videos transcoded H.264 720p CRF26 +faststart (`prep_concert*.py`). To add a night: raw → gitignored stash, transcode into a committed `concertN/`, add a `CONCERTS` entry in build_site.py.

## 7. Key files & environment
- `POOL.md` (read it) · `Pool_Log.xlsx` (22-col log) · `append_log.py` + `format_log.py` (append/format; run format after every append) · `dose.py` (band + L24, `cya_current=144`) · `weather.py` (Open-Meteo; **reports clear-sky sun — CANNOT see wildfire smoke**, so discount its sun number on smoky days) · `build_site.py` · `prep_concert*.py`.
- Windows; use **`python`** (not python3). `gh` CLI installed (`/c/Program Files/GitHub CLI/gh.exe`; PATH not always refreshed in new shells). `ffmpeg`/`ffprobe` available (winget Gyan.FFmpeg). Pool level ruler mounted on the skimmer, read in cm. Pump on a 7am–10pm timer; ~500 gal/inch; pump-to-waste ~8 min/inch.

## 8. Immediate next steps
1. **Fri 7/17:** TF-100 arrives → real FC → **recalibrate the over-read factor, resume normal FC logging + precise dosing.**
2. Keep the **twice-daily cadence** (noon TEST + evening DOSE; log + deploy each).
3. **Watch the water level** (12.2, ~0.4/day drop) — suggest a hose top-up if it keeps falling; log fills as `EVENT - Fill` (hose min × ~3.1 GPM).
4. **Wildfire smoke** is reducing UV right now → FC loss cut ~in half (7/16); expect the drop-count decline to stay flat a day or two, then resume once it clears.
5. Standing task: keep analyzing FC-loss vs weather on **fresh-reagent** data only (storm/bather/dog/smoke days are confounded — see POOL.md "Loss-model calibration").
