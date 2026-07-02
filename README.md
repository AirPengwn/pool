# Pool Maintenance — Claude Code Project

This folder is a self-contained pool-maintenance system for Claude Code. It replaces
the old chat workflow of uploading/downloading `Pool_Log.xlsx`. Here, the file lives
on your disk and Claude Code edits it in place.

## What's in here

| File | Purpose |
|------|---------|
| `POOL.md` | **Standing rules & context.** The source of truth. Claude reads this every session. Your specs, dosing rules, FC band table, open items. |
| `Pool_Log.xlsx` | Your living log (Log + Reference tabs). Edited in place. This is the durable memory. |
| `dose.py` | Chlorine dose calculator. All chemistry constants are in the `CONFIG` block at the top — tune there. |
| `append_log.py` | Appends row(s) to the Log tab in place, preserving formatting. |
| `README.md` | This file. |

---

## One-time setup

1. **Install Node.js** (Claude Code needs it): https://nodejs.org — LTS version. Verify with `node --version`.
2. **Install Claude Code:**
   ```
   npm install -g @anthropic-ai/claude-code
   ```
   (If you hit a permissions error on global install, the Claude Code docs cover the fix — or use a Node version manager.)
3. **Put this folder somewhere permanent**, e.g. `~/pool-maintenance/`.
4. **Make sure Python 3 + openpyxl are available** (the scripts use them):
   ```
   python3 -m pip install openpyxl
   ```
5. **Start Claude Code from inside the folder:**
   ```
   cd ~/pool-maintenance
   claude
   ```
   It will pick up `POOL.md` as context. The first time, tell it: "Read POOL.md and confirm you understand the workflow."

---

## Daily routine (the whole thing)

1. Run your noon titration (25 mL fill, drops × 0.2). Get FC and CC.
2. From inside the folder, run `claude` (or keep a session open) and paste the **daily prompt** below.
3. Claude calculates the dose, appends the row(s) in place, and tells you the dose + any flags.
4. Add the chlorine. Done. No upload, no download.

### Daily prompt (copy/paste, fill in the blanks)

```
Noon pool check for [DATE].
Readings (25mL x0.2): FC = [FC], CC = [CC].
[Full panel if taken: pH, TA, CH, CYA = ...]
Weather: [e.g. Sunny 82F dry / Mostly cloudy 69F / Rain all day].
Pump: [on 7AM / off time / any water adds].

Please: run dose.py for the recommendation, then append a TEST row and a DOSE row
(and PUMP/EVENT rows if relevant) to Pool_Log.xlsx in place. Give me the dose and
any flags. Round the dose to whatever makes sense for a low/high-loss day.
```

### What a good response from Claude Code looks like
- Runs `python3 dose.py --fc .. --cc .. --weather ..`
- Appends rows with `append_log.py`
- Confirms: dose in gallons, projected FC, and flags (CC creep, FC out of band, CYA recheck due)

---

## Other things you can just ask for

- **"Log pump off at 10:10 PM"** → it appends a PUMP row, no calc needed.
- **"Did a backwash and added ~300 gal well water"** → EVENT row + sequestrant reminder.
- **"CYA retest came back 140"** → it updates `cya_current` in `dose.py` CONFIG and `POOL.md`, and the band auto-adjusts.
- **"Show me FC trend for the last two weeks"** → it can read the log and chart/summarize.
- **"What's still open?"** → it reads the Open Items in POOL.md.

---

## Tuning the math

Everything chemistry lives in `dose.py` → `CONFIG`:
- `ppm_per_gallon` (3.5) — your pool's dose response.
- `cya_current` (160) — update after each CYA recheck; the FC band recalculates automatically.
- `fc_floor_frac` / `fc_target_top_frac` — the band as a fraction of CYA (7% / 10%).
- `weather_loss` — ppm/day loss by sky condition.
- `dose_round` (0.5) — half-gallon increments (quarters are impractical to judge in an opaque jug).

Change a number there and every future calc uses it. No logic editing needed.

---

## Why this fixes the old problems

- **No version drift:** one file on disk, edited in place. No "which upload is current."
- **Math is locked:** the dosing logic is a script with explicit, tunable constants — not re-derived each time.
- **Real memory:** `POOL.md` + the log file persist on your disk and are read every session. That's the durable "memory" that a chat workspace couldn't give you.
- **You're still the sensor:** you run the test and approve the dose. Claude handles calc + logging only.
