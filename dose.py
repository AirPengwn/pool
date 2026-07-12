#!/usr/bin/env python3
"""
dose.py — Chlorine dose calculator for John's pool.

Usage:
    python3 dose.py --fc 12.0 --cc 0.4 --weather cloudy
    python3 dose.py --fc 12.0 --cc 0.4 --weather cloudy --cya 160

All chemistry constants live in CONFIG below — tune them there, not in the logic.
Output is a plain-text recommendation you (or Claude Code) can read and log.
"""

import argparse

# ----------------------------------------------------------------------
# CONFIG — tune these as the pool/season changes. This is the only block
# you should normally need to edit.
# ----------------------------------------------------------------------
CONFIG = {
    "pool_gallons": 34400,

    # 1 gallon of 12.5% liquid chlorine raises FC by this many ppm in this pool
    "ppm_per_gallon": 3.5,

    # Current stabilizer level. Update after each CYA recheck.
    "cya_current": 144,

    # FC band as a fraction of CYA. floor ~7%, target top ~10% of CYA.
    "fc_floor_frac": 0.07,
    "fc_target_top_frac": 0.10,

    # Weather -> estimated FC loss over the rest of the day (ppm)
    "weather_loss": {
        "sunny":   4.0,
        "partly":  2.75,
        "cloudy":  1.75,
        "rain":    1.0,
    },

    # CC level at/above which to raise a flag
    "cc_flag": 0.5,

    # Round dose to nearest this many gallons. Half-gallon, not quarter —
    # a quarter gallon is impractical to judge by eye in an opaque jug.
    "dose_round": 0.5,

    # ---- Evening-dose 24h loss model (ppm) ----
    # Calibrated 2026-07-01 from fresh-reagent EVENING-dose data (6/25–7/1);
    # pre-6/24 old-reagent data excluded as unreliable. Keyed off the dose-day's
    # full-day sunshine hours (+ water temp). Predict next-noon FC as:
    #   next_noon_FC = today_noon_FC + dose_gal*ppm_per_gallon - L24
    # Validated within ~0.35 ppm on the recent clean hot/sunny days.
    "evening_l24": {
        "rain_sun_max": 3.0,     # sun hrs < this -> rainy/overcast day
        "partly_sun_max": 8.0,   # sun hrs < this -> partly
        "hot_temp": 80.0,        # water temp (F) at/above which sunny days lose more
        # Buckets NARROWED 2026-07-06. Analysis of all fresh-reagent daily losses (see POOL.md
        # "Loss-model calibration") found actual loss barely tracks sun -- slope ~0.08 ppm per
        # sun-hour vs the ~0.25 these buckets implied (3x too steep). At CYA ~150 the stabilizer
        # shields UV, so loss is a near-constant ~3 ppm/day background demand +/- large day-to-day
        # noise (sunny days ALONE ranged 1.2-5.1). The old spread over-predicted sun (all +errors)
        # and under-predicted rain (7/6's first -error). Compressed toward the ~3 mean:
        "l24_rain": 2.5,         # was 1.5 -- rain/low-sun days actually averaged ~2.5 (storm demand + dilution offset the lack of UV)
        "l24_partly": 3.0,
        "l24_sunny": 3.5,        # was 4.0
        "l24_hot_sunny": 3.5,    # was 4.0 (5.0 before 7/4). Hot days ran a touch higher (~3.85 avg); bump back toward ~3.8
                                 # if that holds as more fresh-reagent hot-day data accrues. Keep re-analyzing loss vs weather
                                 # AND vs reagent era (pre-6/24 old-reagent data stays excluded as unreliable).
    },
}


def band_for_cya(cya, cfg):
    floor = round(cya * cfg["fc_floor_frac"])
    top = round(cya * cfg["fc_target_top_frac"])
    # target band is floor..top; "aim" is the top
    return floor, top


def round_to(x, step):
    return round(x / step) * step


def evening_l24(sun_hrs, water_temp, cfg):
    """Estimated 24h FC loss (ppm) for an EVENING dose, from the dose-day's
    full-day sunshine hours and water temp. See CONFIG['evening_l24']."""
    e = cfg["evening_l24"]
    if sun_hrs is None:
        return e["l24_partly"]
    if sun_hrs < e["rain_sun_max"]:
        return e["l24_rain"]
    if sun_hrs < e["partly_sun_max"]:
        return e["l24_partly"]
    if water_temp is not None and water_temp >= e["hot_temp"]:
        return e["l24_hot_sunny"]
    return e["l24_sunny"]


def evening_plan(fc_now, sun_hrs, water_temp, cfg, dose_override=None):
    """Evening-dose recommendation that targets the top of the band at NEXT
    noon (not just 'now'), accounting for the 24h loss. Returns dict.

    Pass dose_override (gallons) to project a SPECIFIC dose instead of the
    model's own raw/rounded suggestion — use this once the final pour amount
    is decided (judgment calls, half-gallon jug rounding, etc. can make the
    actual pour differ from this function's own recommendation), so the
    logged 'Pred next-noon FC' matches what was actually poured."""
    _, aim = band_for_cya(cfg["cya_current"], cfg)
    l24 = evening_l24(sun_hrs, water_temp, cfg)
    ppm = cfg["ppm_per_gallon"]
    raw = max(0.0, (aim + l24 - fc_now) / ppm)
    dose = dose_override if dose_override is not None else round_to(raw, cfg["dose_round"])
    projected = fc_now + dose * ppm - l24
    return {"aim": aim, "l24": l24, "raw_gal": raw,
            "dose": dose, "projected_next_noon": projected}


def calc(fc, cc, weather, cya, cfg):
    floor, aim = band_for_cya(cya, cfg)
    loss = cfg["weather_loss"].get(weather, cfg["weather_loss"]["partly"])

    gap = aim - fc
    raw_gal = max(0.0, gap / cfg["ppm_per_gallon"])
    dose = round_to(raw_gal, cfg["dose_round"])
    projected_now = fc + dose * cfg["ppm_per_gallon"]
    projected_tomorrow = projected_now - loss

    lines = []
    lines.append(f"--- Dose calc (CYA {cya}) ---")
    lines.append(f"FC now: {fc}  |  CC: {cc}")
    lines.append(f"Band: floor {floor}, aim ~{aim}")
    if fc < floor:
        lines.append(f"** FC {fc} is BELOW floor {floor} — dose up promptly. **")
    elif fc > aim:
        lines.append(f"FC {fc} already at/above aim {aim} — no dose needed today.")
    lines.append(f"Weather '{weather}' -> est. loss ~{loss} ppm rest of day")
    lines.append(f"Gap to aim: {gap:+.1f} ppm  ->  raw {raw_gal:.2f} gal")
    lines.append(f"RECOMMENDED DOSE: {dose:.2f} gallon(s)")
    lines.append(f"  -> FC after dose ~{projected_now:.1f}")
    lines.append(f"  -> FC ~{projected_tomorrow:.1f} at next noon (est.)")
    if projected_tomorrow < floor:
        lines.append(f"  NOTE: projected to dip below floor {floor} by tomorrow — that's expected on high-loss days; retest at noon.")
    if cc >= cfg["cc_flag"]:
        lines.append(f"** CC {cc} >= {cfg['cc_flag']} — watch for persistent combined chlorine. Higher FC should clear it; flag if it holds. **")
    return "\n".join(lines)


def main():
    p = argparse.ArgumentParser(description="Pool chlorine dose calculator")
    p.add_argument("--fc", type=float, required=True, help="Free chlorine ppm (25mL, drops x0.2)")
    p.add_argument("--cc", type=float, default=0.0, help="Combined chlorine ppm")
    p.add_argument("--weather", default="partly",
                   choices=["sunny", "partly", "cloudy", "rain"],
                   help="Sky condition for loss estimate")
    p.add_argument("--cya", type=int, default=CONFIG["cya_current"],
                   help="Override current CYA (else uses CONFIG)")
    p.add_argument("--evening", action="store_true",
                   help="Evening-dose mode: print the calibrated L24 projection "
                        "instead of the daytime calc")
    p.add_argument("--sun", type=float, help="[--evening] Sun hours today (for L24 bucket)")
    p.add_argument("--water_temp", type=float, help="[--evening] Water temp F (for L24 bucket)")
    p.add_argument("--dose", type=float,
                   help="[--evening] Project this specific dose (gal) instead of the "
                        "model's own suggestion — use once the final pour amount is decided")
    args = p.parse_args()

    if args.evening:
        plan = evening_plan(args.fc, args.sun, args.water_temp, CONFIG,
                             dose_override=args.dose)
        print(f"--- Evening-dose plan (CYA {args.cya}) ---")
        print(f"FC now: {args.fc}")
        print(f"L24 (24h loss est.): {plan['l24']} ppm  [sun {args.sun}h, water {args.water_temp}F]")
        if args.dose is not None:
            print(f"Dose (override): {plan['dose']:.2f} gal")
        else:
            print(f"Recommended dose: {plan['dose']:.2f} gal (raw {plan['raw_gal']:.2f})")
        print(f"Projected NEXT-NOON FC: ~{plan['projected_next_noon']:.1f}"
              "  -> log as 'Pred next-noon FC' on tonight's DOSE row")
    else:
        print(calc(args.fc, args.cc, args.weather, args.cya, CONFIG))


if __name__ == "__main__":
    main()
