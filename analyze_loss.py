#!/usr/bin/env python3
"""analyze_loss.py -- rebuild the daily chlorine-loss dataset on the CORRECTED
volume and test which model actually predicts loss best.

Chlorine is tracked as MASS (ppm*gal), not concentration, so rain dilution and
level changes are handled exactly instead of being smeared into "loss":

    mass       = volume_gal * FC_ppm
    volume_gal = FULL_GAL + (level_cm - FULL_LEVEL_CM) * GAL_PER_CM
    loss_mass  = mass_prev + dose_gal*CL_MASS_PER_GAL - mass_next
    loss_ppm   = loss_mass / avg_volume

Only 1-day-apart noon pairs with trustworthy FC on both ends are used
(pre-2026-06-24 old reagents and the 7/11-7/16 degraded-reagent era are excluded
automatically -- the latter has no FC logged at all).

Run: python analyze_loss.py
"""
import os
from datetime import date, timedelta
import openpyxl

FULL_GAL       = 31400.0    # measured geometry, pool full
FULL_LEVEL_CM  = 15.81      # skimmer-ruler reading at "full" (36 in shallow)
GAL_PER_CM     = 196.3      # 800 sq ft surface, vertical walls in the top 3 ft
CL_MASS_PER_GAL = 3.5 * 34400.0   # ppm*gal delivered by 1 gal of 12.5% (volume-independent)

LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Pool_Log.xlsx")
C = dict(date=1, type=3, fc=5, sun=15, rain=16, temp=18, level=19, dose=13, load=23)
GOOD_FROM = date(2026, 6, 24)     # fresh FAS-DPD onwards


def as_date(v):
    if hasattr(v, "year"):
        return date(v.year, v.month, v.day)
    if isinstance(v, str) and "-" in v:
        y, m, d = v.split("-")[:3]
        return date(int(y), int(m), int(d))
    return None


def num(v):
    return v if isinstance(v, (int, float)) else None


def fit(xs, ys):
    """least squares y = a + b x; returns (b, a, r)"""
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    b = sxy / sxx if sxx else 0.0
    return b, my - b * mx, (sxy / (sxx * syy) ** 0.5 if sxx and syy else 0.0)


def mae(errs):
    return sum(abs(e) for e in errs) / len(errs)


def main():
    ws = openpyxl.load_workbook(LOG, data_only=True)["Log"]
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r[0]]

    tests, doses, level_seen = {}, {}, {}
    for r in rows:
        d = as_date(r[C["date"] - 1])
        if not d:
            continue
        typ = str(r[C["type"] - 1] or "")
        if typ.startswith("TEST") and d not in tests:      # first (noon) test wins
            tests[d] = dict(fc=num(r[C["fc"] - 1]), sun=num(r[C["sun"] - 1]),
                            rain=num(r[C["rain"] - 1]), temp=num(r[C["temp"] - 1]),
                            level=num(r[C["level"] - 1]), load=r[C["load"] - 1])
        if typ.startswith("DOSE"):
            doses[d] = doses.get(d, 0.0) + (num(r[C["dose"] - 1]) or 0.0)

    # carry the last known level forward when a day lacks one
    last = None
    for d in sorted(tests):
        if tests[d]["level"] is not None:
            last = tests[d]["level"]
            level_seen[d] = True
        else:
            tests[d]["level"] = last
            level_seen[d] = False

    vol = lambda cm: FULL_GAL + (cm - FULL_LEVEL_CM) * GAL_PER_CM

    pairs = []
    for d in sorted(tests):
        n = d + timedelta(days=1)
        a, b = tests.get(d), tests.get(n)
        if not b or d < GOOD_FROM:
            continue
        if a["fc"] is None or b["fc"] is None or a["level"] is None or b["level"] is None:
            continue
        v0, v1 = vol(a["level"]), vol(b["level"])
        dose = doses.get(d, 0.0)
        loss_mass = v0 * a["fc"] + dose * CL_MASS_PER_GAL - v1 * b["fc"]
        vavg = (v0 + v1) / 2
        pairs.append(dict(d=d, fc0=a["fc"], fc1=b["fc"], dose=dose,
                          v0=v0, v1=v1, added=v1 - v0,
                          loss=loss_mass / vavg, fcavg=(a["fc"] + b["fc"]) / 2,
                          sun=a["sun"], temp=a["temp"], rain=a["rain"],
                          # Load is recorded on the ENDING test -- it describes what
                          # the pool experienced during the window just measured.
                          load=b.get("load"),
                          exact=level_seen.get(d) and level_seen.get(n)))

    print(f"{'day':<12}{'FC0':>6}{'FC1':>6}{'dose':>6}{'water+':>8}{'loss':>7}"
          f"{'sun':>6}{'temp':>6}  lvl")
    for p in pairs:
        s = f"{p['sun']:.1f}" if p["sun"] is not None else "  - "
        t = f"{p['temp']:.0f}" if p["temp"] is not None else " - "
        print(f"{p['d'].isoformat():<12}{p['fc0']:>6.1f}{p['fc1']:>6.1f}{p['dose']:>6.1f}"
              f"{p['added']:>+8.0f}{p['loss']:>7.2f}{s:>6}{t:>6}  {'exact' if p['exact'] else 'carried'}")

    losses = [p["loss"] for p in pairs]
    n = len(losses)
    mean_loss = sum(losses) / n
    sd = (sum((l - mean_loss) ** 2 for l in losses) / (n - 1)) ** 0.5
    print(f"\nn = {n} pairs   mean loss {mean_loss:.2f} ppm/24h   stdev {sd:.2f}   "
          f"range {min(losses):.2f}-{max(losses):.2f}")

    # --- what drives it? ---
    print("\n-- correlations vs loss --")
    for name, key in [("sun hours", "sun"), ("water temp", "temp"), ("FC carried", "fcavg")]:
        sub = [(p[key], p["loss"]) for p in pairs if p[key] is not None]
        if len(sub) < 3:
            continue
        b, a, r = fit([x for x, _ in sub], [y for _, y in sub])
        print(f"  {name:<12} n={len(sub):<3} slope {b:+.3f}   r {r:+.2f}")

    # --- candidate models ---
    print("\n-- model comparison (MAE, ppm) --")
    def bucket(p):                      # HISTORICAL sun-buckets (retired 2026-07-19)
        s, t = p["sun"], p["temp"]
        if s is None:
            return 3.0
        if s < 3:
            return 2.5
        if s < 8:
            return 3.0
        return 3.5
    import dose as _D                   # score the LIVE constant, not a stale copy
    _live = _D.CONFIG["evening_l24"]["l24_sunny"]
    cands = {"old sun-buckets 2.5/3/3.5": [bucket(p) for p in pairs],
             f"LIVE dose.py flat {_live}": [_live] * n,
             "flat mean": [mean_loss] * n}
    sub = [p for p in pairs if p["fcavg"] is not None]
    k, d0, _ = fit([p["fcavg"] for p in sub], [p["loss"] for p in sub])
    cands[f"first-order {k:+.3f}*FC{d0:+.2f}"] = [k * p["fcavg"] + d0 for p in pairs]
    kk = sum(p["loss"] for p in sub) / sum(p["fcavg"] for p in sub)
    cands[f"proportional {kk:.4f}*FC"] = [kk * p["fcavg"] for p in pairs]

    for name, pred in sorted(cands.items(), key=lambda kv: mae([l - q for l, q in zip(losses, kv[1])])):
        errs = [l - q for l, q in zip(losses, pred)]
        bias = sum(errs) / n
        print(f"  {name:<28} MAE {mae(errs):.2f}   bias {bias:+.2f}")

    # --- loss by load class: the driver that actually matters ---
    byload = {}
    for p in pairs:
        # Load is "canonical[+canonical]; optional human qualifier" -- aggregate on
        # the canonical part only, so detail in the qualifier does not fragment classes.
        cls = str(p["load"] or "(unrecorded)").split(";")[0].strip()
        byload.setdefault(cls, []).append(p["loss"])
    print("\n-- loss by LOAD class (the real driver) --")
    for k, v in sorted(byload.items(), key=lambda kv: -sum(kv[1]) / len(kv[1])):
        m = sum(v) / len(v)
        sd = (sum((x - m) ** 2 for x in v) / (len(v) - 1)) ** 0.5 if len(v) > 1 else float("nan")
        print(f"  {k:<14} n={len(v):<3} mean {m:.2f}" + (f"   sd {sd:.2f}" if len(v) > 1 else ""))
    print()
    print("NOTE: load class was TESTED as a predictor and FAILED -- see POOL.md,")
    print("2026-08-12/13. Classes overlap heavily (quiet mean 2.29 vs storm 2.06,")
    print("both sd ~1.3), and leave-one-out CV showed every predictor -- alone AND")
    print("combined -- is WORSE than a flat mean (analyze_multivar.py). These class")
    print("means are DESCRIPTIVE ONLY; do NOT fit L24 per class.")


if __name__ == "__main__":
    main()
