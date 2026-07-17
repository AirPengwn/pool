#!/usr/bin/env python3
"""backtest_volume.py -- does the corrected pool volume fit the season's data?

Re-runs every good-reagent evening dose -> next-noon prediction pair with the
corrected dose response (dose.py's current ppm_per_gallon) against the OLD 3.5
the predictions were logged with, and reports:

  implied_R = 3.5 + (actual - predicted)/gal
      the dose response the chlorine data ALONE "wanted" for each pair (loss-model
      noise averages out over many days, so the mean estimates the true response --
      independent of the pool-geometry measurement).
  mean signed error & MAE, old vs corrected R
      whether applying the corrected response removes the season-long bias.

Confounded dose-days (next-day rain / a deliberately small 1-gal dose / storm /
dog+bather load) are flagged and excluded from the "clean" aggregate -- those
misses are real demand events, not the dose response.

Run: python backtest_volume.py
"""
import os
from datetime import date, timedelta
import openpyxl
import dose as D

OLD_R = 3.5                               # what the logged pred_fc values used
NEW_R = D.CONFIG["ppm_per_gallon"]        # corrected response (live from dose.py)
LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Pool_Log.xlsx")
C = dict(date=1, type=3, fc=5, pred_fc=8, dose=13)

# Confounded dose-days per the log notes -> excluded from the "clean" aggregate.
CONFOUND = {
    date(2026, 6, 25): "next-day heavy rain",
    date(2026, 6, 28): "only 1 gal dosed (low FC carried)",
    date(2026, 7, 6):  "storm day",
    date(2026, 7, 9):  "dog + bather load next day",
}


def as_date(v):
    if hasattr(v, "year"):
        return date(v.year, v.month, v.day)
    if isinstance(v, str) and "-" in v:
        y, m, dd = v.split("-")[:3]
        return date(int(y), int(m), int(dd))
    return v


def main():
    ws = openpyxl.load_workbook(LOG, data_only=True)["Log"]
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r[0]]

    noon = {}                             # date -> first (noon) TEST fc
    for r in rows:
        if str(r[C["type"] - 1] or "").startswith("TEST"):
            dt = as_date(r[C["date"] - 1])
            if dt not in noon:
                fc = r[C["fc"] - 1]
                noon[dt] = fc if isinstance(fc, (int, float)) else None

    pairs = []
    for r in rows:
        if not str(r[C["type"] - 1] or "").startswith("DOSE"):
            continue
        dt, pred, gal = as_date(r[C["date"] - 1]), r[C["pred_fc"] - 1], r[C["dose"] - 1]
        if pred is None or not isinstance(gal, (int, float)) or gal <= 0:
            continue
        F1 = noon.get(dt + timedelta(days=1))
        if not isinstance(F1, (int, float)):
            continue                      # no next-noon reading (e.g. reagent-blank era)
        err = round(F1 - pred, 2)
        pairs.append(dict(dd=dt, gal=gal, pred=pred, act=F1, old=err,
                          new=round(err - gal * (NEW_R - OLD_R), 2),
                          impR=round(OLD_R + err / gal, 2),
                          conf=CONFOUND.get(dt)))

    def stats(ps):
        o = [p["old"] for p in ps]; n = [p["new"] for p in ps]
        imp = sorted(p["impR"] for p in ps)
        mean = lambda x: sum(x) / len(x)
        mae = lambda x: sum(abs(v) for v in x) / len(x)
        return (len(ps), mean(o), mean(n), mae(o), mae(n), mean(imp), imp[len(imp) // 2])

    print(f"{'dose day':<12}{'gal':>4}{'pred':>6}{'actual':>7}{'old_err':>8}"
          f"{'new_err':>8}{'impliedR':>9}  flag")
    for p in pairs:
        print(f"{p['dd'].isoformat():<12}{p['gal']:>4}{p['pred']:>6}{p['act']:>7}"
              f"{p['old']:>+8.1f}{p['new']:>+8.1f}{p['impR']:>9.2f}  {p['conf'] or ''}")

    print(f"\nOLD R = {OLD_R} (as logged)   NEW R = {NEW_R} (corrected volume)")
    for name, ps in [("ALL pairs", pairs),
                     ("CLEAN pairs", [p for p in pairs if not p["conf"]])]:
        k, mo, mn, maeo, maen, impm, impmed = stats(ps)
        print(f"\n{name}  (n={k}):")
        print(f"  mean signed error: {mo:+.2f} -> {mn:+.2f} ppm   (0 = unbiased)")
        print(f"  mean |error| (MAE): {maeo:.2f} -> {maen:.2f} ppm")
        print(f"  implied response the data wanted: mean {impm:.2f}  median {impmed:.2f} ppm/gal")


if __name__ == "__main__":
    main()
