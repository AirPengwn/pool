#!/usr/bin/env python3
"""analyze_multivar.py -- can the weak predictors COMBINED beat a flat mean?

Each predictor tested alone has failed (sun, water temp, FC carried, load class,
visible debris). This asks whether a multivariate fit does better.

The catch: adding predictors ALWAYS improves in-sample fit, even for pure noise
-- with k predictors and n points you can drive residuals to zero once k = n-1.
So in-sample R^2 is meaningless here. This uses LEAVE-ONE-OUT CROSS-VALIDATION:
refit the model n times, each time holding out one day, and score only the
held-out prediction. That is the honest measure of whether it would help on a
day we have not seen.

Baseline to beat: predicting the flat mean of the training days.

Run: python analyze_multivar.py
"""
import os
from datetime import date, timedelta
import openpyxl

FULL_GAL, FULL_LEVEL_CM, GAL_PER_CM = 31400.0, 15.81, 196.3
CL_PER_GAL = 125000.0
LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Pool_Log.xlsx")
C = dict(date=1, type=3, fc=5, dose=13, sun=15, rain=16, temp=18, level=19, load=23)
GOOD_FROM = date(2026, 6, 24)


def as_date(v):
    if hasattr(v, "year"):
        return date(v.year, v.month, v.day)
    if isinstance(v, str) and "-" in v:
        y, m, d = v.split("-")[:3]
        return date(int(y), int(m), int(d))


def num(v):
    return v if isinstance(v, (int, float)) else None


def solve(A, b):
    """Gaussian elimination with partial pivoting. Returns None if singular."""
    n = len(A)
    M = [row[:] + [b[i]] for i, row in enumerate(A)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(M[r][c]))
        if abs(M[p][c]) < 1e-12:
            return None
        M[c], M[p] = M[p], M[c]
        for r in range(n):
            if r == c:
                continue
            f = M[r][c] / M[c][c]
            for k in range(c, n + 1):
                M[r][k] -= f * M[c][k]
    return [M[i][n] / M[i][i] for i in range(n)]


def fit(X, y):
    """Least squares with intercept. X = list of feature rows."""
    n, k = len(X), len(X[0])
    D = [[1.0] + list(row) for row in X]
    A = [[sum(D[r][i] * D[r][j] for r in range(n)) for j in range(k + 1)] for i in range(k + 1)]
    b = [sum(D[r][i] * y[r] for r in range(n)) for i in range(k + 1)]
    return solve(A, b)


def predict(coef, row):
    return coef[0] + sum(c * v for c, v in zip(coef[1:], row))


def loocv(X, y):
    """Leave-one-out CV mean absolute error. None if any fold is singular."""
    n = len(X)
    errs = []
    for i in range(n):
        Xt = [X[j] for j in range(n) if j != i]
        yt = [y[j] for j in range(n) if j != i]
        coef = fit(Xt, yt)
        if coef is None:
            return None
        errs.append(abs(y[i] - predict(coef, X[i])))
    return sum(errs) / n


def loocv_mean(y):
    n = len(y)
    return sum(abs(y[i] - (sum(y) - y[i]) / (n - 1)) for i in range(n)) / n


def main():
    ws = openpyxl.load_workbook(LOG, data_only=True)["Log"]
    tests, doses = {}, {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r[0]:
            continue
        d, t = as_date(r[C["date"] - 1]), str(r[C["type"] - 1] or "")
        if t.startswith("TEST") and d not in tests:
            tests[d] = {k: num(r[C[k] - 1]) for k in ("fc", "sun", "rain", "temp", "level")}
        if t.startswith("DOSE"):
            doses[d] = doses.get(d, 0.0) + (num(r[C["dose"] - 1]) or 0.0)

    last = None
    for d in sorted(tests):
        if tests[d]["level"] is not None:
            last = tests[d]["level"]
        else:
            tests[d]["level"] = last

    vol = lambda cm: FULL_GAL + (cm - FULL_LEVEL_CM) * GAL_PER_CM
    rows = []
    for d in sorted(tests):
        n_ = d + timedelta(days=1)
        a, b = tests.get(d), tests.get(n_)
        if not b or d < GOOD_FROM:
            continue
        if None in (a["fc"], b["fc"], a["level"], b["level"], a["sun"], a["temp"]):
            continue
        v0, v1 = vol(a["level"]), vol(b["level"])
        dz = doses.get(d, 0.0)
        loss = (v0 * a["fc"] + dz * CL_PER_GAL - v1 * b["fc"]) / ((v0 + v1) / 2)
        rows.append(dict(d=d, loss=loss, sun=a["sun"], temp=a["temp"],
                         fc=(a["fc"] + b["fc"]) / 2, rain=a["rain"] or 0.0, dose=dz))

    y = [r["loss"] for r in rows]
    n = len(y)
    print(f"n = {n} day-pairs with a complete predictor set\n")

    base = loocv_mean(y)
    print(f"{'model':<34}{'LOOCV MAE':>10}{'vs flat':>10}")
    print(f"{'flat mean (baseline)':<34}{base:>10.2f}{'--':>10}")

    feats = [("sun", "sun hours"), ("temp", "water temp"), ("fc", "FC carried"),
             ("rain", "rain"), ("dose", "gal dosed")]
    for key, label in feats:
        X = [[r[key]] for r in rows]
        m = loocv(X, y)
        if m:
            print(f"{'single: ' + label:<34}{m:>10.2f}{m - base:>+10.2f}")

    for combo in (["sun", "temp"], ["sun", "temp", "fc"],
                  ["sun", "temp", "fc", "rain"], ["sun", "temp", "fc", "rain", "dose"]):
        X = [[r[k] for k in combo] for r in rows]
        m = loocv(X, y)
        label = f"combined: {len(combo)} predictors"
        if m:
            print(f"{label:<34}{m:>10.2f}{m - base:>+10.2f}")

    # in-sample, to show the illusion
    print("\n-- the same models scored IN-SAMPLE (the illusion) --")
    mean_y = sum(y) / n
    print(f"{'flat mean':<34}{sum(abs(v - mean_y) for v in y) / n:>10.2f}")
    for combo in (["sun", "temp", "fc"], ["sun", "temp", "fc", "rain", "dose"]):
        X = [[r[k] for k in combo] for r in rows]
        coef = fit(X, y)
        ins = sum(abs(y[i] - predict(coef, X[i])) for i in range(n)) / n
        print(f"{'combined: ' + str(len(combo)) + ' predictors':<34}{ins:>10.2f}")


if __name__ == "__main__":
    main()
