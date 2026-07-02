#!/usr/bin/env python3
"""
weather.py — Pull current weather for the pool's location (Guilford, CT) and
map it to dose.py's loss bucket.

Source: Open-Meteo (https://open-meteo.com) — free, no API key, stdlib only.

Usage:
    python weather.py
    python weather.py --json        # machine-readable, for chaining into logging

Output (default): a log-ready Weather string + the recommended dose.py --weather
bucket (sunny / partly / cloudy / rain), so the same conditions drive both the
log entry and the dose calc. The "sensor is you" rule still holds for FC/CC —
this only fills the Weather column and the loss estimate.

All location/tuning constants live in CONFIG below.
"""

import argparse
import json
import sys
import urllib.request

CONFIG = {
    # Pool location — Guilford, CT
    "latitude": 41.2834,
    "longitude": -72.6818,
    "timezone": "America/New_York",

    # Daily precip (inches) at/above which the day counts as a "rain" loss day
    "rain_precip_in": 0.10,

    # PRIMARY classifier: fraction of elapsed daylight hours with bright sunshine
    # (sunshine_duration). This tracks actual UV far better than cloud cover,
    # which counts thin high cloud the sun shines straight through.
    "sunny_min_sun_frac": 0.60,   # >= this -> sunny
    "partly_min_sun_frac": 0.30,  # >= this -> partly, else cloudy

    # FALLBACK only (if hourly sunshine is unavailable): cloud-cover (%) cutoffs
    "sunny_max_cloud": 30,    # < this -> sunny
    "partly_max_cloud": 70,   # < this -> partly, else cloudy
}

# WMO weather code -> (short text, is_precip)
WMO = {
    0: ("Clear", False), 1: ("Mainly clear", False), 2: ("Partly cloudy", False),
    3: ("Overcast", False), 45: ("Fog", False), 48: ("Rime fog", False),
    51: ("Light drizzle", True), 53: ("Drizzle", True), 55: ("Heavy drizzle", True),
    56: ("Freezing drizzle", True), 57: ("Freezing drizzle", True),
    61: ("Light rain", True), 63: ("Rain", True), 65: ("Heavy rain", True),
    66: ("Freezing rain", True), 67: ("Freezing rain", True),
    71: ("Light snow", True), 73: ("Snow", True), 75: ("Heavy snow", True),
    77: ("Snow grains", True), 80: ("Rain showers", True), 81: ("Rain showers", True),
    82: ("Violent rain showers", True), 85: ("Snow showers", True),
    86: ("Snow showers", True), 95: ("Thunderstorm", True),
    96: ("Thunderstorm w/ hail", True), 99: ("Thunderstorm w/ hail", True),
}


def fetch_history(cfg, past_days):
    """Daily completed totals for the last `past_days` days (+ today, partial).

    Returns list of dicts: date, sun_hrs, uv_max, rad_mj, precip_in, hi_f.
    """
    url = (
        "https://api.open-meteo.com/v1/forecast"
        f"?latitude={cfg['latitude']}&longitude={cfg['longitude']}"
        "&daily=sunshine_duration,uv_index_max,shortwave_radiation_sum,"
        "precipitation_sum,temperature_2m_max"
        "&temperature_unit=fahrenheit&precipitation_unit=inch"
        f"&timezone={cfg['timezone']}&past_days={past_days}&forecast_days=1"
    )
    with urllib.request.urlopen(url, timeout=20) as r:
        dy = json.load(r)["daily"]
    out = []
    for i, t in enumerate(dy["time"]):
        out.append({
            "date": t,
            "sun_hrs": round((dy["sunshine_duration"][i] or 0) / 3600.0, 1),
            "uv_max": dy["uv_index_max"][i],
            "rad_mj": dy["shortwave_radiation_sum"][i],
            "precip_in": dy["precipitation_sum"][i],
            "hi_f": dy["temperature_2m_max"][i],
        })
    return out


def fetch(cfg):
    url = (
        "https://api.open-meteo.com/v1/forecast"
        f"?latitude={cfg['latitude']}&longitude={cfg['longitude']}"
        "&current=temperature_2m,relative_humidity_2m,precipitation,weather_code,"
        "cloud_cover,wind_speed_10m,is_day"
        "&hourly=cloud_cover,shortwave_radiation,sunshine_duration,precipitation"
        "&daily=weather_code,temperature_2m_max,temperature_2m_min,"
        "precipitation_sum,uv_index_max"
        "&temperature_unit=fahrenheit&wind_speed_unit=mph&precipitation_unit=inch"
        f"&timezone={cfg['timezone']}&forecast_days=1"
    )
    with urllib.request.urlopen(url, timeout=20) as r:
        return json.load(r)


def day_so_far(cur, hourly):
    """Summarize the SUN ACTUALLY RECEIVED today up to now, from hourly data.

    A single instantaneous snapshot at test time can misread the whole day
    (e.g. a cloud rolling in 10 min before the noon test makes a sunny morning
    look 'overcast'). This sums/averages the realized daytime hours instead.

    Returns dict: {sunshine_hrs, avg_day_cloud, n_day_hours, precip_today_so_far}
    or None if hourly data is unavailable.
    """
    if not hourly or "time" not in hourly:
        return None
    now = cur.get("time", "")
    now_date, now_hour = now[:10], int(now[11:13] or 0)

    sun_secs, clouds, precip = 0.0, [], 0.0
    for i, t in enumerate(hourly["time"]):
        if t[:10] != now_date or int(t[11:13]) > now_hour:
            continue
        sun_secs += (hourly["sunshine_duration"][i] or 0.0)
        precip += (hourly["precipitation"][i] or 0.0)
        if (hourly["shortwave_radiation"][i] or 0.0) > 0:   # daylight hour
            clouds.append(hourly["cloud_cover"][i] or 0.0)
    return {
        "sunshine_hrs": sun_secs / 3600.0,
        "avg_day_cloud": (sum(clouds) / len(clouds)) if clouds else None,
        "n_day_hours": len(clouds),
        "precip_today_so_far": precip,
    }


def classify(cur, daily, cfg, dsf=None):
    """Return one of dose.py's buckets: sunny / partly / cloudy / rain.

    Uses the realized daytime cloud average so far (dsf) when available, so one
    stray cloud at test time can't mislabel the day. Falls back to the
    instantaneous cloud_cover snapshot if hourly data is missing.
    """
    code = cur.get("weather_code")
    code_is_precip = WMO.get(code, ("", False))[1]
    precip_today = daily["precipitation_sum"][0] or 0.0
    raining_now = (cur.get("precipitation") or 0.0) > 0.0

    if raining_now or precip_today >= cfg["rain_precip_in"] or (
            code_is_precip and (cur.get("cloud_cover") or 0) >= cfg["partly_max_cloud"]):
        return "rain"

    # Primary: realized sunshine fraction of daylight elapsed so far.
    if dsf and dsf["n_day_hours"] >= 2:
        frac = dsf["sunshine_hrs"] / dsf["n_day_hours"]
        if frac >= cfg["sunny_min_sun_frac"]:
            return "sunny"
        if frac >= cfg["partly_min_sun_frac"]:
            return "partly"
        return "cloudy"

    # Fallback: instantaneous cloud snapshot (hourly sunshine unavailable).
    cloud = cur.get("cloud_cover")
    if cloud is None:
        return "partly"
    if cloud < cfg["sunny_max_cloud"]:
        return "sunny"
    if cloud < cfg["partly_max_cloud"]:
        return "partly"
    return "cloudy"


def log_string(cur, daily, dsf=None):
    code = cur.get("weather_code")
    sky = WMO.get(code, ("Unknown", False))[0]
    temp = cur.get("temperature_2m")
    rh = cur.get("relative_humidity_2m")
    precip_today = daily["precipitation_sum"][0] or 0.0
    hi = daily["temperature_2m_max"][0]
    uv = daily["uv_index_max"][0]
    sun = ""
    if dsf is not None and dsf["n_day_hours"] >= 1:
        frac = dsf["sunshine_hrs"] / dsf["n_day_hours"]
        sun = (f"; sun so far {dsf['sunshine_hrs']:.1f}h of {dsf['n_day_hours']}h "
               f"daylight ({frac*100:.0f}%)")
    return (f"{sky} now, {temp:.0f}F, {rh:.0f}% RH{sun}; "
            f"today high {hi:.0f}F, {precip_today:.2f}in precip, UV max {uv:.1f}")


def main():
    p = argparse.ArgumentParser(description="Pull pool-site weather (Open-Meteo)")
    p.add_argument("--json", action="store_true",
                   help="Emit JSON (weather string + bucket + raw fields)")
    p.add_argument("--history", type=int, metavar="N",
                   help="Print JSON of daily sun totals for the last N days "
                        "(+ today, partial). For the Sun tab in Pool_Log.xlsx.")
    args = p.parse_args()

    if args.history is not None:
        try:
            print(json.dumps(fetch_history(CONFIG, args.history), indent=2))
        except Exception as e:
            sys.exit(f"history fetch failed: {e}")
        return

    try:
        d = fetch(CONFIG)
    except Exception as e:
        sys.exit(f"weather fetch failed: {e}\n"
                 "(No network? Fall back to the user's own observation.)")

    cur, daily = d["current"], d["daily"]
    dsf = day_so_far(cur, d.get("hourly"))
    bucket = classify(cur, daily, CONFIG, dsf)
    summary = log_string(cur, daily, dsf)

    if args.json:
        print(json.dumps({
            "weather_string": summary,
            "dose_bucket": bucket,
            "observed_at": cur.get("time"),
            "day_so_far": dsf,
            "current": cur,
            "daily": daily,
        }, indent=2))
    else:
        print(f"Weather (Guilford, CT) @ {cur.get('time')}")
        print(f"  Log string : {summary}")
        print(f"  dose.py bucket: --weather {bucket}")
        if dsf is not None and dsf["n_day_hours"] >= 1:
            frac = dsf["sunshine_hrs"] / dsf["n_day_hours"]
            snap = cur.get("cloud_cover")
            print(f"  (realized: {dsf['sunshine_hrs']:.1f} sunshine hrs / "
                  f"{dsf['n_day_hours']}h daylight = {frac*100:.0f}% sun; "
                  f"snapshot now {snap:.0f}% cloud — classifier uses the sun fraction)")


if __name__ == "__main__":
    main()
