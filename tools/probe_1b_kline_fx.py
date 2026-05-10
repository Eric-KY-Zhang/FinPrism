"""Phase 4f Step 1B — Xueqiu kline historical FX probe.

Same lessons from 1A apply: ``USDCNY`` returns silently empty, must be
``USDCNY.FX``; needs warmup hit on ``xueqiu.com/hq`` and ``Accept-Encoding:
gzip, deflate`` instead of ``identity``.

This probe also measures the count cap (Xueqiu refuses counts > ~2147 for FX
day series — see ``count_cap_probe`` in the dump).
"""
from __future__ import annotations

import datetime
import json
import os
import time

import requests

API = "https://stock.xueqiu.com/v5/stock/chart/kline.json"
WARMUP_URL = "https://xueqiu.com/hq"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": UA,
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Referer": "https://xueqiu.com/",
    "Origin": "https://xueqiu.com",
}

OUT_PATH = "samples/xueqiu_kline_USDCNY.json"

# 2024-01-01 00:00:00 +08:00 in ms
BEGIN_2024_CST_MS = 1704038400000
# 2026-05-03 00:00:00 +08:00 in ms (today, today's date per env)
BEGIN_TODAY_CST_MS = int(
    (datetime.datetime(2026, 5, 3) - datetime.datetime(1970, 1, 1)).total_seconds() * 1000
)


def warmup(session: requests.Session) -> dict:
    h = dict(HEADERS)
    h["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    h.pop("Origin", None)
    h.pop("Referer", None)
    r = session.get(WARMUP_URL, headers=h, timeout=30)
    return {
        "url": WARMUP_URL,
        "status": r.status_code,
        "cookies_obtained": sorted(session.cookies.keys()),
    }


def fetch_kline(session: requests.Session, symbol: str, begin_ms: int,
                count: int, cookie: str | None) -> tuple[str, int, str]:
    h = dict(HEADERS)
    if cookie:
        h["Cookie"] = cookie
    params = {
        "symbol": symbol,
        "begin": begin_ms,
        "period": "day",
        "type": "before",
        "count": count,
    }
    r = session.get(API, headers=h, params=params, timeout=30)
    return r.url, r.status_code, r.text


def parse(body: str | None):
    if not body or not body.lstrip().startswith("{"):
        return None
    try:
        return json.loads(body)
    except Exception:
        return None


def ts_to_dt_str(ts_ms: int, tz_offset_hours: int = 0) -> str:
    return datetime.datetime.utcfromtimestamp(ts_ms / 1000 + tz_offset_hours * 3600).isoformat()


def summarize(parsed):
    if not parsed:
        return {"error_code": None, "item_count": 0}
    data = parsed.get("data") or {}
    items = data.get("item") or []
    return {
        "error_code": parsed.get("error_code"),
        "error_description": parsed.get("error_description"),
        "column_names": data.get("column"),
        "item_count": len(items),
        "first_item": items[0] if items else None,
        "last_item": items[-1] if items else None,
    }


def main() -> None:
    cookie = os.environ.get("XUEQIU_COOKIE", "") or None
    attempts = []

    session = requests.Session()
    warm = warmup(session)
    time.sleep(1.2)

    # 1) USDCNY 365 trading-days sample (no-cookie + with-cookie if env)
    cookie_modes = [False] + ([True] if cookie else [])
    parsed_primary = None
    for use_c in cookie_modes:
        url, st, body = fetch_kline(
            session, "USDCNY.FX", BEGIN_2024_CST_MS, -365, cookie if use_c else None
        )
        p = parse(body)
        if not use_c:
            parsed_primary = p
        s = summarize(p)
        s.update({
            "label": f"USDCNY.FX_365d_cookie={use_c}",
            "url": url, "status": st, "body_len": len(body) if body else 0,
        })
        attempts.append(s)
        time.sleep(1.2)

    # 1b) Plan-spec form (USDCNY without .FX) for documentation: should be empty
    url, st, body = fetch_kline(
        session, "USDCNY", BEGIN_2024_CST_MS, -365, cookie
    )
    p = parse(body)
    s = summarize(p)
    s.update({
        "label": "USDCNY_plain_no_FX_suffix (control)",
        "url": url, "status": st, "body_len": len(body) if body else 0,
        "body_excerpt": (body[:300] if body else None),
    })
    attempts.append(s)
    time.sleep(1.2)

    # 2) Long pull from today back, count=-1500 (~5.7 yr trading days)
    url, st, body = fetch_kline(
        session, "USDCNY.FX", BEGIN_TODAY_CST_MS, -1500, cookie
    )
    parsed_long = parse(body)
    items_long = (parsed_long or {}).get("data", {}).get("item", [])
    long_summary = {
        "label": "USDCNY.FX_long_count=-1500_from_today",
        "url": url, "status": st, "body_len": len(body) if body else 0,
        "begin_ms": BEGIN_TODAY_CST_MS,
        "begin_meaning_cst": ts_to_dt_str(BEGIN_TODAY_CST_MS, 8),
        "earliest_ts_ms": items_long[0][0] if items_long else None,
        "latest_ts_ms": items_long[-1][0] if items_long else None,
        "earliest_date_utc": ts_to_dt_str(items_long[0][0], 0) if items_long else None,
        "earliest_date_cst": ts_to_dt_str(items_long[0][0], 8) if items_long else None,
        "latest_date_cst": ts_to_dt_str(items_long[-1][0], 8) if items_long else None,
        "item_count": len(items_long),
    }
    attempts.append(long_summary)
    time.sleep(1.2)

    # 3) HKDCNY + KRWCNY 365 day sanity
    for sym in ("HKDCNY.FX", "KRWCNY.FX"):
        url, st, body = fetch_kline(
            session, sym, BEGIN_2024_CST_MS, -365, cookie
        )
        p = parse(body)
        s = summarize(p)
        s.update({
            "label": f"{sym}_365d",
            "url": url, "status": st, "body_len": len(body) if body else 0,
        })
        attempts.append(s)
        time.sleep(1.2)

    # 4) Count cap probe — push to find true upper bound
    cap_probe = []
    for cnt in (-2000, -2500, -3000):
        url, st, body = fetch_kline(
            session, "USDCNY.FX", BEGIN_TODAY_CST_MS, cnt, cookie
        )
        p = parse(body)
        items = (p or {}).get("data", {}).get("item") or []
        cap_probe.append({
            "count_param": cnt,
            "status": st,
            "actual_items_returned": len(items),
            "earliest_date_cst": ts_to_dt_str(items[0][0], 8) if items else None,
        })
        time.sleep(1.5)

    # 5) Time-zone clarification: use begin = exactly 2024-01-01 00:00:00 UTC
    BEGIN_2024_UTC_MS = 1704067200000  # 2024-01-01 00:00 UTC = 2024-01-01 08:00 CST
    url, st, body = fetch_kline(
        session, "USDCNY.FX", BEGIN_2024_UTC_MS, -5, cookie
    )
    p = parse(body)
    items_tz = (p or {}).get("data", {}).get("item") or []
    tz_check = {
        "label": "begin = 2024-01-01 00:00:00 UTC (=08:00 CST)",
        "begin_ms": BEGIN_2024_UTC_MS,
        "expected_first_cst_day_if_begin_is_CST": "2023-12-25 area",
        "expected_first_cst_day_if_begin_is_UTC": "2023-12-25 area (similar)",
        "first_3_items": items_tz[:3] if items_tz else None,
    }

    out = {
        "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "cookie_present_in_env": bool(cookie),
        "primary_sample": {
            "symbol": "USDCNY.FX",
            "period": "day",
            "count": -365,
            "begin_ms": BEGIN_2024_CST_MS,
            "begin_meaning_utc": ts_to_dt_str(BEGIN_2024_CST_MS, 0),
            "begin_meaning_cst": ts_to_dt_str(BEGIN_2024_CST_MS, 8),
        },
        "warmup": warm,
        "attempts": attempts,
        "count_cap_probe": cap_probe,
        "tz_cross_check": tz_check,
        "full_USDCNY_365d_parsed": parsed_primary,
    }
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # Console summary
    print(f"OUT: {OUT_PATH}")
    for a in attempts:
        print(
            f"  [{a.get('label')}] status={a.get('status')} items={a.get('item_count','-')} "
            f"ec={a.get('error_code','-')}"
        )
    print(f"count_cap_probe: {cap_probe}")


if __name__ == "__main__":
    main()
