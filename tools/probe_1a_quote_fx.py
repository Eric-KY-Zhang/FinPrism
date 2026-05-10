"""Phase 4f Step 1A — Xueqiu realtime quote FX probe.

Findings while building this probe (recorded for 1D report):
  * The plan-suggested symbol form ``USDCNY,HKDCNY,KRWCNY`` returns HTTP 200 with
    ``{"data":[],"error_code":0}`` — i.e. silently empty. Correct form is the
    ``.FX`` suffix: ``USDCNY.FX,HKDCNY.FX,KRWCNY.FX``.
  * ``Accept-Encoding: identity`` (per the global header recipe) silently
    returns an empty body (Content-Length: 0). Switching to
    ``Accept-Encoding: gzip, deflate`` resolves it. Anti-bot fingerprint.
  * A session warmup hit on ``https://xueqiu.com/hq`` is required to obtain the
    ``xq_a_token`` / ``xqat`` cookies. Without it the API also returns empty.

This probe records BOTH the plan-suggested call (no warmup, identity encoding,
no .FX) and the working call, so the diagnosis is reproducible from the dump.
"""
from __future__ import annotations

import datetime
import json
import os
import time

import requests

API = "https://stock.xueqiu.com/v5/stock/realtime/quotec.json"
WARMUP_URL = "https://xueqiu.com/hq"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

# Plan-spec headers (Accept-Encoding: identity)
HEADERS_PLAN = {
    "User-Agent": UA,
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "identity",
    "Referer": "https://xueqiu.com/",
    "Origin": "https://xueqiu.com",
}

# Working headers (gzip allowed)
HEADERS_WORK = dict(HEADERS_PLAN)
HEADERS_WORK["Accept-Encoding"] = "gzip, deflate"

OUT_PATH = "samples/xueqiu_quote_currency.json"


def safe_parse(body: str | None):
    if not body:
        return None
    if not body.lstrip().startswith("{"):
        return None
    try:
        return json.loads(body)
    except Exception:
        return None


def call(session: requests.Session, headers: dict, symbol: str, cookie: str | None):
    h = dict(headers)
    if cookie:
        h["Cookie"] = cookie
    r = session.get(API, headers=h, params={"symbol": symbol}, timeout=30)
    return {
        "url": r.url,
        "status": r.status_code,
        "len": len(r.text),
        "body_excerpt": r.text[:1200] if r.text else None,
        "parsed": safe_parse(r.text),
    }


def warmup(session: requests.Session) -> dict:
    h = dict(HEADERS_WORK)
    h["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    h.pop("Origin", None)
    h.pop("Referer", None)
    r = session.get(WARMUP_URL, headers=h, timeout=30)
    return {
        "url": WARMUP_URL,
        "status": r.status_code,
        "cookies_obtained": sorted(session.cookies.keys()),
    }


def main() -> None:
    cookie_env = os.environ.get("XUEQIU_COOKIE", "") or None

    # 1) Plan-spec call: no session, no warmup, identity encoding, no .FX suffix
    plan_session = requests.Session()
    plan_call = call(
        plan_session,
        HEADERS_PLAN,
        "USDCNY,HKDCNY,KRWCNY",
        cookie_env,
    )
    time.sleep(1.2)

    # 2) Working call (no XUEQIU_COOKIE env): warmup + .FX suffix + gzip
    work_session = requests.Session()
    warm = warmup(work_session)
    time.sleep(1.2)
    work_call = call(
        work_session,
        HEADERS_WORK,
        "USDCNY.FX,HKDCNY.FX,KRWCNY.FX",
        None,  # rely on session cookies, not env
    )
    time.sleep(1.2)

    # 3) With-cookie variant (only if env present)
    cookie_call = None
    if cookie_env:
        cookie_session = requests.Session()
        cookie_call = call(
            cookie_session,
            HEADERS_WORK,
            "USDCNY.FX,HKDCNY.FX,KRWCNY.FX",
            cookie_env,
        )
        time.sleep(1.2)

    # 4) Cross-check: plain-symbol vs .FX, single CNY variant
    extra_calls = {}
    for symbol in ("USDCNY", "USDCNY.FX", "CNY"):
        s = requests.Session()
        warmup(s)
        time.sleep(1.0)
        extra_calls[symbol] = call(s, HEADERS_WORK, symbol, None)
        time.sleep(1.2)

    out = {
        "url_template": API,
        "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "cookie_present_in_env": bool(cookie_env),
        "diagnosis": {
            "plan_symbol_form_works": False,
            "needed_symbol_suffix": ".FX",
            "needed_warmup": WARMUP_URL,
            "needed_accept_encoding": "gzip, deflate (NOT 'identity')",
            "schema_note": (
                "Response shape is {\"data\": [...], \"error_code\": 0}. "
                "Items are at parsed['data'], NOT parsed['data']['items']."
            ),
        },
        "plan_call": {
            "label": "plan-spec: identity encoding, no warmup, no .FX",
            "headers": HEADERS_PLAN,
            "result": plan_call,
        },
        "warmup": warm,
        "no_cookie_working": {
            "label": "working: warmup + .FX + gzip",
            "headers_diff": {"Accept-Encoding": "gzip, deflate"},
            "result": work_call,
        },
        "with_cookie": {
            "label": "with XUEQIU_COOKIE env",
            "result": cookie_call,
        },
        "symbol_variants": extra_calls,
    }
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # Console summary
    print(f"OUT: {OUT_PATH}")
    print(f"plan_call status={plan_call['status']} len={plan_call['len']}")
    print(f"work_call status={work_call['status']} len={work_call['len']}")
    parsed = work_call["parsed"]
    if parsed:
        items = parsed.get("data") or []
        print(f"work_call error_code={parsed.get('error_code')}, items={len(items)}")
        for it in items:
            if isinstance(it, dict):
                print(
                    f"  {it.get('symbol')}: current={it.get('current')} "
                    f"last_close={it.get('last_close')} ts={it.get('timestamp')}"
                )


if __name__ == "__main__":
    main()
