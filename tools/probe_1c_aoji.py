"""Phase 4f Step 1C — Aoji (傲基) stock code confirmation probe.

Runs 5 candidate lookups (3 xueqiu, 1 a-share xueqiu, 1 sina A-share HTML),
extracts company_name / quote_name / currency / exchange, and dumps:
  - samples/aoji_probe_summary.json (normalized table)
  - samples/aoji_probe_<label>.json|.html (raw bodies)

Uses the same Xueqiu warmup (xueqiu.com/hq) and gzip Accept-Encoding pattern
discovered in 1A/1B — without warmup, Xueqiu returns 200 + empty body.
"""
from __future__ import annotations

import datetime
import json
import os
import re
import time

import requests

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

XUEQIU_HEADERS = {
    "User-Agent": UA,
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Referer": "https://xueqiu.com/",
    "Origin": "https://xueqiu.com",
}

SINA_HEADERS = {
    "User-Agent": UA,
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Accept-Encoding": "gzip, deflate",
}

CANDIDATES = [
    ("HK_02519_xueqiu_quote", "hk_quote", "02519",
     "https://stock.xueqiu.com/v5/stock/quote.json?symbol=02519&extend=detail"),
    ("HK_02519_xueqiu_finance_balance", "hk_fin", "02519",
     "https://stock.xueqiu.com/v5/stock/finance/hk/balance.json"
     "?symbol=02519&type=all&is_detail=true&count=8"),
    ("HK_09927_xueqiu_quote", "hk_quote", "09927",
     "https://stock.xueqiu.com/v5/stock/quote.json?symbol=09927&extend=detail"),
    ("A_003031_xueqiu_quote", "a_quote", "SZ003031",
     "https://stock.xueqiu.com/v5/stock/quote.json?symbol=SZ003031&extend=detail"),
    ("A_003031_sina_balance", "a_sina", "sz003031",
     "https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/"
     "stockid/003031/ctrl/part/displaytype/4.phtml"),
]


def warmup_xueqiu(session: requests.Session) -> dict:
    h = dict(XUEQIU_HEADERS)
    h["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    h.pop("Origin", None)
    h.pop("Referer", None)
    r = session.get("https://xueqiu.com/hq", headers=h, timeout=30)
    return {"status": r.status_code, "cookies": sorted(session.cookies.keys())}


def safe_json(body: str | None):
    if not body or not body.lstrip().startswith("{"):
        return None
    try:
        return json.loads(body)
    except Exception:
        return None


def main() -> None:
    cookie = os.environ.get("XUEQIU_COOKIE", "") or None
    xueqiu_session = requests.Session()
    warm = warmup_xueqiu(xueqiu_session)
    time.sleep(1.2)

    results = []
    for label, kind, sym, url in CANDIDATES:
        if "xueqiu" in label.lower():
            h = dict(XUEQIU_HEADERS)
            if cookie:
                h["Cookie"] = cookie
            sess = xueqiu_session
        else:
            h = dict(SINA_HEADERS)
            sess = requests.Session()

        try:
            r = sess.get(url, headers=h, timeout=30)
            if "sina" in label.lower():
                r.encoding = "gb18030"  # sina financial pages are GB18030
            body = r.text
            status = r.status_code
        except Exception as exc:  # pylint: disable=broad-except
            body = f"EXCEPTION: {exc}"
            status = None

        summary = {
            "label": label,
            "kind": kind,
            "url": url,
            "status": status,
            "body_len": len(body) if isinstance(body, str) else 0,
            "body_excerpt": body[:800] if isinstance(body, str) else None,
        }

        if "xueqiu" in label.lower():
            parsed = safe_json(body)
            if parsed:
                summary["error_code"] = parsed.get("error_code")
                summary["error_description"] = parsed.get("error_description")
                if "quote.json" in url:
                    quote = (parsed.get("data") or {}).get("quote") or {}
                    summary["company_name"] = quote.get("name")
                    summary["currency"] = quote.get("currency")
                    summary["exchange"] = quote.get("exchange")
                    summary["type"] = quote.get("type")
                    summary["status_field"] = quote.get("status")
                    summary["isin"] = quote.get("isin")
                    summary["fin_report_name"] = quote.get("fin_report_name")
                    summary["listed_date"] = quote.get("listed_date")
                elif "balance" in url:
                    data = parsed.get("data") or {}
                    summary["quote_name"] = data.get("quote_name")
                    summary["currency"] = data.get("currency")
                    summary["last_report_name"] = data.get("last_report_name")
                    summary["annual_settle_date"] = data.get("annual_settle_date")
                    summary["list_len"] = len(data.get("list") or [])
        if "sina" in label.lower() and isinstance(body, str):
            summary["html_contains_aoji"] = "傲基" in body
            summary["html_contains_aoji_pinyin"] = bool(re.search(r"Aoji", body, re.I))
            summary["html_contains_unknown"] = "无效" in body or "未知" in body or "无此" in body
            m = re.search(r"<title>(.*?)</title>", body, re.IGNORECASE | re.DOTALL)
            summary["title"] = m.group(1).strip() if m else None
            # try to find company name
            m2 = re.search(r"<h1[^>]*>(.*?)</h1>", body, re.IGNORECASE | re.DOTALL)
            summary["h1"] = m2.group(1).strip() if m2 else None

        # Dump raw body
        ext = "html" if "sina" in label.lower() else "json"
        safe_label = re.sub(r"[^A-Za-z0-9_]", "_", label)
        raw_path = f"samples/aoji_probe_{safe_label}.{ext}"
        try:
            with open(raw_path, "w", encoding="utf-8") as f:
                f.write(body if isinstance(body, str) else "")
            summary["dump"] = raw_path
        except Exception as exc:  # pylint: disable=broad-except
            summary["dump_error"] = str(exc)

        results.append(summary)
        time.sleep(1.2)

    out = {
        "fetched_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "cookie_present_in_env": bool(cookie),
        "warmup": warm,
        "candidates": results,
        "preexisting_evidence": {
            "samples/xueqiu_HK_02519_balance.json": (
                "Phase 4c HK probe residue (UTF-16-LE encoded): header already "
                "shows quote_name=傲基控股, currency=CNY, "
                "annual_settle_date=12-31. Strong evidence 02519.HK = 傲基股份."
            ),
        },
    }
    with open("samples/aoji_probe_summary.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("OUT: samples/aoji_probe_summary.json")
    for r in results:
        print(
            f"  [{r.get('label')}] status={r.get('status')} "
            f"name={r.get('company_name') or r.get('quote_name') or r.get('h1') or '-'} "
            f"currency={r.get('currency','-')} exchange={r.get('exchange','-')}"
        )


if __name__ == "__main__":
    main()
