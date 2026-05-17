from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def avg(a: float, b: float | None) -> float:
    return a if b is None else (a + b) / 2


def ratio(a: float, b: float) -> float:
    return a / b


def calc_metrics(s: dict[str, float], days: int = 365) -> dict[str, float]:
    ni_for_roe = s.get("pni", s["ni"])
    return {
        "NPM": ratio(s["ni"], s["rev"]),
        "GPM": ratio(s.get("gp", s["rev"] - s["cogs"]), s["rev"]),
        "OER": ratio(s["opex"], s["rev"]),
        "ROA": ratio(s["ni"], avg(s["ta"], s["ta_p"])),
        "ROE": ratio(ni_for_roe, avg(s["eq"], s["eq_p"])),
        "TAGR": ratio(s["ta"], s["ta_p"]) - 1,
        "RGR": ratio(s["rev"], s["rev_p"]) - 1,
        "NPGR": ratio(s["ni"], s["ni_p"]) - 1,
        "CR": ratio(s["ca"], s["cl"]),
        "QR": ratio(s["ca"] - s["inv"], s["cl"]),
        "CASHR": ratio(s["cash"], s["cl"]),
        "DAR": ratio(s["tl"], s["ta"]),
        "DIO": ratio(avg(s["inv"], s["inv_p"]) * days, s["cogs"]),
        "DSO": ratio(avg(s["ar"], s["ar_p"]) * days, s["rev"]),
        "DPO": ratio(avg(s["ap"], s["ap_p"]) * days, s["cogs"]),
        "CAT": ratio(s["rev"], avg(s["ca"], s["ca_p"])),
        "TAT": ratio(s["rev"], avg(s["ta"], s["ta_p"])),
    }


def add_ccc(m: dict[str, float]) -> dict[str, float]:
    m["CCC"] = m["DIO"] + m["DSO"] - m["DPO"]
    return m


SAMPLES: dict[str, tuple[int, dict[str, float]]] = {
    # Source: Sina public financial tables for 300866 Anker Innovations, CNY 10k.
    "A_300866_ANKER": (
        360,
        {
            "rev": 3051440.34,
            "rev_p": 2471008.03,
            "cogs": 1676297.55,
            "gp": 3051440.34 - 1676297.55,
            "opex": 682680.28 + 109305.85 + 5261.85 + 289278.55,
            "ni": 261719.39,
            "ni_p": 221112.40,
            "pni": 254513.19,
            "ta": 2006689.25,
            "ta_p": 1660370.73,
            "tl": 935522.24,
            "eq": 1052760.33,
            "eq_p": 895804.34,
            "ca": 1608652.66,
            "ca_p": 1236754.93,
            "cl": 676851.15,
            "inv": 499711.90,
            "inv_p": 323355.42,
            "ar": 187262.35,
            "ar_p": 165420.01,
            "ap": 187050.66,
            "ap_p": 177835.92,
            "cash": 365650.93,
        },
    ),
    # Source: StockAnalysis public TSLA FY2025 tables, USD millions.
    "US_TSLA": (
        365,
        {
            "rev": 94827.0,
            "rev_p": 97690.0,
            "cogs": 77733.0,
            "gp": 17094.0,
            "opex": 12739.0,
            "ni": 3855.0,
            "ni_p": 7153.0,
            "ta": 137806.0,
            "ta_p": 122070.0,
            "tl": 54941.0,
            "eq": 82137.0,
            "eq_p": 72913.0,
            "ca": 68642.0,
            "ca_p": 58360.0,
            "cl": 31714.0,
            "inv": 12392.0,
            "inv_p": 12017.0,
            "ar": 4576.0,
            "ar_p": 4418.0,
            "ap": 13371.0,
            "ap_p": 12474.0,
            "cash": 16513.0,
        },
    ),
    # Source: HKEX Aoji FY2025 annual report / Xueqiu public table, RMB millions.
    "HK_02519_AOJI": (
        365,
        {
            "rev": 13698.738,
            "rev_p": 10709.648,
            "cogs": 10090.568,
            "gp": 3608.170,
            "opex": 3202.153,
            "ni": 161.139,
            "ni_p": 504.299,
            "ta": 10134.961,
            "ta_p": 8779.550,
            "tl": 6982.303,
            "eq": 3132.425,
            "eq_p": 3079.521,
            "ca": 5997.382,
            "ca_p": 4768.520,
            "cl": 4013.762,
            "inv": 1627.440,
            "inv_p": 1445.386,
            "ar": 1560.190,
            "ar_p": 1269.396,
            "ap": 1698.174,
            "ap_p": 1438.444,
            "cash": 1613.126,
        },
    ),
    # Source: StockAnalysis public 013890 FY2025 tables, KRW millions.
    "KR_013890_ZINUS": (
        365,
        {
            "rev": 913180.0,
            "rev_p": 920399.0,
            "cogs": 595725.0,
            "gp": 317455.0,
            "opex": 292027.0,
            "ni": -18506.0,
            "ni_p": -6798.0,
            "ta": 1073121.0,
            "ta_p": 1192229.0,
            "tl": 423750.0,
            "eq": 649371.0,
            "eq_p": 672764.0,
            "ca": 522353.0,
            "ca_p": 656034.0,
            "cl": 322463.0,
            "inv": 157589.0,
            "inv_p": 230167.0,
            "ar": 103979.0,
            "ar_p": 266223.0,
            "ap": 54511.0,
            "ap_p": 83966.0,
            "cash": 127479.0,
        },
    ),
    # Source: TSMC 2025Q4 official financial statements / FinMind public API, NT$ millions.
    "TW_2330_TSMC": (
        365,
        {
            "rev": 3809054.272,
            "rev_p": 2894307.699,
            "cogs": 1527760.293,
            "gp": 2281293.979,
            "opex": 345649.630,
            "ni": 1717882.627,
            "ni_p": 1173267.703,
            "ta": 7933023.878,
            "ta_p": 6691938.000,
            "tl": 2472228.595,
            "eq": 5419595.994,
            "eq_p": 4288545.167,
            "ca": 3817130.817,
            "ca_p": 3088352.120,
            "cl": 1458019.289,
            "inv": 288109.485,
            "inv_p": 287868.810,
            "ar": 279051.553,
            "ar_p": 270683.235,
            "ap": 82551.595,
            "ap_p": 72800.558,
            "cash": 2767856.402,
        },
    ),
}


def assert_close(label: str, actual: float, expected: float, tol: float = 1e-9) -> None:
    if abs(actual - expected) > tol:
        raise AssertionError(f"{label}: expected {expected}, got {actual}")


def check_vba_formula_text() -> None:
    tool_module = None
    for path in (ROOT / "modules").glob("*.bas"):
        text = path.read_text(encoding="utf-8-sig", errors="ignore")
        if "Private Function StandardIndicatorFormula" in text:
            tool_module = text
            break
    if tool_module is None:
        raise AssertionError("cannot locate StandardIndicatorFormula")

    required_snippets = [
        'StandardIndicatorFormula = RatioFormula(rev & "-" & cogs, rev)',
        'StandardRef(rowMap, "FIN", isCol), StandardRef(rowMap, "RD", isCol)',
        'StandardIndicatorFormula = RatioFormula(pni, AverageExpression(eq, eqP))',
        'rowNum = StandardRawDumpRow(sheetName, rowNum, colNum)',
        'Private Function StandardRawDumpRow',
    ]
    for snippet in required_snippets:
        if snippet not in tool_module:
            raise AssertionError(f"missing VBA formula fix: {snippet}")


def main() -> int:
    check_vba_formula_text()

    for name, (days, sample) in SAMPLES.items():
        metrics = add_ccc(calc_metrics(sample, days))
        if name == "A_300866_ANKER":
            assert_close("A GPM excludes tax surcharge", metrics["GPM"], (sample["rev"] - sample["cogs"]) / sample["rev"])
            assert_close("A OER includes R&D", metrics["OER"], sample["opex"] / sample["rev"])
            assert_close("A ROE uses average equity", metrics["ROE"], sample["pni"] / avg(sample["eq"], sample["eq_p"]))
        print(
            f"{name}: "
            f"NPM={metrics['NPM']:.6f} GPM={metrics['GPM']:.6f} "
            f"OER={metrics['OER']:.6f} ROA={metrics['ROA']:.6f} "
            f"ROE={metrics['ROE']:.6f} CR={metrics['CR']:.4f} "
            f"DIO={metrics['DIO']:.2f} DSO={metrics['DSO']:.2f} "
            f"DPO={metrics['DPO']:.2f} CCC={metrics['CCC']:.2f}"
        )

    print("indicator formula logic check PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
