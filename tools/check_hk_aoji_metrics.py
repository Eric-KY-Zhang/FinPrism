from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def load_sample(name: str) -> dict:
    raw = (ROOT / "samples" / name).read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "utf-16-le", "gb18030"):
        try:
            return json.loads(raw.decode(encoding))
        except Exception:
            continue
    raise ValueError(f"cannot decode sample: {name}")


def record(name: str, year: int) -> dict:
    data = load_sample(name)
    for item in data["data"]["list"]:
        period_end = str(item.get("ed") or item.get("report_date") or "")
        if period_end.startswith(str(year)):
            return item
    raise ValueError(f"missing year {year} in {name}")


def first(value) -> float:
    if isinstance(value, list):
        value = value[0] if value else 0
    return float(value or 0)


def millions(value) -> float:
    return first(value) / 1_000_000


def assert_close(name: str, actual: float, expected: float, tolerance: float = 0.001) -> None:
    if abs(actual - expected) > tolerance:
        raise AssertionError(f"{name}: expected {expected}, actual {actual}")


def main() -> int:
    balance_2025 = record("xueqiu_HK_02519_balance.json", 2025)
    balance_2024 = record("xueqiu_HK_02519_balance.json", 2024)
    income_2025 = record("xueqiu_HK_02519_income.json", 2025)
    income_2024 = record("xueqiu_HK_02519_income.json", 2024)
    cashflow_2025 = record("xueqiu_HK_02519_cash_flow.json", 2025)

    revenue = millions(income_2025["tto"])
    cogs = millions(income_2025["slgcost"])
    net_income = millions(income_2025["ploashh"])
    current_assets = millions(balance_2025["ca"])
    current_liabilities = millions(balance_2025["clia"])
    inventory = millions(balance_2025["iv"])
    receivables = millions(balance_2025["trrb"])
    payables = millions(balance_2025["trpy"])
    equity = millions(balance_2025["shhfd"])
    capex = millions(cashflow_2025["adtfxda"])
    cash_end = millions(cashflow_2025["cceqeyr"])
    eps = first(income_2025["beps_aju"])

    assert_close("Revenue", revenue, 13698.738)
    assert_close("COGS", cogs, 10090.568)
    assert_close("Inventory", inventory, 1627.44)
    assert_close("Accounts receivable", receivables, 1560.19)
    assert_close("Accounts payable", payables, 1698.174)
    assert_close("Stockholders equity", equity, 3132.425)
    assert_close("Basic EPS", eps, 0.3909801455, tolerance=0.0000001)
    assert_close("Capex", capex, -252.696)
    assert_close("Cash ratio", cash_end / current_liabilities, 0.4019, tolerance=0.0001)

    avg_inventory = (inventory + millions(balance_2024["iv"])) / 2
    avg_receivables = (receivables + millions(balance_2024["trrb"])) / 2
    avg_payables = (payables + millions(balance_2024["trpy"])) / 2
    dio = avg_inventory * 365 / cogs
    dso = avg_receivables * 365 / revenue
    dpo = avg_payables * 365 / cogs

    assert_close("Gross margin", (revenue - cogs) / revenue, 0.263394, tolerance=0.000001)
    assert_close("Net margin", net_income / revenue, 0.011763, tolerance=0.000001)
    assert_close("DIO", dio, 55.5757, tolerance=0.0001)
    assert_close("DSO", dso, 37.6969, tolerance=0.0001)
    assert_close("DPO", dpo, 56.7295, tolerance=0.0001)
    assert_close("CCC", dio + dso - dpo, 36.5431, tolerance=0.0001)

    print("HK Aoji FY2025 metric check PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
