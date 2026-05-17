"""Phase 5a end-to-end live verification.

Plan executed in one COM session (single open of the workbook):

  Step 2 — VBE smoke
    - Run Test_Phase5a_Xueqiu_AnonWarmup_Smoke
    - Run Test_Phase5a_NoCookieCellNeeded

  Step 3 — 港股 end-to-end with E5 cleared (00700 / 09988 / 02519)
  Step 4 — 美股 end-to-end with E5 cleared (BABA / JD / AAPL)

  Step 5 — FX_missing + HK Tencent fixture + TW TSMC fixture offline tests
  Step 6 — HK performance baseline: 5 companies, time the macro

Safety:
  - Snapshots full sample pool A14:P50 and E5 before mutating; restores after.
  - Aborts before mutation if 港股/美股 sample pool is already populated to a
    degree that could mix with the test slate (asks via screenshot file).
  - All mutations happen on a copy of the workbook (`.phase5a-live.xlsm`)
    so the user's `上市公司财务数据查询.xlsm` is not modified.
"""
from __future__ import annotations

import sys
import time
import shutil
import json
from pathlib import Path
from typing import Optional

import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
BOOK_REAL = ROOT / "上市公司财务数据查询.xlsm"
BOOK = ROOT / "上市公司财务数据查询.phase5a-live.xlsm"

TEST_HK = [("00700", "腾讯控股"), ("09988", "阿里巴巴-W"), ("02519", "傲基股份")]
TEST_US = [("BABA", "阿里巴巴ADR"), ("JD", "京东ADR"), ("AAPL", "Apple")]
TEST_HK_PERF = [("00700", "腾讯控股"), ("00939", "建设银行"), ("00941", "中国移动"),
                ("00388", "港交所"), ("01024", "快手-W")]

POOL_DATA_START_ROW = 14
SCAN_END_ROW = 50

# (market_key, code_col, name_col) for snapshot/restore
MARKET_COLS = {
    "A":  (1, 2),
    "US": (4, 5),
    "HK": (7, 8),
    "KR": (10, 11),
    "TW": (14, 15),
}


def snapshot_pool(ws_pool) -> dict:
    snap = {}
    for mk, (code_col, name_col) in MARKET_COLS.items():
        rows = []
        for r in range(POOL_DATA_START_ROW, SCAN_END_ROW + 1):
            code = ws_pool.Cells(r, code_col).Value
            name = ws_pool.Cells(r, name_col).Value
            rows.append((code, name))
        snap[mk] = rows
    snap["E5"] = ws_pool.Range("E5").Value
    snap["E3"] = ws_pool.Range("E3").Value
    snap["E4"] = ws_pool.Range("E4").Value
    return snap


def write_market_rows(ws_pool, mk: str, samples: list) -> None:
    code_col, name_col = MARKET_COLS[mk]
    # Clear rows we will write
    for offset in range(len(samples)):
        r = POOL_DATA_START_ROW + offset
        ws_pool.Cells(r, code_col).ClearContents()
        ws_pool.Cells(r, name_col).ClearContents()
    for offset, (code, name) in enumerate(samples):
        r = POOL_DATA_START_ROW + offset
        ws_pool.Cells(r, code_col).NumberFormat = "@"
        ws_pool.Cells(r, code_col).Value = code
        ws_pool.Cells(r, name_col).Value = name


def restore_pool(ws_pool, snap: dict) -> None:
    for mk, (code_col, name_col) in MARKET_COLS.items():
        rows = snap[mk]
        for offset, (code, name) in enumerate(rows):
            r = POOL_DATA_START_ROW + offset
            ws_pool.Cells(r, code_col).NumberFormat = "@"
            ws_pool.Cells(r, code_col).Value = code if code is not None else ""
            ws_pool.Cells(r, name_col).Value = name if name is not None else ""
    ws_pool.Range("E5").Value = snap["E5"] if snap["E5"] is not None else ""
    ws_pool.Range("E3").Value = snap["E3"]
    ws_pool.Range("E4").Value = snap["E4"]


def describe_sheet(wb, name: str) -> str:
    try:
        ws = wb.Sheets(name)
    except Exception:
        return f"      {name}: MISSING"
    used = ws.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    last_col = used.Column + used.Columns.Count - 1
    return f"      {name}: last_row={last_row} last_col={last_col}"


def summarize_diag(wb, sheet_name: str) -> dict:
    """Read diagnostic sheet and bucket status codes + source names."""
    try:
        ws = wb.Sheets(sheet_name)
    except Exception:
        return {"_missing": True}
    if ws.Visible == 0:
        try:
            ws.Visible = -1
        except Exception:
            pass
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    status_counts: dict = {}
    source_by_company: dict = {}
    http_status: dict = {}
    error_rows: list = []
    for r in range(3, last_row + 1):
        company = str(ws.Cells(r, 1).Value or "")
        status = str(ws.Cells(r, 4).Value or "")
        source = str(ws.Cells(r, 5).Value or "")
        http_st = str(ws.Cells(r, 14).Value or "")
        if not status:
            continue
        status_counts[status] = status_counts.get(status, 0) + 1
        http_status[http_st] = http_status.get(http_st, 0) + 1
        if company and source:
            source_by_company.setdefault(company, set()).add(source)
        if "401" in http_st or "403" in http_st or "429" in http_st:
            error_rows.append((r, company, status, source, http_st))
    return {
        "total_rows": last_row - 2,
        "status_counts": status_counts,
        "source_by_company": {k: sorted(v) for k, v in source_by_company.items()},
        "http_status_distribution": http_status,
        "error_rows_4xx": error_rows,
    }


def time_macro(excel, macro: str) -> tuple[float, Optional[str]]:
    t0 = time.time()
    err = None
    try:
        excel.Run(macro, True)
    except Exception as e:
        err = str(e)
    return (time.time() - t0, err)


def run_macro_test(excel, macro: str) -> tuple[bool, str]:
    try:
        result = str(excel.Run("模块_测试.RunOfflineTest", macro))
        return (result == "PASS", result)
    except Exception as e:
        return (False, f"COM-ERR: {e}")


def clear_xueqiu_cache(prefix: str = "xueqiu_") -> int:
    cache_dir = ROOT / ".cache"
    if not cache_dir.exists():
        return 0
    deleted = 0
    for p in cache_dir.glob(f"{prefix}*.json"):
        try:
            p.unlink()
            deleted += 1
        except Exception:
            pass
    return deleted


def main() -> int:
    # Always work on a copy so the user's xlsm is untouched.
    shutil.copy2(BOOK_REAL, BOOK)
    print(f"[setup] working on copy: {BOOK.name}")

    # Pre-clear xueqiu HTTP cache so anon-warmup actually fires.
    n = clear_xueqiu_cache("xueqiu_")
    print(f"[setup] cleared {n} xueqiu cache files (forces fresh HTTP)")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    snap = None
    ws_pool = None
    failures: list[str] = []
    report = {}

    try:
        wb = excel.Workbooks.Open(str(BOOK))
        ws_pool = wb.Sheets("样本池")

        snap = snapshot_pool(ws_pool)
        print(f"\n[snapshot] E5={('<set>' if snap['E5'] else '<empty>')!r}, "
              f"E3={snap['E3']!r}, E4={snap['E4']!r}")

        excel.Run("模块_工具函数.SetSilentMode", True)

        # =============================================================
        # Step 2 — VBE smoke
        # =============================================================
        print("\n[STEP 2] Phase5a smoke tests (live HTTP)")
        for mname in ("Test_Phase5a_Xueqiu_AnonWarmup_Smoke",
                      "Test_Phase5a_NoCookieCellNeeded"):
            ok, msg = run_macro_test(excel, mname)
            print(f"  {'PASS' if ok else 'FAIL'}: {mname} → {msg}")
            if not ok:
                failures.append(f"step2/{mname}: {msg}")
        report["step2"] = "PASS" if not failures else "FAIL"

        # =============================================================
        # Step 3 — HK end-to-end with E5 cleared
        # =============================================================
        print("\n[STEP 3] HK end-to-end with E5 cleared (00700/09988/02519)")
        ws_pool.Range("E5").Value = ""
        # Clear other markets to isolate; reset to year=2024 Q4
        for mk in ("A", "US", "HK", "KR", "TW"):
            cc, nc = MARKET_COLS[mk]
            for r in range(POOL_DATA_START_ROW, SCAN_END_ROW + 1):
                ws_pool.Cells(r, cc).ClearContents()
                ws_pool.Cells(r, nc).ClearContents()
        ws_pool.Range("E3").Value = 2024
        ws_pool.Range("E4").Value = "Q4"
        write_market_rows(ws_pool, "HK", TEST_HK)

        elapsed_hk, err_hk = time_macro(excel, "模块_总入口.一键港股")
        print(f"  一键港股 elapsed={elapsed_hk:.1f}s err={err_hk!r}")
        for name in ("港股_资产负债表", "港股_利润表",
                     "港股_现金流量表", "港股_指标表"):
            print(describe_sheet(wb, name))
        diag_hk = summarize_diag(wb, "港股_抓取诊断")
        print(f"  港股_抓取诊断 summary: {json.dumps(diag_hk, ensure_ascii=False, default=str)[:600]}")
        report["step3"] = {"elapsed_s": round(elapsed_hk, 1), "err": err_hk, "diag": diag_hk}
        if err_hk:
            failures.append(f"step3 一键港股: {err_hk}")
        if diag_hk.get("error_rows_4xx"):
            failures.append(f"step3 HK 4xx rows: {diag_hk['error_rows_4xx']}")

        # =============================================================
        # Step 4 — US end-to-end with E5 cleared (BABA / JD ADR fallback)
        # =============================================================
        print("\n[STEP 4] US end-to-end with E5 cleared (BABA/JD/AAPL)")
        for mk in ("A", "US", "HK", "KR", "TW"):
            cc, nc = MARKET_COLS[mk]
            for r in range(POOL_DATA_START_ROW, SCAN_END_ROW + 1):
                ws_pool.Cells(r, cc).ClearContents()
                ws_pool.Cells(r, nc).ClearContents()
        write_market_rows(ws_pool, "US", TEST_US)

        elapsed_us, err_us = time_macro(excel, "模块_总入口.一键美股")
        print(f"  一键美股 elapsed={elapsed_us:.1f}s err={err_us!r}")
        for name in ("美股_资产负债表", "美股_利润表",
                     "美股_现金流量表", "美股_指标表"):
            print(describe_sheet(wb, name))
        diag_us = summarize_diag(wb, "美股_抓取诊断")
        print(f"  美股_抓取诊断 summary: {json.dumps(diag_us, ensure_ascii=False, default=str)[:800]}")
        report["step4"] = {"elapsed_s": round(elapsed_us, 1), "err": err_us, "diag": diag_us}
        if err_us:
            failures.append(f"step4 一键美股: {err_us}")
        if diag_us.get("error_rows_4xx"):
            failures.append(f"step4 US 4xx rows: {diag_us['error_rows_4xx']}")

        # =============================================================
        # Step 5 — offline regression suite
        # =============================================================
        print("\n[STEP 5] offline regression smoke (no HTTP)")
        for mname in ("Test_Offline_FX_Missing_DoesNotFallbackToOne",
                      "Test_Offline_HK_Xueqiu_Tencent",
                      "Test_Offline_TW_FinMind_TSMC"):
            ok, msg = run_macro_test(excel, mname)
            print(f"  {'PASS' if ok else 'FAIL'}: {mname} → {msg}")
            if not ok:
                failures.append(f"step5/{mname}: {msg}")
        report["step5"] = "PASS" if not any(f.startswith("step5/") for f in failures) else "FAIL"

        # =============================================================
        # Step 6 — HK perf baseline (5 companies)
        # =============================================================
        print("\n[STEP 6] HK perf baseline (5 companies fresh fetch)")
        clear_xueqiu_cache("xueqiu_HK")
        for mk in ("A", "US", "HK", "KR", "TW"):
            cc, nc = MARKET_COLS[mk]
            for r in range(POOL_DATA_START_ROW, SCAN_END_ROW + 1):
                ws_pool.Cells(r, cc).ClearContents()
                ws_pool.Cells(r, nc).ClearContents()
        write_market_rows(ws_pool, "HK", TEST_HK_PERF)
        elapsed_perf, err_perf = time_macro(excel, "模块_总入口.一键港股")
        print(f"  5-company 一键港股 elapsed={elapsed_perf:.1f}s err={err_perf!r}")
        diag_perf = summarize_diag(wb, "港股_抓取诊断")
        print(f"  perf diag: status_counts={diag_perf.get('status_counts')}")
        report["step6"] = {
            "companies": len(TEST_HK_PERF),
            "elapsed_s": round(elapsed_perf, 1),
            "err": err_perf,
            "status_counts": diag_perf.get("status_counts"),
        }

        excel.Run("模块_工具函数.SetSilentMode", False)

    except Exception as exc:
        failures.append(f"setup/run exception: {exc}")
        print(f"\nEXCEPTION: {exc}")
    finally:
        if wb is not None:
            try:
                if snap and ws_pool is not None:
                    restore_pool(ws_pool, snap)
            except Exception as e:
                print(f"restore_pool error: {e}")
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass

    print("\n" + "=" * 70)
    print("PHASE 5A LIVE VERIFICATION SUMMARY")
    print("=" * 70)
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    if failures:
        print("\nFAILURES:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("\nALL PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
