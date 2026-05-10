"""
Reviewer spot-check: open the main-repo xlsm, dump:
  - 样本池 B6 toggle current value
  - 汇率 sheet rows
  - 4 market sheets: A1 comment text + R1 first 3 codes
  - 美股_抓取诊断 / 港股_抓取诊断 / 韩股_抓取诊断: header + first 3 rows incl K11 FX_Rate
  - presence of 模块_测试 module + 模块_抓汇率 module
"""
from __future__ import annotations
import sys
from pathlib import Path
import win32com.client as win32

XLSM = Path(r"E:\Claude+CODEX Project\FS Capture\VBA Captor\上市公司财务数据查询.xlsm")


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def safe_comment(ws):
    try:
        c = ws.Range("A1").Comment
        if c is None:
            return None
        return c.Text()
    except Exception as e:
        return f"<err: {e}>"


def main():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(XLSM))
        try:
            # Pool B6
            try:
                pool = wb.Sheets("样本池")
                log(f"\n=== 样本池 B6 = {pool.Range('B6').Value!r}")
                log(f"=== 样本池 A2 (year) = {pool.Range('A2').Value!r}")
                log(f"=== 样本池 A4 (qtr) = {pool.Range('A4').Value!r}")
                # samples
                for col_pair, market in [(("A","B"),"A"), (("E","F"),"US"), (("I","J"),"HK"), (("M","N"),"KR")]:
                    log(f"  {market} samples (first 3):")
                    for r in range(10, 13):
                        code = pool.Range(f"{col_pair[0]}{r}").Value
                        name = pool.Range(f"{col_pair[1]}{r}").Value
                        log(f"    R{r}: {code!r} / {name!r}")
            except Exception as e:
                log(f"!! pool read: {e}")

            # 汇率 sheet
            try:
                fx = wb.Sheets("汇率")
                log(f"\n=== 汇率 sheet:")
                last = fx.Cells(fx.Rows.Count, 1).End(-4162).Row
                for r in range(1, max(last, 1) + 1):
                    row = [fx.Cells(r, c).Value for c in range(1, 9)]
                    log(f"  R{r}: {row}")
            except Exception as e:
                log(f"!! fx sheet read: {e}")

            # 4 market sheets - A1 comment + R1 first 3 cells
            for market in ["A股", "美股", "港股", "韩股"]:
                for kind in ["资产负债表", "利润表"]:
                    name = f"{market}_{kind}"
                    try:
                        ws = wb.Sheets(name)
                        comm = safe_comment(ws)
                        log(f"\n=== {name}")
                        log(f"  A1 comment: {comm!r}")
                        log(f"  R1 first 6 cols: {[ws.Cells(1, c).Value for c in range(1, 7)]}")
                    except Exception as e:
                        log(f"!! {name}: {e}")

            # diagnostic sheets
            for diag in ["美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"]:
                try:
                    ws = wb.Sheets(diag)
                    # unhide just for inspection
                    ws.Visible = -1
                    log(f"\n=== {diag}")
                    log(f"  R2 headers (11 cols): {[ws.Cells(2, c).Value for c in range(1, 12)]}")
                    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                    log(f"  data rows: 3..{last}")
                    for r in range(3, min(last + 1, 6)):
                        log(f"    R{r}: {[ws.Cells(r, c).Value for c in range(1, 12)]}")
                    ws.Visible = 0
                except Exception as e:
                    log(f"!! {diag}: {e}")

            # VB modules
            try:
                vbproj = wb.VBProject
                names = sorted(vbproj.VBComponents.Item(i).Name for i in range(1, vbproj.VBComponents.Count + 1))
                log(f"\n=== VBComponents ({len(names)}):")
                for n in names:
                    flag = ""
                    if "测试" in n: flag = " <-- TEST MODULE"
                    if "抓汇率" in n: flag = " <-- FX MODULE"
                    log(f"  {n}{flag}")
            except Exception as e:
                log(f"!! vbproject: {e}")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
