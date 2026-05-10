"""
Phase 4g Round 1 reviewer probe:
  1) Cross-market sheet exists, R1 headers, freeze panes, comment
  2) Run BuildCrossMarketIndicatorSheet on empty state -> A3 should show '(no data)' msg
  3) After running it again, R1/R2 layout matches collected source sheets
  4) Diagnostic sheets: K2 == 'FX_Rate', col K width ~ 12
  5) Tab order: cross-market sheet sits between 韩股_抓取诊断 and 汇率
"""
from __future__ import annotations
import sys
from pathlib import Path
import win32com.client as win32

XLSM = Path(r"E:\Claude+CODEX Project\FS Capture\VBA Captor\上市公司财务数据查询.xlsm")


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def main():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(XLSM))
        try:
            # 1. Cross-market sheet existence + structure
            log("\n=== 1) 跨市场_指标表 baseline structure ===")
            try:
                ws = wb.Sheets("跨市场_指标表")
                log(f"  R1: {[ws.Cells(1, c).Value for c in range(1, 6)]}")
                log(f"  Col widths A/B/C: {ws.Columns('A').ColumnWidth}/{ws.Columns('B').ColumnWidth}/{ws.Columns('C').ColumnWidth}")
                log(f"  Row 1 / 2 heights: {ws.Rows(1).RowHeight}/{ws.Rows(2).RowHeight}")
            except Exception as e:
                log(f"  !! sheet missing or broken: {e}")
                return

            # 2. Run BuildCrossMarketIndicatorSheet on potentially empty / partial state
            log("\n=== 2) call BuildCrossMarketIndicatorSheet ===")
            try:
                excel.Run("模块_工具函数.BuildCrossMarketIndicatorSheet")
                log("  + macro returned without raising")
            except Exception as e:
                log(f"  !! macro raised: {e}")

            # 3. Inspect post-build state
            log("\n=== 3) post-build state ===")
            ws = wb.Sheets("跨市场_指标表")
            r1 = [ws.Cells(1, c).Value for c in range(1, 12)]
            r2 = [ws.Cells(2, c).Value for c in range(1, 12)]
            log(f"  R1 (cols 1-11): {r1}")
            log(f"  R2 (cols 1-11): {r2}")
            log(f"  A3 (empty msg or first indicator label): {ws.Cells(3, 1).Value!r}")
            # First 3 formula cells D3, E3, F3 if any
            for col in range(4, 8):
                cell = ws.Cells(3, col)
                log(f"  R3C{col}: value={cell.Value!r}  formula={cell.Formula!r}")

            # 4. Diagnostic sheets self-heal verification
            log("\n=== 4) diagnostic sheets header self-heal ===")
            for diag in ["美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断"]:
                try:
                    ws_d = wb.Sheets(diag)
                    ws_d.Visible = -1  # unhide for inspection
                    headers = [ws_d.Cells(2, c).Value for c in range(1, 12)]
                    k_width = ws_d.Columns("K").ColumnWidth
                    title_a1 = ws_d.Cells(1, 1).Value
                    log(f"  {diag}: R2 headers (1-11) = {headers}")
                    log(f"    K col width = {k_width}, A1 title = {title_a1!r}")
                    log(f"    K2 == 'FX_Rate' ? {headers[10] == 'FX_Rate'}")
                    ws_d.Visible = 0
                except Exception as e:
                    log(f"  !! {diag}: {e}")

            # 5. Tab order
            log("\n=== 5) tab order (cross-market between diag and FX) ===")
            order = [wb.Sheets(i+1).Name for i in range(wb.Sheets.Count)]
            log(f"  全部 {len(order)} sheets:")
            for n in order:
                log(f"    {n}")

            # Find positions
            try:
                pos_kr_diag = order.index("韩股_抓取诊断")
                pos_cross = order.index("跨市场_指标表")
                pos_fx = order.index("汇率")
                log(f"\n  韩股_抓取诊断 idx = {pos_kr_diag}")
                log(f"  跨市场_指标表  idx = {pos_cross}  (expect just after 韩股_抓取诊断 group)")
                log(f"  汇率           idx = {pos_fx}   (expect after 跨市场_指标表)")
                log(f"  Order OK ? {pos_kr_diag < pos_cross < pos_fx}")
            except ValueError as e:
                log(f"  !! ordering check missing sheet: {e}")

            # 6. Button verification
            log("\n=== 6) Button '合并跨市场指标表' on 样本池 ===")
            try:
                pool = wb.Sheets("样本池")
                found = False
                for shape in pool.Shapes:
                    txt = ""
                    try:
                        txt = shape.TextFrame2.TextRange.Text
                    except Exception:
                        pass
                    if "合并跨市场指标表" in txt or shape.Name.startswith("BtnBuildCrossInd"):
                        found = True
                        try:
                            macro = shape.OnAction
                        except Exception:
                            macro = "<n/a>"
                        log(f"  shape '{shape.Name}': txt={txt!r}, OnAction={macro}")
                if not found:
                    log("  !! 合并跨市场指标表 button not found")
            except Exception as e:
                log(f"  !! button scan: {e}")

            # 7. 一键全抓 macro auto-refresh hook
            log("\n=== 7) 一键全抓 source contains BuildCrossMarketIndicatorSheet call ===")
            try:
                vbproj = wb.VBProject
                comp = vbproj.VBComponents.Item("模块_总入口")
                src = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                hit = "BuildCrossMarketIndicatorSheet" in src
                log(f"  source contains call ? {hit}")
                # Find the line
                if hit:
                    for line_no, line in enumerate(src.splitlines(), 1):
                        if "BuildCrossMarketIndicatorSheet" in line:
                            log(f"    L{line_no}: {line.strip()}")
            except Exception as e:
                log(f"  !! source check: {e}")

        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
