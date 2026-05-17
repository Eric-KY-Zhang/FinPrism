"""
Phase 4g Round 2 reviewer probe (Step 3):
  1) Sample pool: row 9-12 = action button rows, row 13 = headers, row 14+ = data
  2) Existing companies preserved at row 14+ (migration worked)
  3) 6 hide-tab buttons placed: 5 market buttons + global button
  4) Toggle macro works: 切换A股tabs hides A股_* sheets, second call unhides
  5) Const POOL_DATA_START_ROW = 11 in source
"""
from __future__ import annotations
import sys
from pathlib import Path
import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
XLSM = ROOT / "上市公司财务数据查询.xlsm"


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def main():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(XLSM))
        try:
            pool = wb.Sheets("样本池")

            # 1. Layout: rows 8/9/10/11
            log("\n=== 1) 样本池 row-by-row layout ===")
            for r in (7, 8, 9, 10, 11, 12, 13):
                row = [pool.Cells(r, c).Value for c in (1, 2, 5, 6, 9, 10, 13, 14)]
                rh = pool.Rows(r).RowHeight
                log(f"  R{r} (h={rh}): A/B/E/F/I/J/M/N = {row}")

            # 2. Const + freeze pane
            log(f"\n  Freeze pane SplitRow = {excel.ActiveWindow.SplitRow if pool.Application.ActiveSheet.Name == '样本池' else 'n/a'}")

            # 3. Buttons: scan for hide-tab Shapes
            log("\n=== 2) Hide-tab buttons on 样本池 ===")
            expected = {
                "BtnHideA": "A11:C12",
                "BtnHideUS": "D11:F12",
                "BtnHideHK": "G11:I12",
                "BtnHideKR": "J11:M12",
                "BtnHideTW": "N11:P12",
                "BtnHideAll": "R6:U7",
                "BtnBuildCrossInd": "R4:U5",
            }
            found_buttons = {}
            for shape in pool.Shapes:
                try:
                    txt = shape.TextFrame2.TextRange.Text
                except Exception:
                    txt = ""
                if shape.Name in expected:
                    found_buttons[shape.Name] = (txt, shape.OnAction, shape.TopLeftCell.Address, shape.BottomRightCell.Address)
            for name, want in expected.items():
                if name in found_buttons:
                    txt, action, tl, br = found_buttons[name]
                    log(f"  + {name} @ {tl}-{br}: '{txt}' -> {action}")
                else:
                    log(f"  !! MISSING: {name} (expected {want})")

            # 4. Toggle test: hide A股, verify, unhide, verify
            log("\n=== 3) Toggle 切换A股tabs behavior ===")
            a_sheets = [s for s in (wb.Sheets(i+1) for i in range(wb.Sheets.Count)) if s.Name.startswith("A股_")]
            initial_a_visibility = {s.Name: s.Visible for s in a_sheets}
            log(f"  A股_* sheets ({len(a_sheets)}):")
            for s in a_sheets:
                log(f"    {s.Name}: visible={s.Visible} (-1=visible, 0=hidden)")

            log("  -> calling 模块_总入口.切换A股tabs (1st time, toggles current state)")
            excel.Run("模块_总入口.切换A股tabs")
            for s in a_sheets:
                log(f"    {s.Name}: visible={s.Visible}")
            changed_once = any(s.Visible != initial_a_visibility.get(s.Name) for s in a_sheets)
            log(f"  A股_* visibility changed after 1st call? {changed_once}")

            log("  -> calling 模块_总入口.切换A股tabs (2nd time, restores current state)")
            excel.Run("模块_总入口.切换A股tabs")
            for s in a_sheets:
                log(f"    {s.Name}: visible={s.Visible}")
            restored = all(s.Visible == initial_a_visibility.get(s.Name) for s in a_sheets)
            log(f"  A股_* visibility restored after 2nd call? {restored}")

            # 5. Global toggle
            log("\n=== 4) Toggle 切换所有分市场tabs (toggles official market sheets + cross indicator) ===")
            all_market_sheets = []
            for prefix in ("A股_", "美股_", "港股_", "韩股_", "台股_"):
                for i in range(wb.Sheets.Count):
                    s = wb.Sheets(i+1)
                    if s.Name.startswith(prefix) and "抓取诊断" not in s.Name:
                        all_market_sheets.append(s)
            all_market_sheets.append(wb.Sheets("跨市场_指标表"))
            log(f"  Official 5-market sheets + cross indicator: {len(all_market_sheets)} total")

            log("  -> 1st call (hide all)")
            excel.Run("模块_总入口.切换所有分市场tabs")
            hidden_count = sum(1 for s in all_market_sheets if s.Visible == 0)
            log(f"  Hidden after 1st call: {hidden_count}/{len(all_market_sheets)}")

            log("  -> 2nd call (unhide all)")
            excel.Run("模块_总入口.切换所有分市场tabs")
            visible_count = sum(1 for s in all_market_sheets if s.Visible == -1)
            log(f"  Visible after 2nd call: {visible_count}/{len(all_market_sheets)}")

            # 6. Const verification
            log("\n=== 5) POOL_DATA_START_ROW const ===")
            try:
                vbproj = wb.VBProject
                comp = vbproj.VBComponents.Item("模块_工具函数")
                src = comp.CodeModule.Lines(1, 50)
                for line_no, line in enumerate(src.splitlines(), 1):
                    if "POOL_DATA_START_ROW" in line:
                        log(f"  L{line_no}: {line.strip()}")
            except Exception as e:
                log(f"  !! source check: {e}")

            # 7. Verify 汇率 / 样本池 NOT in toggle scope
            log("\n=== 6) Shared sheets must NOT be toggled ===")
            shared = ["样本池", "汇率"]
            log("  Calling 切换所有分市场tabs once (hide), check shared sheets stay visible:")
            excel.Run("模块_总入口.切换所有分市场tabs")
            for n in shared:
                try:
                    s = wb.Sheets(n)
                    log(f"    {n}: visible={s.Visible} (should be -1)")
                except Exception as e:
                    log(f"    {n}: !! {e}")
            # restore
            excel.Run("模块_总入口.切换所有分市场tabs")

        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
