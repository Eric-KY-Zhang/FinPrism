"""Inspect every shape on 样本池: name, OnAction, whether the target Sub
actually exists in any VBComponent. Diagnoses 'buttons don't react'."""
from __future__ import annotations

from pathlib import Path
import re

import win32com.client as win32

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "上市公司财务数据查询.xlsm"


def main():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(BOOK))
        try:
            ws = wb.Sheets("样本池")
            print(f"=== Shapes on 样本池 ({ws.Shapes.Count} total) ===")
            shapes_info = []
            for sh in ws.Shapes:
                try:
                    name = sh.Name
                except Exception:
                    name = "?"
                try:
                    action = sh.OnAction
                except Exception:
                    action = "?"
                try:
                    cap = sh.TextFrame2.TextRange.Text[:40]
                except Exception:
                    cap = ""
                shapes_info.append((name, action, cap))
                print(f"  {name:20s} OnAction={action!r:60s} caption={cap!r}")

            print(f"\n=== VBA components ===")
            try:
                vb = wb.VBProject
            except Exception as e:
                print(f"  CANNOT ACCESS VBProject: {e}")
                return
            comps = {}
            for c in vb.VBComponents:
                comps[c.Name] = c.CodeModule
            for n in sorted(comps.keys()):
                print(f"  {n}  ({comps[n].CountOfLines} lines)")

            print(f"\n=== Cross-check button OnAction → Sub existence ===")
            for name, action, cap in shapes_info:
                if not action or action in ("?", "0"):
                    continue
                # e.g. "模块_总入口.一键港股"
                if "." in action:
                    modname, subname = action.split(".", 1)
                else:
                    modname, subname = None, action
                if modname is None:
                    # search all modules
                    found_in = [m for m, cm in comps.items() if subname in _module_subs(cm)]
                    print(f"  {name:20s} → {action}: found in {found_in or 'NONE'}")
                    continue
                if modname not in comps:
                    print(f"  {name:20s} → {action}: MODULE MISSING ({modname})")
                    continue
                subs = _module_subs(comps[modname])
                if subname in subs:
                    print(f"  {name:20s} → {action}: OK")
                else:
                    print(f"  {name:20s} → {action}: SUB MISSING (module has: {sorted(subs)[:5]} ...)")
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def _module_subs(code_module) -> set[str]:
    try:
        lines = code_module.Lines(1, code_module.CountOfLines)
    except Exception:
        return set()
    subs = set()
    for m in re.finditer(r"^\s*(?:Public\s+|Private\s+)?Sub\s+([^\s(]+)", lines, re.MULTILINE):
        subs.add(m.group(1))
    for m in re.finditer(r"^\s*(?:Public\s+|Private\s+)?Function\s+([^\s(]+)", lines, re.MULTILINE):
        subs.add(m.group(1))
    return subs


if __name__ == "__main__":
    main()
