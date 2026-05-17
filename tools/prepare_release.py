"""
Build the public release folder from the current FinPrism workbook.

Outputs:
  release/FinPrism-v1.1.xlsm
  release/FinPrism-v1.1-source.xlsm
  release/RELEASE_NOTES-v1.1.md
  release/SHA256SUMS.txt
  release/README.md
  release/LICENSE
"""
from __future__ import annotations

import hashlib
import shutil
import sys
import zipfile
from pathlib import Path

import win32com.client as win32

VERSION = "v1.1"
ROOT = Path(__file__).resolve().parent.parent
SRC_XLSM = ROOT / "上市公司财务数据查询.xlsm"
RELEASE_DIR = ROOT / "release"
PUBLIC_XLSM = RELEASE_DIR / f"FinPrism-{VERSION}.xlsm"
SOURCE_XLSM = RELEASE_DIR / f"FinPrism-{VERSION}-source.xlsm"
README = ROOT / "README.md"
LICENSE = ROOT / "LICENSE"
NOTES = RELEASE_DIR / f"RELEASE_NOTES-{VERSION}.md"
CHECKSUMS = RELEASE_DIR / "SHA256SUMS.txt"

MARKETS = ("A股", "美股", "港股", "韩股", "台股")
KINDS = ("资产负债表", "利润表", "现金流量表", "指标表")
DIAGNOSTICS = ("美股_抓取诊断", "港股_抓取诊断", "韩股_抓取诊断", "台股_抓取诊断")


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def assert_inside_root(path: Path) -> None:
    root = ROOT.resolve()
    resolved = path.resolve()
    if resolved != root and root not in resolved.parents:
        raise RuntimeError(f"Refusing to operate outside project root: {resolved}")


def reset_release_dir() -> None:
    assert_inside_root(RELEASE_DIR)
    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    RELEASE_DIR.mkdir(parents=True, exist_ok=True)


def strip_webextensions(path: Path) -> None:
    tmp = path.with_suffix(".tmp")
    removed = 0
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                lowered = item.filename.lower()
                if "webextension" in lowered:
                    removed += 1
                    continue
                zout.writestr(item, zin.read(item.filename))
    tmp.replace(path)
    log(f"  + stripped {removed} webextension entries from {path.name}")


def clear_range_safely(ws, addr: str) -> None:
    try:
        rng = ws.Range(addr)
        merged = bool(rng.MergeCells)
        if merged:
            rng.UnMerge()
        rng.ClearContents()
        if merged:
            rng.Merge()
    except Exception:
        pass


def reset_wide_sheet(ws, is_indicator: bool) -> None:
    try:
        ws.Cells.Clear()
    except Exception:
        pass
    ws.Cells.Font.Name = "微软雅黑"
    ws.Cells.Font.Size = 10
    ws.Cells.HorizontalAlignment = -4131
    ws.Cells.VerticalAlignment = -4108
    ws.Range("A1").Value = "指标类型" if is_indicator else "大类"
    ws.Range("B1").Value = "指标名称"
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Font.Color = 0xFFFFFF
    ws.Range("A1:B1").Interior.Color = 0xC47244
    ws.Columns("A").ColumnWidth = 18 if is_indicator else 30
    ws.Columns("B").ColumnWidth = 28 if is_indicator else 40
    if is_indicator:
        ws.Range("C1").Value = "英文指标名"
        ws.Range("C1").Font.Bold = True
        ws.Range("C1").Font.Color = 0xFFFFFF
        ws.Range("C1").Interior.Color = 0xC47244
        ws.Columns("C").ColumnWidth = 34
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 20


def clean_workbook(path: Path, *, public_release: bool) -> None:
    log(f"Cleaning {path.name}")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(path))
        try:
            try:
                ws = wb.Sheets("使用说明")
                ws.Delete()
                log("  + removed legacy 使用说明 sheet")
            except Exception:
                pass

            pool = wb.Sheets("样本池")
            for addr in ("B5", "B7", "B8", "O5", "O6", "E5:P5", "W2:Z8"):
                clear_range_safely(pool, addr)
            pool.Range("E6").Value = "原币"
            for addr in ("E5", "E6"):
                try:
                    if pool.Range(addr).Comment is not None:
                        pool.Range(addr).Comment.Delete()
                except Exception:
                    pass

            for diag_name in DIAGNOSTICS:
                try:
                    ws = wb.Sheets(diag_name)
                    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                    if last_row >= 3:
                        ws.Range(ws.Cells(3, 1), ws.Cells(last_row, 17)).Clear()
                    ws.Visible = 0
                except Exception as exc:
                    log(f"  ! diagnostic cleanup skipped for {diag_name}: {exc}")

            for market in MARKETS:
                for kind in KINDS:
                    name = f"{market}_{kind}"
                    try:
                        reset_wide_sheet(wb.Sheets(name), kind == "指标表")
                    except Exception as exc:
                        log(f"  ! report cleanup skipped for {name}: {exc}")

            try:
                reset_wide_sheet(wb.Sheets("跨市场_指标表"), True)
            except Exception as exc:
                log(f"  ! cross indicator cleanup skipped: {exc}")

            try:
                fx = wb.Sheets("汇率")
                last_row = fx.Cells(fx.Rows.Count, 1).End(-4162).Row
                if last_row >= 2:
                    fx.Range(fx.Cells(2, 1), fx.Cells(last_row, 10)).Clear()
                fx.Visible = -1
            except Exception as exc:
                log(f"  ! fx cleanup skipped: {exc}")

            for market in MARKETS:
                for kind in KINDS:
                    try:
                        wb.Sheets(f"{market}_{kind}").Visible = 0 if public_release else -1
                    except Exception:
                        pass
            wb.Sheets("样本池").Visible = -1
            wb.Sheets("跨市场_指标表").Visible = -1
            wb.Sheets("汇率").Visible = -1
            wb.Sheets("样本池").Activate()

            props = wb.BuiltinDocumentProperties
            try:
                props.Item("Author").Value = "Eric Zhang"
                props.Item("Last Author").Value = "Eric Zhang"
                props.Item("Company").Value = ""
                props.Item("Manager").Value = ""
                props.Item("Title").Value = f"FinPrism {VERSION}"
                props.Item("Comments").Value = (
                    f"{VERSION} release - FinPrism - contact: 214978902@qq.com"
                )
            except Exception:
                pass

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    strip_webextensions(path)


def write_release_notes() -> None:
    NOTES.write_text(
        f"""# FinPrism {VERSION} Release Notes

Date: 2026-05-17

## Highlights

- Added Taiwan market support: sample-pool columns, 4 Taiwan report sheets, diagnostics, and cross-market indicator integration.
- Added Phase 5a Xueqiu anonymous session warmup: users no longer need to paste or maintain a manual Xueqiu cookie in the sample pool.
- Moved workbook usage instructions and FX explanations into README; Excel no longer has a separate 使用说明 sheet or FX instruction panel.
- Simplified 汇率 sheet to a pure data table: report period plus USDCNY/HKDCNY/KRWCNY/TWDCNY EOP/AVG columns.
- Preserved data-accuracy behavior: missing FX remains blank and diagnostic-driven, never fallback-to-1.
- Preserved the clean v1.1 workbook layout: sample-pool row 5 is intentionally blank, row 14+ remains the user input area, and 5 markets feed the cross-market indicator sheet.

## Files

- `FinPrism-{VERSION}.xlsm`: clean end-user workbook.
- `FinPrism-{VERSION}-source.xlsm`: clean workbook with VBA project/source modules retained.
- `README.md`: user guide.
- `LICENSE`: MIT license.
- `SHA256SUMS.txt`: checksums for release files.
""",
        encoding="utf-8",
    )


def write_checksums(paths: list[Path]) -> None:
    lines = []
    for path in paths:
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        lines.append(f"{digest}  {path.name}")
    CHECKSUMS.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    if not SRC_XLSM.exists():
        raise SystemExit(f"Source workbook missing: {SRC_XLSM}")

    reset_release_dir()
    shutil.copy2(SRC_XLSM, PUBLIC_XLSM)
    shutil.copy2(SRC_XLSM, SOURCE_XLSM)
    shutil.copy2(README, RELEASE_DIR / "README.md")
    shutil.copy2(LICENSE, RELEASE_DIR / "LICENSE")

    clean_workbook(PUBLIC_XLSM, public_release=True)
    clean_workbook(SOURCE_XLSM, public_release=False)
    write_release_notes()
    write_checksums([PUBLIC_XLSM, SOURCE_XLSM, RELEASE_DIR / "README.md", RELEASE_DIR / "LICENSE", NOTES])

    log("\nRelease folder ready:")
    for path in sorted(RELEASE_DIR.iterdir()):
        log(f"  + {path.name} ({path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
