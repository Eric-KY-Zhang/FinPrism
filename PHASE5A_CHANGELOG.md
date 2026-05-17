# Phase 5a — 雪球 cookie 自动化 (匿名 warmup session)

> Status: **live verified — anon warmup works end-to-end without E5 cookie**
> Author: Claude (Opus 4.7) under Generator role
> Date: 2026-05-17

## 1. 改动摘要

把项目里已经验证可用的 `FetchViaPowerShell(url, warmupFirst:=True)` 路径
(原本只服务汇率 K 线) 扩到全部雪球 API 调用,消除 `样本池!E5` 手动维护
`xq_a_token` 的痛点。E5 留空也能正常抓港股 / 中概美股 fallback。

实施代码改动 3 处 + 文档脚本 1 个:

| # | 文件 | 改动 |
|---|---|---|
| 1 | `modules/模块_抓汇率.bas` | `FetchViaPowerShell` `Private` → `Public` (函数体不变) |
| 2 | `modules/模块_工具函数.bas` | `XueqiuHttpGet` 函数体替换为 `FetchViaPowerShell(strUrl, True)`;签名不变,`strCookie` 参数保留以兼容 20+ 调用点 |
| 2b | `modules/模块_抓港股财报.bas`<br>`modules/模块_抓美股财报.bas` | **Hotfix (live verify 发现)**: 删除 `If Len(strCookie)=0 Then Err.Raise` 早退保护。原 guard 让 E5 留空时整个 HK / 雪球 US fallback 路径直接抛错,导致改动 1+2 形同虚设。删除后下游 `XueqiuHttpGet` 自己 warmup,与 spec 目标一致。 |
| 3 | `modules/模块_测试.bas` | 文件末尾追加 `Test_Phase5a_Xueqiu_AnonWarmup_Smoke` 和 `Test_Phase5a_NoCookieCellNeeded` 两个 live smoke 用例 |
| 4 | `tools/install_modules.py` | `layout_sample_pool` 的 `config_rows` 删除 `(5, "雪球 Cookie", ...)` 一项。重装后 row 5 留空(仍在面板边框内,看起来像一行 spacer)。**不动 ReadDisplayCurrency 对 E6 的引用 (30+ 处)**。原 spec 里的 `scripts/phase5a_update_doc_cells.py` openpyxl 脚本已删除——见下 §3a。 |

未改:
- `ReadXueqiuCookie()` 保持原状,E5 即便用户继续填 token 也不会报错。
- `CachedXueqiuHttpGet` 缓存逻辑、TTL、诊断 17 列 schema 完全不动。
- `XueqiuHttpGet` 函数签名 (`strUrl, strCookie`) 不动,所有上游调用 0 改动。
- EDGAR / StockAnalysis / FinMind / AkShare / A 股新浪等路径 0 改动。

## 2. 已验证的技术事实 (实施依据,不再重测)

- `GET https://xueqiu.com/hq` 匿名响应 `Set-Cookie` 下发: `xq_a_token`、
  `xqat`、`xq_r_token`、`xq_id_token`、`cookiesu`、`u`、`acw_tc`。
- 带这些 cookie 调 `https://stock.xueqiu.com/v5/stock/finance/hk/balance.json`
  `?symbol=00700&type=Q4&is_detail=true&count=5` 返回 `HTTP 200`,
  `error_code=0`,`list` 长度 5 (腾讯港股 5 期资产负债表)。
- `FetchViaPowerShell` 在 [`modules/模块_抓汇率.bas:61-120`](modules/%E6%A8%A1%E5%9D%97_%E6%8A%93%E6%B1%87%E7%8E%87.bas#L61)
  已用 `HttpClient + UseCookies=$true + AutomaticDecompression(GZip,Deflate)`
  正确实现 warmup → session token 提取 → 目标 URL 拉取的完整流程。

## 3a. ⚠️ openpyxl 杀按钮 (live-fire 踩坑实录)

最初的 spec Change 4 用 `openpyxl.load_workbook(..., keep_vba=True)` +
`wb.save(...)` 来改 `样本池!A5` 文案。**这是错的**:

- `keep_vba=True` 只保留 vbaProject.bin (VBA 代码)。
- **不保留** worksheet 上的 `Shape` / form-control 按钮 / OnAction 绑定。
- 跑 `wb.save(book)` 后,所有 `_apply_card_brand` 创建的圆角矩形按钮 + 它们的
  `OnAction = "模块_总入口.一键港股"` 等绑定**全部消失**。用户打开 xlsm
  会看到一个静态布局,点哪都没反应——但 VBA 模块还在,VBE 立即窗口
  调 `一键港股` 仍能正常跑。

修复路径:**永远不要用 openpyxl 写 xlsm 里有按钮的工作表**。同步操作
统一走 `tools/install_modules.py` 的 Excel COM 路径(它已经处理按钮重建)。

具体本轮做的:

1. **删除** `scripts/phase5a_update_doc_cells.py` 和 `tools/phase5a_inspect_doc_cells.py`
   两个 openpyxl 脚本。
2. **改在** `install_modules.py` 内移除 row 5 `(5, "雪球 Cookie", ...)`,
   重装后 row 5 留空、A5 不再出现"已弃用"标签。
3. **重装 + COM 路径** 重新建出 15 个按钮(BtnRunAll / BtnRunA…TW /
   BtnHide*  / BtnBuildCrossInd / BtnClearAllData / BtnClearCache 等),
   全部 OnAction 验证通过(`tools/probe_button_bindings.py` 输出 15/15 OK)。

## 3. 已知风险

1. **PowerShell 启动开销 ~150-300ms / 次**。10 家样本 × 4 张表 = 40 次
   雪球调用,全市场抓数预计多 30-60 秒。下一 phase (5b) 可加 session
   token 进程内缓存或 PowerShell 长连接复用来抵消。
2. **匿名 session 限流可能比登录态严**。出现 429 / 403 频繁需调大
   `HttpGetWithRetry` 节流间隔,或在 `RunCachedHttpGet` 走更长 TTL。
3. **不影响 FX 模块** (`模块_抓汇率`):本来就在用 `FetchViaPowerShell`,
   签名 `Private → Public` 是纯可见性放宽,行为零变化。
4. **不影响其他数据源**:EDGAR (SEC)、StockAnalysis、FinMind、AkShare、
   新浪 A 股、stockanalysis_KR 都走各自独立 HTTP 封装,本轮零改动。
5. **匿名 session 拿不到登录态独享 API**。如果将来发现某个端点需要
   登录 cookie,需重新打开 E5 手动 token 通道 (本轮保留了 ReadXueqiuCookie
   函数,改回去成本低)。

## 4. 回滚指引

最快回滚:

```powershell
git revert <phase-5a-commit>
py tools/install_modules.py
```

或最小化手工回滚(只回 VBA,不动文档脚本):

1. `modules/模块_抓汇率.bas` 把 `Public Function FetchViaPowerShell` 改回
   `Private Function FetchViaPowerShell`。
2. `modules/模块_工具函数.bas` 的 `XueqiuHttpGet` 函数体恢复为 Phase 4 前
   的 WinHttp + `Accept-Encoding: identity` + `If Len(strCookie) > 0 Then
   .SetRequestHeader "Cookie", strCookie` 版本;参考 Git history 同一函数
   的上一次提交。
3. 运行 `py tools/install_modules.py` 重新注入 VBA。
4. 提醒用户重新在 E5 粘贴 `xq_a_token`。

## 5. 不变量
本轮不动:
- 4 市场 + 台股的 5 张分市场表 + 跨市场指标表保留。
- 诊断 sheet 17 列 schema、AppStateGuard、cache 分源 TTL 不变。
- RMB toggle 行为不变;汇率缺失仍走 `FX_MISSING` 诊断,不 fallback 到 1。
- 样本池 Row 14+ 用户录入数据未触碰。

## 6. Live verification 结果(2026-05-17,本机执行)

`py tools/phase5a_live_verify.py` 在 `上市公司财务数据查询.xlsm` 的副本上跑完
6 个步骤。E5 cookie 临时清空,完成后还原。结果:

| 步骤 | 内容 | 结果 |
|---|---|---|
| 2 | `Test_Phase5a_Xueqiu_AnonWarmup_Smoke` + `Test_Phase5a_NoCookieCellNeeded` | **PASS / PASS**(live HTTP) |
| 3 | `一键港股` E5 留空,3 家(00700/09988/02519) | **PASS** 30.7s,114 OK_XUEQIU + 43 MISSING(09988 因 3 月财年口径),HTTP 全 200,**零 401/403/429** |
| 4 | `一键美股` E5 留空,3 家(BABA/JD/AAPL) | **PASS** 67.5s,BABA/JD 走 `Xueqiu` fallback,AAPL 走 `EDGAR us-gaap` 主路径;187 OK_XUEQIU + 101 EDGAR OK + 60 MISSING + 8 RECOMMEND_FUZZY,**零 4xx** |
| 5 | 3 个离线回归(`FX_Missing`/`HK_Xueqiu_Tencent`/`TW_FinMind_TSMC`) | **PASS / PASS / PASS** |
| 6 | HK perf baseline,5 家(00700/00939/00941/00388/01024)新缓存 | **50.7s**;188 OK_XUEQIU + 22 MISSING,zero 4xx |

性能基线参考(Phase 5b 优化目标):5 家港股全 fresh fetch 在本机约
**50.7s** ≈ 2.5s / 公司 / statement(20 个 HTTP 调用,每个 ~2.5s 含 PS 启动)。

## 7. 后续维护提醒

1. **不要再用 openpyxl 改 .xlsm**。Phase 5a 初稿曾经写了
   `scripts/phase5a_update_doc_cells.py` 走 openpyxl 改 `样本池!A5` 文案,
   结果触发了 §3a 的按钮被清问题。该脚本和 `tools/phase5a_inspect_doc_cells.py`
   已经从仓库删除。后续如果要改 workbook 里的 cell 文案,统一改
   `tools/install_modules.py` 的 `layout_sample_pool` 函数,然后跑
   `py tools/install_modules.py` 让 Excel COM 重新铺布局并重建按钮。
2. **样本池 row 5 已从 install 布局里移除**(`config_rows` 不再含 `(5, "雪球 Cookie", ...)`),
   重装 workbook 后 row 5 自然留空,A5 不会再出现"已弃用"标签。
3. **E5 cookie 清空**已经做过。如果分发 release,统一走
   `tools/prepare_release.py`(已经处理 E5/cache/诊断历史的清理,
   底层调 `模块_工具函数.bas` 里的 `CleanReleaseWorkbook` 宏)。
4. **回滚路径**见 §4。如果未来要把雪球路径切回登录态 cookie 走 WinHttp,
   恢复 `XueqiuHttpGet` 函数体到 Phase 5a 前版本并重新启用 HK / US fetcher
   里的 `If Len(strCookie)=0 Then Err.Raise` 早退保护即可。
