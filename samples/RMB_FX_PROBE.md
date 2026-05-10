# RMB FX Probe — Xueqiu quote + kline + Aoji confirmation

Date: 2026-05-03
Cookie source: `XUEQIU_COOKIE` env var (NOT set during this probe — all data
collected without auth cookie)
Endpoint families:
- Realtime quote: `https://stock.xueqiu.com/v5/stock/realtime/quotec.json?symbol=USDCNY.FX,HKDCNY.FX,KRWCNY.FX`
- Kline: `https://stock.xueqiu.com/v5/stock/chart/kline.json?symbol=USDCNY.FX&begin=<ms>&period=day&type=before&count=-N`
- HK quote / finance: `https://stock.xueqiu.com/v5/stock/quote.json?symbol=02519&extend=detail` and `.../finance/hk/balance.json?symbol=02519`

## TL;DR

1. **Quote endpoint通**: 不需登录 cookie。但要 (a) 把 plan-spec 的 `USDCNY,HKDCNY,KRWCNY` 改成 `USDCNY.FX,HKDCNY.FX,KRWCNY.FX`, (b) 把 `Accept-Encoding: identity` 改成 `gzip, deflate`, (c) 先 GET 一次 `https://xueqiu.com/hq` 拿 `xq_a_token` cookie。否则 200 + 空 body。USDCNY.FX = 6.8282, HKDCNY.FX = 0.87156, KRWCNY.FX = 0.00464。
2. **K 线 6 年回溯通**: USDCNY.FX 在 `count=-2000` 拿到 2018-08-31 起 2000 个交易日, 远早于 2020-01-01。`count` 硬上限 ~2147 条 (数据起 2018-02-07)。HKDCNY.FX / KRWCNY.FX 同 365d 均 OK。
3. **傲基股份 = 02519.HK 傲基股份, 港股 (exchange=HK, type=30), 报告币种 CNY (人民币), 年报截止 12-31, 已有 6 期年报**。03031 (中瓷电子) / 09927 (赛力斯) 均与傲基无关。

---

## 1A — Realtime Quote FX

### Endpoint

```
GET https://stock.xueqiu.com/v5/stock/realtime/quotec.json?symbol=USDCNY.FX,HKDCNY.FX,KRWCNY.FX
```

### 关键修正 (相对 plan §1A)

| 项 | plan 给的 | 实测必须改成 | 不改的后果 |
|---|---|---|---|
| symbol | `USDCNY,HKDCNY,KRWCNY` | `USDCNY.FX,HKDCNY.FX,KRWCNY.FX` | HTTP 200, `data: []` (空数组, 不报错) |
| `Accept-Encoding` | `identity` | `gzip, deflate` | HTTP 200, body 长度 0, `Content-Length: 0` |
| Session warmup | 无 | 先 GET `https://xueqiu.com/hq` | 同上, 空 body — 缺 `xq_a_token` cookie |
| 响应 schema | `parsed.data.items[]` | `parsed.data[]` (直接是数组) | KeyError |

### 实测结果 (无 cookie)

| symbol | current | last_close | percent | timestamp ms | timestamp CST |
|---|---:|---:|---:|---:|---|
| USDCNY.FX | 6.8282 | 6.8282 | 0.0% | 1777582801977 | 2026-05-01 05:00:01 |
| HKDCNY.FX | 0.87156 | 0.8718 | -0.03% | 1777669192206 | 2026-05-02 04:59:52 |
| KRWCNY.FX | 0.00464 | 0.00463 | +0.22% | 1777669191575 | 2026-05-02 04:59:51 |

`error_code = 0`, `data.length = 3`, all 3 symbols returned.

### Cookie 必需性

- **不必需**。无 `XUEQIU_COOKIE` env 时, 仅靠 session warmup 拿到的访客 cookie (`xq_a_token`) 就足够拉到完整 3 行数据。
- 控制对照: `plan_call`(无 warmup, identity 编码, plain 符号) 状态 200, body 长 0 — 三个问题任一存在都会失败。

### 文件

- `samples/xueqiu_quote_currency.json` — 包含 plan_call (失败 control), warmup, no_cookie_working (生产), symbol_variants (USDCNY / USDCNY.FX / CNY 三种对照)
- `tools/probe_1a_quote_fx.py` — 可重跑

---

## 1B — Kline Historical FX (USDCNY / HKDCNY / KRWCNY)

### Endpoint

```
GET https://stock.xueqiu.com/v5/stock/chart/kline.json
    ?symbol=USDCNY.FX&begin=<ms>&period=day&type=before&count=-N
```

### Column index (实测)

```
0: timestamp     (ms; UTC 04:00 = CST 12:00 — 见时区段)
1: volume        (FX 永远 null)
2: open
3: high
4: low
5: close         <-- 期间均值算法用此列
6: chg           (与上一交易日 close 之差)
7: percent       (chg %)
8: turnoverrate  (FX 永远 null)
9: amount        (FX 永远 null)
10: volume_post
11: amount_post
```

完整 `column` 数组: `["timestamp","volume","open","high","low","close","chg","percent","turnoverrate","amount","volume_post","amount_post"]` (12 列)。

### 时区结论 — `begin` 解读

实测 `begin=1704038400000` (= `2024-01-01T00:00:00+08:00 CST` = `2023-12-31T16:00:00 UTC`):
- 返回的 `last_item[0]` ts = `1703826000000` = `2023-12-29T13:00:00 CST` = `2023-12-29T05:00:00 UTC` (上一个交易日 CST 13:00)。
- 说明 `begin` 是**含义为 CST 端点的毫秒时间戳**, 接口语义是"取早于这个 begin 的 N 个 day bar"。如果传"2024-01-01 UTC"的 ms (`1704067200000`), 拉到的最新一条会是 2024-01-01 当天的 day bar。
- **每条 day bar 的 ts 一律是 04:00:00 UTC = 12:00:00 CST 当日中午**, 不是 00:00。Step 2/3 切片时按 `dt.date()`(取 UTC 04:00 的日期, 与 CST 当日相同)对齐即可。
- 因为同一历史日 UTC-day 与 CST-day 一致 (差 8h, day bar 落在中午), 时区差对"按日切片"无影响。

### 时间范围 / count 上限

| 测试 | count 参数 | 实际返回 | 最早 day bar (CST) |
|---|---:|---:|---|
| USDCNY.FX_365d (begin=2024-01-01 CST) | -365 | 365 | 2022-08-08 |
| USDCNY.FX_long (begin=2026-05-03 CST) | -1500 | 1500 | 2020-07-31 |
| count_cap_probe | -2000 | 2000 | 2018-08-31 |
| count_cap_probe | -2500 | **2147** (截断) | 2018-02-07 |
| count_cap_probe | -3000 | **2147** (截断) | 2018-02-07 |

**结论**: 单次 kline 请求硬上限 ≈ 2147 day bar (~ 8.5 trading yr)。FX 历史最早 2018-02-07。要拉 2020-01-01 之后任意年报期间, 一次 `count=-2000 + begin=今天 CST` 完全覆盖。

### 三大币种 sanity (begin=2024-01-01 CST, count=-365)

| symbol | error_code | item_count | first_item_close | last_item_close |
|---|---:|---:|---:|---:|
| USDCNY.FX | 0 | 365 | 6.7505 | 7.0996 |
| HKDCNY.FX | 0 | 365 | 0.85994 | 0.90923 |
| KRWCNY.FX | 0 | 365 | 0.00519 | 0.00549 |

Cookie 同 1A 不必需。

### Plan-spec control (没改 .FX 后缀)

`USDCNY` (无 .FX) → 200 + `data.item: []` 空数组。同 1A 现象。**Production code 必须用 .FX 后缀**。

### 文件

- `samples/xueqiu_kline_USDCNY.json` — `attempts` 含 5 个调用 + `count_cap_probe` + `tz_cross_check` + `full_USDCNY_365d_parsed` 完整 365 条
- `tools/probe_1b_kline_fx.py`

---

## 1C — 傲基股份代码确认

### 5 个候选实测

| label | url | status | name / quote_name | currency | exchange | type | list_len |
|---|---|---:|---|---|---|---:|---:|
| HK_02519_xueqiu_quote | `quote.json?symbol=02519` | 200 | **傲基股份** | HKD (报价) | HK | 30 | — |
| HK_02519_xueqiu_finance_balance | `finance/hk/balance.json?symbol=02519&count=8` | 200 | **傲基股份** | **CNY (报告)** | — | — | 6 |
| HK_09927_xueqiu_quote | `quote.json?symbol=09927` | 200 | 赛力斯 (Seres) | HKD | HK | 30 | — |
| A_003031_xueqiu_quote | `quote.json?symbol=SZ003031` | 200 | 中瓷电子 | CNY | SZ | 11 | — |
| A_003031_sina_balance | `money.finance.sina.com.cn/.../sz003031/...` | 200 | (title) 中瓷电子(003031)资产负债表_新浪财经_新浪网 | — | — | — | — |

### 结论 (拒不带 `?`)

> **傲基股份 = 02519.HK 傲基股份, 港股 SEHK (xueqiu exchange=HK, type=30), 报告币种 CNY (人民币), 年报截止日 12-31, last_report_name = 2025年报, 已收录 6 期 (2020–2025) 年报。**

补充观察:
- 02519 quote 端点的 `currency=HKD` 是**交易报价币种** (港股按港币挂牌交易)。
- 02519 balance 端点的 `currency=CNY` / `currency_name=人民币` 是**报告披露币种** (公司财报以人民币编制)。
- 这是港股"trading currency ≠ reporting currency"的典型案例; Step 2/3 必须把 quote currency 和 finance currency 分开存。
- 09927 (赛力斯) 是无关公司, 但同为 HK / type=30, 排除。
- 003031 (中瓷电子, A 股) 全无关。Sina HTML 显式 `<title>` 已确认, `傲基` 字符串不存在于该页 (`html_contains_aoji=False`)。
- `xueqiu_HK_02519_balance.json` (Phase 4c 残留, UTF-16-LE 编码) 早期 quote_name 为"傲基控股"; 2025-04 更名为"傲基股份"。

### 文件

- `samples/aoji_probe_summary.json`
- `samples/aoji_probe_HK_02519_xueqiu_quote.json`
- `samples/aoji_probe_HK_02519_xueqiu_finance_balance.json`
- `samples/aoji_probe_HK_09927_xueqiu_quote.json`
- `samples/aoji_probe_A_003031_xueqiu_quote.json`
- `samples/aoji_probe_A_003031_sina_balance.html`
- `tools/probe_1c_aoji.py`

---

## 1D — 期间平均汇率算法 (供 Step 2/3)

### 选定算法: 报告期间所有可用 day bar 的算术平均 close

```python
def avg_rate_for_period(kline_items, period_start, period_end):
    """kline_items: [[ts_ms, vol, open, high, low, close, ...], ...]
    period_start/period_end: datetime.date, 含端点
    回报: 该期间 close 列的算术平均, 无样本时 None。
    """
    rates = []
    for it in kline_items:
        # ts = ms UTC; +8h 偏移取 CST 当日日期
        d = datetime.datetime.utcfromtimestamp(it[0] / 1000 + 8 * 3600).date()
        if period_start <= d <= period_end and it[5] is not None:
            rates.append(it[5])
    return sum(rates) / len(rates) if rates else None
```

### 取舍说明 — 为什么不用月度加权 / 期初期末加权

1. **月度加权需另一份月度交易额数据** (Xueqiu FX 没有 `volume`/`amount`, 全 null), 无法实施。
2. **期初期末加权** = `(rate(start) + rate(end)) / 2`, 简单但对人民币这种全年波动相对小的币种, 把中间月份波动全抹掉, 不如算术平均忠实。
3. **算术平均 day-close** 是会计实务"average rate for the period"的标准近似, 也是 IFRS / US GAAP 中"Average exchange rate for the period"的最常见实操。
4. 假节假日 (CST 春节、HKD 港股圣诞) 自然落在 `kline_items` 缺口上 — 不参与平均, 无需手工剔除。
5. **期末汇率** 单独跑一次取 `period_end` 当日 (或往前最近一个交易日) 的 close 即可, 用于资产负债表行项目。

### 取数模式 (Step 3 落地建议)

每个公司、每个报告期 (`sd` / `ed` 由 finance API 提供, 见 1C 02519 balance):
1. 按公司"报告币种" (从 `finance/hk/balance.json` 的 `currency` 字段读) 决定要拉哪个 FX symbol。
2. 一次 `count=-2000 begin=今天 CST` 把整段 USDCNY.FX (或 HKDCNY.FX / KRWCNY.FX) 拉到内存。
3. 同一进程内多公司、多报告期复用这条 kline cache, 不重复拉。
4. 每个报告期调一次 `avg_rate_for_period(items, sd, ed)`, 得期间平均汇率。
5. 期末汇率 `eop_rate(items, ed)` = ed 当日 close, 缺则 `next_earlier_close(items, ed)`。

---

## 1E — 给 Step 2 / Step 3 的可执行输入

### Step 2 (FX cache 模块设计)

1. **生产端点参数固定**:
   - quote: `symbol=USDCNY.FX,HKDCNY.FX,KRWCNY.FX` (3 in 1, 一次拿全)
   - kline: `symbol={ONE}.FX&period=day&type=before&begin={今天 CST 0 点 ms}&count=-2000`
2. **HTTP 客户端基线**:
   - `Accept-Encoding: gzip, deflate` (NOT `identity`)
   - 每次 process 启动时 GET 一次 `https://xueqiu.com/hq` 暖 session
   - cookie env `XUEQIU_COOKIE` 作 enhancement, 不作 dependency (无 cookie 已通)
   - 调用间 `time.sleep(1.2)`
3. **响应 schema 解析点**:
   - quote: `parsed["data"]` 是 list, 每项有 `symbol`, `current`, `last_close`, `timestamp` (ms UTC)
   - kline: `parsed["data"]["item"]` 是 list, 每项是 12 列 list, **timestamp = item[0], close = item[5]**

### Step 3 (傲基扫描脚本)

1. **股票代码硬编码 = `02519`** (xueqiu HK 端点用裸 5 位码, 不带 `.HK`)
2. **报告币种硬编码 = `CNY`** (从 1C 实测确认), Step 3 不需要再做 currency detection
3. **走 HK plugin 现有 `finance/hk/balance.json|income.json|cash_flow.json|indicator.json` 4 端点, count=8** — 已有 6 期 (2020–2025), 完整覆盖

### Step 2/3 共用风险与可调参数 (Risks)

| # | 风险 | 触发条件 | 可调参数 / 兜底 |
|---:|---|---|---|
| 1 | quote 端点 silent empty body | (a) `Accept-Encoding: identity`, (b) symbol 没 `.FX`, (c) 没 warmup | Step 2 FX 模块 retry 1 次 + 检查 `len(items) == 期望数量`, 不通则 raise |
| 2 | kline `count` 上限 ~2147 | 想拉 8.5 yr 之前的数据 | Step 3 用 `begin=今天 CST` + `count=-2000` 单次足够 (覆盖到 2018-08); 不要循环分段 |
| 3 | 傲基公司更名 | 02519 quote_name 从"傲基控股"→"傲基股份" (2025-04) | Step 3 显示用 `quote_name` 字段实时取, 不要硬编码字符串 |
| 4 | Xueqiu 反爬限速 | 高频 quote / kline 拉取 | `time.sleep(1.2)` 间隔; 一次 process FX cache 复用; 429 响应时退避到 2.0s |
| 5 | trading currency vs reporting currency | 02519 quote 端 = HKD, balance 端 = CNY | Step 3 显式从 `balance.json` `currency` 字段读"报告币种", 不要复用 quote 端的 `currency` |
| 6 | UTF-16 BOM 污染 | 之前 PowerShell `Out-File` 留下的残留 (`xueqiu_HK_02519_balance.json` 仍是 UTF-16-LE) | 所有 sample / 缓存写入用 Python `open(..., "w", encoding="utf-8")`, 永远不要 PS `Out-File` |
| 7 | FX symbol kline 在节假日缺口 | 圣诞、春节 | `avg_rate_for_period` 直接跳过缺失日; 期末汇率用 `next_earlier_close` 回溯 |

---

## Files Dumped

### Code (可重跑)

- `tools/probe_1a_quote_fx.py`
- `tools/probe_1b_kline_fx.py`
- `tools/probe_1c_aoji.py`

### Data

- `samples/xueqiu_quote_currency.json` (覆盖之前 PS 失败残留)
- `samples/xueqiu_kline_USDCNY.json` (覆盖之前 PS 失败残留)
- `samples/aoji_probe_summary.json`
- `samples/aoji_probe_HK_02519_xueqiu_quote.json`
- `samples/aoji_probe_HK_02519_xueqiu_finance_balance.json`
- `samples/aoji_probe_HK_09927_xueqiu_quote.json`
- `samples/aoji_probe_A_003031_xueqiu_quote.json`
- `samples/aoji_probe_A_003031_sina_balance.html`
