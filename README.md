# 上市公司财务数据查询 · FinPrism

> **FinPrism** — A multi-market listed-company financial data prism for cross-comparable wide-table analysis.

把 A 股 / 美股 / 港股 / 韩股 / 台股 5 个市场的资产负债表、利润表、现金流量表和 18 项标准财务指标,一键整理成横向同业对标宽表,支持原币和统一 RMB 双口径。

打开 Excel,在「样本池」录入公司代码,选择期间和币种口径,点击一键按钮即可生成报表。

[![Download](https://img.shields.io/badge/Download-v1.1-blue?style=for-the-badge&logo=microsoft-excel)](https://github.com/Eric-KY-Zhang/FinPrism/releases/latest)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)

作者: Eric Zhang · 联系邮箱: 214978902@qq.com

---

## 数据源覆盖

| 市场 | 数据源 | 备注 |
|---|---|---|
| A 股 | 新浪财经 | 不需要登录凭证 |
| 美股 | SEC EDGAR + stockanalysis.com / 雪球备用路径 | 部分中概股字段不完整时会使用备用来源 |
| 港股 | 雪球 HK API | 需要在样本池填写雪球 `xq_a_token` |
| 韩股 | stockanalysis.com KRX 财报页面 | 不需要登录凭证 |
| 台股 | 公开台股财务报表数据 | 不需要登录凭证 |
| 汇率 | 雪球 quote / k-line API | 用于 USDCNY / HKDCNY / KRWCNY / TWDCNY |

---

## 快速开始

### 1. 下载

去 [Releases 页面](https://github.com/Eric-KY-Zhang/FinPrism/releases/latest) 下载 `FinPrism-v1.1.xlsm`。

如果你想看或修改 VBA 源码,下载 `FinPrism-v1.1-source.xlsm`。

### 2. 启用宏

打开 xlsm 后,Excel 顶部会出现黄色安全栏:

```text
受保护的视图 / 已禁用宏 -> 点「启用编辑」和「启用内容」
```

第一次启用后,后续打开同一文件通常不会再询问。

### 3. 设置样本池

所有日常操作都在「样本池」完成。

| 单元格 | 内容 | 说明 |
|---|---|---|
| `E3` | 年度 | 例如 `2025`;留空则取每家公司各自最新可得报告期 |
| `E4` | 季度 | `Q1` / `Q2` / `Q3` / `Q4` / `全部` |
| `E5` | 雪球 cookie | 抓港股或部分中概美股备用路径时填写 `xq_a_token` |
| `E6` | 货币口径 | `原币` 或 `统一RMB` |

从第 14 行起按市场录入公司:

| 市场 | 代码列 | 简称列 | 示例 |
|---|---|---|---|
| A 股 | `A` | `B` | `300866` / `安克创新` |
| 美股 | `D` | `E` | `AAPL` / `Apple` |
| 港股 | `G` | `H` | `00700` / `腾讯控股` |
| 韩股 | `J` | `K` | `005930` / `三星电子` |
| 台股 | `N` | `O` | `2330` / `台积电` |

只填需要抓取的市场即可,其他市场可以留空。

### 4. 抓取数据

样本池顶部提供这些按钮:

| 按钮 | 作用 |
|---|---|
| 一键 A 股 / 美股 / 港股 / 韩股 / 台股 | 只更新对应市场的 4 张报表 |
| 一键全抓 5 市场 | 顺序更新 5 个市场并刷新「跨市场_指标表」 |
| 一键抓取跨市场指标表 | 不重抓原始报表,只用已抓取数据刷新合并指标表 |
| 显示 / 隐藏 X 数据 | 控制各市场正式报表的显隐 |
| 一键清空所有数据 | 清空已生成报表,保留样本池公司和参数 |
| 清空 HTTP 缓存 | 清除本地 24 小时缓存,下次重新从公开数据源取数 |

---

## 输出表

Excel 内不再保留单独的「使用说明」sheet;使用说明集中维护在本 README。

| 表 | 用途 |
|---|---|
| 样本池 | 录入公司、设置参数、执行抓数 |
| 跨市场_指标表 | 5 个市场 18 项标准指标的合并对照视图 |
| A股 / 美股 / 港股 / 韩股 / 台股报表 | 每个市场 4 张表:资产负债表、利润表、现金流量表、指标表 |
| 汇率 | 只保留折算使用的汇率数据表,不再放说明文字 |
| 抓取诊断 | 默认隐藏,用于排查字段命中、单位、汇率和缓存状态 |

宽表布局:

- 第 1 行:公司名 + 代码
- 第 2 行:报告期
- A / B 列:大类、指标名;指标表额外有英文指标名列
- A 股按报告期并集对齐;美股 / 港股 / 韩股 / 台股按公司自身可用期间展开
- 港股单位:百万,以公司报告币种为准
- 韩股单位:十亿韩元
- 台股单位:百万新台币

---

## 汇率与人民币折算

「汇率」sheet 现在是纯数据表,列结构如下:

| 列 | 含义 |
|---|---|
| `A` | 报告期 |
| `B:C` | USDCNY 期末 / 期均 |
| `D:E` | HKDCNY 期末 / 期均 |
| `F:G` | KRWCNY 期末 / 期均 |
| `H:I` | TWDCNY 期末 / 期均 |
| `J` | 备注 / override |

折算规则:

- `E6 = 原币`:按各公司原始报告币种显示,不做人民币折算。
- `E6 = 统一RMB`:非人民币报告币种自动折算为人民币,便于跨市场横向比较。
- 资产负债表使用期末汇率。
- 利润表和现金流量表使用期间平均汇率。
- RMB / CNY 报告币种不折算。

汇率使用说明:

- USDCNY 用于美股,HKDCNY 用于港股,KRWCNY 用于韩股,TWDCNY 用于台股。
- 工具会从公开历史行情获取报告期对应汇率,首次取得后写入「汇率」sheet。
- 同一报告期、同一币种的汇率会本地缓存 24 小时;需要强制刷新时点击「清空 HTTP 缓存」。
- 如需手工调整,直接修改「汇率」sheet 的 `B:I` 对应单元格,并在 `J` 列写明来源或原因;修改后重跑对应市场按钮。
- 若汇率获取失败,统一 RMB 模式下相关单元格会留空,诊断表会标记 `FX_MISSING`;不要把缺失汇率默认为 1。

---

## 获取雪球 Cookie

抓港股或部分中概美股备用路径需要雪球 `xq_a_token`:

1. 浏览器打开 `https://xueqiu.com` 并登录。
2. 按 `F12` 打开开发者工具,进入 Network,刷新页面。
3. 任选一个 `xueqiu.com` 请求,查看 Headers 里的 Cookie。
4. 找到 `xq_a_token=...;`,复制 token 值。
5. 粘贴到「样本池」`E5`。

cookie 过期时重新复制最新 token 即可。

---

## 常见问题

### 抓不到数据怎么办?

先取消隐藏对应市场的「抓取诊断」sheet,查看状态、数据源、字段命中、单位、汇率和备注。港股或中概美股最常见原因是雪球 cookie 过期。

### 切换原币 / 统一 RMB 后数字变化正常吗?

正常。统一 RMB 会把外币报告币种按汇率折算成人民币。资产负债表走期末汇率,利润表和现金流量表走期间平均汇率。

### 美股 fiscal year 不是 12 月怎么办?

工具会按披露报告期处理。增长率公式找 prior period 时先精确日期匹配,匹配不到再退到相近报告期,用于覆盖 AAPL FY 9 月底、MSFT FY 6 月底等情况。

### 港股为什么抓得慢?

港股指标表需要当年和前一年数据来计算增长率,所以「一键港股」会在抓资产负债表 / 利润表时同时拉当年 + 前一年。

---

## License

本项目基于 [MIT License](LICENSE) 开源,Copyright (c) 2026 Eric Zhang。

### 第三方组件

- [`modules/JsonConverter.bas`](modules/JsonConverter.bas) — VBA-JSON v2.3.1 by Tim Hall ([VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)),MIT License。

### 数据来源声明

本工具仅作为公开数据的检索整理客户端,不附带任何数据本身。各数据源版权归原站点所有。使用者需自行遵守各数据源的服务条款和 robots.txt 限速要求。

---

## 开发者文档

- [`ARCHITECTURE.md`](ARCHITECTURE.md) — 整体架构与模块划分
- [`STATUS.md`](STATUS.md) — 完整开发记录与状态收口
- [`AGENTS.md`](AGENTS.md) — 多 Agent 协作约定
- [`docs/plans/`](docs/plans/) — 各 Phase 设计文档
- [`tools/build_template.py`](tools/build_template.py) + [`tools/install_modules.py`](tools/install_modules.py) — 从空白模板重建 xlsm 的脚本
- [`tools/prepare_release.py`](tools/prepare_release.py) — 生成 `release/` 发布包
