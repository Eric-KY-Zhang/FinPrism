# 上市公司财务数据查询 · FinPrism

> **FinPrism** — A multi-market listed-company financial data prism for cross-comparable wide-table analysis.

把 A 股 / 美股 / 港股 / 韩股四个市场的**资产负债表、利润表、现金流量表和 18 项标准财务指标**,一键整理成横向同业对标宽表,支持**原币**和**统一 RMB**双口径。

打开 Excel,填几个股票代码,点一下按钮 —— 报表就出来了。

[![Download](https://img.shields.io/badge/Download-v1.0-blue?style=for-the-badge&logo=microsoft-excel)](https://github.com/Eric-KY-Zhang/FinPrism/releases/latest)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)

作者: Eric Zhang · 联系邮箱: 214978902@qq.com

---

## 数据源覆盖

| 市场 | 数据源 |
|---|---|
| A 股 | 新浪财经 |
| 美股 | SEC EDGAR + stockanalysis.com (中概股 fallback,已验证 BABA / JD / PDD) |
| 港股 | 雪球 HK API |
| 韩股 | stockanalysis.com KRX 财报页面 |
| 汇率 | 雪球 quote / k-line API |

---

## 快速开始

### 1. 下载

去 **[Releases 页面](https://github.com/Eric-KY-Zhang/FinPrism/releases/latest)** 下载 `FinPrism-v1.0.xlsm`(575 KB)。

> 如果你想看 / 改 VBA 源码,下载 `FinPrism-v1.0-source.xlsm`(702 KB)。

### 2. 启用宏

打开 xlsm 后,Excel 顶部会出现黄色安全栏:

```
受保护的视图 / 已禁用宏 → 点「启用编辑」和「启用内容」
```

第一次启用后,后续打开就不会再问。

### 3. 用「样本池」sheet 抓数

打开后默认就在「样本池」sheet。整个工具只需要在这张 sheet 上操作:

#### 3.1 设置抓取参数（顶部）

| 单元格 | 内容 | 说明 |
|---|---|---|
| `E3` | 年度 | 例如 `2024`;留空则取每家公司各自的最新可得报告期 |
| `E4` | 季度 | `Q1` / `Q2` / `Q3` / `Q4` / `全部` |
| `E5` | 雪球 cookie | 抓**港股 / 部分中概美股**时必填,见下文「获取雪球 cookie」 |
| `E6` | 货币口径 | `原币` 或 `统一RMB`(自动按汇率折算) |

#### 3.2 录入公司（第 13 行起,按市场分栏）

| 列 | 市场 | 例如 |
|---|---|---|
| `A:B` | A 股 | `贵州茅台` `600519` |
| `E:F` | 美股 | `Apple` `AAPL` |
| `I:J` | 港股 | `腾讯控股` `00700` |
| `M:N` | 韩股 | `三星电子` `005930` |

只填想抓的市场即可,其他列留空。

#### 3.3 点按钮抓取

样本池顶部有以下一键按钮:

| 按钮 | 作用 |
|---|---|
| **一键 A 股** / 美股 / 港股 / 韩股 | 只更新对应市场的 4 张报表 |
| **一键全抓 4 市场** | 顺序更新 16 张正式报表,并刷新「跨市场_指标表」 |
| **显示 / 隐藏 X 数据** | 控制各市场正式报表的显隐 |
| **清空 HTTP 缓存** | 清除本地 24h 缓存,下次抓数从数据源重新取 |

抓取完成后,对应市场的 4 张报表会自动显示。

---

## 输出表

抓取后会得到以下 17 张报表(诊断 sheet 默认隐藏):

```
1.  使用说明
2.  样本池                    ← 操作入口
3.  跨市场_指标表             ← 4 个市场 18 项指标的合并对照视图
4-7.   A 股_资产负债表 / 利润表 / 现金流量表 / 指标表
8-11.  美股_资产负债表 / 利润表 / 现金流量表 / 指标表
12-15. 港股_资产负债表 / 利润表 / 现金流量表 / 指标表
16-19. 韩股_资产负债表 / 利润表 / 现金流量表 / 指标表
```

**宽表布局**:

- 第 1 行:公司名 + 代码
- 第 2 行:报告期
- A / B 列:大类、指标名(指标表多一列英文)
- A 股按报告期并集对齐(便于横向对比),美 / 港 / 韩按各家公司自身可用期间展开
- **原币模式**:各市场按原口径(美元 / 港元 / 韩元等)
- **统一 RMB 模式**:非人民币按报告期汇率折算
  - 资产负债表用**期末**汇率
  - 利润表 / 现金流量表用**期间平均**汇率
- 港股单位:百万(各家公司报告币种)
- 韩股单位:十亿韩元(KRW billions,原表为百万,工具内除 1000)

**指标表**:18 项标准指标,4 个市场共用同一套口径。

---

## 常见问题

### 怎么获取雪球 cookie?

抓港股或部分中概美股(主路径失败时)需要雪球 cookie:

1. 浏览器打开 https://xueqiu.com 并登录
2. F12 打开开发者工具 → Network → 刷新页面
3. 任选一个 xueqiu.com 的请求 → Headers → Cookie
4. 找到 `xq_a_token=...;` 那一段,复制 token 值
5. 粘贴到「样本池」`E5`

cookie 过期会报错,重新复制最新的就行。

### 抓不到数据怎么办?

打开「**美股_抓取诊断**」/「**港股_抓取诊断**」/「**韩股_抓取诊断**」sheet(右键 → 取消隐藏),里面有每家公司每个字段的命中情况、单位、汇率、fallback 路径。

A 股一般直接抓新浪不会有问题。

### 美股 fiscal year 不是 12 月怎么办?

工具自动处理:增长率公式找 prior period 时先精确日期匹配,匹配不到退到 ±31 天最近期间(覆盖 AAPL FY 9 月底 / MSFT FY 6 月底等场景),不会跨年误命中。

### 港股为什么抓得慢?

为了让指标表的增长率公式有可比期数据,「一键港股」会在抓资产负债表 / 利润表时同时拉**当年 + 前一年**,所以耗时约 2x。现金流量表不双拉。

---

## License

本项目基于 [MIT License](LICENSE) 开源,Copyright (c) 2026 Eric Zhang。

### 第三方组件

- [`modules/JsonConverter.bas`](modules/JsonConverter.bas) — VBA-JSON v2.3.1 by Tim Hall ([VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)),MIT License。

### 数据来源声明

本工具仅作为公开数据的检索整理客户端,不附带任何数据本身。各数据源版权归原站点所有(新浪财经 / SEC EDGAR / 雪球 / stockanalysis.com)。使用者需自行遵守各数据源的服务条款和 robots.txt 限速要求。

---

## 开发者文档

如果你想从源码构建 / 修改 VBA 模块,见:

- [`ARCHITECTURE.md`](ARCHITECTURE.md) — 整体架构与模块划分
- [`STATUS.md`](STATUS.md) — 完整开发记录与 commit 历史
- [`AGENTS.md`](AGENTS.md) — 多 Agent 协作约定
- [`docs/plans/`](docs/plans/) — 各 Phase 设计文档(4b ~ 4n)
- [`tools/build_template.py`](tools/build_template.py) + [`tools/install_modules.py`](tools/install_modules.py) — 从空白模板重建 xlsm 的脚本
