# AGENTS.md

> 本文件是 FS Capture / VBA Captor 项目的 agent 工作手册。它基于 `PROJECT_RETROSPECTIVE.md` 与 `STATUS.md` 总结，目标不是复述项目历史，而是沉淀一套可迁移到后续本地 Excel / VBA / Python 自动化项目的协作规则。

---

## 1. Agent 进入项目的第一动作

先读以下文件，不要直接改代码：

1. `PROJECT_RETROSPECTIVE.md`  
   了解真实用户、交付目标、协作模式、踩坑记录和可复用模板。

2. `STATUS.md` 最新收口段  
   只从末尾最新阶段读起，再按需回溯。不要默认早期章节仍代表当前状态。

3. `ARCHITECTURE.md`  
   做代码改动时必须读，尤其是模块依赖、sheet inventory、数据流、关键 invariants。

4. 当前 phase plan 或用户最新 prompt  
   如果 plan 与旧文档冲突，以用户最新明确决策为准，但要在报告里说明冲突点。

进入后先判断本轮属于哪类工作：

| 类型 | 默认策略 |
|---|---|
| Bug 修复 | 先复现，再做最小补丁，再跑相关回归 |
| Phase 实施 | 严格按 plan step 顺序做，不自由重排 |
| UX / 视觉 | 不碰业务逻辑，按截图、cell range、颜色、行高列宽精确执行 |
| 文档收口 | 只更新当前事实，不重写历史决策 |
| Release / 清理 | 保留用户数据，清理可再生数据和敏感缓存 |

---

## 2. 项目协作模型

本项目最有效的模式是三角协作：

```text
Planner / Reviewer  ->  写 plan、定义边界、独立验收
Generator           ->  端到端实现、装表、跑回归、提交报告
Static Auditor      ->  扫全量代码，给 backlog，Reviewer 取舍
```

Agent 要遵守自己的角色边界：

- Generator 不重新设计 Planner 已锁定的方案。
- Reviewer 不只看 Generator 自报，必须独立验证关键路径。
- Static auditor 给建议，不等于必须全做；个人工具项目要按 ROI 砍掉过度工程。
- 多 agent 协作只拆互不冲突的工作面，例如文档、inspect、VBA wrapper、Python installer；不要让多个 agent 同时改同一 VBA 大模块的重叠区域。

---

## 3. 冻结边界: true frozen vs state-bound

这是后续项目最重要的规则之一。

| 文件 / 对象类型 | 分类 | 规则 |
|---|---|---|
| 核心业务 helper、公开函数签名、fetch/cache/write/FX 主逻辑 | True frozen | 除非用户明确要求，否则不改 |
| 集成回归驱动，例如 `test_fx_live.py`、`diff_*` | True frozen | 不为让测试通过而改基线 |
| `inspect_phase*_state.py` | State-bound | 当主线功能明确删除或移动 sheet/button/cell 时，必须同步 inspect |
| `STATUS.md` / `PHASE_*_PLAN.md` | 历史追溯 | 追加收口段，不改写已闭环历史 |
| `README.md` / `ARCHITECTURE.md` | 当前说明 | 跟随最终产品状态更新 |

如果 plan 写了 frozen，但用户新需求明确删除被检查对象，要停下或明确说明后再同步 state-bound inspect。不要把过时 inspect 当成产品真相。

---

## 4. 本项目核心 invariants

任何改动都不能破坏以下事实：

- 工具是本地 Excel 桌面工具，不是网络服务，不是采集平台。
- 用户是财务 / 审计专业人士，界面和文档要用业务语言，不要堆技术黑话。
- 5 市场正式输出保留：A 股、美股、港股、韩股、台股各 4 张分市场表。
- 跨市场只保留 `跨市场_指标表`，不恢复已删除的跨市场 BS / IS / CF 合表。
- RMB toggle 必须保持：原币与统一 RMB 可切换，汇率缺失时写空并写诊断，不 fallback 到 1。
- `GetFxRate` 旧签名保持兼容；新逻辑通过扩展 helper 实现。
- 诊断 sheet 当前是 17 列 schema，不要回退到 11 列。
- `.cache/` 是可再生 HTTP 缓存；清空数据不等于清空缓存，除非用户明确点清 HTTP 缓存。
- `样本池` 用户录入区必须保留，尤其是 Row 14+ 公司代码和简称。Row 5 已弃用 (Phase 5a),保持空白,不要放任何标签或值。
- `XueqiuHttpGet` 内部已委托 `FetchViaPowerShell` 匿名 warmup;不要在 HK/US fetcher 里恢复 `If Len(strCookie)=0 Then Err.Raise` 早退保护;不要用 openpyxl 改 .xlsm (会清掉所有 Shape / OnAction 绑定)。

---

## 5. Excel / VBA 安全操作规则

Excel COM 是本项目最高频风险源。所有 agent 必须按下面规则处理：

1. 避免模态弹窗卡死  
   在会触发 merge/unmerge、删除 sheet、清空合并区域的宏里保存并关闭 `Application.DisplayAlerts`，退出时恢复原值。

2. 入口宏使用 AppStateGuard  
   对耗时入口保存 / 恢复 `ScreenUpdating`、`DisplayAlerts`、`DisplayStatusBar`、`EnableEvents`、`Calculation`、`StatusBar`、`DisplayPageBreaks`。异常路径也必须恢复。

3. 合并单元格不能直接当普通 Range 写  
   清空或重排合并区域时，先确认 merge 状态；必要时 unmerge -> clear -> merge。

4. COM 卡死先排查 Excel 状态污染  
   如果连续回归突然 0x800AC472 或一直 hang，优先检查是否有残留 Excel.exe，而不是立刻怀疑业务代码。

5. Silent 模式不能弹 MsgBox  
   COM smoke 和回归脚本调用入口宏时，必须支持 `blnSilent=True` 或等效静默路径。

6. 不要在 hidden Excel 中依赖可见 UI 行为  
   如果必须排查 UI 弹窗，临时 `Visible=True` 可以用于定位，但不要把它作为产品逻辑。

---

## 6. 代码改动原则

保持补丁窄而可验：

- 先定位调用链，再改最小入口。
- 优先加 wrapper / guard / diagnostic，不直接重写稳定 fetch 主路径。
- 公开函数签名要向后兼容，除非 plan 明确允许破坏。
- 不要把 UX 改动和数据逻辑改动混在一个无边界补丁里。
- VBA 大模块容易冲突，改前用搜索确认所有调用点。
- 对结构化数据优先用既有 parser/helper，不要临时字符串拼接。
- 任何数据准确性问题优先显式诊断，不要静默兜底成看似正确的数字。

对本项目尤其注意：

- 不要恢复单表辅助按钮。
- 不要恢复跨市场 BS / IS / CF / 字段映射 sheet。
- 不要把 stockanalysis fallback 改回手动 toggle。
- 不要让 FX missing fallback 到 `1`。
- 不要改 5 市场 fetch 字段映射，除非本轮就是数据源维护。

---

## 7. 测试与回归栈

本项目的验证不是单一测试，而是分层栈：

```text
联网 / 集成回归     test_fx_live.py, diff_phase4f_step3_lite.py
状态检查            inspect_phase4g/h/k/l/m_state.py
离线 fixture 测试   run_offline_tests.py
手工 / COM smoke    针对本轮入口、弹窗、Excel 状态恢复
```

常用完整回归：

```powershell
py tools/test_fx_live.py --skip-install
py -u tools/diff_phase4f_step3_lite.py
py -u tools/inspect_phase4g_state.py
py -u tools/inspect_phase4h_state.py
py -u tools/inspect_phase4k_state.py
py -u tools/inspect_phase4l_state.py
py -u tools/inspect_phase4m_state.py
py tools/run_offline_tests.py
```

运行原则：

- 任一既有 PASS 变 FAIL，立即停下，不继续叠改。
- 如果 inspect 因产品状态变化失效，先判断它是不是 state-bound；同步后要在报告里说明。
- 离线 fixture 测试不得发 HTTP。
- Live HTTP 测试要尊重缓存、限流和 retry 设计，不要为了快绕过产品路径。
- UI-only 文档改动可以不跑完整宏，但要说明未跑原因。

---

## 8. Phase plan 模板

后续类似项目建议每个 phase 都包含这些块：

```markdown
> 版本 / 状态 / 作者 / 背景

## 项目语境
[说明这是本地工具、已有数据源延续、不是新采集系统]

## 用户已锁定决策
| 决策点 | 结论 |

## Step 总览
| Step | 内容 | 验收 |

## Step 详细
- 实施位置
- 函数签名 / cell range / 颜色 / 行高列宽 / 文件路径
- 验证方式
- 不要做什么

## 严禁动
| 文件或区域 | 原因 |

## 触发停下条件
- 任一 frozen 回归退化
- VBE Compile 失败
- 性能超过预算
- 命中率低于阈值
- 真实数据明显异常

## State-bound inspect 同步规则
[明确哪些检查可随产品状态迁移]
```

Plan 要具体到足够让 Generator 不自由发挥。视觉任务必须给目标截图、颜色 hex、cell range、列宽、行高。

---

## 9. 报告格式

完成后报告要短，但必须可复核：

```text
完成: <phase / task>
Commit: <hash> <message>

改动:
- <文件/函数>: <做了什么>

验证:
- <命令>: PASS / FAIL,关键尾行
- <COM smoke>: <关键数值>

边界:
- 未跑什么 / 为什么
- 是否有 state-bound inspect 同步
```

如果没有 commit，不要伪造 commit hash。  
如果测试没跑，不要写 PASS。  
如果 Excel/网络/COM 受环境影响，要把环境状态写清楚。

---

## 10. 用户视角文档规则

本项目的最终用户不是开发者。任何用户面向文档、sheet 文案、按钮 caption 都要按下面规则写：

- 用“抓数 / 报表 / 对比 / 汇率 / 缓存 / 诊断”这类业务词。
- 少用或不用 `WinHttp`、`UDF`、`fallback`、`TTL`、`schema`、`COM` 等技术词。
- 不解释实现细节，解释用户该怎么判断和操作。
- 错误提示要指向可行动作，例如更新 cookie、清 HTTP 缓存、重跑某个市场。
- 成功提示不要带特殊符号，避免 VBA/Excel 弹窗乱码。

内部文档可以技术化，但要可追溯、可验证。

---

## 11. 数据准确性优先级

财务工具里，错误数字比空值更危险。

因此：

- 汇率缺失 -> 空值 + `FX_MISSING` 诊断。
- 字段缺失 -> 空值 + QA / diagnostic，而不是猜测。
- HTTP 失败 -> 明确来源、状态、重试次数、缓存状态。
- 跨市场 BS / IS / CF 行项目不可比时，不强行合并；只合并已标准化的 18 项指标。
- QA 规则以低误报为先，WARN 不阻断用户工作流。

---

## 12. 可复用到后续项目的经验

最值得迁移的不是具体 VBA，而是这套交付方法：

1. 先锁用户决策，再写 phase plan。
2. 每个 phase 都有严禁动清单和停下条件。
3. 把测试分成 true frozen 和 state-bound inspect。
4. 对 Excel / Office 自动化统一做 AppStateGuard。
5. 复杂数据源项目必须有离线 fixture。
6. 数据异常显式诊断，不做静默伪正确。
7. 视觉任务用像素级 spec，不让 agent 自由发挥。
8. 用户文档按目标读者重写，不从实现细节翻译。
9. 每个收口都写 STATUS 段，长期项目靠状态文档续航。
10. 静态审阅给 backlog，Reviewer 按 ROI 取舍，不盲目 enterprise 化。

---

## 13. 当前发布基线

当前项目已进入 v1.0 release 收尾状态。长期维护时，优先从以下入口恢复上下文：

- `PROJECT_RETROSPECTIVE.md`: 项目复盘和协作方法。
- `STATUS.md` §EE: Phase 4n 最终优化收口。STATUS.md §FF (v1.1 release / TW 台股接入) 和 §GG (Phase 5a 雪球 cookie 自动化) 是当前末尾基线。
- `ARCHITECTURE.md`: 当前架构、数据流、invariants。
- `README.md`: 用户使用说明和 release notes。
- `PHASE5A_CHANGELOG.md`: 雪球 cookie 自动化 (E5 弃用、`XueqiuHttpGet` 委托 `FetchViaPowerShell`) 的事实来源。
- `tools/run_offline_tests.py`: 离线 parser / QA / fixture 快速检查。

对后续维护者的简短提醒：

> 不要为了“看起来更完整”恢复已删功能。这个项目最终选择的是少而准：5 市场分表 + 1 张跨市场指标表 + RMB toggle + 诊断 / QA / 缓存，而不是全字段跨市场万能合表。

