---
name: paper-docx-optimizer
display_name: 论文 DOCX 优化器
display_name_en: Paper DOCX Optimizer
description: 中文论文 DOCX 一体化优化与排版。当用户要降 AI 痕迹（降 AIGC 率）、降查重率、学术润色、按章节批量改写、回填 Word、体检论文格式、统一三线表字体字号，或为系统设计章节画 ER 图时触发。关键词：降AI、降AIGC、降重、查重率、论文润色、论文修改、DOCX 优化、按章节改写、论文格式、格式体检、三线表、学位论文、毕业论文、硕士论文、博士论文、ER图、数据库设计图。
description_zh: 论文 DOCX 一体化优化器：结构解析 → AI 痕迹扫描 → 三线优化（降AI/降重/润色）→ run 级回填 Word → 格式体检与规整 → ER 图生成，全程最小干预、术语零改动。
description_en: All-in-one optimizer for Chinese thesis DOCX. Parses structure, scans AI-trace patterns, applies minimal-intervention rewriting (AI-reduction / plagiarism-reduction / academic polishing), writes results back into the original Word layout at run level, audits and fixes thesis formatting, and renders thesis-style ER diagrams.
category: 效率工具
version: 1.0.0
author: Luwen
---

# 论文 DOCX 优化器

把「论文导出成 docx，要降 AI、降重、润色、还要符合学校格式」这件事做成一条可重复执行的流水线。核心是三件事：**定位**（哪里风险高）、**改写**（只动该动的句子）、**回填**（版式一个字都不变）。

## 何时使用

触发场景：

- 给一份论文 docx（可选附带 AIGC / 查重报告），要求自动优化并输出 docx
- 要求降低 AI 检测率、降低查重率、提升学术表达
- 要求按章节批处理：只改摘要和结论，或跳过绪论
- 要在不破坏原有 Word 版式的前提下改正文
- 要体检论文格式（纸张、边距、字体字号、行距、题注、三线表、页码），或一键规整
- 要为系统设计 / 数据库设计章节生成中文 ER 图（整体图或单表图）

不适用：

- 只要手工润色一段文字，不需要成文流程
- 从零生成整篇论文（本技能只做优化，不做代写）
- 只处理 PDF / Markdown，不需要回到 docx
- UML 类图、流程图、时序图（ER 图以外）

## 铁律

1. **原稿只读**。任何情况下都不覆盖输入 docx，产物一律写入 `output/final/`。
2. **最小干预**。降 AI 只做句内重组、语序调整、局部替换；禁止整段推翻重写、禁止扩写、禁止洗稿。
3. **术语与编号零改动**。模型名、模块名、缩写、图/表/公式/文献编号、数值、单位必须逐字保留。
4. **保版式**。回填走 run 级替换（保留字体、字号、加粗、上标），不重建段落、不重建表格。
5. **不编造**。不新增文献、数据、实验结果、结论；缺依据的地方只做语言收紧，不补内容。
6. **不降质**。禁止为降 AI 引入语法错误、口语词、错误术语；禁止删除必要的局限性说明。
7. **改完必复检**。改写后重新扫描，指标不改善就回退该段，不靠感觉交付。

## 全链路

```
input/论文.docx（+ 可选报告）
   │
   ├─ ① parse     结构解析：标题层级、章节段落区间、包完整性   → analysis/headings.json、manifest.json
   ├─ ② scan      AI 痕迹扫描：模式命中、词表、句长分布        → analysis/ai_scan.md
   ├─ ③ optimize  章节规划 + 规则改写 + 术语守卫 + run 级回填  → analysis/chapter_*.md、final/*_v1.docx
   ├─ ④ refine    按 references/ 策略逐段精修（人工 / AI 协作）  → 章节任务文档 + 逐版 docx
   ├─ ⑤ finalize  另存最终版 + 完整性复检                       → final/*_最终版.docx
   └─ ⑥ audit     格式体检 / 规整 + 交付汇总                    → format_check.md、delivery_report.md
```

四道门禁：解析必须拿到可用标题结构；改写必须通过术语守卫；定稿必须通过包完整性；体检严格问题必须清零或逐条说明。

## 执行步骤

### Step 0 确认输入

收集三件事，缺什么问什么，不要猜：

1. 论文 docx 路径（若目录内有多个 docx，用 `--docx-path` 显式指定）
2. 是否有 AIGC / 查重报告（有就一起解析，没有就走默认风险规则 `references/risk-rules.md`）
3. 处理范围：全篇 / 指定章节白名单 / 排除某些章节
4. 学校格式要求是否有特殊项（边距、字号、章标题体例），有则先改 `scripts/format_check.py` 的 `SPEC`

### Step 1 解析结构

```bash
python scripts/paper_pipeline.py parse --input-dir <input_dir> --output-dir <output_dir>
```

读 `analysis/thesis_structure.md` 确认三件事：标题是否成层级、章节数是否合理、`docx_integrity_ok` 是否为 true。若标题为空 → 论文没用「标题 1/2/3」样式，先跑 `format_check.py --fix` 规整样式，或改用 `--chapter-allowlist` 显式圈定范围。

### Step 2 扫描定位

```bash
python scripts/ai_scan.py --docx <论文.docx> --out <output_dir>/analysis
```

读 `ai_scan.md` 的章节热力图与高危清单。判定：`overall_score` 高、`cv < 0.28`、T1 词密集的段落优先。**只改「高」段落**，中低风险段落最多做词级清理。扫描器口径说明见 `references/perplexity-burstiness.md`。

### Step 3 生成规划并执行规则轮

```bash
# 先预览，确认改写幅度
python scripts/paper_pipeline.py optimize --input-dir <input_dir> --output-dir <output_dir> --dry-run

# 确认后落盘
python scripts/paper_pipeline.py optimize --input-dir <input_dir> --output-dir <output_dir>
```

必看 `analysis/thesis_master_overview.md`（章节矩阵 + 优先级 + 危险片段）和 `analysis/chapter_*.md`（每章任务）。规则轮只处理强标记词与明确的模板句，`skipped` 里是术语守卫拦截的段落，需人工判断。

### Step 4 逐章精修

这是真正的主体工作，规则轮只是清障。对每个 P0/P1 章节：

1. 打开 `analysis/chapter_XX_*.md`，按其中「本轮目标」定位段落
2. 按 `references/ai-pattern-taxonomy.md` 判模式，按对应策略改：降 AI → `@references/strategy-ai-reduction.md`；降重 → `@references/strategy-plagiarism.md`；润色 → `@references/strategy-polishing.md`
3. 改完在 Word / 文档中落笔，或整理「原文 → 改后」对照表后一次性回填
4. 该章改完立刻重扫该章，`score` 应下降、`cv` 应上升
5. 在章节任务文档的「处理记录」写清改了什么、为什么

段落级操作纪律：

- 单段改动量控制在 30% 以内，只动句子内部与句间衔接
- 一段内不同时做「降 AI + 降重 + 润色」三件事，先结构后词句
- 遇到术语密集的方法段，只做词级清理，不动句式
- 摘要、结论、展望这三处可以放开改；实验数据段一个字不动

### Step 5 定稿

```bash
python scripts/paper_pipeline.py finalize --output-dir <output_dir>
```

输出 `final/*_最终版.docx` 并复检包完整性。多轮迭代时保留 `_v1` `_v2`，最终版单独命名。

### Step 6 格式体检与交付

```bash
# 只体检
python scripts/format_check.py --docx <最终版.docx> --out <output_dir>/analysis

# 体检 + 规整页面/样式/题注（可选统一三线表）
python scripts/format_check.py --docx <最终版.docx> --out <output_dir>/analysis --fix --fix-tables

# 汇总交付报告
python scripts/paper_pipeline.py report --output-dir <output_dir>
```

`format_check.md` 中「❌ 必须修改」必须清零；「⚠️ 建议核对」要逐条确认或说明为何不改（例如学校要求不同）。判定依据见 `references/thesis-format-spec.md`。

### 附带：ER 图

```bash
python scripts/render_er_diagram.py --input schema.yaml --output-dir <output_dir>/diagrams \
  --mode system-simple --format svg,png --imagegen-pack
```

先写 `schema.yaml`（模板 `templates/schema.example.yaml`），再按 `@references/er-diagram-spec.md` 的构图规范校验，最后用图像模型把结构图重绘为论文插图风。图内中文必须逐字核对。

## 改写执行规范

### 五种允许的动作

1. 删套话：`综上所述`、`值得注意的是`、`具有重要的现实意义` —— 直接删
2. 换具体：`显著提升` → `MSE 由 0.312 降到 0.245`
3. 调语序：把「通过 A 实现了 B」改成「B 的实现依赖 A」
4. 拆合句：等长句群拆出短句，长句加插入语
5. 去对称：`首先/其次/最后`、`一方面/另一方面` 只留一侧

### 四种禁止的动作

1. 整段重写（哪怕意思没变）
2. 用同义词替换术语（`注意力机制` → `关注力机制`）
3. 补写原文没有的内容（文献、数据、结论）
4. 改动图表、公式、引用编号与数值

### 术语守卫

脚本在回填前逐段比对：Latin 术语、图表公式编号、数值三类 token 若在新文本中丢失，该段自动回退并记入 `skipped`。看到回退说明改写动了不该动的东西，改回来再跑。人工复核时同样按这三类检查。

## 输出规范

| 产物 | 位置 | 说明 |
| --- | --- | --- |
| 结构清单 | `analysis/thesis_structure.md` | 标题层级 + 包完整性 |
| 总揽规划 | `analysis/thesis_master_overview.md` | 章节矩阵、优先级、进度、危险片段 |
| 章节任务 | `analysis/chapter_*.md` | 每章目标、风险、处理记录 |
| 扫描报告 | `analysis/ai_scan.md` | 热力图 + 高危段落 + 指标 |
| 格式体检 | `analysis/format_check.md` | 合规分 + 必须改 / 建议核对 |
| 交付报告 | `analysis/delivery_report.md` | 全链路汇总 |
| 优化稿 | `final/*_优化稿_v1.docx` | 规则轮产物 |
| 最终版 | `final/*_最终版.docx` | 交付文件 |
| 规整版 | `final/*_格式规整.docx` | `--fix` 产物 |

交付时必须同时给出：改了多少段、哪些段落被守卫回退、扫描指标变化（score / cv）、体检剩余问题。数字不粉饰。

## 子资源

**知识文档（references/）**

| 文件 | 用途 |
| --- | --- |
| `@references/workflow-and-io-contract.md` | 全链路、输入输出契约、命令行、失败降级 |
| `@references/risk-rules.md` | 章节优先级、默认风险清单、段落保护清单 |
| `@references/ai-pattern-taxonomy.md` | 6 大类 33 种 AI 模式与对应改写动作（P01–P10 已量化） |
| `@references/ai-vocabulary-blacklist.md` | T1/T2/T3 三级词表（脚本直接读取） |
| `@references/perplexity-burstiness.md` | 检测原理、指标口径、自检流程 |
| `@references/strategy-ai-reduction.md` | 降 AI 痕迹技术规程 |
| `@references/strategy-plagiarism.md` | 降重与引用规范化 |
| `@references/strategy-polishing.md` | 学术润色与表达规范 |
| `@references/evaluation-criteria.md` | 5 维评分与评估流程 |
| `@references/thesis-format-spec.md` | 中文论文格式规范速查与 SPEC 自定义 |
| `@references/er-diagram-spec.md` | ER 图风格、schema 写法、美化流程 |

**脚本（scripts/）**

| 脚本 | 作用 | 依赖 |
| --- | --- | --- |
| `scripts/paper_pipeline.py` | parse / optimize / finalize / report 四子命令 | python-docx |
| `scripts/ai_scan.py` | 离线 AI 痕迹扫描（docx 或纯文本） | python-docx（读 docx 时） |
| `scripts/format_check.py` | 格式体检 + `--fix` 规整 + 三线表 | python-docx |
| `scripts/render_er_diagram.py` | ER 图渲染与美化包 | pyyaml、Graphviz |
| `scripts/docx_ops.py` | 内部共享库（文件发现、run 级替换、完整性校验），不单独调用 | python-docx |

**模板（templates/）**

| 文件 | 用途 |
| --- | --- |
| `templates/master-overview-template.md` | 总揽规划文档骨架（章节矩阵、优先级、进度） |
| `templates/chapter-task-template.md` | 单章任务文档骨架（目标、诊断、处理记录） |
| `templates/config.example.json` | `--config` 配置文件示例 |
| `templates/schema.example.yaml` | ER 图 schema 示例 |
| `templates/imagegen_prompt_table.md` | 单表 ER 图绘图 prompt |
| `templates/imagegen_prompt_system_simple.md` | 整体 ER 图绘图 prompt（简版） |
| `templates/imagegen_prompt_system_full.md` | 整体 ER 图绘图 prompt（详版） |

**素材（assets/）**：`assets/icon.png`（技能市场图标，1024×1024，非运行时依赖）

## 环境与降级

```bash
pip install python-docx pyyaml
```

- 缺 python-docx：改写与体检不可用，只能走 `ai_scan.py --text` 做纯文本扫描
- 缺 pyyaml：ER 图只能吃 `.json` schema
- 缺 Graphviz：ER 图无法出结构参考图，可直接用模板 prompt 让图像模型按描述画，但需人工核拓扑
- 论文 `.doc` 旧格式：先用 Word 另存为 `.docx`
- 段落含超链接或域代码：脚本会跳过并记录，这类段落请在 Word 里手工改

## 帮助命令

| 命令 | 回复内容 |
| --- | --- |
| `论文帮助` / `paper help` | 列出全链路六步与四个脚本用法 |
| `论文格式` / `paper format` | 摘录 `references/thesis-format-spec.md` 的字体字号页边距速查表 |
| `论文扫描` / `paper scan` | 说明 ai_scan 的指标含义与判定线 |
| `论文策略` / `paper strategy` | 列出三篇策略文档与适用场景 |
| `论文配置` / `paper config` | 说明 `templates/config.example.json` 各项含义 |

## 事实纪律

1. 不承诺任何检测平台的具体数值，风险分只作内部相对指标
2. 改写结果必须由作者复核，事实、数据、结论由作者负责
3. 学校格式以本校研究生院模板为准，本技能默认参数仅供对照
4. 涉及降重的具体操作只做语言层处理，不提供规避查重系统的技术手段
