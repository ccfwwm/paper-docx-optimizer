# 论文 DOCX 优化器 · paper-docx-optimizer

把「论文导出成 docx，要降 AI 痕迹、降重、润色、还得符合学校格式」做成一条可重复执行的流水线。

核心只有三件事：**定位**（哪里风险高）、**改写**（只动该动的句子）、**回填**（版式一个字都不变）。

适用于中文学位论文（本科 / 硕士 / 博士）与期刊投稿稿件的成稿优化，不代写、不洗稿、不编造数据。

---

## 它解决什么问题

| 痛点 | 本技能的做法 |
| --- | --- |
| 通篇改写会把论文改坏 | 只改高风险段落，改动量控制在 30% 以内，术语与数值逐字保留 |
| 改完 Word 版式全乱 | 回填走 run 级替换，字体、字号、加粗、上标、题注全部保留 |
| 不知道 AI 味从哪来 | 33 种可观察模式 + 三级词表 + 句长分布量化，给出定位清单 |
| 改完不知道有没有变好 | 改写前后各扫一次，风险分与变异系数必须改善，否则回退该段 |
| 格式检查靠肉眼 | 纸张 / 边距 / 字体 / 字号 / 行距 / 缩进 / 题注 / 三线表 / 页码自动体检并可一键规整 |
| 系统设计章节缺 ER 图 | schema → Graphviz 结构图 → 图像模型重绘为论文插图风 |

---

## 全链路

```
input/论文.docx（+ 可选 AIGC/查重报告）
   │
   ├─ ① parse     结构解析：标题层级、章节段落区间、包完整性   → analysis/headings.json、manifest.json
   ├─ ② scan      AI 痕迹扫描：模式命中、词表、句长分布        → analysis/ai_scan.md
   ├─ ③ optimize  章节规划 + 规则改写 + 术语守卫 + run 级回填  → analysis/chapter_*.md、final/*_优化稿_v1.docx
   ├─ ④ refine    按 references/ 策略逐段精修（人工 / AI 协作） → 章节任务文档 + 逐版 docx
   ├─ ⑤ finalize  另存最终版 + 完整性复检                     → final/*_最终版.docx
   └─ ⑥ audit     格式体检 / 规整 + 交付汇总                  → format_check.md、delivery_report.md
```

四道门禁：解析必须拿到可用标题结构；改写必须通过术语守卫；定稿必须通过包完整性；体检严格问题必须清零或逐条说明。

---

## 快速开始

```bash
pip install python-docx pyyaml
```

```bash
# 1. 先预览改写幅度（不写盘）
python scripts/paper_pipeline.py optimize --input-dir ./input --output-dir ./out --dry-run

# 2. 确认后落盘
python scripts/paper_pipeline.py optimize --input-dir ./input --output-dir ./out

# 3. 定稿
python scripts/paper_pipeline.py finalize --output-dir ./out

# 4. 格式体检（加 --fix 可一键规整，--fix-tables 统一三线表）
python scripts/format_check.py --docx ./out/final/论文_最终版.docx --out ./out/analysis --fix --fix-tables

# 5. 交付汇总
python scripts/paper_pipeline.py report --output-dir ./out
```

单独扫一份文档：

```bash
python scripts/ai_scan.py --docx 论文.docx --out ./out/analysis
python scripts/ai_scan.py --text 论文.md --out ./out/analysis   # 纯文本模式
```

ER 图：

```bash
python scripts/render_er_diagram.py --input schema.yaml --output-dir ./out/diagrams \
  --mode system-simple --format svg,png --imagegen-pack
```

---

## 目录结构

```
paper-docx-optimizer/
├── SKILL.md                              # 技能定义与执行步骤（AI 读这个）
├── references/
│   ├── workflow-and-io-contract.md       # 全链路、输入输出契约、失败降级
│   ├── risk-rules.md                     # 章节优先级、默认风险清单、段落保护清单
│   ├── ai-pattern-taxonomy.md            # 6 大类 33 种 AI 模式 → 改写动作
│   ├── ai-vocabulary-blacklist.md        # T1/T2/T3 三级词表（脚本直接读取）
│   ├── perplexity-burstiness.md          # 检测原理、指标口径、自检流程
│   ├── strategy-ai-reduction.md          # 降 AI 痕迹技术规程
│   ├── strategy-plagiarism.md            # 降重与引用规范化
│   ├── strategy-polishing.md             # 学术润色与表达规范
│   ├── evaluation-criteria.md            # 5 维评分与评估流程
│   ├── thesis-format-spec.md             # 中文论文格式规范速查
│   └── er-diagram-spec.md                # ER 图风格与 schema 写法
├── scripts/
│   ├── paper_pipeline.py                 # parse / optimize / finalize / report
│   ├── ai_scan.py                        # 离线 AI 痕迹扫描
│   ├── format_check.py                   # 格式体检 + 规整 + 三线表
│   ├── render_er_diagram.py              # ER 图渲染与美化包
│   └── docx_ops.py                       # 共享库：文件发现、run 级替换、完整性校验
├── templates/                            # 规划文档模板、配置示例、ER schema 示例、绘图 prompt
└── assets/icon.png                       # 市场图标
```

---

## 铁律

1. **原稿只读** —— 任何情况下不覆盖输入 docx，产物一律写入 `output/final/`
2. **最小干预** —— 只做句内重组、语序调整、局部替换；禁止整段推翻、禁止扩写
3. **术语与编号零改动** —— 模型名、模块名、缩写、图表公式文献编号、数值、单位逐字保留
4. **保版式** —— 回填走 run 级替换，不重建段落与表格
5. **不编造** —— 不新增文献、数据、实验结论；缺依据处只做语言收紧
6. **不降质** —— 禁止为降 AI 引入语法错误、口语词、错误术语
7. **改完必复检** —— 指标不改善就回退该段，不靠感觉交付

---

## 实测效果（仓库自带样例）

| 指标 | 改写前 | 规则轮后 | 说明 |
| --- | --- | --- | --- |
| AI 痕迹风险分 | 43.7 | 24.2 | 同一扫描器口径，越低越好 |
| 高风险段落 | 2 | 1 | 剩余项需人工逐段精修 |
| 格式合规分 | 85 | 94 | `--fix --fix-tables` 之后 |
| 严格格式问题 | 0 | 0 | 提醒项为学校可选要求 |

风险分是本技能内置扫描器的相对指标，仅用于前后对比，**不承诺任何第三方检测平台的具体数值**。

---

## 环境与降级

| 缺什么 | 影响 | 兜底 |
| --- | --- | --- |
| python-docx | 改写与体检不可用 | 只能 `ai_scan.py --text` 做纯文本扫描 |
| pyyaml | ER 图读不了 YAML schema | 改用 `.json` schema |
| Graphviz | 出不了结构参考图 | 直接用模板 prompt 让图像模型按描述画，需人工核拓扑 |
| `.doc` 旧格式 | 无法读取 | 先用 Word 另存为 `.docx` |

段落含超链接或域代码时脚本会跳过并记录原因，这类段落请在 Word 里手工修改。

---

## 边界与免责

- 只做**语言层**处理，不提供规避查重 / 检测系统的技术手段
- 事实、数据、结论与引用关系必须由作者复核并负责
- 学校格式以本校研究生院模板为准，内置参数仅供对照（可在 `scripts/format_check.py` 的 `SPEC` 中修改）
- 从零生成整篇论文、PDF/Markdown 单格式处理、UML 类图 / 流程图 / 时序图不属于本技能范围

---

## 许可

MIT License，见 [LICENSE](LICENSE)。
