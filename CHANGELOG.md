# 更新日志

## v1.0.0 — 2026-10-28

首个整合版本。把原先分散的 4 个技能（docx 优化流水线、AI 痕迹改写策略、docx 转 tex、ER 图渲染）收敛为一个技能，去掉与 docx 优化无关的 LaTeX 导出链路，统一了脚本入口与文档口径。

### 新增

- `scripts/paper_pipeline.py` 统一四子命令：`parse` / `optimize` / `finalize` / `report`
- `scripts/docx_ops.py` 共享库：文件发现、run 级替换、术语守卫、包完整性校验
- `scripts/ai_scan.py` 离线扫描器：P01–P10 模式正则、T1/T2/T3 词表、句长变异系数、章节热力图
- `scripts/format_check.py` 格式体检：13 类检查项、`--fix` 规整、`--fix-tables` 三线表
- 文档 `references/ai-pattern-taxonomy.md`（6 大类 33 种模式）
- 文档 `references/ai-vocabulary-blacklist.md`（128 词三级词表，脚本直接读取）
- 文档 `references/perplexity-burstiness.md`（指标口径与自检口径）
- 文档 `references/risk-rules.md`（章节优先级与段落保护清单）
- 文档 `references/thesis-format-spec.md`（中文论文格式规范速查与 SPEC 自定义）
- 文档 `references/er-diagram-spec.md`（ER 图风格与 schema 写法）
- 章节级任务文档模板，按 P0/P1/P2 自动生成改写目标与处理记录位

### 改进

- 回填方式从「整段替换」改为「run 内手术式替换」，混合格式段落（加粗术语、上标引用）不再丢格式
- 术语守卫覆盖三类 token（Latin 术语 / 图表公式编号 / 数值），拦截即回退并记录原因
- 扫描器最小段落字数与流水线口径统一为 12 字，短段落不再漏检
- 章节任务文档只引用本章区间内的报告片段，不再跨章借调
- 无标题层级文档、无命中段落等边界情况给出明确降级提示
- 交付报告补齐句长分布、变异系数判定与全部 docx 产物清单

### 移除

- LaTeX / tex 导出链路（与 docx 优化目标无关）
- 重复的 bootstrap / convert 脚本（能力合并进 `paper_pipeline.py`）
- 嵌套子目录引用，全部文件平铺在二级目录内，符合平台两级目录硬约束
