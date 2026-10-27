# 工作流与输入输出契约

## 一、全链路

```
输入目录                     工作目录                      交付
────────────────────────────────────────────────────────────────────────
论文.docx          ┌─→ ① 解析 parse      analysis/
AIGC报告.md（可选）─┤      结构清单          ├── headings.json
查重报告.md（可选）  │      章节区间          ├── thesis_structure.md
                    │      完整性校验        └── thesis_master_overview.md
                    │
                    ├─→ ② 扫描 ai_scan    analysis/
                    │      风险热力图        ├── ai_scan.json
                    │      高危段落          └── ai_scan.md
                    │
                    ├─→ ③ 优化 optimize   analysis/chapter_*.md
                    │      规划文档          analysis/rewrite_preview.md（dry-run）
                    │      规则改写 + 术语守卫 final/xxx_优化稿_v1.docx
                    │
                    ├─→ ④ 精修（人工/AI） 章节任务文档逐段改写
                    │      references/ 策略
                    │
                    ├─→ ⑤ 定稿 finalize   final/xxx_最终版.docx
                    │      完整性复检
                    │
                    └─→ ⑥ 体检 format_check  analysis/format_check.md
                           格式规整（--fix）  final/xxx_格式规整.docx
                           交付报告           analysis/delivery_report.md（report）
```

四道强制门禁：① 解析必须产出可用的标题结构；③ 改写必须通过术语守卫；⑤ 定稿必须通过包完整性校验；⑥ 体检严格问题必须清零或逐条说明。

## 二、输入契约

```text
input/
├── 论文原文.docx        # 必需，唯一主论文
├── AIGC检测报告.md      # 可选
└── 查重报告.md          # 可选
```

**主论文发现规则**

1. 输入目录只有一个 `*.docx` → 直接用它
2. 多个 docx → 按文件名命中 `论文` / `设计与实现` / `实现与设计` / `thesis` / `毕业论文` / `学位论文` 优先，其次取文件名最短者
3. 自动跳过 `~$` 开头的 Word 临时锁文件
4. 仍有歧义 → 用 `--docx-path` 显式指定（推荐在多文件目录中总是显式指定）

**报告发现规则**（可选）

按文件名特征排序，取第一个命中：`aigc` > `检测报告` > `查重` > `similarity` > `report`。报告支持三种结构：

```markdown
<!-- 1. Markdown 表格 -->
| 序号 | 片段 | 风险 |
| --- | --- | --- |
| 1 | 近年来，随着…… | 高 |

<!-- 2. HTML 表格（检测平台导出） -->
<tr><td>1</td><td>近年来，随着……</td><td>高</td></tr>

<!-- 3. 项目符号 -->
- [高] 近年来，随着……
```

## 三、输出契约

```text
output/
├── analysis/
│   ├── headings.json               # 标题条目（index/style/level/text）
│   ├── thesis_structure.md         # 人类可读结构清单
│   ├── thesis_master_overview.md   # 总揽：章节矩阵、优先级、进度、危险片段
│   ├── chapter_00_摘要.md          # 章节任务文档（每章一份）
│   ├── rewrite_preview.md          # 仅 dry-run 生成
│   ├── ai_scan.json / ai_scan.md   # AI 痕迹扫描
│   ├── format_check.json / .md     # 格式体检
│   └── delivery_report.md          # 汇总交付报告
├── intermediate/                   # 预留：外部转换产物（docx2tex 等）
├── final/
│   ├── xxx_优化稿_v1.docx
│   ├── xxx_最终版.docx
│   └── xxx_格式规整.docx           # 仅 --fix 生成
├── logs/                           # 预留：外部工具日志
├── manifest.json                   # 输入快照与完整性结论
├── pipeline_result.json            # 改写明细（changes / skipped）
└── finalize_result.json            # 定稿路径与复检结论
```

**manifest.json 字段**

| 字段 | 说明 |
| --- | --- |
| `docx_path` | 主论文绝对路径 |
| `report_path` | 检测报告路径，无则 `null` |
| `input_dir` / `output_dir` | 工作目录 |
| `headings_path` | 标题 JSON 路径 |
| `heading_count` | 标题条目数（含各级） |
| `docx_size_bytes` | 原始文件大小 |
| `docx_integrity_ok` | 包完整性（zip + document.xml 可解析） |
| `docx_integrity_errors` | 完整性问题列表 |
| `created_at` | 解析时间 |

**pipeline_result.json 关键字段**：`output_path`、`dry_run`、`chapter_count`、`enabled_chapters`、`changed_paragraphs`、`skipped_paragraphs`、`changes[]`（含 `patterns` / `before` / `after`）、`skipped[]`（含回退原因）。

## 四、命令行契约

```bash
# ① 解析
python scripts/paper_pipeline.py parse --input-dir input --output-dir out

# ② 扫描（可对原稿或任一版本扫描）
python scripts/ai_scan.py --docx input/论文.docx --out out/analysis

# ③ 优化：先预览，再落盘
python scripts/paper_pipeline.py optimize --input-dir input --output-dir out --dry-run
python scripts/paper_pipeline.py optimize --input-dir input --output-dir out

# 只处理部分章节
python scripts/paper_pipeline.py optimize --input-dir input --output-dir out \
  --chapter-allowlist "摘要,相关技术,总结与展望"
python scripts/paper_pipeline.py optimize --input-dir input --output-dir out \
  --chapter-blocklist "绪论,需求分析"

# ④ 定稿
python scripts/paper_pipeline.py finalize --output-dir out

# ⑤ 体检 / 规整
python scripts/format_check.py --docx out/final/论文_最终版.docx --out out/analysis
python scripts/format_check.py --docx out/final/论文_最终版.docx --out out/analysis --fix --fix-tables

# ⑥ 汇总
python scripts/paper_pipeline.py report --output-dir out

# 附图
python scripts/render_er_diagram.py --input schema.yaml --output-dir out/diagrams \
  --mode system-simple --format svg,png --imagegen-pack
```

所有子命令都支持 `--config`，配置文件键名与命令行参数同名（下划线形式），示例见 `templates/config.example.json`。

## 五、章节过滤语义

- `--chapter-allowlist`：只保留标题**包含**任一关键词的章节
- `--chapter-blocklist`：排除标题包含任一关键词的章节
- 两者同时给出：先白名单过滤，再黑名单排除
- 过滤只影响改写范围，不影响结构抽取与总览生成（被排除的章节在矩阵中标记为「跳过」）

## 六、失败与降级

| 故障 | 表现 | 处理 |
| --- | --- | --- |
| 未安装 python-docx | 脚本报 ImportError | `pip install python-docx`；或改用 `ai_scan.py --text` 走纯文本链路 |
| docx 是 `.doc` 旧格式 | 无法打开 | 先用 Word 另存为 .docx |
| 未识别到标题样式 | 结构清单为空 | 在 Word 中给章节套用「标题 1/2/3」；或按 `references/thesis-format-spec.md` 用 `--fix` 统一 |
| 段落含超链接/域 | 该段被跳过并记入 `skipped` | 手工在 Word 中改，不要用脚本强改 |
| 改写丢失术语/编号 | 该段回退，记入 `skipped` | 检查是否误删了术语；调整词表后重跑 |
| 论文体积过大（>50MB，多图） | 处理变慢 | 先删图片副本做流程验证，最后再对正式文件执行 |
| 格式体检严格问题未清零 | 退出码 1 | 按报告逐项修复；每校细则不同，以本校模板为准 |

## 七、并发与版本

- 同一输出目录不要并行跑多个 `optimize`（manifest 与 pipeline_result 会互相覆盖）
- 迭代产出一律 `_优化稿_v1.docx` → `_v2` → `_最终版.docx`，**永不覆盖原始 docx**
- 需要回退：`final/` 下保留每一版，不看 `pipeline_result.json` 也能找到历史
