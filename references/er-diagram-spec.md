# 论文 ER 图规范与生成流程

> 用途：系统设计章节的数据库图。风格统一为「中文论文截图风」：白底黑线、矩形实体、椭圆属性、菱形关系、1/N/M 基数、主键短斜杠。

## 一、何时画、画哪种

| 场景 | 模式 | 说明 |
| --- | --- | --- |
| 介绍单张表的字段构成 | `table` | 一个实体 + 环绕属性椭圆，每个实体一张图 |
| 展示整体数据库设计（完整版） | `system-full` | 实体 + 关系菱形 + 属性椭圆 + 基数 |
| 展示整体数据库设计（精简版） | `system-simple` | 只保留实体、关系、基数，去掉属性椭圆 |

学位论文常用组合：**整体精简图（一章一张）+ 关键单表图（2–3 张）**。整体完整图容易拥挤，除非表少于 5 张，否则不建议。

不适用：UML 类图、流程图、时序图、架构图。

## 二、固定风格

- 白底黑线，无颜色、无阴影、无渐变、无装饰、无箭头、无水印
- 实体/表名 → **矩形**；属性 → **椭圆**；关系 → **菱形**
- 基数写在连线旁：`1`、`N`、`M`
- 主键：只在「实际主键属性」到实体的连线上加一道短斜杠；不要把所有编号类属性都当主键，斜杠也不画在文字上
- 中文字体偏宋体/论文截图风格，字号偏大，留白充分
- 中文标签必须与 schema 完全一致，一个字都不能错

### 单表图构图

- 中心实体矩形居中，宽高比约 1.6:1
- 属性椭圆围绕实体**放射状环布**，不得上下堆叠成列表
- 连线为细直线或轻微斜线，不带箭头

### 整体图构图

- 优先用真实数据表作节点，**不要**用「资源汇总」「业务对象」等抽象分组替代表
- 布局依据真实外键关系，避免把多张表画成彼此孤立的模块
- 关系菱形分布在实体之间，不得重叠
- 基数靠近关系或实体端
- 想控制左-中-右布局时，在 schema 里给实体加 `rank: left | middle | right`

## 三、schema 写法定

```yaml
entities:
  - name: 用户
    primary_key: 编号          # 必须出现在 attributes 中，且字面完全一致
    rank: right                # 可选：left / middle / right
    attributes: [编号, 用户名, 密码, 头像, 角色]

relationships:
  - name: 收藏
    from: 用户
    to: 服装
    from_cardinality: "1"
    to_cardinality: "N"
    attributes: [收藏时间]      # 可选：关系上的属性
```

约束（脚本会校验并报错）：

- `primary_key` 必须在 `attributes` 里
- 实体名唯一；`from` / `to` 必须指向已声明实体
- 只接受 `.yaml` / `.yml` / `.json`；YAML 需要 `pyyaml`

模板见 `templates/schema.example.yaml`。

## 四、生成命令

```bash
# 单表图（第 5 个实体）
python scripts/render_er_diagram.py --input schema.yaml --output-dir out/diagrams \
  --mode table --entity-index 5 --format svg,png --imagegen-pack

# 整体完整图
python scripts/render_er_diagram.py --input schema.yaml --output-dir out/diagrams \
  --mode system-full --format svg,png --imagegen-pack

# 整体精简图（论文推荐）
python scripts/render_er_diagram.py --input schema.yaml --output-dir out/diagrams \
  --mode system-simple --format svg,png --imagegen-pack
```

常用参数：

| 参数 | 作用 |
| --- | --- |
| `--mode` | `table` / `system-full` / `system-simple` |
| `--entity` / `--entity-index` | 单表模式下指定实体（索引从 1 开始） |
| `--format` | `svg`、`png` 或 `svg,png` |
| `--imagegen-pack` | 额外产出结构参考图、`imagegen_prompt.md`、`labels.json` |
| `--dot` | Graphviz `dot.exe` 路径（默认 `C:\Program Files\Graphviz\bin\dot.exe`） |
| `--layout` | `dot` / `neato` |
| `--font` | 默认 `Microsoft YaHei`，正式插图建议 `SimSun` |
| `--safe-filenames` | 输出纯 ASCII 文件名，避免中文路径在终端里出乱码 |

产出：

```text
out/diagrams/
├── table_05.png / .svg          # 结构图（可用于直接插入论文）
├── imagegen_prompt.md           # 美化用提示词
├── labels.json                  # 全部中文标签清单，用于校对
└── reference/                   # 交给图像模型的结构参考图
```

## 五、美化工作流（结构图 → 论文插图）

Graphviz 输出用于**结构校验与追溯**，最终插图用图像模型重绘：

1. 带 `--imagegen-pack` 跑一遍脚本，得到结构参考图与 `imagegen_prompt.md`
2. 读 `imagegen_prompt.md`，以结构参考图作参考图向图像模型提交 prompt
3. 检查四点：形状对不对、连线对不对、主键斜杠在不在、中文标签与 `1/N/M` 是否逐字一致
4. 文字有错 → 用同一 prompt 加强「文字必须逐字一致」重生成一次；仍不稳定 → 保留线稿风格，用后处理覆盖文字
5. 通过后按论文题注规范命名（如 `图4-3 数据库 ER 图`），插入正文

Prompt 必备五要素（脚本已写入模板，`templates/imagegen_prompt_*.md`）：

- Use case: `infographic-diagram`
- Asset type: `Chinese thesis ER diagram`
- Input image: Graphviz 结构参考图，仅用于拓扑
- Style: 黑白论文截图风、干净手绘感
- Constraints: 中文标签逐字一致、无色、无阴影、无箭头、无水印

## 六、常见问题

| 现象 | 处理 |
| --- | --- |
| 找不到 Graphviz | 安装 Graphviz，或用 `--dot` 指定 `dot.exe` |
| 中文字体发虚/乱码 | `--font "SimSun"` 或 `--font "Microsoft YaHei"` |
| 主键没斜杠 | 检查 `primary_key` 与 `attributes` 是否字面完全一致 |
| 整体图太挤 | 改用 `system-simple`，或在 schema 中给实体加 `rank` |
| 终端里中文实体名显示异常 | 用 `--entity-index` 选表，并加 `--safe-filenames` |
| 图像模型把中文写错 | 用 `labels.json` 逐字核对，重生成或后处理覆盖文字 |
| 关系方向搞反 | `from` 是「1」端，`to` 是「N」端时写 `from_cardinality: "1"`、`to_cardinality: "N"` |
