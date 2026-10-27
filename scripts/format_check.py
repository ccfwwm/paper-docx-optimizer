#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文学位论文 DOCX 格式体检与规整。

两件事：
1. 体检：对照《研究生学位论文格式的规定》默认参数，逐项检查页面、字体字号、
   行距缩进、标题层级、图表题注、三线表、页眉页码、结构完整性，输出可读报告与 JSON。
2. 规整（--fix）：把页面、正文样式、标题样式、题注缩进改回规范值；
   --fix-tables 额外把表格统一为三线表。

用法：
    python format_check.py --docx paper.docx --out out/
    python format_check.py --docx paper.docx --out out/ --fix
    python format_check.py --docx paper.docx --out out/ --fix --fix-tables --save-as out/fixed.docx

依赖：python-docx
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

from docx_ops import is_caption_style, normalize_style_name  # noqa: E402

# ---------------------------------------------------------------- 规范参数

SPEC = {
    "page": {"width_twips": 11906, "height_twips": 16838, "tolerance_twips": 40},
    "margins_cm": {"top": 2.2, "bottom": 2.2, "left": 2.7, "right": 2.2, "tolerance_cm": 0.15},
    "body": {"east_asia": "宋体", "latin": "Cambria Math", "size_pt": 12, "line_multiple": 1.5},
    "headings": {
        1: {"east_asia": "黑体", "size_pt": 18, "bold": True, "align": "center"},
        2: {"east_asia": "黑体", "size_pt": 16, "bold": True, "align": "left"},
        3: {"east_asia": "黑体", "size_pt": 15, "bold": False, "align": "left"},
    },
    "caption_size_pt": 10.5,
    "first_line_indent_pt": 24,  # 小四字号下 2 字符
}

CAPTION_TEXT_RE = re.compile(r"^\s*(图|表|Fig\.?|Table)\s*[0-9０-９]+(?:[\-\.－][0-9０-９]+)*")
CHAPTER_TITLE_RE = re.compile(r"^\s*(第[一二三四五六七八九十百]+章|\d+(?:\.\d+)?\s+\S)")
CITATION_RE = re.compile(r"\[\d+(?:[,\-–]\d+)*\]")
# 前后置部分：不参与「第X章」编号要求
MATTER_TITLES = (
    "摘要", "摘 要", "摘　要", "abstract", "目录", "目 录",
    "参考文献", "致谢", "后记", "附录", "结论", "总结与展望", "绪论",
)

findings: list[dict] = []


def add(level: str, item: str, detail: str, count: int = 1) -> None:
    findings.append({"level": level, "item": item, "detail": detail, "count": count})


# ---------------------------------------------------------------- 度量工具

def east_asia_font(style) -> str | None:
    try:
        rPr = style.element.rPr
        if rPr is None:
            return None
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            return None
        return rFonts.get(qn("w:eastAsia"))
    except Exception:
        return None


def size_pt(style) -> float | None:
    try:
        if style.font.size is None:
            return None
        return style.font.size.pt
    except Exception:
        return None


def line_spacing(style) -> float | None:
    try:
        value = style.paragraph_format.line_spacing
        return value if isinstance(value, float) else None
    except Exception:
        return None


def cm(twips: int) -> float:
    return round(twips / 566.93, 2)


def paragraph_text(paragraph) -> str:
    return paragraph.text.strip()


def get_heading_style(document, level: int):
    """按名称或 style_id 定位标题样式，兼容中英文 Word。"""
    for name in (f"Heading {level}", f"标题 {level}", f"标题{level}"):
        try:
            return document.styles[name]
        except KeyError:
            continue
    for style in document.styles:
        if getattr(style, "style_id", "") == f"Heading{level}":
            return style
    return None


# ---------------------------------------------------------------- 各项检查

def check_page(section) -> None:
    spec_page = SPEC["page"]
    width, height = section.page_width.twips, section.page_height.twips
    if abs(width - spec_page["width_twips"]) > spec_page["tolerance_twips"] or abs(
        height - spec_page["height_twips"]
    ) > spec_page["tolerance_twips"]:
        add("严格", "纸张尺寸", f"当前 {cm(width)}×{cm(height)} cm，规范为 A4 21×29.7 cm")
    else:
        add("通过", "纸张尺寸", "A4")

    spec_m = SPEC["margins_cm"]
    actual = {
        "top": cm(section.top_margin.twips),
        "bottom": cm(section.bottom_margin.twips),
        "left": cm(section.left_margin.twips),
        "right": cm(section.right_margin.twips),
    }
    bad = [
        f"{key} {actual[key]}cm≠{spec_m[key]}cm"
        for key in spec_m
        if key != "tolerance_cm" and abs(actual[key] - spec_m[key]) > spec_m["tolerance_cm"]
    ]
    if bad:
        add("提醒", "页边距", "；".join(bad))
    else:
        add("通过", "页边距", "上2.2/下2.2/左2.7/右2.2 cm")


def check_body_style(document) -> None:
    try:
        normal = document.styles["Normal"]
    except KeyError:
        add("严格", "正文样式", "缺少 Normal 样式，无法判定正文基准格式")
        return

    font = east_asia_font(normal)
    expected = SPEC["body"]["east_asia"]
    if font is None:
        add("提醒", "正文字体", "未在样式层显式设置中文字体（可能继承主题或段落直排格式）")
    elif font != expected:
        add("提醒", "正文字体", f"当前 {font}，规范为 {expected}")
    else:
        add("通过", "正文字体", expected)

    body_size = size_pt(normal)
    if body_size is None:
        add("提醒", "正文字号", "样式层未设置字号，需人工确认正文为小四（12pt）")
    elif abs(body_size - SPEC["body"]["size_pt"]) > 0.6:
        add("提醒", "正文字号", f"当前 {body_size}pt，规范为小四 12pt")
    else:
        add("通过", "正文字号", "小四 12pt")

    spacing = line_spacing(normal)
    if spacing is None:
        add("提醒", "正文行距", "样式层未设置行距，需人工确认是否为 1.5 倍")
    elif abs(spacing - SPEC["body"]["line_multiple"]) > 0.06:
        add("提醒", "正文行距", f"当前 {spacing} 倍，规范为 1.5 倍")
    else:
        add("通过", "正文行距", "1.5 倍")


def check_body_paragraphs(document) -> dict:
    total = 0
    indented = 0
    long_no_indent = []
    for paragraph in document.paragraphs:
        text = paragraph_text(paragraph)
        if len(text) < 20 or CAPTION_TEXT_RE.match(text):
            continue
        style = normalize_style_name(paragraph.style.name if paragraph.style else "")
        if style.startswith("Heading") or is_caption_style(style):
            continue
        total += 1
        indent = paragraph.paragraph_format.first_line_indent
        if indent is not None and indent.pt > 6:
            indented += 1
        elif len(long_no_indent) < 5:
            long_no_indent.append(text[:40])

    if total == 0:
        add("提醒", "首行缩进", "未找到可比对的正文段落")
        return {"total": 0, "indented": 0}
    ratio = indented / total
    if ratio >= 0.85:
        add("通过", "首行缩进", f"{indented}/{total} 段落含首行缩进（{ratio:.0%}）")
    elif ratio >= 0.5:
        add("提醒", "首行缩进", f"仅 {indented}/{total} 段落含首行缩进，示例：{long_no_indent[:1]}")
    else:
        add("严格", "首行缩进", f"仅 {indented}/{total} 段落含首行缩进，正文普遍缺 2 字符缩进")
    return {"total": total, "indented": indented}


def check_headings(document) -> None:
    levels: dict[int, int] = {}
    h1_texts = []
    for paragraph in document.paragraphs:
        style = normalize_style_name(paragraph.style.name if paragraph.style else "")
        m = re.match(r"^Heading (\d)$", style)
        if not m:
            continue
        level = int(m.group(1))
        levels[level] = levels.get(level, 0) + 1
        if level == 1:
            h1_texts.append(paragraph_text(paragraph))

    if not levels:
        add("严格", "标题层级", "未使用「标题 1/2/3」样式，目录与交叉引用无法自动生成")
        return
    add("通过", "标题层级", "；".join(f"H{lv} {cnt} 个" for lv, cnt in sorted(levels.items())))

    if levels.get(1, 0) == 0:
        add("严格", "章标题", "缺少一级标题（章），无法生成论文目录")

    odd = [
        t for t in h1_texts
        if not CHAPTER_TITLE_RE.match(t)
        and t.strip().lower() not in MATTER_TITLES
        and not t.strip().startswith(("附录", "结论"))
    ]
    if odd:
        add("提醒", "章标题序号", f"{len(odd)} 个章标题未采用「第X章 / 数字」序号，如：{odd[0][:20]}")

    for level, spec in SPEC["headings"].items():
        style = get_heading_style(document, level)
        if style is None:
            continue
        font = east_asia_font(style)
        head_size = size_pt(style)
        issues = []
        if font and font != spec["east_asia"]:
            issues.append(f"字体 {font}≠{spec['east_asia']}")
        if head_size and abs(head_size - spec["size_pt"]) > 0.6:
            issues.append(f"字号 {head_size}pt≠{spec['size_pt']}pt")
        if issues:
            add("提醒", f"H{level} 标题格式", "；".join(issues))


def check_captions(document) -> dict:
    captions = []
    misplaced = 0
    for paragraph in document.paragraphs:
        text = paragraph_text(paragraph)
        if not CAPTION_TEXT_RE.match(text):
            continue
        style = normalize_style_name(paragraph.style.name if paragraph.style else "")
        is_fig = text.lstrip().startswith(("图", "Fig"))
        captions.append({"text": text[:40], "style": style, "kind": "图" if is_fig else "表"})
        if paragraph.paragraph_format.first_line_indent is not None and paragraph.paragraph_format.first_line_indent.pt > 6:
            misplaced += 1
    if not captions:
        add("提醒", "图表题注", "未识别到「图X-X」「表X.X」题注")
    else:
        add("通过", "图表题注", f"识别到 {len(captions)} 条题注")
        if misplaced:
            add("提醒", "题注缩进", f"{misplaced} 条题注带首行缩进，题注应顶格居中")
    return {"count": len(captions), "items": captions[:6]}


def check_tables(document) -> dict:
    if not document.tables:
        add("提醒", "三线表", "文档中没有表格")
        return {"count": 0, "three_line": 0}
    three_line = 0
    samples = []
    for table in document.tables:
        tblPr = table._tbl.tblPr
        borders = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        if borders is None:
            samples.append("未设置表格边框（可能是默认网格线）")
            continue
        top = borders.find(qn("w:top"))
        bottom = borders.find(qn("w:bottom"))
        inside_h = borders.find(qn("w:insideH"))
        inside_v = borders.find(qn("w:insideV"))

        def is_on(node):
            return node is not None and node.get(qn("w:val")) not in (None, "none", "nil")

        if is_on(top) and is_on(bottom) and not is_on(inside_h) and not is_on(inside_v):
            three_line += 1
        else:
            samples.append("存在内框线或缺少顶/底框线")
    if three_line == len(document.tables):
        add("通过", "三线表", f"{three_line}/{len(document.tables)} 张表符合三线表")
    else:
        add("提醒", "三线表", f"{three_line}/{len(document.tables)} 张表符合三线表；{samples[0] if samples else ''}")
    return {"count": len(document.tables), "three_line": three_line}


def check_sections_and_references(document) -> None:
    all_text = "\n".join(paragraph_text(p) for p in document.paragraphs)
    required = {
        "中文摘要": ("摘要",),
        "英文摘要": ("Abstract", "ABSTRACT"),
        "目录": ("目录", "目　录", "目 录"),
        "参考文献": ("参考文献",),
        "致谢": ("致谢", "后记"),
    }
    missing = [name for name, keys in required.items() if not any(k in all_text for k in keys)]
    if missing:
        add("提醒", "结构完整性", f"未检出：{'、'.join(missing)}")
    else:
        add("通过", "结构完整性", "摘要/Abstract/目录/参考文献/致谢 齐备")

    citations = CITATION_RE.findall(all_text)
    if not citations:
        add("提醒", "引用标注", "正文未检出 [n] 形式引用标注")
    else:
        add("通过", "引用标注", f"检出 {len(citations)} 处 [n] 引用标注")

    for section in document.sections:
        has_field = any(
            "PAGE" in (elem.text or "")
            for t in section.footer.paragraphs
            for elem in t._p.iter(qn("w:instrText"))
        )
        if not has_field:
            add("提醒", "页码", "页脚未检测到 PAGE 页码域")
            break
    else:
        add("通过", "页码", "页脚含 PAGE 域")


# ---------------------------------------------------------------- 规整

def set_style_font(style, east_asia: str, latin: str, size_pt_value: float, bold: bool | None = None) -> None:
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement

        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), east_asia)
    style.font.size = Pt(size_pt_value)
    if bold is not None:
        style.font.bold = bold


def fix_page(document) -> list[str]:
    changed = []
    for section in document.sections:
        spec_m = SPEC["margins_cm"]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(spec_m["top"])
        section.bottom_margin = Cm(spec_m["bottom"])
        section.left_margin = Cm(spec_m["left"])
        section.right_margin = Cm(spec_m["right"])
    changed.append("页面统一为 A4、边距 2.2/2.2/2.7/2.2 cm")
    return changed


def fix_styles(document) -> list[str]:
    changed = []
    body = SPEC["body"]
    try:
        normal = document.styles["Normal"]
        set_style_font(normal, body["east_asia"], body["latin"], body["size_pt"])
        normal.paragraph_format.line_spacing = body["line_multiple"]
        changed.append("正文样式：宋体/Cambria Math、小四、1.5 倍行距")
    except KeyError:
        pass

    for level, spec in SPEC["headings"].items():
        style = get_heading_style(document, level)
        if style is None:
            continue
        set_style_font(style, spec["east_asia"], "Cambria Math", spec["size_pt"], spec["bold"])
        style.paragraph_format.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER if spec["align"] == "center" else WD_ALIGN_PARAGRAPH.LEFT
        )
        style.paragraph_format.first_line_indent = Pt(0) if level == 1 else Pt(24)
        changed.append(f"H{level} 标题：{spec['east_asia']} {spec['size_pt']}pt")
    return changed


def fix_captions(document) -> list[str]:
    fixed = 0
    for paragraph in document.paragraphs:
        text = paragraph_text(paragraph)
        if not CAPTION_TEXT_RE.match(text):
            continue
        style = normalize_style_name(paragraph.style.name if paragraph.style else "")
        if style.startswith("Heading"):
            continue
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fixed += 1
    return [f"{fixed} 条图表题注取消缩进并居中"] if fixed else []


def fix_tables(document) -> list[str]:
    from docx.oxml import OxmlElement

    fixed = 0
    for table in document.tables:
        tblPr = table._tbl.tblPr
        existing = tblPr.find(qn("w:tblBorders"))
        if existing is not None:
            tblPr.remove(existing)
        borders = OxmlElement("w:tblBorders")
        for edge, size in (("top", 12), ("bottom", 12), ("left", 0), ("right", 0), ("insideH", 0), ("insideV", 0)):
            node = OxmlElement(f"w:{edge}")
            if size:
                node.set(qn("w:val"), "single")
                node.set(qn("w:sz"), str(size))
                node.set(qn("w:space"), "0")
                node.set(qn("w:color"), "000000")
            else:
                node.set(qn("w:val"), "none")
                node.set(qn("w:sz"), "0")
            borders.append(node)
        tblPr.append(borders)

        if table.rows:
            header = table.rows[0]
            for cell in header.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                old = tcPr.find(qn("w:tcBorders"))
                if old is not None:
                    tcPr.remove(old)
                tcBorders = OxmlElement("w:tcBorders")
                bottom = OxmlElement("w:bottom")
                bottom.set(qn("w:val"), "single")
                bottom.set(qn("w:sz"), "6")
                bottom.set(qn("w:space"), "0")
                bottom.set(qn("w:color"), "000000")
                tcBorders.append(bottom)
                tcPr.append(tcBorders)
        fixed += 1
    return [f"{fixed} 张表统一为三线表"] if fixed else []


# ---------------------------------------------------------------- 主流程

def run_checks(document) -> None:
    findings.clear()
    if document.sections:
        check_page(document.sections[0])
    check_body_style(document)
    check_body_paragraphs(document)
    check_headings(document)
    check_captions(document)
    check_tables(document)
    check_sections_and_references(document)


def summarize() -> dict:
    errors = sum(1 for f in findings if f["level"] == "严格")
    warnings = sum(1 for f in findings if f["level"] == "提醒")
    passed = sum(1 for f in findings if f["level"] == "通过")
    score = max(0, 100 - errors * 8 - warnings * 3)
    return {"score": score, "errors": errors, "warnings": warnings, "passed": passed}


def build_markdown(result: dict) -> str:
    lines = [
        "# 论文格式体检报告",
        "",
        f"- 文档: `{result['docx_path']}`",
        f"- 合规分: **{result['score']} / 100**",
        f"- 严格问题 {result['errors']} 项 / 提醒 {result['warnings']} 项 / 通过 {result['passed']} 项",
        "",
    ]
    if result.get("fixes"):
        lines += ["## 已执行的规整", ""] + [f"- {item}" for item in result["fixes"]] + [""]

    groups = [("严格", "❌ 必须修改"), ("提醒", "⚠️ 建议核对"), ("通过", "✅ 通过项")]
    for level, title in groups:
        items = [f for f in result["findings"] if f["level"] == level]
        if not items:
            continue
        lines += [f"## {title}", ""]
        for item in items:
            lines.append(f"- **{item['item']}**：{item['detail']}")
        lines.append("")

    lines += [
        "## 规范依据",
        "",
        "- 详见 references/thesis-format-spec.md（《研究生学位论文格式的规定》默认参数速查）",
        "- 各校细则不同，以本校研究生院模板为准；`--fix` 前务必确认参数",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="中文学位论文 DOCX 格式体检与规整")
    parser.add_argument("--docx", required=True, help="待检查的 .docx")
    parser.add_argument("--out", help="报告输出目录（默认与输入同目录）")
    parser.add_argument("--fix", action="store_true", help="规整页面、正文与标题样式、题注缩进")
    parser.add_argument("--fix-tables", action="store_true", help="把表格统一为三线表（需同时 --fix）")
    parser.add_argument("--save-as", help="规整后的另存路径（默认 <原名>_格式规整.docx）")
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"文件不存在: {docx_path}", file=sys.stderr)
        return 2

    document = Document(str(docx_path))
    run_checks(document)
    result = {"docx_path": str(docx_path), "findings": findings, **summarize(), "fixes": []}

    if args.fix:
        result["fixes"] += fix_page(document)
        result["fixes"] += fix_styles(document)
        result["fixes"] += fix_captions(document)
        if args.fix_tables:
            result["fixes"] += fix_tables(document)
        target = Path(args.save_as) if args.save_as else docx_path.with_name(f"{docx_path.stem}_格式规整.docx")
        document.save(str(target))
        result["saved_as"] = str(target)
        run_checks(document)
        result.update(summarize())
        result["findings"] = findings

    out_dir = Path(args.out).resolve() if args.out else docx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "format_check.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "format_check.md").write_text(build_markdown(result), encoding="utf-8")

    print(
        json.dumps(
            {
                "score": result["score"],
                "errors": result["errors"],
                "warnings": result["warnings"],
                "fixes": result["fixes"],
                "saved_as": result.get("saved_as"),
                "report": str(out_dir / "format_check.md"),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0 if result["errors"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
