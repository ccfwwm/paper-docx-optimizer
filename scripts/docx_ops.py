#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""论文 DOCX 通用操作库（paper-docx-optimizer 内部共享模块）。

职责：
- 输入文件发现（主论文 docx / 检测报告）
- 标题结构抽取（python-docx 优先，OOXML 兜底）
- run 级安全文本替换（保留原有字体、字号、加粗、上标等格式）
- DOCX 包完整性校验

依赖：python-docx（pip install python-docx）
"""

from __future__ import annotations

import json
import re
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# 主论文候选文件名特征（按优先级）
DOCX_HINTS = ["论文", "设计与实现", "实现与设计", "thesis", "毕业论文", "学位论文"]

# 检测报告候选文件名特征（按优先级）
REPORT_HINTS = ["aigc", "检测报告", "查重", "similarity", "report"]

# 非正文样式（不做文本改写）
HEADING_STYLE_RE = re.compile(r"^(heading|标题)\s*(\d)$", re.IGNORECASE)
CAPTION_STYLE_HINTS = ("caption", "题注", "插图标题", "三线表标题", "表标题", "图标题")
SKIP_STYLE_HINTS = ("toc", "目录", "页眉", "页脚", "header", "footer", "reference", "参考文献")

CJK_RE = re.compile(r"[\u4e00-\u9fff]")


# ---------------------------------------------------------------- 基础 IO

def load_json(path: str | Path) -> dict:
    p = Path(path)
    if not p.exists():
        return {}
    return json.loads(p.read_text(encoding="utf-8"))


def dump_json(path: str | Path, data) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_config(config_path: str | None) -> dict:
    if not config_path:
        return {}
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"config not found: {path}")
    return load_json(path)


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def mkdirs(output_dir: Path) -> dict[str, Path]:
    dirs = {
        "analysis": output_dir / "analysis",
        "intermediate": output_dir / "intermediate",
        "final": output_dir / "final",
        "logs": output_dir / "logs",
    }
    for path in dirs.values():
        path.mkdir(parents=True, exist_ok=True)
    return dirs


# ---------------------------------------------------------------- 文件发现

def pick_docx(input_dir: Path, explicit: str | None = None) -> Path:
    """在输入目录中定位主论文 docx。"""
    if explicit:
        path = Path(explicit)
        if not path.exists():
            raise FileNotFoundError(f"docx not found: {path}")
        return path

    docx_files = [
        p for p in sorted(input_dir.glob("*.docx"))
        if not p.name.startswith("~$")  # 排除 Word 临时锁文件
    ]
    if not docx_files:
        raise FileNotFoundError(f"no docx found in {input_dir}")
    if len(docx_files) == 1:
        return docx_files[0]

    ranked = sorted(
        docx_files,
        key=lambda p: (
            0 if any(h.lower() in p.name.lower() for h in DOCX_HINTS) else 1,
            len(p.name),
            p.name.lower(),
        ),
    )
    return ranked[0]


def pick_report(input_dir: Path, explicit: str | None = None) -> Path | None:
    """在输入目录中定位 AIGC / 查重报告（可选）。"""
    if explicit:
        path = Path(explicit)
        if not path.exists():
            raise FileNotFoundError(f"report not found: {path}")
        return path

    ranked = []
    for path in input_dir.iterdir():
        if not path.is_file() or path.suffix.lower() == ".docx":
            continue
        lower = path.name.lower()
        for index, hint in enumerate(REPORT_HINTS):
            if hint in lower:
                ranked.append((index, len(path.name), path))
                break
    if not ranked:
        return None
    ranked.sort(key=lambda item: (item[0], item[1], item[2].name.lower()))
    return ranked[0][2]


# ---------------------------------------------------------------- 样式工具

def normalize_style_name(name: str) -> str:
    """把「标题 1」「heading 1」统一成「Heading 1」，保留原样返回不易识别的名称。"""
    if not name:
        return ""
    text = name.strip()
    m = HEADING_STYLE_RE.match(text)
    if m:
        return f"Heading {m.group(2)}"
    return text


def heading_level(style_name: str) -> int | None:
    """返回标题级别（1-9），非标题样式返回 None。"""
    m = HEADING_STYLE_RE.match((style_name or "").strip())
    if m:
        return int(m.group(2))
    return None


def is_caption_style(style_name: str) -> bool:
    low = (style_name or "").lower()
    return any(hint.lower() in low for hint in CAPTION_STYLE_HINTS)


def is_skippable_style(style_name: str) -> bool:
    low = (style_name or "").lower()
    return any(hint in low for hint in SKIP_STYLE_HINTS)


# ---------------------------------------------------------------- 结构抽取

def extract_headings(docx_path: Path) -> list[dict]:
    """抽取标题条目。优先 python-docx，失败时回退 OOXML 解析。"""
    try:
        return _extract_headings_python_docx(docx_path)
    except Exception:
        return extract_headings_ooxml(docx_path)


def _extract_headings_python_docx(docx_path: Path) -> list[dict]:
    from docx import Document

    document = Document(str(docx_path))
    headings = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        style = normalize_style_name(paragraph.style.name if paragraph.style else "")
        level = heading_level(style)
        is_front_title = style == "Title" and text.replace(" ", "") in {
            "摘要", "摘　要", "Abstract", "ABSTRACT",
        }
        if level is None and not is_front_title:
            continue
        headings.append(
            {
                "index": index,
                "style": style,
                "level": level or 0,
                "text": text,
            }
        )
    return headings


def _ooxml_paragraph_text(paragraph: ET.Element) -> str:
    parts = []
    for node in paragraph.findall(".//w:t", WORD_NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def extract_headings_ooxml(docx_path: Path) -> list[dict]:
    """不依赖 python-docx 的兜底抽取（直接读 word/document.xml）。"""
    with zipfile.ZipFile(docx_path) as package:
        xml_bytes = package.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    headings = []
    for index, paragraph in enumerate(root.findall(".//w:body/w:p", WORD_NS)):
        text = _ooxml_paragraph_text(paragraph)
        if not text:
            continue
        style_node = paragraph.find("./w:pPr/w:pStyle", WORD_NS)
        style = ""
        if style_node is not None:
            style = style_node.attrib.get(f"{{{WORD_NS['w']}}}val", "")
        style = normalize_style_name(style)
        level = heading_level(style)
        normalized = text.replace(" ", "")
        inferred_chapter = bool(re.match(r"^第[一二三四五六七八九十百]+章", normalized))
        if level is None and not inferred_chapter:
            continue
        headings.append(
            {
                "index": index,
                "style": style or "OOXML inferred",
                "level": level or 1,
                "text": text,
            }
        )
    return headings


def build_spans(headings: list[dict], paragraph_count: int) -> list[dict]:
    """把标题条目转成「章节段落区间」列表。"""
    spans = []
    for idx, item in enumerate(headings):
        start = item["index"]
        end = headings[idx + 1]["index"] if idx + 1 < len(headings) else paragraph_count
        spans.append(
            {
                "title": item["text"],
                "style": item["style"],
                "level": item.get("level", 0),
                "start": start,
                "end": end,
            }
        )
    return spans


# ---------------------------------------------------------------- 文本替换

def paragraph_plain_text(paragraph) -> str:
    """仅由 run 组成的纯文本（用于判断是否含超链接/域，避免替换串行）。"""
    return "".join(run.text for run in paragraph.runs)


def _locate_span(runs, start: int, end: int):
    """在 run 序列中定位字符区间 [start, end)，返回 (run_idx, local_start, local_end)。"""
    if end <= start:
        return None
    offset = 0
    hit = None
    for index, run in enumerate(runs):
        length = len(run.text)
        run_start, run_end = offset, offset + length
        offset = run_end
        if run_start <= start < run_end:
            if end <= run_end:
                hit = (index, start - run_start, end - run_start)
        elif run_start < end <= run_end:
            return None  # 跨 run，无法局部替换
        if hit and end <= run_end:
            return hit
    return None


def replace_paragraph_text(paragraph, new_text: str, old_text: str | None = None) -> tuple[bool, str]:
    """在保留格式的前提下替换段落文本。

    两级策略：
    1. surgical（优先）：比对新旧文本的公共前后缀，若差异段恰好落在单个 run 内，
       只改这个 run 的局部文本，段落内其余 run 的加粗/上标/字体全部原样保留。
    2. host（兜底）：差异段跨 run 时，把新文本整体写入最长 run，其余 run 清空文本
       但保留元素（避免删除域代码、脚注引用、批注锚点）。此类段落会标记为「格式风险」。

    含超链接或域代码的段落整体跳过，不做改写。
    返回 (是否已改写, 模式或跳过原因)。
    """
    if paragraph._p.findall(".//" + f"{{{WORD_NS['w']}}}hyperlink"):
        return False, "含超链接，跳过以免破坏链接"
    if paragraph._p.findall(".//" + f"{{{WORD_NS['w']}}}instrText"):
        return False, "含域代码，跳过"

    runs = list(paragraph.runs)
    original = paragraph.text
    if not runs:
        paragraph.text = new_text
        return True, "surgical"
    if paragraph_plain_text(paragraph) != original:
        return False, "run 文本与段落文本不一致（可能存在超链接/域），跳过"

    old = old_text if old_text is not None else original
    if old and old != new_text:
        prefix = 0
        limit = min(len(old), len(new_text))
        while prefix < limit and old[prefix] == new_text[prefix]:
            prefix += 1
        suffix = 0
        while (
            suffix < (limit - prefix)
            and old[len(old) - 1 - suffix] == new_text[len(new_text) - 1 - suffix]
        ):
            suffix += 1
        old_end = len(old) - suffix
        new_end = len(new_text) - suffix
        span = _locate_span(runs, prefix, old_end)
        if span:
            run_index, local_start, local_end = span
            run = runs[run_index]
            run.text = run.text[:local_start] + new_text[prefix:new_end] + run.text[local_end:]
            return True, "surgical"

    host = max(runs, key=lambda r: len(r.text))
    host.text = new_text
    for run in runs:
        if run is not host and run.text:
            run.text = ""
    return True, "host"


# ---------------------------------------------------------------- 完整性校验

def docx_integrity(docx_path: Path) -> dict:
    """检查 DOCX 是否为合法 OOXML 包（zip 完整性 + 关键部件可解析）。"""
    result = {"path": str(docx_path), "ok": False, "errors": [], "parts": []}
    if not docx_path.exists():
        result["errors"].append("文件不存在")
        return result
    try:
        with zipfile.ZipFile(docx_path) as package:
            bad = package.testzip()
            if bad:
                result["errors"].append(f"zip 部件损坏: {bad}")
            names = package.namelist()
            result["parts"] = names
            for required in ("[Content_Types].xml", "word/document.xml"):
                if required not in names:
                    result["errors"].append(f"缺少必需部件: {required}")
            if "word/document.xml" in names:
                try:
                    ET.fromstring(package.read("word/document.xml"))
                except Exception as exc:
                    result["errors"].append(f"document.xml 无法解析: {exc}")
    except zipfile.BadZipFile as exc:
        result["errors"].append(f"不是合法 zip: {exc}")
    result["ok"] = not result["errors"]
    return result
