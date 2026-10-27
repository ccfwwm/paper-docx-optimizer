#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI 痕迹离线扫描器。

不做「是否 AI 生成」的联网检测，只做可解释的文本特征量化，为改写提供定位依据：
- 句式分布：句长均值 / 标准差 / 变异系数 / 突发性（burstiness）
- 模式命中：并列枚举、三段式序列、套话收束、泛化评价、模板化背景、连接词堆积等
- 词表命中：三级 AI 高频词（实时读取 references/ai-vocabulary-blacklist.md）

用法：
    python ai_scan.py --docx paper.docx --out out/
    python ai_scan.py --text draft.txt --out out/
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_ops import (  # noqa: E402
    extract_headings,
    heading_level,
    is_caption_style,
    normalize_style_name,
)

SKILL_DIR = Path(__file__).resolve().parents[1]
DEFAULT_LEXICON = SKILL_DIR / "references" / "ai-vocabulary-blacklist.md"

SENTENCE_SPLIT_RE = re.compile(r"[。！？；!?;]+|\n+")
CJK_OR_WORD_RE = re.compile(r"[\u4e00-\u9fff]|[A-Za-z]+")
CONNECTIVES = ("因此", "从而", "进而", "同时", "此外", "并且", "而且", "另外", "基于此")

# 模式库：id / 名称 / 正则 / 权重 / 处理方向
PATTERNS = [
    ("P01", "并列枚举", r"[（(][1-9１-９][）)]|[①②③④⑤]", 8, "拆成递进或因果，不用编号罗列"),
    ("P02", "三段式序列", r"首先[^。]{0,120}(?:其次|然后)[^。]{0,120}(?:最后|最终)", 12, "合并为两句，保留最关键的先后关系"),
    ("P03", "套话收束", r"综上所述|由此可见|总而言之|总的来说|不难看出|一言以蔽之", 9, "直接给结论，删掉引导套话"),
    ("P04", "泛化评价", r"具有重要(?:的)?(?:意义|作用|价值)|极大(?:地)?(?:提高|提升|方便|改善)|有效(?:地)?(?:提升|改善|解决|缓解)|为用户提供(?:了)?(?:便利|方便)|一目了然", 8, "换成可验证的具体事实或删去评价"),
    ("P05", "模板化背景", r"随着[^。]{2,24}的(?:快速|迅猛|不断|持续)发展|在[^。]{2,12}的(?:大)?背景下|日益(?:增长|突出|重要|普及)", 8, "限定具体时间、场景与数据来源"),
    ("P06", "全知视角", r"众所周知|显而易见|不可否认|我们(?:可以|能够)(?:看到|发现|认为)", 6, "改为作者视角的具体观察"),
    ("P07", "完美对称", r"不仅[^。]{0,60}而且|既[^。]{0,40}又|一方面[^。]{0,80}另一方面", 7, "只保留一侧，另一侧改写为独立短句"),
    ("P08", "空洞定义", r"所谓[^。]{2,24}是指|是一种(?:以|基于)[^。]{0,30}的(?:技术|方法|模式)", 5, "补上具体机制或直接引用文献定义"),
    ("P09", "排比堆叠", r"(?:[，、][^，。]{2,8}){4,}[。]", 7, "打散为长短句交错"),
]

LEVEL_HINTS = [
    (60, "高"),
    (35, "中"),
    (0, "低"),
]


# ---------------------------------------------------------------- 词表

def load_lexicon(path: Path) -> list[dict]:
    """解析词表 Markdown 表格：| 词 | T1/T2/T3 | 替换方向 |"""
    if not path.exists():
        return []
    entries = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 2:
            continue
        tier = cells[1].upper().replace("级", "")
        if tier not in {"T1", "T2", "T3"}:
            continue
        word = cells[0].strip("`").strip()
        if not word or word.startswith("-"):
            continue
        entries.append({"word": word, "tier": tier, "hint": cells[2] if len(cells) > 2 else ""})
    return entries


# ---------------------------------------------------------------- 文本切分

def split_sentences(text: str) -> list[str]:
    parts = [p.strip() for p in SENTENCE_SPLIT_RE.split(text)]
    return [p for p in parts if len(p) >= 2]


def char_len(text: str) -> int:
    return len(CJK_OR_WORD_RE.findall(text))


def sentence_stats(sentences: list[str]) -> dict:
    lengths = [char_len(s) for s in sentences] or [0]
    count = len(lengths)
    mean = sum(lengths) / count
    var = sum((x - mean) ** 2 for x in lengths) / count
    std = math.sqrt(var)
    cv = std / mean if mean else 0.0
    burstiness = (std - mean) / (std + mean) if (std + mean) else 0.0
    return {
        "count": count,
        "mean": round(mean, 2),
        "std": round(std, 2),
        "cv": round(cv, 3),
        "burstiness": round(burstiness, 3),
        "short_ratio": round(sum(1 for x in lengths if x < 12) / count, 3),
        "long_ratio": round(sum(1 for x in lengths if x > 45) / count, 3),
    }


# ---------------------------------------------------------------- 段落分析

def analyze_paragraph(text: str, lexicon: list[dict]) -> dict:
    sentences = split_sentences(text)
    stats = sentence_stats(sentences)
    total_chars = max(1, char_len(text))

    patterns = []
    pattern_score = 0
    for pid, name, regex, weight, hint in PATTERNS:
        hits = re.findall(regex, text)
        if hits:
            patterns.append({"id": pid, "name": name, "count": len(hits), "hint": hint})
            pattern_score += weight * min(2, len(hits))

    connective_hits = sum(text.count(word) for word in CONNECTIVES)
    connective_density = connective_hits / total_chars * 100
    if connective_density >= 3:
        patterns.append(
            {
                "id": "P10",
                "name": "连接词堆积",
                "count": connective_hits,
                "hint": f"每百字连接词 {connective_density:.1f} 个，删去一半",
            }
        )
        pattern_score += 8

    tier_hits = {"T1": [], "T2": [], "T3": []}
    for entry in lexicon:
        if entry["word"] and entry["word"] in text:
            tier_hits[entry["tier"]].append(entry["word"])

    lexicon_score = 0.0
    lexicon_score += 12 * len(tier_hits["T1"]) + 4 * len(tier_hits["T2"]) + 1.5 * len(tier_hits["T3"])
    lexicon_score = min(30.0, lexicon_score / total_chars * 100)

    uniformity = 0.0
    if stats["count"] >= 3:
        if stats["cv"] < 0.18:
            uniformity = 22.0
        elif stats["cv"] < 0.28:
            uniformity = 14.0
        elif stats["cv"] < 0.38:
            uniformity = 6.0

    score = min(100.0, pattern_score + lexicon_score + uniformity)
    level = next(label for threshold, label in LEVEL_HINTS if score >= threshold)

    return {
        "chars": total_chars,
        "sentence": stats,
        "patterns": patterns,
        "tier1": tier_hits["T1"],
        "tier2": tier_hits["T2"],
        "tier3": tier_hits["T3"],
        "connective_density": round(connective_density, 2),
        "score": round(score, 1),
        "level": level,
    }


# ---------------------------------------------------------------- 文档读取

def read_docx_paragraphs(docx_path: Path) -> tuple[list[dict], list[dict]]:
    from docx import Document

    document = Document(str(docx_path))
    headings = {item["index"]: item for item in extract_headings(docx_path)}
    chapter = "（前置部分）"
    paragraphs = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        style = normalize_style_name(paragraph.style.name if paragraph.style else "")
        level = heading_level(style)
        if index in headings and level in (0, 1):
            chapter = text
            continue
        if level is not None or is_caption_style(style) or not text:
            continue
        paragraphs.append({"index": index, "chapter": chapter, "style": style, "text": text})
    return paragraphs, list(headings.values())


def read_text_paragraphs(path: Path) -> list[dict]:
    blocks = [b.strip() for b in re.split(r"\n\s*\n", path.read_text(encoding="utf-8")) if b.strip()]
    return [{"index": i, "chapter": "（全文）", "style": "", "text": b} for i, b in enumerate(blocks)]


# ---------------------------------------------------------------- 输出

def build_markdown(result: dict) -> str:
    lines = [
        "# AI 痕迹扫描报告",
        "",
        f"- 文档: `{result['source']}`",
        f"- 扫描段落: {result['paragraph_count']}（有效段落 {result['scored_count']}）",
        f"- 全文风险分: **{result['overall_score']} / 100**（越高越像机器文本）",
        f"- 句长均值 {result['sentence']['mean']} 字，标准差 {result['sentence']['std']}，"
        f"变异系数 {result['sentence']['cv']}，突发性 {result['sentence']['burstiness']}",
        f"- 高风险段落: {result['high_risk_count']} 个，中风险: {result['medium_risk_count']} 个",
        "",
        "## 章节热力图",
        "",
        "| 章节 | 风险分 | 等级 | 分布 | 主要问题 |",
        "| --- | --- | --- | --- | --- |",
    ]
    for chapter, info in result["chapters"].items():
        filled = min(10, max(0, round(info["score"] / 10)))
        bar = "█" * filled + "░" * (10 - filled)
        lines.append(
            f"| {chapter} | {info['score']} | {info['level']} | {bar} | {info['top_issue']} |"
        )

    lines += ["", "## 高风险段落清单", ""]
    if result["top_paragraphs"]:
        for item in result["top_paragraphs"]:
            issues = "、".join(p["name"] for p in item["patterns"]) or "—"
            tiers = "、".join(item["tier1"]) or "—"
            lines += [
                f"### P{item['index']} · {item['chapter']} · 风险 {item['score']}（{item['level']}）",
                "",
                f"- 模式: {issues}",
                f"- 强标记词: {tiers}",
                f"- 原文: {item['text'][:180]}",
                "",
            ]
    else:
        lines += ["- 未发现高风险段落。", ""]

    lines += [
        "## 使用建议",
        "",
        "1. 只对「高」等级段落动手，按 references/strategy-ai-reduction.md 做句内重组。",
        "2. 每改完一段重新扫描一次，观察变异系数是否上升（句长应该更参差）。",
        "3. 风险分下降不等于质量提升，改完仍要按 references/evaluation-criteria.md 复核。",
        "",
    ]
    return "\n".join(lines)


def cmd_scan(args) -> int:
    source = Path(args.docx or args.text)
    if not source.exists():
        print(f"输入不存在: {source}", file=sys.stderr)
        return 2

    lexicon = load_lexicon(Path(args.lexicon) if args.lexicon else DEFAULT_LEXICON)
    if source.suffix.lower() == ".docx" and not args.text:
        try:
            paragraphs, _ = read_docx_paragraphs(source)
        except ImportError:
            print("缺少 python-docx，请先: pip install python-docx", file=sys.stderr)
            return 2
    else:
        paragraphs = read_text_paragraphs(source)

    scored = []
    for item in paragraphs:
        if char_len(item["text"]) < args.min_length:
            continue
        analysis = analyze_paragraph(item["text"], lexicon)
        scored.append({**item, **analysis})

    chapters: dict[str, dict] = {}
    for item in scored:
        bucket = chapters.setdefault(item["chapter"], {"scores": [], "chars": 0, "issues": {}})
        bucket["scores"].append(item["score"] * item["chars"])
        bucket["chars"] += item["chars"]
        for pattern in item["patterns"]:
            bucket["issues"][pattern["name"]] = bucket["issues"].get(pattern["name"], 0) + 1

    chapter_summary = {}
    for name, bucket in chapters.items():
        score = round(bucket["scores"] and sum(bucket["scores"]) / max(1, bucket["chars"]) or 0, 1)
        top = max(bucket["issues"], key=bucket["issues"].get) if bucket["issues"] else "未检出明显模式"
        chapter_summary[name] = {
            "score": score,
            "level": next(label for threshold, label in LEVEL_HINTS if score >= threshold),
            "top_issue": top,
            "paragraphs": len(bucket["scores"]),
        }

    overall = round(
        sum(item["score"] * item["chars"] for item in scored) / max(1, sum(item["chars"] for item in scored)),
        1,
    ) if scored else 0.0

    all_sentences = [s for item in scored for s in split_sentences(item["text"])]
    result = {
        "source": str(source),
        "paragraph_count": len(paragraphs),
        "scored_count": len(scored),
        "overall_score": overall,
        "sentence": sentence_stats(all_sentences),
        "high_risk_count": sum(1 for item in scored if item["level"] == "高"),
        "medium_risk_count": sum(1 for item in scored if item["level"] == "中"),
        "chapters": chapter_summary,
        "top_paragraphs": sorted(scored, key=lambda x: x["score"], reverse=True)[: args.top],
        "lexicon_size": len(lexicon),
        "scanned_at": __import__("datetime").datetime.now().isoformat(timespec="seconds"),
    }

    out_dir = Path(args.out).resolve() if args.out else source.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "ai_scan.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (out_dir / "ai_scan.md").write_text(build_markdown(result), encoding="utf-8")

    print(
        json.dumps(
            {
                "overall_score": overall,
                "high_risk_count": result["high_risk_count"],
                "medium_risk_count": result["medium_risk_count"],
                "sentence": result["sentence"],
                "report": str(out_dir / "ai_scan.md"),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="离线扫描论文的 AI 文本特征")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--docx", help="待扫描的 .docx")
    group.add_argument("--text", help="待扫描的纯文本/Markdown")
    parser.add_argument("--out", help="报告输出目录（默认与输入同目录）")
    parser.add_argument("--lexicon", help="词表路径（默认 references/ai-vocabulary-blacklist.md）")
    parser.add_argument("--min-length", type=int, default=12, help="参与打分的段落最小字数（默认 12，与流水线口径一致）")
    parser.add_argument("--top", type=int, default=12, help="列出前 N 个高风险段落")
    args = parser.parse_args()
    return cmd_scan(args)


if __name__ == "__main__":
    raise SystemExit(main())
