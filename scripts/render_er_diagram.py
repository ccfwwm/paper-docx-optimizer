#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""论文 ER 图渲染器。

用 Graphviz 生成结构准确的 ER 图（实体矩形 / 属性椭圆 / 关系菱形 / 1-N-M 基数 /
主键短斜杠），并可附带产出「图像模型美化包」（结构参考图 + prompt + 标签清单）。

用法：
    python render_er_diagram.py --input schema.yaml --output-dir out/diagrams \
        --mode system-simple --format svg,png --imagegen-pack

模式：
    table          单实体 ER 图（一个实体 + 环绕属性）
    system-full    整体 ER 图（含属性椭圆）
    system-simple  整体精简图（只保留实体、关系、基数）

依赖：pyyaml（YAML 输入）、Graphviz（dot.exe）
构图规范与美化流程见 references/er-diagram-spec.md
"""

import argparse
import json
import math
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import yaml
except ImportError:  # pragma: no cover - handled at runtime
    yaml = None


DEFAULT_DOT = r"C:\Program Files\Graphviz\bin\dot.exe"
SKILL_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = SKILL_DIR / "templates"
SVG_NS = "http://www.w3.org/2000/svg"
ET.register_namespace("", SVG_NS)


@dataclass
class Attribute:
    name: str
    primary: bool = False


@dataclass
class Entity:
    name: str
    primary_key: str
    attributes: list[Attribute]
    rank: str | None = None


@dataclass
class Relationship:
    name: str
    source: str
    target: str
    source_cardinality: str
    target_cardinality: str
    attributes: list[str]


def load_schema(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"schema file not found: {path}")
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        return json.loads(text)
    if path.suffix.lower() in {".yaml", ".yml"}:
        if yaml is None:
            raise RuntimeError("PyYAML is required for YAML input. Install pyyaml or use JSON.")
        return yaml.safe_load(text)
    raise ValueError("schema must be .json, .yaml, or .yml")


def as_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    raise ValueError(f"expected list, got {type(value).__name__}")


def parse_schema(raw: dict) -> tuple[list[Entity], list[Relationship]]:
    entities = []
    seen = set()
    for item in as_list(raw.get("entities")):
        name = require_text(item, "name")
        if name in seen:
            raise ValueError(f"duplicate entity name: {name}")
        seen.add(name)
        primary_key = require_text(item, "primary_key")
        raw_attrs = as_list(item.get("attributes"))
        if primary_key not in raw_attrs:
            raise ValueError(f"primary_key '{primary_key}' is not listed in attributes for entity '{name}'")
        attributes = [Attribute(str(attr), str(attr) == primary_key) for attr in raw_attrs]
        entities.append(Entity(name=name, primary_key=primary_key, attributes=attributes, rank=item.get("rank")))

    names = {entity.name for entity in entities}
    relationships = []
    for item in as_list(raw.get("relationships")):
        name = require_text(item, "name")
        source = require_text(item, "from")
        target = require_text(item, "to")
        if source not in names:
            raise ValueError(f"relationship '{name}' references missing entity: {source}")
        if target not in names:
            raise ValueError(f"relationship '{name}' references missing entity: {target}")
        relationships.append(
            Relationship(
                name=name,
                source=source,
                target=target,
                source_cardinality=str(item.get("from_cardinality", "")),
                target_cardinality=str(item.get("to_cardinality", "")),
                attributes=[str(attr) for attr in as_list(item.get("attributes"))],
            )
        )
    return entities, relationships


def require_text(item: dict, key: str) -> str:
    value = item.get(key)
    if value is None or str(value).strip() == "":
        raise ValueError(f"missing required field: {key}")
    return str(value)


def q(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def node_id(*parts: str) -> str:
    raw = "__".join(parts)
    safe = re.sub(r"[^0-9A-Za-z_]+", "_", raw)
    return f"n_{safe.strip('_') or 'node'}"


def output_name(prefix: str, display_name: str, index: int | None, safe_filenames: bool) -> str:
    if safe_filenames:
        if index is None:
            return prefix
        return f"{prefix}_{index:02d}"
    if index is None:
        return prefix
    return f"{prefix}_{index:02d}_{display_name}"


def graph_header(font: str, layout: str) -> list[str]:
    overlap = "false" if layout == "neato" else "true"
    return [
        "graph ER {",
        "  graph [bgcolor=\"white\", rankdir=LR, splines=line, overlap=" + q(overlap) + ", outputorder=edgesfirst, nodesep=0.7, ranksep=1.0, margin=0.15];",
        f"  node [fontname={q(font)}, fontsize=18, color=\"black\", penwidth=1.2, margin=\"0.08,0.04\"];",
        f"  edge [fontname={q(font)}, fontsize=16, color=\"black\", penwidth=1.1];",
    ]


def entity_node(entity: Entity, simple: bool = False) -> str:
    if simple:
        label = f"{entity.name}\\n/{entity.primary_key}/"
        return f"  {node_id('entity', entity.name)} [label={q(label)}, shape=box, width=1.4, height=0.55];"
    return f"  {node_id('entity', entity.name)} [label={q(entity.name)}, shape=box, width=1.35, height=0.55];"


def attr_node(entity: Entity, attribute: Attribute) -> str:
    return f"  {node_id('attr', entity.name, attribute.name)} [label={q(attribute.name)}, shape=ellipse, width=1.25, height=0.62];"


def rel_node(index: int, rel: Relationship) -> str:
    return f"  {node_id('rel', str(index), rel.name, rel.source, rel.target)} [label={q(rel.name)}, shape=diamond, width=0.9, height=0.9, fixedsize=false];"


def build_table_dot(entity: Entity, font: str, layout: str) -> tuple[str, set[str]]:
    lines = graph_header(font, layout)
    entity_id = node_id("entity", entity.name)
    lines.append(f"  graph [rankdir=TB];")
    lines.append(entity_node(entity))
    pk_edges = set()
    for idx, attr in enumerate(entity.attributes):
        attr_id = node_id("attr", entity.name, attr.name)
        lines.append(attr_node(entity, attr))
        edge_id = f"pk_{entity.name}_{attr.name}"
        if attr.primary:
            pk_edges.add(edge_id)
            lines.append(f"  {attr_id} -- {entity_id} [id={q(edge_id)}];")
        else:
            lines.append(f"  {attr_id} -- {entity_id};")
    lines.append("}")
    return "\n".join(lines) + "\n", pk_edges


def format_entity_list(entities: list[Entity]) -> str:
    return "、".join(entity.name for entity in entities)


def format_primary_keys(entities: list[Entity]) -> str:
    return "；".join(f"{entity.name}: {entity.primary_key}" for entity in entities)


def format_relationships(relationships: list[Relationship]) -> str:
    if not relationships:
        return "无"
    return "；".join(
        f"{rel.source} -({rel.source_cardinality})- {rel.name} -({rel.target_cardinality})- {rel.target}"
        for rel in relationships
    )


def labels_for_table(entity: Entity) -> dict:
    return {
        "mode": "table",
        "entities": [entity.name],
        "attributes": {entity.name: [attr.name for attr in entity.attributes]},
        "primary_keys": {entity.name: entity.primary_key},
        "relationships": [],
        "cardinalities": [],
    }


def labels_for_system(entities: list[Entity], relationships: list[Relationship], include_attrs: bool) -> dict:
    labels = {
        "mode": "system-full" if include_attrs else "system-simple",
        "entities": [entity.name for entity in entities],
        "primary_keys": {entity.name: entity.primary_key for entity in entities},
        "relationships": [
            {
                "name": rel.name,
                "from": rel.source,
                "to": rel.target,
                "from_cardinality": rel.source_cardinality,
                "to_cardinality": rel.target_cardinality,
                "attributes": rel.attributes,
            }
            for rel in relationships
        ],
        "cardinalities": sorted({rel.source_cardinality for rel in relationships} | {rel.target_cardinality for rel in relationships}),
    }
    if include_attrs:
        labels["attributes"] = {entity.name: [attr.name for attr in entity.attributes] for entity in entities}
        labels["relationship_attributes"] = {rel.name: rel.attributes for rel in relationships if rel.attributes}
    else:
        labels["attributes"] = {}
    return labels


def render_template(template_name: str, values: dict) -> str:
    template_path = TEMPLATE_DIR / template_name
    template = template_path.read_text(encoding="utf-8")
    return template.format(**values)


def write_imagegen_pack(
    output_dir: Path,
    name: str,
    prompt: str,
    labels: dict,
    reference_outputs: list[Path],
) -> list[Path]:
    prompt_path = output_dir / f"{name}.imagegen_prompt.md"
    labels_path = output_dir / f"{name}.labels.json"
    reference_path = output_dir / f"{name}.reference_paths.json"
    resolved_references = [str(path.resolve()) for path in reference_outputs]
    prompt_with_references = "\n".join(
        [
            prompt.rstrip(),
            "",
            "Reference files generated by Graphviz:",
            *[f"- {path}" for path in resolved_references],
            "",
            "Use the PNG/SVG reference only for topology and exact labels; the final image should be a cleaner thesis-style redraw.",
            "",
        ]
    )
    prompt_path.write_text(prompt_with_references, encoding="utf-8")
    labels_path.write_text(json.dumps(labels, ensure_ascii=False, indent=2), encoding="utf-8")
    reference_path.write_text(
        json.dumps({"reference_outputs": resolved_references}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return [prompt_path, labels_path, reference_path]


def table_prompt(entity: Entity) -> str:
    return render_template(
        "imagegen_prompt_table.md",
        {
            "entity": entity.name,
            "primary_key": entity.primary_key,
            "attributes": "、".join(attr.name for attr in entity.attributes),
        },
    )


def system_prompt(entities: list[Entity], relationships: list[Relationship], include_attrs: bool) -> str:
    return render_template(
        "imagegen_prompt_system_full.md" if include_attrs else "imagegen_prompt_system_simple.md",
        {
            "entities": format_entity_list(entities),
            "relationships": format_relationships(relationships),
            "primary_keys": format_primary_keys(entities),
        },
    )


def build_system_dot(entities: list[Entity], relationships: list[Relationship], font: str, layout: str, include_attrs: bool) -> tuple[str, set[str]]:
    lines = graph_header(font, layout)
    pk_edges = set()
    for entity in entities:
        lines.append(entity_node(entity, simple=not include_attrs))
        if include_attrs:
            entity_id = node_id("entity", entity.name)
            for attr in entity.attributes:
                attr_id = node_id("attr", entity.name, attr.name)
                lines.append(attr_node(entity, attr))
                edge_id = f"pk_{entity.name}_{attr.name}"
                if attr.primary:
                    pk_edges.add(edge_id)
                    lines.append(f"  {attr_id} -- {entity_id} [id={q(edge_id)}, len=1.0];")
                else:
                    lines.append(f"  {attr_id} -- {entity_id} [len=1.0];")

    for idx, rel in enumerate(relationships):
        rid = node_id("rel", str(idx), rel.name, rel.source, rel.target)
        sid = node_id("entity", rel.source)
        tid = node_id("entity", rel.target)
        lines.append(rel_node(idx, rel))
        lines.append(f"  {sid} -- {rid} [label={q(rel.source_cardinality)}];")
        lines.append(f"  {rid} -- {tid} [label={q(rel.target_cardinality)}];")
        if include_attrs:
            for attr in rel.attributes:
                aid = node_id("relattr", str(idx), rel.name, attr)
                lines.append(f"  {aid} [label={q(attr)}, shape=ellipse, width=1.1, height=0.58];")
                lines.append(f"  {aid} -- {rid};")

    rank_groups: dict[str, list[str]] = {}
    for entity in entities:
        if entity.rank:
            rank_groups.setdefault(entity.rank, []).append(node_id("entity", entity.name))
    for group_nodes in rank_groups.values():
        if len(group_nodes) > 1:
            lines.append("  { rank=same; " + "; ".join(group_nodes) + "; }")
    lines.append("}")
    return "\n".join(lines) + "\n", pk_edges


def run_graphviz(dot_path: Path, out_path: Path, dot_exe: Path, layout: str, fmt: str) -> None:
    if not dot_exe.exists():
        raise FileNotFoundError(f"Graphviz executable not found: {dot_exe}")
    exe = dot_exe
    if layout != "dot":
        candidate = dot_exe.with_name(f"{layout}.exe")
        if candidate.exists():
            exe = candidate
    result = subprocess.run(
        [str(exe), f"-T{fmt}", str(dot_path), "-o", str(out_path)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        raise RuntimeError(f"Graphviz failed for {dot_path}:\n{result.stderr or result.stdout}")


def dot_with_png_pk_labels(dot: str, pk_edges: set[str]) -> str:
    """Add slash text labels for PNG output, where SVG post-processing is unavailable."""
    updated = dot
    for edge_id in pk_edges:
        pattern = f"id={q(edge_id)}"
        replacement = f"id={q(edge_id)}, xlabel=\"/\", labeldistance=0.2"
        updated = updated.replace(pattern, replacement)
    return updated


def parse_points(value: str) -> list[tuple[float, float]]:
    points = []
    for pair in value.strip().split():
        if "," not in pair:
            continue
        x, y = pair.split(",", 1)
        points.append((float(x), float(y)))
    return points


def slash_for_polyline(points: list[tuple[float, float]], size: float = 12.0) -> tuple[float, float, float, float] | None:
    if len(points) < 2:
        return None
    lengths = []
    total = 0.0
    for a, b in zip(points, points[1:]):
        length = math.dist(a, b)
        lengths.append(length)
        total += length
    if total == 0:
        return None
    halfway = total / 2
    walked = 0.0
    for index, length in enumerate(lengths):
        if walked + length >= halfway:
            ratio = (halfway - walked) / length if length else 0
            x1, y1 = points[index]
            x2, y2 = points[index + 1]
            cx = x1 + (x2 - x1) * ratio
            cy = y1 + (y2 - y1) * ratio
            dx = x2 - x1
            dy = y2 - y1
            norm = math.hypot(dx, dy) or 1.0
            px = -dy / norm
            py = dx / norm
            return (cx - px * size / 2, cy - py * size / 2, cx + px * size / 2, cy + py * size / 2)
        walked += length
    return None


def add_pk_slashes(svg_path: Path, pk_edges: set[str]) -> None:
    if not pk_edges:
        return
    tree = ET.parse(svg_path)
    root = tree.getroot()
    for group in root.findall(f".//{{{SVG_NS}}}g"):
        gid = group.attrib.get("id", "")
        if gid not in pk_edges:
            continue
        path = group.find(f".//{{{SVG_NS}}}path")
        polygon = group.find(f".//{{{SVG_NS}}}polygon")
        slash = None
        if polygon is not None and "points" in polygon.attrib:
            slash = slash_for_polyline(parse_points(polygon.attrib["points"]))
        if slash is None and path is not None:
            coords = [(float(x), float(y)) for x, y in re.findall(r"(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)", path.attrib.get("d", ""))]
            slash = slash_for_polyline(coords)
        if slash is None:
            continue
        x1, y1, x2, y2 = slash
        ET.SubElement(
            group,
            f"{{{SVG_NS}}}line",
            {
                "x1": f"{x1:.2f}",
                "y1": f"{y1:.2f}",
                "x2": f"{x2:.2f}",
                "y2": f"{y2:.2f}",
                "stroke": "black",
                "stroke-width": "1.8",
                "stroke-linecap": "round",
            },
        )
    tree.write(svg_path, encoding="utf-8", xml_declaration=True)


def write_and_render(name: str, dot: str, pk_edges: set[str], output_dir: Path, dot_exe: Path, layout: str, formats: list[str]) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    dot_path = output_dir / f"{name}.dot"
    dot_path.write_text(dot, encoding="utf-8")
    outputs = [dot_path]
    svg_path = output_dir / f"{name}.svg"
    if "svg" in formats:
        run_graphviz(dot_path, svg_path, dot_exe, layout, "svg")
        add_pk_slashes(svg_path, pk_edges)
        outputs.append(svg_path)
    if "png" in formats:
        png_path = output_dir / f"{name}.png"
        png_dot_path = output_dir / f".{name}.png.tmp.dot"
        png_dot_path.write_text(dot_with_png_pk_labels(dot, pk_edges), encoding="utf-8")
        try:
            run_graphviz(png_dot_path, png_path, dot_exe, layout, "png")
        finally:
            png_dot_path.unlink(missing_ok=True)
        outputs.append(png_path)
    return outputs


def parse_formats(value: str) -> list[str]:
    formats = [item.strip().lower() for item in value.split(",") if item.strip()]
    allowed = {"svg", "png"}
    bad = [item for item in formats if item not in allowed]
    if bad:
        raise ValueError(f"unsupported format(s): {', '.join(bad)}")
    return formats or ["svg"]


def main() -> None:
    parser = argparse.ArgumentParser(description="Render thesis-style ER diagrams with Graphviz.")
    parser.add_argument("--input", required=True, help="Path to schema .yaml/.yml/.json")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--mode", choices=["table", "system-full", "system-simple"], required=True)
    parser.add_argument("--format", default="svg", help="svg, png, or svg,png")
    parser.add_argument("--entity", help="When --mode table is used, only render this entity")
    parser.add_argument("--entity-index", type=int, help="When --mode table is used, render one entity by 1-based index")
    parser.add_argument("--safe-filenames", action="store_true", help="Use ASCII-only output filenames such as table_01")
    parser.add_argument("--imagegen-pack", action="store_true", help="Also write imagegen_prompt.md, labels.json, and reference path metadata")
    parser.add_argument("--dot", default=DEFAULT_DOT)
    parser.add_argument("--layout", choices=["dot", "neato"], default="dot")
    parser.add_argument("--font", default="Microsoft YaHei")
    args = parser.parse_args()

    dot_exe = Path(args.dot)
    raw = load_schema(Path(args.input))
    entities, relationships = parse_schema(raw)
    output_dir = Path(args.output_dir)
    formats = parse_formats(args.format)

    outputs = []
    if args.mode == "table":
        selected_entities = list(enumerate(entities, start=1))
        if args.entity:
            selected_entities = [(index, entity) for index, entity in selected_entities if entity.name == args.entity]
            if not selected_entities:
                raise ValueError(f"entity not found: {args.entity}")
        if args.entity_index is not None:
            if args.entity_index < 1 or args.entity_index > len(entities):
                raise ValueError(f"entity-index out of range: {args.entity_index}; valid range is 1..{len(entities)}")
            selected_entities = [(args.entity_index, entities[args.entity_index - 1])]
        for index, entity in selected_entities:
            dot, pk_edges = build_table_dot(entity, args.font, args.layout)
            name = output_name("table", entity.name, index, args.safe_filenames)
            rendered = write_and_render(name, dot, pk_edges, output_dir, dot_exe, args.layout, formats)
            outputs.extend(rendered)
            if args.imagegen_pack:
                outputs.extend(write_imagegen_pack(output_dir, name, table_prompt(entity), labels_for_table(entity), rendered))
    else:
        include_attrs = args.mode == "system-full"
        dot, pk_edges = build_system_dot(entities, relationships, args.font, args.layout, include_attrs)
        name = "system_full" if include_attrs else "system_simple"
        rendered = write_and_render(name, dot, pk_edges, output_dir, dot_exe, args.layout, formats)
        outputs.extend(rendered)
        if args.imagegen_pack:
            outputs.extend(write_imagegen_pack(output_dir, name, system_prompt(entities, relationships, include_attrs), labels_for_system(entities, relationships, include_attrs), rendered))

    print(json.dumps({"outputs": [str(path.resolve()) for path in outputs]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
