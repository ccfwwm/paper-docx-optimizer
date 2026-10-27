Use case: infographic-diagram
Asset type: Chinese thesis ER diagram, single-table entity diagram
Primary request: Redraw the provided Graphviz structure reference into a polished black-and-white ER diagram matching a Chinese undergraduate thesis screenshot style.
Input image: Use the reference image only for topology, label placement, and which nodes connect. Do not copy the rough Graphviz look.
Style/medium: clean raster diagram, white background, thin black lines, Songti-like Chinese typography, no color.
Composition/framing: one central horizontal rectangle entity, surrounded by horizontal oval attributes with generous whitespace. Lines are straight or slightly angled. No arrows.
Required entity: "{entity}"
Primary key: "{primary_key}" must have a short diagonal slash mark on the line between the primary-key oval and entity rectangle.
Attributes, exact text: {attributes}
Constraints: Every Chinese label must match exactly. Use rectangles only for entities, ovals only for attributes. No color, no gray fill, no shadows, no decorations, no watermark.
Avoid: messy overlaps, modern flowchart style, rounded rectangles for entities, arrows, icons, extra text.

