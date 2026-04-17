"""
Parse a Quarto manuscript.qmd into structured content blocks.

Returns a list of blocks that the DocxBuilder can consume directly.
The .qmd file remains the single source of truth for content;
this parser extracts structure without depending on Quarto.

Block types: 'yaml', 'heading', 'paragraph', 'table', 'equation',
             'figure', 'list_item', 'blank'
"""
from __future__ import annotations

import re
import yaml
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Block:
    type: str
    text: str = ""
    level: int = 0
    rows: list[list[str]] = field(default_factory=list)
    headers: list[str] = field(default_factory=list)
    path: str = ""
    caption: str = ""
    label: str = ""
    latex: str = ""
    numbered: bool = True


def parse_qmd(qmd_path: Path) -> tuple[dict, list[Block]]:
    text = qmd_path.read_text(encoding="utf-8")

    # Split YAML frontmatter
    if text.startswith("---"):
        end = text.find("---", 3)
        yaml_text = text[3:end]
        body = text[end + 3:]
        try:
            meta = yaml.safe_load(yaml_text) or {}
        except yaml.YAMLError:
            meta = {}
    else:
        meta = {}
        body = text

    blocks: list[Block] = []
    lines = body.split("\n")
    i = 0
    in_table = False
    table_lines: list[str] = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Blank line
        if not stripped:
            if in_table and table_lines:
                blocks.append(_parse_table(table_lines))
                table_lines = []
                in_table = False
            i += 1
            continue

        # Heading
        m = re.match(r'^(#{1,3})\s+(.+?)(?:\s*\{.*\})?\s*$', line)
        if m:
            if in_table and table_lines:
                blocks.append(_parse_table(table_lines))
                table_lines = []
                in_table = False
            level = len(m.group(1))
            text_h = m.group(2).strip()
            numbered = "{.unnumbered}" not in line
            blocks.append(Block(type="heading", text=text_h, level=level, numbered=numbered))
            i += 1
            continue

        # Display equation ($$...$$)
        if stripped.startswith("$$"):
            eq_lines = [stripped[2:]] if len(stripped) > 2 else []
            i += 1
            while i < len(lines):
                if lines[i].strip().startswith("$$"):
                    label_match = re.search(r'\{#eq-(\w+)\}', lines[i])
                    label = label_match.group(1) if label_match else ""
                    break
                eq_lines.append(lines[i].strip())
                i += 1
            latex = " ".join(eq_lines).strip()
            blocks.append(Block(type="equation", latex=latex, label=label))
            i += 1
            continue

        # Figure: ![caption](path){#fig-label}
        m = re.match(r'^!\[(.+?)\]\((.+?)\)(?:\{.*\})?\s*$', stripped)
        if m:
            blocks.append(Block(
                type="figure",
                caption=m.group(1),
                path=m.group(2),
            ))
            i += 1
            continue

        # Table row
        if stripped.startswith("|"):
            in_table = True
            table_lines.append(stripped)
            i += 1
            continue

        # Table caption (: text {#tbl-label})
        if stripped.startswith(":") and not stripped.startswith(":::"):
            if in_table and table_lines:
                tbl_block = _parse_table(table_lines)
                tbl_block.caption = stripped.lstrip(": ").split("{")[0].strip()
                blocks.append(tbl_block)
                table_lines = []
                in_table = False
            i += 1
            continue

        # ::: refs block (skip)
        if stripped.startswith(":::"):
            i += 1
            continue

        # Regular paragraph (accumulate consecutive non-blank lines)
        if in_table and table_lines:
            blocks.append(_parse_table(table_lines))
            table_lines = []
            in_table = False

        para_lines = [stripped]
        i += 1
        while i < len(lines):
            s = lines[i].strip()
            if not s or s.startswith("#") or s.startswith("|") or s.startswith("$$") or s.startswith("!["):
                break
            para_lines.append(s)
            i += 1
        blocks.append(Block(type="paragraph", text=" ".join(para_lines)))

    if in_table and table_lines:
        blocks.append(_parse_table(table_lines))

    return meta, blocks


def _parse_table(lines: list[str]) -> Block:
    data_rows = []
    headers = []
    for line in lines:
        cells = [c.strip() for c in line.strip("|").split("|")]
        if all(set(c) <= {"-", ":", " "} for c in cells):
            continue
        if not headers:
            headers = cells
        else:
            data_rows.append(cells)
    return Block(type="table", headers=headers, rows=data_rows)
