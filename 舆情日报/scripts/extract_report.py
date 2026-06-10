#!/usr/bin/env python3
"""Extract Huike public-opinion daily DOCX content into a structured JSON manifest.

This script reads a Huike DOCX draft and produces a JSON file containing:
- Section structure and numbered headings
- Each data row with title, source, summary text, and row index
- Metadata (file path, row count, section info)

The JSON manifest is meant for LLM review: the model reads the manifest alongside
filtering rules and produces a decisions.json for process_report.py to apply.

Cross-platform: uses pathlib for path handling, works on Windows/macOS/Linux.

Usage:
    python extract_report.py <input.docx> [-o output.json]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    from docx import Document
except ImportError:
    print("ERROR: python-docx is required. Install with: pip install python-docx", file=sys.stderr)
    sys.exit(3)

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SECTION_RE = re.compile(r"^(\d+)[、.．]\s*(.+)$")


def _qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


@dataclass
class RowData:
    """Represents one row in the Huike report table."""

    index: int  # 0-based row index in the table
    title: str  # Content of the title/media column
    summary: str  # Content of the summary column
    is_heading: bool  # True if this is a section heading row
    section_num: int | None  # Section number if heading, e.g. 1, 2, 3
    section_title: str | None  # Section title text if heading, e.g. 公司新闻
    shading: str | None  # Current shading fill color or None
    media_source: str | None  # Extracted media source from title column


@dataclass
class ReportManifest:
    """Complete extracted content from a Huike DOCX draft."""

    input_path: str
    total_rows: int
    data_rows: int
    heading_rows: int
    sections: list[dict[str, Any]]
    rows: list[dict[str, Any]]
    has_summary_column: bool
    forbidden_links_found: list[str]


def cell_text(cell: Any) -> str:
    """Extract text from a python-docx cell."""
    return cell.text.strip() if cell else ""


def row_shading(cell: Any) -> str | None:
    """Extract shading fill color from a cell's XML."""
    tc_pr = cell._tc.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr")
    if tc_pr is not None:
        shd = tc_pr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd")
        if shd is not None:
            fill = shd.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
            if fill and fill.lower() != "auto":
                return fill.upper()
    return None


def extract_media_source(title: str) -> str | None:
    """Extract media source from title column text.

    The title column typically has format:
        Title text
        （来源：XXX）

    Returns XXX or the raw title if no match.
    """
    if not title:
        return None
    # Try to find （来源：XXX） or (来源：XXX)
    m = re.search(r"[（(]\s*来源\s*[:：]\s*([^）)]+)\s*[）)]", title)
    if m:
        return m.group(1).strip()
    return None


FORBIDDEN_PATTERNS = [
    "163.com", "qq.com", "sohu.com", "sina.com.cn",
    "mp.weixin.qq.com", "weixin",
    "app-web.chnfund.com", "api3.cls.cn", "h.xinhuaxmt.com",
]


def find_forbidden_links(text: str) -> list[str]:
    """Find any forbidden link patterns in text."""
    lowered = text.lower()
    found: list[str] = []
    for pattern in FORBIDDEN_PATTERNS:
        if pattern.lower() in lowered:
            found.append(pattern)
    return found


def extract_docx(path: Path) -> ReportManifest:
    """Extract all rows and structure from a Huike DOCX file."""
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    if path.suffix.lower() not in (".docx",):
        raise ValueError(f"Expected .docx input, got: {path}")

    doc = Document(str(path))
    rows_data: list[RowData] = []
    all_text = ""
    forbidden_links_found: list[str] = []

    for table_idx, table in enumerate(doc.tables):
        all_text += cell_text(table.cell(0, 0)) if table.rows else ""

        for row_idx, row in enumerate(table.rows):
            cells = row.cells
            if not cells:
                continue

            title_cell = cells[0]
            title = cell_text(title_cell)
            all_text += title + "\n"

            summary = cell_text(cells[1]) if len(cells) > 1 else ""
            all_text += summary + "\n"

            # Check section heading
            sec_match = SECTION_RE.match(title.replace(" ", ""))
            is_heading = bool(sec_match)
            sec_num = None
            sec_title = None
            if sec_match:
                sec_num = int(sec_match.group(1))
                sec_title = sec_match.group(2).strip()

            # Shading from first cell
            shade = row_shading(cells[0])

            row_obj = RowData(
                index=row_idx,
                title=title,
                summary=summary,
                is_heading=is_heading,
                section_num=sec_num,
                section_title=sec_title,
                shading=shade,
                media_source=extract_media_source(title),
            )
            rows_data.append(row_obj)

    # Scan all text for forbidden links
    forbidden_links_found = find_forbidden_links(all_text)

    # Build section list
    sections: list[dict[str, Any]] = []
    for row in rows_data:
        if row.is_heading and row.section_num is not None:
            sections.append({
                "num": row.section_num,
                "title": row.section_title,
                "row_index": row.index,
            })

    manifest = ReportManifest(
        input_path=str(path.resolve()),
        total_rows=len(rows_data),
        data_rows=sum(1 for r in rows_data if not r.is_heading),
        heading_rows=sum(1 for r in rows_data if r.is_heading),
        sections=sections,
        rows=[asdict(r) for r in rows_data],
        has_summary_column=any(r.summary for r in rows_data),
        forbidden_links_found=forbidden_links_found,
    )

    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract Huike DOCX content to structured JSON manifest."
    )
    parser.add_argument("input", type=Path, help="Path to Huike .docx draft")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output JSON path (default: <input_stem>_manifest.json)"
    )
    args = parser.parse_args()

    try:
        manifest = extract_docx(args.input)
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"ERROR: Failed to extract report: {exc}", file=sys.stderr)
        return 2

    output_path = args.output or args.input.with_name(args.input.stem + "_manifest.json")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(asdict(manifest), f, ensure_ascii=False, indent=2)

    # Summary to stdout
    print(f"Extracted {manifest.total_rows} rows ({manifest.data_rows} data, {manifest.heading_rows} headings)")
    print(f"Sections: {[f'{s['num']}.{s['title']}' for s in manifest.sections]}")
    if manifest.forbidden_links_found:
        print(f"WARNING: Forbidden links found: {manifest.forbidden_links_found}")
    print(f"Manifest written to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
