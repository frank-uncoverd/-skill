#!/usr/bin/env python3
"""Apply filtering decisions to a Huike DOCX report in-place.

This script takes two inputs:
  1. The original Huike .docx draft
  2. A decisions.json file produced by the LLM after reviewing the manifest

It mechanically applies the decisions:
  - Delete rows marked "delete" by blacklist/section rules
  - Keep rows marked "keep" (including whitelist-protected)
  - Renumber sections after deletions
  - Alternating row shading (white/gray)
  - Clean English summary cells to 5-paragraph XML structure
  - Replace media source names and URLs
  - Remove forbidden links

Cross-platform: pure Python, works on Windows/macOS/Linux.

Usage:
    python process_report.py <input.docx> <decisions.json> -o <output.docx>
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    print("ERROR: python-docx is required. Install with: pip install python-docx", file=sys.stderr)
    sys.exit(3)


# ─── Constants ────────────────────────────────────────────────────────────────

FORBIDDEN_LINK_PATTERNS = [
    "163.com", "qq.com", "sohu.com", "sina.com.cn",
    "mp.weixin.qq.com", "weixin",
    "app-web.chnfund.com", "api3.cls.cn", "h.xinhuaxmt.com",
]
SECTION_RE = re.compile(r"^(\d+)[、.．]\s*(.+)$")
GRAY = "E6E6E6"


# ─── Helpers ───────────────────────────────────────────────────────────────────


def cell_text(cell: Any) -> str:
    """Get plain text of a table cell."""
    return cell.text.strip() if cell else ""


def find_table(doc: Document) -> Any | None:
    """Find the main Huike data table. Returns None if no suitable table found."""
    if not doc.tables:
        return None
    # Huike tables typically have "标题/媒体" as a header cell
    for t in doc.tables:
        for row in t.rows[:3]:
            for cell in row.cells:
                if "标题" in cell.text and "媒体" in cell.text:
                    return t
    # Fallback: first table with enough rows
    for t in doc.tables:
        if len(t.rows) > 3:
            return t
    return doc.tables[0] if doc.tables else None


def is_section_heading(text: str) -> bool:
    """Check if text looks like a section heading (e.g. '1、公司新闻')."""
    return bool(SECTION_RE.match(text.replace(" ", "")))


def get_section_number(text: str) -> int | None:
    """Extract section number from heading text."""
    m = SECTION_RE.match(text.replace(" ", ""))
    return int(m.group(1)) if m else None


def get_section_title(text: str) -> str | None:
    """Extract section title text from heading."""
    m = SECTION_RE.match(text.replace(" ", ""))
    return m.group(2).strip() if m else None


# ─── Core Operations ────────────────────────────────────────────────────────────


def _remove_shading(cell: Any) -> None:
    """Remove shading from a cell completely."""
    tc_pr = cell._tc.find(qn("w:tcPr"))
    if tc_pr is not None:
        shd = tc_pr.find(qn("w:shd"))
        if shd is not None:
            tc_pr.remove(shd)


def _set_shading(cell: Any, color: str | None) -> None:
    """Set shading on a cell. color=None means no shading/white."""
    tc_pr = cell._tc.find(qn("w:tcPr"))
    if tc_pr is None:
        tc_pr = OxmlElement("w:tcPr")
        cell._tc.insert(0, tc_pr)

    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)

    if color:
        shd.set(qn("w:fill"), color)
        shd.set(qn("w:val"), "clear")
    else:
        shd.set(qn("w:fill"), "FFFFFF")
        shd.set(qn("w:val"), "clear")


def apply_deletions(table: Any, keep_indices: set[int]) -> int:
    """Delete rows not in keep_indices. Works bottom-up to preserve indices."""
    deleted = 0
    rows = table.rows
    for idx in range(len(rows) - 1, -1, -1):
        if idx not in keep_indices:
            tr = rows[idx]._tr
            tr.getparent().remove(tr)
            deleted += 1
    return deleted


def renumber_sections(table: Any) -> list[dict[str, Any]]:
    """Renumber section headings consecutively after deletions.

    Returns the updated section list.
    """
    sections: list[dict[str, Any]] = []
    for row in table.rows:
        text = cell_text(row.cells[0]).strip() if row.cells else ""
        if is_section_heading(text):
            title = get_section_title(text) or ""
            sections.append({
                "title": title,
                "text": text,
                "row": row,
            })

    updated: list[dict[str, Any]] = []
    for i, sec in enumerate(sections, start=1):
        old_num = get_section_number(sec["text"])
        if old_num != i:
            _replace_section_number_in_cell(sec["row"].cells[0], i, sec["title"])
        updated.append({"num": old_num, "title": sec["title"], "new_num": i})

    return updated


def _replace_section_number_in_cell(cell: Any, new_num: int, title: str) -> None:
    """Replace the section number in the first matching run."""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            m = re.match(r"^(\d+)[、.．]\s*", run.text)
            if m:
                rest = run.text[len(m.group(0)):]
                if rest:
                    run.text = f"{new_num}、{rest}"
                else:
                    run.text = f"{new_num}、{title}"
                return


def apply_row_shading(table: Any) -> int:
    """Apply alternating row shading (white, #E6E6E6) to data rows.

    Heading/format rows are skipped and not counted in the alternation.
    """
    # Remove all existing shading first
    for row in table.rows:
        for cell in row.cells:
            _remove_shading(cell)

    data_idx = 0
    rows_shaded = 0
    for row in table.rows:
        text = cell_text(row.cells[0]).strip() if row.cells else ""
        if (is_section_heading(text)
                or "标题/媒体" in text
                or "今日新闻" in text
                or "监测情况" in text):
            continue
        if not row.cells:
            continue

        color = GRAY if data_idx % 2 == 1 else None
        for cell in row.cells:
            _set_shading(cell, color)

        data_idx += 1
        rows_shaded += 1

    return rows_shaded


def _split_english_summary(text: str) -> tuple[str, str]:
    """Split English-news summary into Chinese text and URL."""
    lines = text.split("\n")
    chinese_parts: list[str] = []
    url = ""

    for line in lines:
        line = line.strip()
        if "Read full article:" in line or "read full article:" in line.lower():
            url_part = re.sub(r"(?i)read\s*full\s*article\s*:?\s*", "", line).strip()
            if url_part:
                url = url_part
        elif line and re.search(r"[一-鿿]", line):
            chinese_parts.append(line)

    return " ".join(chinese_parts), url


def _rebuild_cell_paragraphs(cell: Any, chinese_text: str, url: str) -> None:
    """Rebuild a cell's paragraphs to exactly 5 clean <w:p> nodes."""
    # Remove all existing paragraphs
    for p in list(cell._tc.findall(qn("w:p"))):
        cell._tc.remove(p)

    def _add_empty_p() -> None:
        p = OxmlElement("w:p")
        cell._tc.append(p)

    def _add_text_p(text: str) -> None:
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)
        p.append(r)
        cell._tc.append(p)

    # P0: empty
    _add_empty_p()
    # P1: Chinese summary
    _add_text_p(chinese_text) if chinese_text else _add_empty_p()
    # P2: empty
    _add_empty_p()
    # P3: Read full article: URL
    _add_text_p(f"Read full article: {url}")
    # P4: empty
    _add_empty_p()


def clean_english_summaries(table: Any) -> list[int]:
    """Rebuild English-news summary cells to 5-paragraph structure.

    Returns list of modified row indices.
    """
    modified: list[int] = []
    for row_idx, row in enumerate(table.rows):
        if len(row.cells) < 2:
            continue
        summary_text = cell_text(row.cells[1])
        if "read full article:" not in summary_text.lower():
            continue

        chinese_text, url = _split_english_summary(summary_text)
        if not url:
            continue

        # Only process rows with bilingual titles (contain Latin text)
        title_text = cell_text(row.cells[0]) if row.cells else ""
        if not re.search(r"[A-Za-z]", title_text):
            continue

        _rebuild_cell_paragraphs(row.cells[1], chinese_text, url)
        modified.append(row_idx)

    return modified


def replace_media_sources(table: Any, replacements: dict[int, dict[str, str]]) -> int:
    """Replace media source names and URLs for specified row indices."""
    replaced = 0
    for row_idx, info in replacements.items():
        if row_idx >= len(table.rows):
            continue
        row = table.rows[row_idx]
        if not row.cells:
            continue

        title_cell = row.cells[0]
        source = info.get("source", "")
        url = info.get("url", "")

        if source:
            for paragraph in title_cell.paragraphs:
                for run in paragraph.runs:
                    new_text = re.sub(
                        r"[（(]\s*来源\s*[:：]\s*[^）)]+\s*[）)]",
                        f"（来源：{source}）",
                        run.text
                    )
                    if new_text != run.text:
                        run.text = new_text
                        replaced += 1

        if url and len(row.cells) > 1:
            summary_cell = row.cells[1]
            for paragraph in summary_cell.paragraphs:
                for run in paragraph.runs:
                    run.text = re.sub(r"https?://[^\s)）]+", url, run.text)

    return replaced


def check_forbidden_links(doc: Document) -> list[str]:
    """Check for any remaining forbidden links."""
    all_text = ""
    for p in doc.paragraphs:
        all_text += p.text + "\n"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_text += cell.text + "\n"

    lowered = all_text.lower()
    found: list[str] = []
    for pattern in FORBIDDEN_LINK_PATTERNS:
        if pattern.lower() in lowered:
            found.append(pattern)
    return found


def apply_alignment(table: Any) -> None:
    """Apply alignment rules:

    - Section heading rows (e.g. 1、行业动态): left-align text + vertical center
    - 今日新闻监测情况如下：and 标题/media rows: center both cells
    - Data rows: title column center + vertical center; summary column left-align
    """
    for row in table.rows:
        if not row.cells:
            continue
        text = cell_text(row.cells[0]).strip()
        if is_section_heading(text):
            # Section heading: left-align + vertical center
            row.cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in row.cells[0].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for cell in row.cells[1:]:
                if cell_text(cell):
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif "标题/媒体" in text or "今日新闻" in text or "监测情况" in text:
            # Format rows: center both cells
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            # Data rows: title cell center + vertical center; summary left-align
            row.cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in row.cells[0].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if len(row.cells) > 1:
                for paragraph in row.cells[1].paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


# ─── Decision Loading ──────────────────────────────────────────────────────────


def load_decisions(path: Path) -> dict[str, Any]:
    """Load and validate a decisions JSON file.

    Expected format:
    {
        "keep_indices": [0, 1, 2, ...],   # 0-based row indices to keep
        "section_renumber": true,
        "media_replacements": {
            "5": {"source": "中国证券报", "url": "https://..."},
            ...
        },
        "english_rows": [7, 12, ...],     # row indices of English-news rows
        "auto_clean_english": true,        # auto-detect and clean English summaries
        "review_note": "..."              # optional summary of changes
    }
    """
    if not path.exists():
        raise FileNotFoundError(f"Decisions file not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        data: dict[str, Any] = json.load(f)

    if "keep_indices" not in data:
        raise ValueError("decisions.json must contain 'keep_indices' list")

    keep = set(data["keep_indices"])
    if not keep:
        raise ValueError("keep_indices is empty — this would delete everything")

    return data


# ─── Main Processing ───────────────────────────────────────────────────────────


def process_report(
    input_path: Path,
    decisions: dict[str, Any],
    output_path: Path,
) -> dict[str, Any]:
    """Apply decisions to a Huike DOCX report and write output.

    Returns a summary dict with results.
    """
    doc = Document(str(input_path))
    table = find_table(doc)
    if table is None:
        raise ValueError(
            "No suitable table found in the DOCX. "
            "Expected a Huike report with 标题/媒体 columns."
        )

    total_before = len(table.rows)

    # 1. Delete rows
    keep = set(decisions.get("keep_indices", []))
    deleted = apply_deletions(table, keep)

    # Save and reopen so row indices stabilize
    doc.save(str(output_path))
    doc = Document(str(output_path))
    table = find_table(doc)
    if table is None:
        raise ValueError("Table lost after deletion phase — unexpected DOCX corruption.")

    total_after = len(table.rows)

    # 2. Renumber sections
    sections_updated = []
    if decisions.get("section_renumber", True):
        sections_updated = renumber_sections(table)

    # 3. Replace media sources
    replacements = decisions.get("media_replacements", {})
    media_replaced = 0
    if replacements:
        conv: dict[int, dict[str, str]] = {}
        for k, v in replacements.items():
            conv[int(k)] = v
        media_replaced = replace_media_sources(table, conv)

    # 4. Clean English summaries
    auto_clean = decisions.get("auto_clean_english", True)
    english_modified = []
    if auto_clean:
        english_modified = clean_english_summaries(table)

    # 5. Row shading
    rows_shaded = apply_row_shading(table)

    # 6. Apply alignment rules
    apply_alignment(table)

    # Save intermediate for re-read
    doc.save(str(output_path))
    doc_final = Document(str(output_path))
    forbidden = check_forbidden_links(doc_final)
    doc_final.save(str(output_path))

    return {
        "rows_before": total_before,
        "rows_after": total_after,
        "deleted": deleted,
        "sections_renumbered": len(sections_updated) > 0,
        "sections": sections_updated,
        "rows_shaded": rows_shaded,
        "english_cells_cleaned": len(english_modified),
        "media_replacements": media_replaced,
        "forbidden_links_remaining": forbidden,
        "output_path": str(output_path.resolve()),
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Apply LLM decisions to a Huike public-opinion daily DOCX report."
    )
    parser.add_argument("input", type=Path, help="Path to the original Huike .docx draft")
    parser.add_argument("decisions", type=Path, help="Path to decisions.json")
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="Output .docx path (default: <input_stem>_终版.docx)")
    args = parser.parse_args()

    output_path = args.output or args.input.with_name(args.input.stem + "_终版.docx")

    if output_path == args.input:
        print("ERROR: Output path must differ from input path.", file=sys.stderr)
        return 2

    try:
        decisions = load_decisions(args.decisions)
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as exc:
        print(f"ERROR: Failed to load decisions: {exc}", file=sys.stderr)
        return 2

    try:
        stats = process_report(args.input, decisions, output_path)
    except (ValueError, FileNotFoundError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"ERROR: Unexpected processing error: {exc}", file=sys.stderr)
        return 2

    # Report results
    print(f"=== Processing Summary ===")
    print(f"Rows: {stats['rows_before']} → {stats['rows_after']} ({stats['deleted']} deleted)")
    print(f"Sections renumbered: {stats['sections_renumbered']}")
    print(f"Rows shaded: {stats['rows_shaded']}")
    print(f"English cells cleaned: {stats['english_cells_cleaned']}")
    print(f"Media sources replaced: {stats['media_replacements']}")
    print(f"Output: {stats['output_path']}")

    if stats["forbidden_links_remaining"]:
        print(f"WARNING: Forbidden links still present: {stats['forbidden_links_remaining']}")

    return 1 if stats["forbidden_links_remaining"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
