#!/usr/bin/env python3
"""Validate Huike public-opinion daily report DOCX files.

This script performs static checks that are easy to miss after filtering:
section numbering, forbidden links, table row shading, and English-news summary
structure hints. It does not decide which news rows should be kept or deleted.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
FORBIDDEN_LINK_PATTERNS = [
    "163.com",
    "qq.com",
    "sohu.com",
    "sina.com.cn",
    "mp.weixin.qq.com",
    "weixin",
    "app-web.chnfund.com",
    "api3.cls.cn",
    "h.xinhuaxmt.com",
]
SECTION_RE = re.compile(r"^(\d+)、(.+)$")
EXPECTED_GRAY = "E6E6E6"


@dataclass
class Issue:
    level: str
    message: str


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def load_document_xml(path: Path) -> ET.Element:
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Expected .docx input, got: {path}")
    with zipfile.ZipFile(path) as zf:
        with zf.open("word/document.xml") as handle:
            return ET.parse(handle).getroot()


def iter_tables(root: ET.Element) -> Iterable[ET.Element]:
    yield from root.findall(".//w:tbl", NS)


def cell_text(cell: ET.Element) -> str:
    parts = [node.text or "" for node in cell.findall(".//w:t", NS)]
    return "".join(parts).strip()


def row_cells(row: ET.Element) -> list[ET.Element]:
    return row.findall("w:tc", NS)


def row_text(row: ET.Element) -> str:
    return " | ".join(cell_text(cell) for cell in row_cells(row)).strip()


def first_cell_text(row: ET.Element) -> str:
    cells = row_cells(row)
    return cell_text(cells[0]).replace(" ", "") if cells else ""


def is_format_row(row: ET.Element) -> bool:
    text = first_cell_text(row)
    return text in {"今日新闻监测情况如下：", "标题/媒体"} or bool(SECTION_RE.match(text))


def row_shading(row: ET.Element) -> str | None:
    fills: list[str] = []
    for cell in row_cells(row):
        shd = cell.find("w:tcPr/w:shd", NS)
        if shd is not None:
            fill = shd.attrib.get(qn("w:fill"))
            if fill and fill.lower() != "auto":
                fills.append(fill.upper())
    if not fills:
        return None
    unique = sorted(set(fills))
    return unique[0] if len(unique) == 1 else ",".join(unique)


def check_sections(rows: list[ET.Element]) -> list[Issue]:
    issues: list[Issue] = []
    section_numbers: list[int] = []
    for row in rows:
        text = first_cell_text(row)
        match = SECTION_RE.match(text)
        if match:
            section_numbers.append(int(match.group(1)))
    if section_numbers:
        expected = list(range(1, len(section_numbers) + 1))
        if section_numbers != expected:
            issues.append(Issue("ERROR", f"Section numbers are not consecutive: {section_numbers}; expected {expected}"))
    else:
        issues.append(Issue("WARN", "No numbered section headings found."))
    return issues


def check_forbidden_links(root: ET.Element) -> list[Issue]:
    issues: list[Issue] = []
    text = "\n".join(node.text or "" for node in root.findall(".//w:t", NS))
    lowered = text.lower()
    for pattern in FORBIDDEN_LINK_PATTERNS:
        if pattern.lower() in lowered:
            issues.append(Issue("ERROR", f"Forbidden link/source remains: {pattern}"))
    return issues


def check_shading(rows: list[ET.Element]) -> list[Issue]:
    issues: list[Issue] = []
    data_index = 0
    for row_index, row in enumerate(rows, start=1):
        if is_format_row(row):
            shade = row_shading(row)
            if shade is not None:
                issues.append(Issue("ERROR", f"Format row {row_index} has shading {shade}; heading/header rows should be unshaded."))
            continue
        if not row_text(row):
            continue
        expected = None if data_index % 2 == 0 else EXPECTED_GRAY
        actual = row_shading(row)
        if expected is None:
            if actual not in (None, "FFFFFF"):
                issues.append(Issue("ERROR", f"Data row {row_index} should be white/unshaded but has {actual}."))
        elif actual != expected:
            issues.append(Issue("ERROR", f"Data row {row_index} should be gray {expected} but has {actual or 'no shading'}."))
        data_index += 1
    return issues


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()


def is_clean_empty_paragraph(paragraph: ET.Element) -> bool:
    return paragraph_text(paragraph) == "" and len(list(paragraph)) == 0


def check_english_summary_structure(rows: list[ET.Element]) -> list[Issue]:
    issues: list[Issue] = []
    for row_index, row in enumerate(rows, start=1):
        cells = row_cells(row)
        if len(cells) < 2:
            continue
        title_text = cell_text(cells[0])
        summary = cells[1]
        summary_text = cell_text(summary)
        if "Read full article:" not in summary_text:
            continue
        paragraphs = summary.findall("w:p", NS)
        if len(paragraphs) != 5:
            issues.append(Issue("ERROR", f"Row {row_index} English summary has {len(paragraphs)} paragraphs; expected 5."))
            continue
        for position in (0, 2, 4):
            if not is_clean_empty_paragraph(paragraphs[position]):
                issues.append(Issue("ERROR", f"Row {row_index} English summary P{position} is not a clean empty paragraph."))
        if "Read full article:" not in paragraph_text(paragraphs[3]):
            issues.append(Issue("ERROR", f"Row {row_index} English summary P3 does not contain Read full article URL."))
        if title_text and not re.search(r"[A-Za-z]", title_text):
            issues.append(Issue("WARN", f"Row {row_index} has Read full article link but title column does not look bilingual/English."))
    return issues


def validate(path: Path) -> list[Issue]:
    root = load_document_xml(path)
    tables = list(iter_tables(root))
    issues: list[Issue] = []
    if not tables:
        return [Issue("ERROR", "No tables found in document.")]
    if len(tables) > 1:
        issues.append(Issue("WARN", f"Document has {len(tables)} tables; expected one main Huike table."))
    rows = tables[0].findall("w:tr", NS)
    issues.extend(check_sections(rows))
    issues.extend(check_forbidden_links(root))
    issues.extend(check_shading(rows))
    issues.extend(check_english_summary_structure(rows))
    return issues


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a final Huike public-opinion daily DOCX.")
    parser.add_argument("docx", type=Path, help="Path to final .docx file")
    args = parser.parse_args()

    try:
        issues = validate(args.docx)
    except Exception as exc:  # noqa: BLE001 - CLI should report any validation loading problem.
        print(f"ERROR: {exc}")
        return 2

    if not issues:
        print("OK: no validation issues found")
        return 0

    for issue in issues:
        print(f"{issue.level}: {issue.message}")
    return 1 if any(issue.level == "ERROR" for issue in issues) else 0


if __name__ == "__main__":
    raise SystemExit(main())
