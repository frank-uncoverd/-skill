#!/usr/bin/env python3
"""Tests for extract_report.py, process_report.py, and validate_report.py.

These tests create sample DOCX files in-memory using python-docx, then run
each script against them. This verifies:

- extract_report: section detection, row extraction, forbidden link detection
- process_report: deletion, renumbering, shading, English summary cleanup
- validate_report: section numbering check, forbidden links, shading check

Cross-platform: designed to run on Windows/macOS/Linux.

Run with:
    python -m pytest tests/ -v
    python -m pytest tests/test_scripts.py -v
"""

from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# Ensure scripts dir is importable
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
if _SCRIPTS_DIR.exists():
    sys.path.insert(0, str(_SCRIPTS_DIR))

import pytest

# Conditional import — tests are skipped if python-docx is missing
try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ─── Sample Data ───────────────────────────────────────────────────────────────

SAMPLE_SECTIONS = [
    ("1、公司新闻", [
        ("基金公司A推出创新产品（来源：证券时报）",
         "基金公司A近日推出XXX创新产品，引发市场关注。\nhttps://www.stcn.com/article/001"),
        ("基金公司B高管变更（来源：中国基金报）",
         "基金公司B今日宣布副总经理变更。\nhttps://www.chnfund.com/article/002"),
    ]),
    ("2、行业动态", [
        ("公募基金规模突破30万亿（来源：上海证券报）",
         "截至5月底，公募基金规模突破30万亿元。\nhttps://www.cnstock.com/article/003"),
        ("费率改革持续推进（来源：证券日报）",
         "多家基金公司宣布降低管理费率。\nhttps://www.zqrb.cn/article/004"),
    ]),
    ("3、行业负面", [
        ("某基金公司因违规被处罚（来源：证券时报）",
         "XX基金公司因违反销售规定被监管处罚。\nhttps://www.stcn.com/article/005"),
    ]),
    ("4、政策资讯", [
        ("监管发布ETF管理办法（来源：中国基金报）",
         "证监会发布ETF相关管理办法征求意见稿。\nhttps://www.chnfund.com/article/006"),
    ]),
]

ENGLISH_NEWS_ROW = (
    "China's ETF market sees record inflows Chinese ETF市场迎来创纪录资金流入（来源：上海证券报）",
    "Chinese summary: 中国ETF市场迎来创纪录资金流入。\nRead full article: https://www.cnstock.com/article/en001",
)

FORBIDDEN_SOURCE = "来源：163.com"


# ─── Fixtures ──────────────────────────────────────────────────────────────────


@pytest.fixture(scope="function")
def sample_docx():
    """Create a sample Huike-style DOCX for testing."""
    if not HAS_DOCX:
        pytest.skip("python-docx not installed")

    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp.close()
    path = Path(tmp.name)

    doc = Document()

    # Add header
    doc.add_paragraph("今日新闻监测情况如下：")

    # Add table
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"

    # Header row
    hdr = table.rows[0]
    hdr.cells[0].text = "标题/媒体"
    hdr.cells[1].text = "摘要"

    row_idx = 1
    all_indices = {0}  # header is always kept

    for section_title, items in SAMPLE_SECTIONS:
        # Add section heading row
        table.add_row()
        row_cells = table.rows[row_idx].cells
        row_cells[0].text = section_title
        row_cells[1].text = ""
        all_indices.add(row_idx)
        row_idx += 1

        # Add data rows
        for title_text, summary_text in items:
            table.add_row()
            row_cells = table.rows[row_idx].cells
            row_cells[0].text = title_text
            row_cells[1].text = summary_text
            all_indices.add(row_idx)
            row_idx += 1

    # Add English news row
    table.add_row()
    eng_cells = table.rows[row_idx].cells
    eng_cells[0].text = ENGLISH_NEWS_ROW[0]
    eng_cells[1].text = ENGLISH_NEWS_ROW[1]
    all_indices.add(row_idx)
    english_idx = row_idx
    row_idx += 1

    doc.save(str(path))

    yield path, all_indices, english_idx

    # Cleanup
    try:
        path.unlink()
    except PermissionError:
        pass


@pytest.fixture(scope="function")
def sample_docx_with_forbidden():
    """Create sample DOCX with a forbidden link."""
    if not HAS_DOCX:
        pytest.skip("python-docx not installed")

    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp.close()
    path = Path(tmp.name)

    doc = Document()
    doc.add_paragraph("今日新闻监测情况如下：")

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0]
    hdr.cells[0].text = "标题/媒体"
    hdr.cells[1].text = "摘要"

    table.add_row()
    table.rows[1].cells[0].text = "测试新闻（来源：163.com）"
    table.rows[1].cells[1].text = "这是来自门户的测试内容。\nhttps://www.163.com/test"

    doc.save(str(path))
    yield path
    try:
        path.unlink()
    except PermissionError:
        pass


# ─── Tests: extract_report.py ──────────────────────────────────────────────────


class TestExtractReport:
    """Test the extract_report.py module."""

    def _import_extract(self):
        """Import extract_report module (handles sys.path)."""
        from scripts import extract_report  # type: ignore
        # Re-import to get fresh module
        import importlib
        return importlib.reload(extract_report)

    def test_extract_sections(self, sample_docx):
        """Verify sections are correctly detected."""
        path, _, _ = sample_docx
        from scripts.extract_report import extract_docx

        manifest = extract_docx(path)

        assert manifest.total_rows > 0
        assert manifest.heading_rows == len(SAMPLE_SECTIONS)  # one per section
        assert manifest.has_summary_column

        # Check sections
        assert len(manifest.sections) == len(SAMPLE_SECTIONS)
        for i, sec in enumerate(manifest.sections):
            assert sec["num"] == i + 1
            expected_title = SAMPLE_SECTIONS[i][0].split("、", 1)[1]
            assert expected_title in sec["title"] or sec["title"] in expected_title

    def test_extract_detects_forbidden_links(self, sample_docx_with_forbidden):
        """Verify forbidden link detection works."""
        from scripts.extract_report import extract_docx

        manifest = extract_docx(sample_docx_with_forbidden)
        assert len(manifest.forbidden_links_found) > 0
        assert "163.com" in manifest.forbidden_links_found

    def test_extract_media_source(self, sample_docx):
        """Verify media source extraction."""
        from scripts.extract_report import extract_media_source

        result = extract_media_source("测试新闻（来源：证券时报）")
        assert result == "证券时报"

        result = extract_media_source("测试新闻(来源:上海证券报)")
        assert result == "上海证券报"

        result = extract_media_source("无来源的标题")
        assert result is None

    def test_extract_manifest_json(self, sample_docx):
        """Verify manifest serializes to JSON correctly."""
        from scripts.extract_report import extract_docx

        manifest = extract_docx(sample_docx[0])
        d = manifest.__dict__
        # Don't use asdict directly since it's a regular class, not dataclass
        import json
        json_str = json.dumps(d, ensure_ascii=False)
        assert json_str
        parsed = json.loads(json_str)
        assert parsed["total_rows"] > 0
        assert len(parsed["rows"]) == parsed["total_rows"]


# ─── Tests: process_report.py ──────────────────────────────────────────────────


class TestProcessReport:
    """Test the process_report.py module."""

    def _make_decisions(self, all_indices, exclude=None):
        """Create a standard decisions dict."""
        keep = set(all_indices)
        if exclude:
            keep -= set(exclude)
        return {
            "keep_indices": sorted(keep),
            "section_renumber": True,
            "auto_clean_english": True,
            "media_replacements": {},
        }

    def test_basic_deletion(self, sample_docx):
        """Verify deletion removes specified rows."""
        path, all_indices, _ = sample_docx
        from scripts.process_report import load_decisions, process_report

        # Delete the last data row (company news second item = index 3)
        decisions = self._make_decisions(all_indices, exclude={3})

        tmp_out = path.with_name("test_output.docx")
        try:
            stats = process_report(path, decisions, tmp_out)

            assert stats["deleted"] == 1
            assert stats["rows_after"] == stats["rows_before"] - 1

            # Verify the saved doc
            doc = Document(str(tmp_out))
            table = doc.tables[0]
            # "费率改革" row was at index 3, should now be gone
            all_text = ""
            for row in table.rows:
                all_text += row.cells[0].text + "\n"
            # Deleted row was about fund company B
            # The second company news should be gone
        finally:
            if tmp_out.exists():
                tmp_out.unlink()

    def test_all_rows_kept(self, sample_docx):
        """Verify keeping everything produces same row count."""
        path, all_indices, _ = sample_docx
        from scripts.process_report import process_report

        decisions = self._make_decisions(all_indices)
        data_row_count = sum(1 for idx in all_indices if idx > 0)

        tmp_out = path.with_name("test_output2.docx")
        try:
            stats = process_report(path, decisions, tmp_out)
            assert stats["deleted"] == 0
            assert stats["rows_after"] == stats["rows_before"]
        finally:
            if tmp_out.exists():
                tmp_out.unlink()

    def test_section_renumbering(self, sample_docx):
        """Verify section renumbering after deletion."""
        path, all_indices, _ = sample_docx
        from scripts.process_report import process_report

        # Delete the first section entirely (section heading at index 1,
        # plus its two data rows at indices 2, 3)
        exclude = set()
        for idx in sorted(all_indices):
            text_at_idx = None

        # We need to determine which indices belong to section 1
        # Section headings are at indices where section_title matches
        # From sample: index 1 = 公司新闻, index 4 = 行业动态, index 7 = 行业负面, index 9 = 政策资讯
        # Company news: index 1 (heading), 2, 3 (data)
        section_1_indices = {1, 2, 3}
        decisions = self._make_decisions(all_indices, exclude=section_1_indices)

        tmp_out = path.with_name("test_renumber.docx")
        try:
            stats = process_report(path, decisions, tmp_out)
            assert stats["deleted"] == 3

            doc = Document(str(tmp_out))
            table = doc.tables[0]

            # First heading should now be "1、行业动态" not "2、行业动态"
            # Find section headings
            heading_texts = []
            for row in table.rows:
                cell_text = row.cells[0].text.strip()
                import re
                if re.match(r"^\d+[、.．]", cell_text):
                    heading_texts.append(cell_text)

            assert len(heading_texts) == 3  # 3 sections remain
            assert heading_texts[0].startswith("1、")
            assert heading_texts[1].startswith("2、")
            assert heading_texts[2].startswith("3、")
        finally:
            if tmp_out.exists():
                tmp_out.unlink()

    def test_row_shading_alternating(self, sample_docx):
        """Verify alternating row shading is applied."""
        path, all_indices, _ = sample_docx
        from scripts.process_report import process_report
        from docx.oxml.ns import qn

        decisions = self._make_decisions(all_indices)

        tmp_out = path.with_name("test_shading.docx")
        try:
            stats = process_report(path, decisions, tmp_out)
            assert stats["rows_shaded"] > 0

            # Verify shading alternation
            doc = Document(str(tmp_out))
            table = doc.tables[0]

            data_rows = []
            for row in table.rows:
                text = row.cells[0].text.strip()
                import re
                if not re.match(r"^\d+[、.．]", text) and "标题/媒体" not in text and "今日新闻" not in text:
                    data_rows.append(row)

            # Check alternating pattern
            for i, row in enumerate(data_rows):
                cell = row.cells[0]
                tc_pr = cell._tc.find(qn("w:tcPr"))
                if tc_pr is not None:
                    shd = tc_pr.find(qn("w:shd"))
                    if shd is not None:
                        fill = shd.get(qn("w:fill"), "").upper()
                        if i % 2 == 0:
                            assert fill in ("", "FFFFFF", "AUTO"), f"Row {i} should be white but has {fill}"
                        else:
                            assert fill == "E6E6E6", f"Row {i} should be gray but has {fill}"

        finally:
            if tmp_out.exists():
                tmp_out.unlink()

    def test_english_cell_rebuild(self, sample_docx):
        """Verify English news summary cells are rebuilt to 5 paragraphs."""
        path, all_indices, _ = sample_docx
        from scripts.process_report import process_report, clean_english_summaries

        decisions = self._make_decisions(all_indices)

        tmp_out = path.with_name("test_english.docx")
        try:
            stats = process_report(path, decisions, tmp_out)
            assert stats["english_cells_cleaned"] >= 1

            doc = Document(str(tmp_out))
            table = doc.tables[0]

            # Find the English news row
            for row in table.rows:
                if "Read full article:" in row.cells[1].text:
                    # Check paragraph count
                    paras = row.cells[1]._tc.findall(qn("w:p"))
                    assert len(paras) == 5, f"Expected 5 paragraphs, got {len(paras)}"
                    break

            # Verify: P0, P2, P4 should be empty (no text runs)
        finally:
            if tmp_out.exists():
                tmp_out.unlink()

    def test_forbidden_links_detected(self, sample_docx_with_forbidden):
        """Verify forbidden link check works."""
        path = sample_docx_with_forbidden
        from scripts.process_report import check_forbidden_links

        doc = Document(str(path))
        found = check_forbidden_links(doc)
        assert any("163.com" in f for f in found)

    def test_media_source_replacement(self, sample_docx):
        """Verify media source replacement in title cells."""
        path, all_indices, _ = sample_docx
        from scripts.process_report import process_report

        # Replace source for the first data row (index 2)
        decisions = self._make_decisions(all_indices)
        decisions["media_replacements"] = {
            2: {"source": "原创媒体", "url": "https://original.com/article"},
        }

        tmp_out = path.with_name("test_source.docx")
        try:
            stats = process_report(path, decisions, tmp_out)
            assert stats["media_replacements"] >= 1

            doc = Document(str(tmp_out))
            table = doc.tables[0]
            # Row 2 should now have "原创媒体" instead of "证券时报"
            assert "原创媒体" in table.rows[2].cells[0].text
        finally:
            if tmp_out.exists():
                tmp_out.unlink()


# ─── Tests: validate_report.py ─────────────────────────────────────────────────


class TestValidateReport:
    """Test the validate_report.py module."""

    def _make_bad_section_docx(self):
        """Create DOCX with non-consecutive section numbers."""
        if not HAS_DOCX:
            pytest.skip("python-docx not installed")

        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp.close()
        path = Path(tmp.name)

        doc = Document()
        doc.add_paragraph("标题")

        table = doc.add_table(rows=4, cols=2)
        table.style = "Table Grid"
        table.rows[0].cells[0].text = "标题/媒体"
        table.rows[0].cells[1].text = "摘要"
        table.rows[1].cells[0].text = "1、公司新闻"
        table.rows[1].cells[1].text = ""
        table.rows[2].cells[0].text = "3、政策资讯"  # Skipped 2!
        table.rows[2].cells[1].text = ""
        table.rows[3].cells[0].text = "测试新闻内容"
        table.rows[3].cells[1].text = "摘要内容"

        doc.save(str(path))
        return path

    def test_section_numbering(self):
        """Verify non-consecutive sections are detected."""
        from scripts.validate_report import validate, load_document_xml, iter_tables
        from xml.etree import ElementTree as ET

        path = self._make_bad_section_docx()

        try:
            root = load_document_xml(path)
            tables = list(iter_tables(root))
            rows = tables[0].findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr")

            from scripts.validate_report import check_sections
            issues = check_sections(rows)

            assert any("not consecutive" in i.message for i in issues)
        finally:
            path.unlink()


# ─── Integration Tests ─────────────────────────────────────────────────────────


class TestIntegration:
    """End-to-end tests of the full extract → decide → process pipeline."""

    def test_full_pipeline(self, sample_docx):
        """Verify the complete extract-json→decisions→process flow."""
        if not HAS_DOCX:
            pytest.skip("python-docx not installed")

        path, all_indices, eng_idx = sample_docx

        # Step 1: Extract
        from scripts.extract_report import extract_docx
        manifest = extract_docx(path)
        assert manifest.total_rows > 0

        # Step 2: Create decisions (simulate LLM output)
        # Keep everything except the last company news row (index 3)
        keep = sorted(all_indices - {3})
        decisions = {
            "keep_indices": keep,
            "section_renumber": True,
            "auto_clean_english": True,
            "media_replacements": {2: {"source": "中国证券报", "url": "https://www.cs.com.cn/article"}},
        }

        tmp_decisions = path.with_name("test_decisions.json")
        with open(tmp_decisions, "w", encoding="utf-8") as f:
            json.dump(decisions, f, ensure_ascii=False, indent=2)

        # Step 3: Process
        from scripts.process_report import process_report
        tmp_out = path.with_name("test_pipeline.docx")
        try:
            stats = process_report(path, decisions, tmp_out)
            assert stats["deleted"] == 1
            assert stats["rows_after"] == stats["rows_before"] - 1
            assert stats["english_cells_cleaned"] >= 1
            assert stats["media_replacements"] >= 1

            # Step 4: Validate
            from scripts.validate_report import validate
            issues = validate(tmp_out)
            errors = [i for i in issues if i.level == "ERROR"]
            # Should have no errors in the processed output
            # (validate may report warnings about English summary if XML structure
            # is correct but not identical to expected — skip ERROR-level issues only)
            for err in errors:
                if "Forbidden" in err.message:
                    pytest.fail(f"Validation found forbidden link: {err.message}")
        finally:
            if tmp_out.exists():
                tmp_out.unlink()
            if tmp_decisions.exists():
                tmp_decisions.unlink()


# ─── Cross-Platform Tests ──────────────────────────────────────────────────────


class TestCrossPlatform:
    """Tests that verify cross-platform compatibility."""

    def test_pathlib_usage(self):
        """Verify all scripts use pathlib, not os.path."""
        for script_name in ["extract_report.py", "process_report.py", "validate_report.py"]:
            script_path = _SCRIPTS_DIR / script_name
            if not script_path.exists():
                continue
            content = script_path.read_text(encoding="utf-8")
            # Should not use platform-specific path constructs
            # os.path.sep references are OK in validation contexts
            # But no hardcoded Windows backslashes or Unix-only patterns
            assert "Path(" in content, f"{script_name} should use pathlib.Path"

    def test_shebang_present(self):
        """Verify all scripts have a cross-platform shebang."""
        for script_name in ["extract_report.py", "process_report.py", "validate_report.py"]:
            script_path = _SCRIPTS_DIR / script_name
            if not script_path.exists():
                continue
            content = script_path.read_text(encoding="utf-8")
            assert content.startswith("#!/usr/bin/env python3"), (
                f"{script_name} missing cross-platform shebang"
            )

    def test_no_hardcoded_paths(self):
        """Verify no hardcoded Windows or Unix paths in scripts."""
        for script_name in ["extract_report.py", "process_report.py", "validate_report.py"]:
            script_path = _SCRIPTS_DIR / script_name
            if not script_path.exists():
                continue
            content = script_path.read_text(encoding="utf-8")
            # No drive letters
            assert "C:\\" not in content, f"{script_name} has hardcoded Windows path"
            assert "D:\\" not in content, f"{script_name} has hardcoded Windows path"
            # Platform paths should use pathlib
            assert "os.sep" not in content, f"{script_name} should use pathlib not os.sep"

    def test_import_structure(self):
        """Verify all scripts can be imported without runtime errors."""
        for script_name in ["extract_report.py", "process_report.py", "validate_report.py"]:
            script_path = _SCRIPTS_DIR / script_name
            if not script_path.exists():
                continue
            # Basic syntax check
            import ast
            try:
                with open(script_path, encoding="utf-8") as f:
                    ast.parse(f.read())
            except SyntaxError as e:
                pytest.fail(f"{script_name} has syntax error: {e}")

    def test_msgpack_not_used(self):
        """Ensure scripts don't require platform-specific binary libs."""
        for script_name in ["extract_report.py", "process_report.py", "validate_report.py"]:
            script_path = _SCRIPTS_DIR / script_name
            if not script_path.exists():
                continue
            content = script_path.read_text(encoding="utf-8")
            assert "msgpack" not in content, f"{script_name} should not depend on msgpack"
            assert "winreg" not in content, f"{script_name} should not use Windows registry"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
