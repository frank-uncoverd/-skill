---
name: huike-public-opinion-daily
description: >-
  Convert Huike public-opinion daily report drafts into final polished DOCX reports for Chinese mutual-fund industry monitoring. Use when the user provides a Huike/慧科舆情日报 DOCX draft and asks to screen, filter, polish, repair formatting, or produce a final version. Uses deterministic Python scripts for mechanical DOCX operations (hardcoded filtering rules, English summary cleanup, alignment, alternating shading).
---
# 慧科舆情日报终版处理

## Core Workflow

When the user provides a Huike DOCX draft, execute this **three-phase** sequence:

### Phase 1 — Extract & Analyze

1. Confirm the input is a `.docx` Huike report with one main table containing `标题/媒体` and `摘要` columns.
2. Run `python scripts/extract_report.py <input.docx>` to produce a structured JSON manifest (optional, for reference).
3. Read the manifest output or the raw extraction file to understand the report structure — sections, English rows, blacklist candidates.

### Phase 2 — Filter Decision (Hardcoded Rules)

4. Apply filtering rules from `references/filter-rules.md` and the hardcoded `DELETE_ORIGINAL_ROWS` set (in `process_report.py` or a wrapper) to determine which rows to keep:

   - **Whitelist** rules (不可删 rows) — industry system changes, regulatory rules, major fund-flow behavior.
   - **Section handling** — `公司新闻` rows: move official statements/regulatory penalties to the appropriate section; delete routine product tracking / manager interviews / investor education.
   - **Blacklist** — delete rows matching: 公募调研, 基金经理观点, 投教科普, 券商金股, 地方政策, PE/VC, 情绪噱头, 非公募, 上一交易日市场动态, 英文·中国市场主题.
   - **English rows** — keep bilingual title/source; clean summary column to 5-paragraph XML structure.
   - **Deduplication** — same-event rows: prefer authoritative original source.
   - **Media source replacement** — trace to first-author original article; replace source name and URL.
5. **Show removable rows summary to user** and ask for confirmation:

   ```
   === 待删除行审核 ===
   原始行数: 36
   待删除: 14 行
     ├ 非公募（机器人/原油/科技基金）: 3行（Row 3-5）
     ├ 公募市场观点/分红/市场回顾: 5行（Row 6,8,9,11,13）
     ├ 黄金价格宏观分析（与ETF重复）: 1行（Row 10）
     ├ 非金融/非公募话题: 3行（Row 22,23,28）
     └ 市场动态/资金流向: 2行（Row 15,16）
   English新闻摘要清理: 5行
   ─────────────────────────────────
   确认执行? (是/否)
   ```

### Phase 3 — Process & Validate

6. Once user confirms, run `process_report.py` (or the simplified wrapper):

   ```bash
   python scripts/process_report.py <input.docx> decisions.json -o <原文件名>_终版.docx
   ```

   Key processing steps:

   - Delete blacklist/duplicate rows
   - Renumber sections after deletions
   - Clean English news summary cells to 5-paragraph XML
   - Apply alternating row shading (white / #E6E6E6)
   - **Apply alignment rules**:
     - Section heading rows (e.g. `1、行业动态`): **左对齐 + 垂直居中**
     - Format rows (`今日新闻监测情况如下：` / `标题/媒体`): 居中
     - Data rows: 标题列居中 + 垂直居中; 摘要列左对齐
7. Run `python scripts/validate_report.py <终版.docx>` for final quality check.
8. If validation errors found, fix and re-run step 6-7.
9. Report the final path and processing summary.

### Source 巡检 (Optional)

Check missed stories using `references/source-check.md` unless the user asks to skip or network/browser access is unavailable.

## Decision Priority

Apply rules in this priority order:

1. Explicit user instruction in the current conversation.
2. Whitelist and dynamic keep rules in `references/filter-rules.md`.
3. Section deletion and blacklist rules.
4. Deduplication and media-source replacement rules.
5. Formatting and validation rules.

If two written rules appear to conflict, keep the higher-priority rule and mention the ambiguity in the delivery note.

## Important Conflict Resolution

English-news handling has one exception:

- General English news must not be deleted only because it is English. Clean the summary column and keep the bilingual title/source.
- If an English-news row is specifically about low-value China-market/Chinese-ETF viewpoint content and matches the blacklist in `references/filter-rules.md`, it may be deleted as a topical blacklist hit.

## Bundled Resources

Read only the resource needed for the current step:

- `references/filter-rules.md`: keep/delete criteria, whitelist, blacklist, section handling, media replacement.
- `references/docx-format.md`: final DOCX layout, English summary XML, row shading, section renumbering, alignment rules.
- `references/source-check.md`: daily missed-story巡检 sites, fallback methods, recommendation output.
- `references/common-failures.md`: known failure modes and fixes.
- `scripts/process_report.py`: deterministic script to apply filtering decisions (keeps/deletes rows, renumbers sections, applies shading, cleans English summaries, applies alignment).
- `scripts/extract_report.py`: reads the DOCX and outputs a structured JSON manifest for LLM review.
- `scripts/validate_report.py`: static DOCX validation for section numbering, forbidden links, row shading, alignment, and English summary structure hints.

## Alignment Rules

Apply in `apply_alignment()` after row shading:

`今日新闻监测情况如下：` 左对齐

| Row Type                     |  Title/Media Cell |      Summary Cell |
| :--------------------------- | ----------------: | ----------------: |
| Section heading (`1、xxx`) | 左对齐 + 垂直居中 | 左对齐 + 垂直居中 |
| Format row (`标题/媒体`)   |              居中 |              居中 |
| Data row (news item)         |   居中 + 垂直居中 |            左对齐 |

## Environment & Dependencies

### Installation (cross-platform)

```bash
# Required
pip install python-docx>=1.1.0

# Optional — for source巡检 (browser-based):
pip install playwright
playwright install chromium

# For running tests:
pip install pytest
```

### Platform Notes

All scripts (`extract_report.py`, `process_report.py`, `validate_report.py`) are pure Python 3 and work on:

- **Windows** (PowerShell, cmd, Git Bash, WSL)
- **macOS** (Terminal)
- **Linux**

No platform-specific shell commands or path handling. Paths with spaces and non-ASCII characters are handled correctly on all platforms.

## Tooling Scripts

Located in `scripts/` directory:

| Script                 | Purpose                               | When to Run      |
| :--------------------- | :------------------------------------ | :--------------- |
| `extract_report.py`  | Extract DOCX → JSON manifest         | Phase 1          |
| `process_report.py`  | Apply decisions → produce final DOCX | Phase 3          |
| `validate_report.py` | Static quality check                  | After processing |

### decisions.json Format

When producing decisions.json for process_report.py, use this structure:

```json
{
  "keep_indices": [0, 1, 2, 4, 5, 7, 10, ...],
  "section_renumber": true,
  "media_replacements": {
    "5": {"source": "中国证券报", "url": "https://www.cs.com.cn/article/xxx"},
    "12": {"source": "证券时报", "url": "https://www.stcn.com/article/xxx"}
  },
  "english_rows": [8, 15],
  "auto_clean_english": true,
  "review_note": "删除13行（公募调研4/投教科普3/...），来源替换3处"
}
```

- `keep_indices`: **Required**. 0-based row indices of rows to keep. All other rows are deleted.
- `section_renumber`: Auto-renumber sections. Default `true`.
- `media_replacements`: Optional. Maps row index → new source name + URL.
- `auto_clean_english`: Auto-detect and rebuild English summary cells to 5-paragraph structure. Default `true`.
- `review_note`: Optional. Summary of changes for the delivery note.

### Running Tests

```bash
cd scripts/..
python -m pytest tests/ -v
```
