# DOCX Format Rules

Core principle: preserve the draft format. Only change table content and the formatting explicitly listed here.

## Title/Media Column

Each news item should appear as:

```text
标题行

（来源：XXX）
```

- Center the title text for data rows. **Section heading rows (e.g. `1、行业动态`): left-align + vertical center.**
- Keep the source line visually separated under the title.
- Preserve existing table and paragraph styles unless a correction is required.

## Summary Column

- Left-align the summary column for all row types.
- For Chinese rows, keep one leading blank line before summary body and one trailing blank line after the final URL.
- Chinese links should directly follow the body text without an extra blank separator.
- English `Read full article:` links require one blank line before the link.

## Alignment Rules

Apply after row shading in `apply_alignment()`:

| Row Type | Title/Media Cell | Summary Cell |
|:---|---:|---:|
| Section heading (`1、xxx`) | **左对齐 + 垂直居中** | **左对齐 + 垂直居中** |
| Format row (`今日新闻监测情况如下：` / `标题/媒体`) | 居中 | 居中 |
| Data row (news item) | 居中 + 垂直居中 | 左对齐 |

## English-News Summary XML

English-news summary cells must be rebuilt as exactly five `<w:p>` paragraph nodes:

1. P0: empty paragraph, completely clean, no `w:pPr`, visually exactly one blank line.
2. P1: Chinese summary, preserving useful `w:rPr` font/color/spacing properties.
3. P2: empty paragraph, completely clean, no `w:pPr`.
4. P3: `Read full article: URL`, preserving hyperlink-like run properties such as underline/color when present.
5. P4: empty paragraph, completely clean, no `w:pPr`.

Do not `deepcopy` an existing paragraph to create blank paragraphs. It can carry `w:pPr` spacing and render as two visual blank lines. Create a clean paragraph element directly, for example with an OOXML element equivalent to `<w:p/>`.

Preserve the English-news title column exactly, including English title, Chinese title, and source. Only clean the summary column unless the row matches a topical blacklist rule.

## Section Renumbering

After filtering, check whether deleting whole sections caused skipped section numbers.

- Renumber remaining section headings with consecutive Arabic numerals.
- Only update the heading number. Do not rewrite news item content.
- Example: `1、公司新闻`, `2、行业动态`, `4、政策资讯` becomes `1、行业动态`, `2、政策资讯` after deleting the first section.

## Row Shading

Row shading must be checked after all deletions and moves.

1. Save the document and reopen it before final shading if using `python-docx`, so row indexes reflect the final document.
2. Remove old shading from all data rows.
3. Starting from the first data row, apply white then gray alternation: white, `#E6E6E6`, white, `#E6E6E6`.
4. Section heading/format rows must not be shaded and must not participate in the alternation count.
5. If shading is not strictly alternating, rerun the shading step.

This is a high-frequency failure point because row indexes change after deletion.

## Final Visual Checklist

- Title column: every news title is centered, with one blank line before source.
- **Section heading rows: left-aligned + vertical centered.**
- Summary column: left aligned, with expected leading and trailing blank lines.
- Links: Chinese links are adjacent to body text; English links have one blank line before `Read full article:`.
- Shading: first data row white, second gray, section heading rows unshaded.
- Section numbers: continuous `1、2、3...`.
- English summary XML: `[empty clean p, Chinese summary, empty clean p, URL, empty clean p]`.
- Forbidden links: no portal, self-media, or scraping-source URLs.