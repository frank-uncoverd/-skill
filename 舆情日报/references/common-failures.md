# Common Failures

Use this table when validating or repairing final DOCX output.

| Symptom | Likely Cause | Fix |
|:---|:---|:---|
| English-news URL missing | URL line was removed while extracting Chinese summary | Preserve lines containing `http://` or `https://` and lines containing Chinese text |
| Row shading offset | Rows were deleted after shading | Save, reopen, remove old shading, then shade final rows again |
| Summary format inconsistent | Chinese and English link logic diverged | Chinese link adjacent to body; English link preceded by one blank line |
| English summary still present | Mixed bilingual summary was treated as non-English and skipped | Actively strip English paragraphs; keep Chinese summary and URL |
| Extra blank line after URL | Empty XML paragraph nodes were retained accidentally | Remove extra `<w:p>` nodes from the cell XML |
| Section numbers skipped | Whole section was deleted and headings were not renumbered | Renumber remaining section headings consecutively |
| Media source not traced | Source name was replaced without finding original article | Search title and replace with first-author official source/link |
| Forbidden final link remains | Portal/self-media/scraping URL was left in output | Replace with original official source URL or flag for manual handling |
| Low-value China-market viewpoint remains | Mixed Chinese/English foreign-capital viewpoint was not blacklisted | Apply China-market viewpoint blacklist |
| Previous-trading-day market recap remains | Yesterday movement recap was not recognized as low value | Apply previous-trading-day market dynamic blacklist |
| English row wrongly deleted | Script matched English keywords too broadly | Delete only if topical blacklist applies; otherwise clean summary column only |
| Blank paragraph renders as two lines | Existing paragraph was deep-copied with `w:pPr` spacing | Use a clean `<w:p/>` with no children for blank paragraphs |