# Source Check

After filtering the Huike draft, scan recent stories from the core source sites and recommend missed items to the user. If browser/network access fails, continue with the other sites and report failures.

## Source List

The current list contains 6 sources. Keep this count consistent unless a seventh source is added.

| # | Source | URL | Method |
|:---|:---|:---|:---|
| 1 | 中国基金报·公募 | https://www.chnfund.com/fund | Web fetch/browser |
| 2 | 证券时报·基金 | https://www.stcn.com/article/list/fund.html | Web fetch/browser |
| 3 | 上海证券报·基金 | https://www.cnstock.com/channel/10033 | Web fetch/browser |
| 4 | 证券日报·基金 | http://www.zqrb.cn/fund/ | PowerShell/curl fallback if web fetch is blocked |
| 5 | 财联社·深度 | https://www.cls.cn/depth?id=1110 | Browser/Playwright for JS rendering |
| 6 | 中证网·ETF/ESG | https://www.cs.com.cn/jijin.html | Browser/Playwright for JS rendering |

## Search Time Window

The search window is determined by the **report date** extracted from the input filename, NOT by the current system time.

### Step 1: Extract Report Date

Parse the input filename for a date in `YYYYMMDD` format (e.g., `易方达舆情日报20260609.docx` → date = **2026-06-09**).

### Step 2: Determine Search Window

The search covers the 24-hour period ending at **8:00 AM on the report date**:

```
Normal window:
  [报告日期 - 1天 08:00]  ──────────────────────→  [报告日期 08:00]
  (前一日早上8点)                                    (报告日早上8点)
```

### Step 3: Holiday Extension

If the **day immediately before the report date** (`报告日期 - 1天`) falls on a **non-workday** (Saturday, Sunday, or Chinese public holiday), extend the window start back to **8:00 AM of the most recent preceding workday**.

```
Extended window (前一日为节假日):
  [上一个工作日 08:00]  ──···──  [节假日]  ──  [报告日期 08:00]
```

#### Holiday Calendar Reference

| Year | Holiday | Typical Dates (YYYY-MM-DD) |
|:---|:---|:---|
| 2026 | 元旦 | 01-01 (Thu) |
| 2026 | 春节 | 02-17 (Tue) ~ 02-23 (Mon) |
| 2026 | 清明节 | 04-04 (Sat) ~ 04-06 (Mon) |
| 2026 | 劳动节 | 05-01 (Fri) ~ 05-05 (Tue) |
| 2026 | 端午节 | 06-19 (Fri) ~ 06-21 (Sun) |
| 2026 | 国庆+中秋 | 09-27 (Sun) ~ 10-04 (Sun) |

Also check **Weekends**: every Saturday and Sunday is a non-workday.

> **Implementation note**: The simplest approach is to check if `报告日期 - 1天` is Saturday (5) or Sunday (6). If yes, walk backwards day-by-day until a weekday (Mon-Fri) is found. For Chinese public holidays (调休日), manually verify the specific date against the current year's holiday schedule.

### Step 4: Filter Results

Scan each source site for articles published **within the time window** (by article timestamp or page publish date). Extract titles and links, focusing on:

- Public funds (公募基金)
- ETF
- ESG
- Regulatory policy
- Fee reform
- Industry rules
- Meaningful fund-flow or holder-structure changes

### Examples

| Filename Date | Day-of-week | Day-before | Window |
|:---|---:|---:|:---|
| 20260609 (Tue) | 周二 | 周一(workday) | 06-08 08:00 → 06-09 08:00 |
| 20260608 (Mon) | 周一 | 周日(❌ holiday) → 前周五 06-05 | 06-05 08:00 → 06-08 08:00 |
| 20260601 (Mon) | 周一 | 周日(❌ holiday) → 前周五 05-29 | 05-29 08:00 → 06-01 08:00 |
| 20260104 (Sun?) | 周日 | 周六(❌) + 元旦(❌?) → 前周三 or 调休日 | Verify against holiday calendar |

## Scan Goal

Extract titles and links from each source site, limited to the **determined time window**. Focus on public funds, ETF, ESG, regulatory policy, fee reform, industry rules, and meaningful fund-flow or holder-structure changes.

## Fallbacks

- If `zqrb.cn` over HTTP fails, try `https://www.zqrb.cn`.
- If a JS-rendered source fails, try installed Chrome/Chromium through Playwright or the available browser tool.
- If a site returns 403/502 or requires manual verification, record the failed site and continue.

## Cross-Check

Compare source-site findings against the draft/final report:

- Exact duplicate: skip.
- Same event with better angle/source: mark as replace/merge candidate.
- Not present in Huike draft: mark as supplement candidate.

## Recommendation Priority

| Priority | Feature | Action |
|:---:|:---|:---|
| P0 | Major regulatory policy, management measures, fee reform | Directly recommend supplementing |
| P1 | Exclusive data, deep investigation, major ETF/fund-flow change | Recommend with short summary |
| P2 | Routine industry dynamic or product issuance | Briefly mention |
| Skip | Fund manager interview, investor education, event-only story | Do not recommend |

## Output Format

```text
【推荐来源】XXX
【标题】XXX
【与初稿差异】慧科未收录 / 角度不同
【推荐理由】XXX（20字内）
【链接】原始URL
```
