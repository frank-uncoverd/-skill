# Filtering Rules

Apply these rules to the Huike report table after extracting all section headings and rows.

## Input Shape

Expected section headings:

- `1、公司新闻`
- `2、行业动态`
- `3、行业负面`
- `4、政策资讯`
- `5、ETF资讯`
- `6、ESG资讯`

Expected columns: `标题/媒体` and `摘要`.

## Step 1: Whitelist Keep Rules

Scan every title and summary first. Mark matching rows as `不可删`; later blacklist rules must skip these rows unless the user explicitly overrides.

| Priority | Trigger | Reason |
|:---:|:---|:---|
| ★★★★ | `业绩比较基准`, `费率改革`, `降费`, `让利` | 行业制度变革 |
| ★★★★ | `暂停业务` or `暂停产品注册` when industry-level | 监管风向标 |
| ★★★★ | `合规大年`, `退薪`, `绩效薪酬` | 行业主线 |
| ★★★★ | `ESG强制披露`, `ESG评级变动` | 制度性里程碑 |
| ★★★ | `ETF规模` plus month/quarter trend | 非日度趋势 |
| ★★★ | `持有人结构`, `机构持有` | 资金行为分析 |
| ★★★ | `央行货币政策`, `逆周期` | 核心宏观 |
| ★★★ | `管理办法`, `实施细则`, `监管指引` | 监管规则发布 |
| ★★ | `QDII` plus `恢复申购` or `新增额度` | 跨境投资信号 |

Also keep rows with these features even without exact keyword hits:

- Regulatory rules, including titles ending in `办法`, `细则`, `指引`, or `规定`.
- Industry system changes involving fees, settlement, disclosure, or comparable rule changes.
- Major fund-flow behavior, including large subscriptions/redemptions, holder-structure shifts, or cross-border quota changes.

## Step 2: Section Handling

For `公司新闻`:

- Keep and move to the appropriate industry dynamic/negative section when it is an official industry statement, deep analysis, major scale record, material product innovation, regulatory penalty, or senior-management change.
- Delete routine product tracking, daily announcements, ordinary manager interviews, investor education, ETF daily data, and non-policy events.

Institution news shortcut:

| Level | Condition | Action |
|:---:|:---|:---|
| A | 监管处罚, 高管变动, 规模破纪录, 行业正式发声 | Move to industry negative/dynamic and keep |
| B | 常规产品发行, 经理专访, 投教, ETF日常数据 | Delete |

Mnemonic: `处罚首创必须留，产品经理一律删`.

## Step 3: Blacklist Delete Rules

Skip rows marked `不可删`. Delete rows matching these rules:

| Category | Delete Trigger |
|:---|:---|
| 公募调研 | `调研逾`, `调研次数`, `最受宠` |
| 基金经理观点 | `隐形重仓`, `攻防体系`, or `基金经` plus `专访`/`认为`/`展望` |
| 投教科普 | `全景图`, `你会买吗`, `播客`, `投教` |
| 券商金股 | `金股组合`, `金股指数` |
| 地方政策 | `地方` plus `政策`, unless clearly financial/fund-relevant |
| PE/VC | `股权`, `PE`, `VC`, `私募股权`, `一级市场` |
| 情绪噱头 | Title contains `大曝光`, `跑了`, or `来了` as hype wording |
| 非公募 | `医保`, `储能`, `北交所` when the topic is enterprise policy rather than public funds |
| 与公募无关 | `世界杯`, `算电协同`, `电力现货`, and other non-financial/non-public-fund topics |
| 活动论坛 | `大赛`, `论坛成功举办`, `研讨会` unless policy-class |
| 同源拥挤 | Same media has 3 or more non-whitelist rows; keep only 1-2 strongest rows |
| 中国市场观点 | Mixed Chinese/English viewpoint items with `看好中国`, `中国市场`, or `外资` plus `中国` |
| 上一交易日市场动态 | `反弹`, `放量`, `净流入`, `重获抱团`, `领涨`, `普涨`, `下跌` when merely reviewing the previous trading day |
| 英文新闻·中国市场主题 | English-news rows about `China market`, `Chinese ETF`, `中国`, or China ETF/market viewpoint content |

Must-delete examples:

- 港股迎久违放量反弹 机构聚焦AI算力与红利双主线（中国证券报）
- 南向资金持续净流入 机构看好港股中期行情（中国证券报）
- A股市场震荡反弹 硬科技方向重获资金抱团（上海证券报）
- 光通信产业链领涨 A股结构性行情延续（中国证券报）
- 科技股反弹，能否“逢低布局”？机构给出思路（国际金融报）
- 2026年世界杯开赛在即 A股公司多赛道征战（上海证券报）
- “算电协同”电力现货交易开启，互联网大厂会跟吗？（第一财经）
- China's ETF market suffers record quarterly outflow（英文·中国市场主题）

## Step 4: Attribute Cleanup and Deduplication

- Do not delete general English-news rows solely because they are English; clean their summary column according to `docx-format.md`.
- Preserve bilingual English-news title rows exactly, including English title, Chinese title, and source.
- Deduplicate same-event rows. Prefer authoritative original sources and deeper reporting.

## Step 5: Media Source Replacement

Core principle: trace to the first-author original article, replacing both source name and URL. Do not merely replace a display name.

Procedure:

1. Search by title through browser/search engine, including WeChat search or Baidu when useful.
2. Determine first author by priority: original media official site, then WeChat official account, then financial repost.
3. Replace source name and link with the original author official-site link when available.

Forbidden final links:

- Portals: `163.com`, `qq.com`, `sohu.com`, `sina.com.cn`
- Self-media: `mp.weixin.qq.com`, `weixin`
- Scraping sources: `app-web.chnfund.com`, `api3.cls.cn`, `h.xinhuaxmt.com`

Common source-tracing hints:

- `app-web.chnfund.com` should trace to `chnfund.com` original article.
- `api3.cls.cn` should trace to `cls.cn` original article.
- `h5.stcn.com` should trace to `stcn.com` original article.