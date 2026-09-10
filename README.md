# 洛克王国世界 Wiki 图鉴爬虫（nrc_scraper.py）

基于 [Scrapling](https://github.com/D4Vinci/Scrapling) 的洛克王国世界 B 站 wiki（nrc）图鉴爬虫，用于抓取**精灵图鉴**、**技能档案**与**技能来源 / 获取条件**，并生成合并表。

相比传统 `requests + BeautifulSoup`，使用 `Fetcher` / `StealthyFetcher` 自动伪装 TLS 指纹与请求头，抗反爬能力更强；详情页抓取采用并发 Spider 框架（`ThreadPoolExecutor` + 退避重试 + 失败串行补抓），速度更快、更稳。

> 本站只抓新版 wiki：`https://wiki.biligame.com/nrc/`。旧站 `rocom`（数据截至 S3）相关代码与参数已全部移除。

## 环境要求

- Python 3.9+
- 需要浏览器内核（`StealthyFetcher` 依赖 patchright / playwright 的 chromium）

### 安装依赖

```bash
pip install "scrapling[fetchers]"
scrapling install              # 下载 chromium（Fetcher 用）
patchright install chromium    # 下载 patchright 专用 chromium（StealthyFetcher 用）
```

> 注意：`patchright install chromium` 与 `scrapling install` 下载的 chromium 版本不同，两者都需要安装，否则 `StealthyFetcher` 会报浏览器缺失。

## 使用方法

在脚本所在目录运行：

```bash
# 抓取 → 合并 → 复检（有旧数据会先体检并报告差异，只补新增）
python nrc_scraper.py

# 只体检已有数据文件，不联网
python nrc_scraper.py check

# 不联网，用已有 CSV 重算合并表与「技能石课题」列
python nrc_scraper.py rebuild
```

- 第一次没有数据也能直接跑，会全量建立；
- 之后每次运行都会自动保留旧数据、只追加本次新抓到的内容，并打印与上次的差异。

## 输出文件

| 文件 | 行数 | 说明 | 生成模式 |
| --- | --- | --- | --- |
| `spirits_data_all.csv` | 622 | 精灵图鉴完整数据（序号、名称、阶数、进化位置、主副属性、形态、异色、赛季、是否有技能石课题） | 全部 |
| `spirits_data_all.xlsx` | 622 | 同上，带自动筛选与冻结表头的 Excel 版 | 全部 |
| `skills_data.csv` | 579 | 技能档案（技能名、属性、类型、威力、耗能、技能描述、风味文本） | 全部 |
| `skill_sources.csv` | 30772 | 精灵 × 技能 来源（谁会这招、怎么学：自学 / 技能石 / 血脉 / 传说，含获取条件） | 全部 |
| `combined_data.csv` | 203 | 合并表：精灵 → 最多 3 个技能石 + 获取条件 | 全部 / `rebuild` |

## 合并表格式（combined_data.csv）

目标列：`序号, 精灵名称, 属性, 技能名称1, 技能名称2, 技能名称3, 条件`

- 每个精灵最多展示 **3 个**技能石技能。
- `条件` 列 = 该精灵每个技能石技能的获取条件用 `; ` 拼接，例如 `使用1次闪光; 使用1次光刃`。
- **一个序号一行**：
  - 同一精灵的不同形态（如 `鸭吉吉（燃了鸭）`）条件相同，合并成一行，名称取括号前的**基础名**；
  - 同一序号的进化链成员（如 `迪莫 / 圣光迪莫`）课题相同，也合并成一行，保留**本体形态**（有「精灵阶数」的那个）；
  - 条件不同的仍各自成行（如 `191` 的 `棋契陛下` 与 `棋祈督`）。
- 没有技能石课题的精灵不出现在该表，所以 622 个精灵最终只有 203 行。

示例：

```csv
序号,精灵名称,属性,技能名称1,技能名称2,技能名称3,条件
001,迪莫,光,闪光,光刃,,使用1次闪光; 使用1次光刃
010,水灵,水,水炮,潮汐,,使用2次水炮; 使用2次潮汐
011,鸭吉吉,普通,防御,吓退,伺机而动,使用1次防御; 使用1次吓退; 使用1次伺机而动
```

## 实现要点

### 抓取底层（`fetch` / `fetch_with_retry`）
- 默认走 `Fetcher.get(url, impersonate='chrome', timeout=20000)`；
- BWiki 限流时会返回 HTTP 200 的「**简化页**」（内容不完整），因此由调用方用 `is_ok(page)` 判定内容是否完整，失败则退避重试；
- 重试仍失败时换 `StealthyFetcher.fetch(headless=True, network_idle=True)` 兜底一次。

### 并发抓取器（`Spider`）
- `ThreadPoolExecutor` 并发（`MAX_WORKERS = 3`）+ 每个请求间隔 `REQUEST_DELAY = 0.3s`；
- 失败的页面最后改为**串行补抓**（`SERIAL_GAP = 2.0s`），避免继续触发限流。

### 精灵图鉴（`crawl_spirits`）
- 抓取「精灵图鉴」页，解析 `div.npc-card`；
- 从属性中提取 序号 / 阶数 / 进化位置 / 主副属性 / 形态 / 异色 / 赛季 等字段。

### 技能档案（`crawl_skills`）
- 先从「技能图鉴」页解析 `div.nrc-skill-catalog-card`，取去重后的技能链接（跳过 `redlink=1`）；
- 再并发抓技能详情页，解析 `.nrc-skill-tags` / `.nrc-skill-type-chip` / `.nrc-skill-category-chip` / `.nrc-skill-effect-text` / `.nrc-skill-flavor` / `.nrc-skill-stat`；
- 这里**只取技能自身属性**，「谁会这招」由精灵页产出（覆盖面更全）。

### 技能来源与获取条件（`crawl_sources`）
- 逐个抓精灵详情页的 `div.roco-sk-card`，读 `data-source` 区分为 自学 / 技能石 / 血脉 / 传说；
- **获取条件的唯一正确来源**是精灵页的「图鉴课题」面板 `div.roco-quest-content > div.roco-quest-row`：
  - 只保留 `使用N次XX` 这类课题，技能名取 `.roco-quest-text` 内嵌的 `a[title]`；
  - ⚠️ 不能用 `.roco-sk-dcard` 的 `.roco-sk-source-task`，那个任务挂在别的精灵身上，会串数据。

### 合并表（`build_combined`）
- 汇总每只精灵的图鉴课题（技能 + `使用N次` 条件），按基础名归组；
- 代表形态优先取「有阶数」的本体，再按「序号 + 条件」合并同一进化链。

## 配置项（文件顶部）

| 常量 | 默认 | 说明 |
| --- | --- | --- |
| `MAX_WORKERS` | 3 | 并发数，调大会触发 BWiki 限流 |
| `REQUEST_DELAY` | 0.3 | 每个请求前的固定间隔（秒） |
| `RETRY_TIMES` | 3 | 单页重试次数 |
| `SERIAL_GAP` | 2.0 | 串行补抓间隔（秒） |
| `MAX_SKILLS` | 3 | 合并表每个精灵最多展示几个技能 |

## 注意事项

- 网站结构可能随时间变化。若某模块抓取为空，多半是解析选择器需要同步更新，可对照浏览器实际 DOM 调整。
- 运行会在当前目录生成日志 `nrc_scraper_<时间>.log`（已被 `.gitignore` 忽略）。
- Windows 下中文输出已统一为 UTF-8（`sys.stdout.reconfigure`，文件用 `utf-8-sig` 写入）。
