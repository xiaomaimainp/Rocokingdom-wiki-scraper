# 洛克王国世界 Wiki 图鉴爬虫（Rocokingdom-wiki-scraper）

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

# 抓取「图鉴课题」聚合页，生成技能石清单 skill_stones.csv（一次静态请求覆盖全部有课题的物种）
python nrc_scraper.py skillstones

# 用聚合页交叉验证主爬取的技能石课题（按序号比对差异）
python nrc_scraper.py validate
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
| `evo_chains.csv` | 进化链：家族 / 分支 / 阶序 / 源精灵(+序号) → 目标精灵(+序号) + 进化条件 + 是否特殊 | 全部 |
| `skill_stones.csv` | 技能石清单：序号 / 精灵名称 / 类别 / 条件 / 奖励ID（按物种去重，每物种一行） | `skillstones` |
| `skill_stones_diff.csv` | 交叉验证差异明细（仅 `validate` 有差异时生成） | `validate` |

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

## 进化链格式（evo_chains.csv）

目标列：`家族, 分支, 阶序, 源精灵, 源精灵序号, 目标精灵, 目标精灵序号, 进化条件, 是否特殊`

- 每行 = 一条进化步骤（`源精灵` → `目标精灵`）。同一家族的所有形态会铺开成多行；`家族` 取进化链最底层的那个精灵。
- `分支`：来自页面 `data-branch`（普通线性进化都是 `1`；存在分支进化时会出现 `2`、`3`…）。
- `阶序`：来源形态的阶位标签，如 `初阶 · 冰`。
- `源精灵序号` / `目标精灵序号`：对应精灵在 `spirits_data_all.csv` 中的 `序号`（图鉴主 key）。进化链里的精灵名可能带「（…）」形态后缀，而 `spirits_data_all.csv` 中同名精灵常以形态后缀区分，因此这里按「基础名（忽略形态后缀）→ 序号」对齐，方便用序号直接关联两张表；对齐不上的会在体检中报「按序号找不到」。
- `进化条件`：进化所需条件文本，例如 `LV28 突破1次`、`LV40 生涯使用[滚雪球]15次`。
- `是否特殊`：**是** = 该家族存在分支进化（`data-branch>1`），或任一进化步骤的条件**不是纯等级进化**。只有「`LVx`」或「`LVx 进化`」这类纯等级进化算**否**（普通）；突破（`LVx 突破y次`）、天气/时间、性别（雌性）、技能使用（生涯使用[X]N次）、道具（信物首领化）等附加条件一律算 **是**（特殊）。

示例：

```csv
家族,分支,阶序,源精灵,源精灵序号,目标精灵,目标精灵序号,进化条件,是否特殊
小雪人,1,初阶 · 冰,小雪人,066,雪怪,067,LV40 生涯使用[滚雪球]15次,是
钨丝贝贝,1,初阶 · 机械,钨丝贝贝,348,辉光幕机,349,LV28 突破1次,是
钨丝贝贝,1,二阶 · 机械,辉光幕机,349,机幕方舟,350,LV40 突破2次,是
```

> 解析自每个精灵页的「进化链」面板（`.roco-evo-wrap` → `.roco-evo-timeline` → `.roco-evo-node` / `.roco-evo-conn`）。一个家族的完整链在任意成员页都能拿到，脚本按家族去重，只保留一次。

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

### 进化链（`parse_evo_chain` / `build_evo_chains`）
- 解析精灵页「进化链」面板：`.roco-evo-wrap` 下每个 `.roco-evo-timeline`（`data-branch` 标分支）是 `节点 → 连接 → 节点` 的序列；
- 节点名在 `.roco-evo-name-main a[title]`（当前页为 `selflink`，直接取文本），阶位在 `.roco-evo-name-sub`；进化条件在 `.roco-evo-conn > .roco-evo-conn-pill > span`；
- 同一家族的完整链在任意成员页都能拿到，按家族去重（只保留一次）；`是否特殊` = 分支>1 或任一条件不是纯等级进化（突破、天气/时间、性别、技能使用、道具等附加条件都算特殊，只有 `LVx` / `LVx 进化` 为普通）。
- 生成后按「基础名（忽略形态后缀）→ 序号」与 `spirits_data_all.csv` 对齐，补上 `源精灵序号` / `目标精灵序号`，体检改为按序号比对（消除形态后缀造成的误报）。
- 最终 `evo_chains.csv` 按 `源精灵序号` 升序（0→x）输出，家族内步骤与分支保持原有顺序。

### 合并表（`build_combined`）
- **课题（技能石）以官方「图鉴课题」聚合页（`skill_stones.csv`）为准**：按**序号（物种）一行**，技能石归属与官方页完全一致——分支形态（如 191 的「棋契陛下 / 棋祈督」）不再被拆成重复行，技能也不会被错归到其他序号；`属性` 取自精灵表的该序号主属性，`精灵名称` 取官方页的物种名。
- 因此 `combined_data.csv` 的技能石**数量与官方聚合页逐条一致**（无重复、无遗漏）；无聚合页数据时回退到「按基础名归组来源表」的旧逻辑。

### 技能石清单与交叉验证（`skillstones` / `validate`）
- `python nrc_scraper.py skillstones` 抓取 wiki 的「图鉴课题」聚合页，解析全部技能石课题，写出 `skill_stones.csv`（也是 `update` / `rebuild` 构建合并表的数据源）：
  - 数据内嵌在聚合页静态 HTML 的 `book="..."` 属性（HTML 实体转义 JSON），**一次静态请求即可覆盖全部有课题的物种**，无需渲染、无需翻页；
  - 每个物种一行（按序号 + 名称），每一条技能石课题一行；列为 `序号, 精灵名称, 类别, 条件, 奖励ID`；`类别` 固定为 `技能石`，`条件` 即「使用N次XX」，`奖励ID` 为物品 ID（含技能石本体 + 课题研究点）。
- `python nrc_scraper.py validate` 用 `skill_stones.csv` **交叉验证**主爬取（逐精灵页）的技能石课题：按序号对比两者条件集合，列出「仅主爬取有 / 仅聚合页有」的差异，并写出 `skill_stones_diff.csv`（验证归属映射已修正，目前 202 物种全部一致、0 差异）。

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

## 转载与使用

除另有标注外，本站原创文字及有权授权的资料整理成果，采用知识共享「署名—相同方式共享 4.0 国际」协议（CC BY-SA 4.0）发布。

- 协议说明（中文）：<https://creativecommons.org/licenses/by-sa/4.0/deed.zh>
- 协议说明（英文）：<https://creativecommons.org/licenses/by-sa/4.0/>
- 完整法律文本：<https://creativecommons.org/licenses/by-sa/4.0/legalcode>

转载、引用或使用本站的页面内容及整理数据，须注明本站名称，并附上原页面链接。对受上述协议保护的内容，还须保留适用的作者信息、附上许可链接，并说明是否作过修改；改编成果须以相同协议发布。

## 游戏素材

本站展示的游戏角色、立绘、图标、截图、音视频、标识等素材，其著作权、商标权及相关权利归魔方工作室、腾讯游戏及相应权利人所有。其他第三方内容的权利归各自权利人所有。

这些素材不是自由版权内容，不适用本站的 CC BY-SA 4.0 协议。本站仅为学术研究、评论、教学及百科介绍，在适用著作权法允许的合理使用范围内引用相关内容。素材在本站出现，并不意味着本站拥有其权利，或能够授权他人使用。

如对素材引用、署名或使用范围有异议，请提供相关页面与说明，联系维护者核实处理。
