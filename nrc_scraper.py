"""
洛克王国世界 WIKI（nrc）图鉴爬虫
=================================

只抓取新版 wiki：https://wiki.biligame.com/nrc/

基于 Scrapling：
    Fetcher          —— 主抓取（服务端渲染页，快）
    StealthyFetcher  —— 兜底（BWiki 对高频请求返回「简化页」时自动切换）

用法：
    python nrc_scraper.py            # 抓取 → 体检 → 合并 → 复检
    python nrc_scraper.py check      # 只体检已有数据文件，不联网
    python nrc_scraper.py rebuild    # 不联网，用已有 CSV 重算合并表
    python nrc_scraper.py skillstones # 抓取图鉴课题聚合页，生成 skill_stones.csv（技能石）
    python nrc_scraper.py validate   # 用聚合页交叉验证主爬取的技能石课题
    python nrc_scraper.py features   # 抓取特性图鉴，生成 features_data.csv 并补精灵表特性两列
    python nrc_scraper.py normalize  # 不联网，统一来源表的「来源类型」「等级」写法
    python nrc_scraper.py skillcatalog # 抓全部精灵页，生成 skill_catalog.csv（含血脉属性）
"""
import csv
import logging
import os
import random
import re
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote
import html as _html
import json

# Windows 下输出默认 GBK 会导致中文乱码，统一为 UTF-8
for _s in (sys.stdout, sys.stderr):
    if hasattr(_s, 'reconfigure'):
        getattr(_s, 'reconfigure')(encoding='utf-8')

from scrapling.fetchers import Fetcher, StealthyFetcher

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ====================== 配置 ======================
BASE = "https://wiki.biligame.com/nrc/"
SPIRIT_URL = BASE + quote("精灵图鉴")
SKILL_URL = BASE + quote("技能图鉴")
FEATURE_URL = BASE + quote("特性图鉴")

MAX_WORKERS = 3           # 并发数（调大会触发 BWiki 限流）
REQUEST_DELAY = 0.3       # 每个请求前的固定间隔（秒）
RETRY_TIMES = 3           # Fetcher 失败重试次数
SERIAL_GAP = 2.0          # 串行补抓的间隔（秒）

STAGE_MAP = {'1': '一阶', '2': '二阶', '3': '三阶'}
SHINY_MAP = {'no': '否', 'yes': '是'}
SEASON_MAP = {'none': '无'}
POSITION_MAP = {'initial|': '初始', 'initial|final': '初始|最终', '': ''}
SOURCE_MAP = {'level': '自学', 'machine': '技能石', 'blood': '血脉', 'legendary': '传说'}
MAX_SKILLS = 3            # combined_data.csv 每个精灵最多展示几个技能

# 输出表：
#   spirits_data_all.csv  精灵档案
#   skills_data.csv       技能档案（技能自身的属性，与精灵无关）
#   skill_sources.csv     精灵 × 技能 来源（谁在哪、怎么获得）
#   combined_data.csv     精灵 → 最多 3 个技能石 + 获取条件（视图，由聚合页算出）
#   evo_chains.csv / skill_stones.csv  进化链 / 图鉴课题聚合页
#   features_data.csv     特性图鉴（每个特性一行，含介绍与拥有精灵）
SNAP_DIR = '.snapshots'   # 最近两次抓取的键集合，作为「连续两次缺失才淘汰」的基底
SNAP_KEEP = 2
SPIRIT_FIELDS = ['序号', '精灵名称', '精灵阶数', '进化位置', '主属性',
                 '副属性', '形态', '主形态/地区形态', '异色', '赛季', '技能石课题',
                 '特性', '特性介绍']
SKILL_FIELDS = ['技能名称', '技能属性', '技能类型', '威力', '耗能',
                '技能描述', '风味文本']
SOURCE_FIELDS = ['精灵名称', '技能名称', '来源类型', '等级', '获取条件']
COMBINED_FIELDS = ['序号', '精灵名称', '属性', '技能名称1', '技能名称2', '技能名称3', '条件']
EVO_FIELDS = ['家族', '分支', '阶序', '源精灵', '源精灵序号', '目标精灵', '目标精灵序号', '进化条件', '是否特殊']
FEATURE_FIELDS = ['特性编号', '特性名称', '特性介绍', '主属性', '副属性',
                  '拥有精灵数', '拥有精灵']
CATALOG_FIELDS = ['精灵名称', '技能ID', '技能名称', '来源类型', '学习等级',
                  '血脉属性', '技能系别', '技能类别', '威力', '能耗', '技能描述', '顺序']

F_SPIRITS = 'spirits_data_all.csv'
F_SPIRITS_XLSX = 'spirits_data_all.xlsx'
F_SKILLS = 'skills_data.csv'
F_SOURCES = 'skill_sources.csv'
F_COMBINED = 'combined_data.csv'
F_EVO = 'evo_chains.csv'
F_STONES = 'skill_stones.csv'
F_FEATURES = 'features_data.csv'
F_CATALOG = 'skill_catalog.csv'
STONE_FIELDS = ['序号', '精灵名称', '类别', '条件', '奖励ID']

# 每张表的记录键：用于快照比对与合并去重（键变了就视为「旧记录消失 + 新记录出现」）
TABLE_KEYS = {
    'spirits': lambda r: f"{r.get('序号', '')}|{r.get('精灵名称', '')}",
    'skills': lambda r: str(r.get('技能名称', '')),
    'sources': lambda r: f"{r.get('精灵名称', '')}|{r.get('技能名称', '')}",
    'stones': lambda r: f"{r.get('序号', '')}|{r.get('条件', '')}",
    'evo': lambda r: '|'.join(str(r.get(k, '')) for k in
                             ('家族', '分支', '源精灵', '目标精灵', '进化条件')),
    'features': lambda r: str(r.get('特性编号') or r.get('特性名称', '')),
    'catalog': lambda r: f"{r.get('精灵名称', '')}|{r.get('技能名称', '')}|{r.get('来源类型', '')}",
}


# ====================== 日志 ======================
class _Tee:
    """把输出同时写到多个流（控制台 + UTF-8 日志文件）"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            try:
                s.write(data)
                s.flush()
            except (ValueError, OSError):
                pass

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except (ValueError, OSError):
                pass

    def isatty(self):
        return False


def setup_logging():
    log_path = f'nrc_scraper_{time.strftime("%Y%m%d_%H%M%S")}.log'
    log_fh = open(log_path, 'a', encoding='utf-8', buffering=1)
    sys.stdout = _Tee(sys.__stdout__, log_fh)
    sys.stderr = _Tee(sys.__stderr__, log_fh)
    # Scrapling 在导入时就把 handler 绑到了原始 stderr，需要改指向 tee
    for h in logging.getLogger('scrapling').handlers:
        if isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler):
            h.setStream(sys.stderr)
    print(f"运行日志已写入: {log_path}")
    return log_path


# ====================== 抓取底层 ======================
def fetch(url, stealth=False, timeout=20000):
    """统一入口：默认 Fetcher，stealth=True 时走 StealthyFetcher（带浏览器）"""
    if stealth:
        return StealthyFetcher.fetch(url, headless=True, network_idle=True)
    return Fetcher.get(url, impersonate='chrome', timeout=timeout)


def fetch_with_retry(url, is_ok, retries=RETRY_TIMES):
    """抓取直到 is_ok(page) 为真。

    BWiki 限流时返回 HTTP 200 的「简化页」，必须由调用方用 is_ok 判定内容是否完整。
    重试仍失败则换 StealthyFetcher 兜底一次。
    """
    last = None
    for attempt in range(retries):
        if attempt:
            time.sleep(random.uniform(2, 4) * attempt)
        time.sleep(REQUEST_DELAY)
        try:
            page = fetch(url)
            if is_ok(page):
                return page
            last = page
        except Exception as e:
            print(f"  [重试 {attempt + 1}/{retries}] {url} 异常: {type(e).__name__}")
    # 兜底：换浏览器内核再试一次
    try:
        page = fetch(url, stealth=True)
        if is_ok(page):
            return page
        last = page
    except Exception as e:
        print(f"  [StealthyFetcher 失败] {url}: {type(e).__name__}")
    return last


def text_of(node, selector):
    """取节点内首个匹配元素的可见文本"""
    found = node.css(selector)
    if not found:
        return ''
    return ' '.join(t.strip() for t in found[0].css('::text').getall() if t.strip())


def base_name(name):
    """取括号前的基础名（形态后缀不影响归并）"""
    return re.split(r'[（(]', name or '')[0].strip()


def is_placeholder(name):
    """未公开 / 空白占位内容（如 wiki 的「???」），不写入任何数据表"""
    text = (name or '').strip()
    return not text or bool(re.fullmatch(r'[\s?？·—_-]+', text))


# ====================== 并发抓取器 ======================
class Spider:
    """通用并发抓取器：ThreadPoolExecutor + 退避重试 + 失败串行补抓"""

    def __init__(self, name, max_workers=MAX_WORKERS):
        self.name = name
        self.max_workers = max_workers

    def _one(self, item, is_ok, parse):
        key, url = item
        page = fetch_with_retry(url, is_ok)
        return key, (parse(page, key) if page is not None else None)

    def run(self, targets, is_ok, parse):
        """targets: [(key, url), ...] -> {key: parse结果}"""
        print(f"\n[{self.name}] 共 {len(targets)} 个页面，{self.max_workers} 并发开始抓取")
        results, failed = {}, []
        with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
            futures = {ex.submit(self._one, t, is_ok, parse): t[0] for t in targets}
            for i, fut in enumerate(as_completed(futures), 1):
                key, data = fut.result()
                if data:
                    results[key] = data
                else:
                    failed.append(key)
                if i % 50 == 0 or i == len(futures):
                    print(f"  进度 {i}/{len(futures)}，空/失败 {len(failed)}")

        if failed:
            print(f"[{self.name}] {len(failed)} 个疑似被限流，改为串行补抓...")
            url_of = {k: u for k, u in targets}
            still = []
            for i, key in enumerate(failed, 1):
                time.sleep(SERIAL_GAP)
                _, data = self._one((key, url_of[key]), is_ok, parse)
                if data:
                    results[key] = data
                    print(f"  [{i}/{len(failed)}] {key} 补抓成功")
                else:
                    still.append(key)
            if still:
                print(f"[警告] 串行补抓后仍为空 {len(still)} 个: {still}")
        print(f"[{self.name}] 完成：成功 {len(results)}/{len(targets)}")
        return results


# ====================== 精灵图鉴 ======================
def parse_spirit_list(page, _key=None):
    """解析精灵图鉴页"""
    cards = page.css('.npc-card')
    if not cards:
        return []
    rows, seen = [], set()
    for c in cards:
        link = c.css('a')
        name = (link[0].attrib.get('title', '') or '').strip() if link else ''
        if not name:
            name = text_of(c, '.npc-card-copy')
        number = (c.attrib.get('data-number', '') or '').strip()
        if not name or not number or is_placeholder(name):
            continue

        stage_txt = text_of(c, '.npc-stage')
        if stage_txt == '首领':
            stage = ''
        elif stage_txt:
            stage = stage_txt
        else:
            stage = STAGE_MAP.get((c.attrib.get('data-stage', '') or '').strip(), '')

        form_raw = (c.attrib.get('data-form', '') or '').strip()
        if 'lord' in form_raw:
            form = '首领形态'
        elif 'regional' in form_raw:
            form = '地区形态'
        else:
            form = '原始形态'
        main_form = '主形态' if 'main' in form_raw else ''

        pos_raw = (c.attrib.get('data-position', '') or '').strip()
        position = POSITION_MAP.get(pos_raw, '')
        if not position and stage == '三阶' and 'lord' not in form_raw:
            position = '最终'

        dtypes = (c.attrib.get('data-type', '') or '').split('|')
        shiny_raw = (c.attrib.get('data-shiny', '') or '').strip()
        season_raw = (c.attrib.get('data-season', '') or '').strip()

        row = {
            '序号': number,
            '精灵名称': name,
            '精灵阶数': stage,
            '进化位置': position,
            '主属性': dtypes[0].strip() if dtypes else '',
            '副属性': dtypes[1].strip() if len(dtypes) > 1 else '',
            '形态': form,
            '主形态/地区形态': main_form,
            '异色': SHINY_MAP.get(shiny_raw, shiny_raw),
            '赛季': SEASON_MAP.get(season_raw, season_raw),
        }
        key = (row['序号'], row['精灵名称'])
        if key in seen:
            continue
        seen.add(key)
        rows.append(row)
    return rows


def crawl_spirits():
    print(f"\n===== 抓取精灵图鉴 =====")
    page = fetch_with_retry(SPIRIT_URL, lambda p: bool(p.css('.npc-card')))
    if page is None:
        print("精灵图鉴页抓取失败")
        return []
    cards = page.css('.npc-card')
    rows = parse_spirit_list(page)
    ignored = len(cards) - len(rows)
    print(f"解析到 {len(rows)} 条精灵数据" +
          (f"（忽略 {ignored} 条占位/重复卡片）" if ignored else ""))
    return rows


# ====================== 技能图鉴 ======================
def crawl_skill_links():
    """返回 [(技能名, url), ...]"""
    page = fetch_with_retry(SKILL_URL, lambda p: bool(p.css('.nrc-skill-catalog-card')))
    if page is None:
        print("技能图鉴页抓取失败")
        return []
    cards = page.css('.nrc-skill-catalog-card')
    links = {}
    for c in cards:
        a = c.css('a[href]')
        if not a:
            continue
        href = (a[0].attrib.get('href', '') or '').strip()
        name = (a[0].attrib.get('title', '') or '').strip()
        if not name:
            name = text_of(c, '.nrc-skill-catalog-card-name')
        if not name or is_placeholder(name) or not href or 'redlink=1' in href:
            continue
        if not href.startswith('http'):
            href = 'https://wiki.biligame.com' + href
        links.setdefault(name, href)
    print(f"技能图鉴去重后共 {len(links)} 个技能")
    return [(n, u) for n, u in links.items()]


def parse_skill_detail(page, skill_name):
    """解析技能详情页：属性 / 类型 / 威力 / 耗能 / 描述 / 风味。

    兼容两套模板：
    - 旧模板：.nrc-skill-tags / .nrc-skill-type-chip / .nrc-skill-stat
    - 新模板：.sd-s-page（data-sd-element 为属性）/ .sd-s-stat（威力/耗能）
    """
    detail = {'技能名称': skill_name, '技能属性': '', '技能类型': '', '威力': '',
              '耗能': '', '技能描述': '', '风味文本': ''}

    sd = page.css('.sd-s-page')
    if sd:
        el = (sd[0].attrib.get('data-sd-element') or '').strip()
        detail['技能属性'] = el.replace('系', '').strip()
        for st in page.css('.sd-s-stat'):
            label = ''.join(t.strip() for t in st.css('.sd-s-stat-label::text').getall() if t.strip())
            value = ''.join(t.strip() for t in st.css('.sd-s-stat-value::text').getall() if t.strip())
            if '威力' in label:
                detail['威力'] = value
            elif '耗能' in label:
                detail['耗能'] = value
        detail['技能描述'] = text_of(page, '.sd-s-description-text')
        return detail

    tags = page.css('.nrc-skill-tags')
    scope = tags[0] if tags else page

    chip = scope.css('.nrc-skill-type-chip span')
    if chip:
        detail['技能属性'] = ' '.join(
            t.strip() for t in chip[-1].css('::text').getall() if t.strip()).replace('系', '')
    cate = scope.css('.nrc-skill-category-chip span')
    if cate:
        detail['技能类型'] = ' '.join(
            t.strip() for t in cate[-1].css('::text').getall() if t.strip())

    detail['技能描述'] = text_of(page, '.nrc-skill-effect-text')
    detail['风味文本'] = text_of(page, '.nrc-skill-flavor')

    for st in page.css('.nrc-skill-stat'):
        label, value = st.css('.nrc-skill-stat-label'), st.css('.nrc-skill-stat-value')
        if not (label and value):
            continue
        ltxt = ' '.join(t.strip() for t in label[0].css('::text').getall() if t.strip())
        vtxt = ' '.join(t.strip() for t in value[0].css('::text').getall() if t.strip())
        if '威力' in ltxt:
            detail['威力'] = vtxt
        elif '耗能' in ltxt:
            detail['耗能'] = vtxt
    return detail


def skill_ok(page):
    """限流的「简化页」没有属性标签；新模板用 .sd-s-page 判定"""
    return bool(page.css('.nrc-skill-tags') or page.css('.sd-s-page'))


def crawl_skills(old_skills=None):
    """技能档案：只取技能自身的属性。

    「谁能学这个技能」由精灵页产出（skill_sources.csv），覆盖面更全
    （技能页 30733 对 ⊂ 精灵页 30750 对），所以这里不解析来源。

    增量模式（传入 old_skills）：仅抓取「新增技能」与「属性为空的旧技能」，
    其余直接复用旧数据，避免全量重抓 ~600 个技能页触发 wiki 限流。
    """
    print(f"\n===== 抓取技能图鉴 =====")
    links = crawl_skill_links()
    if not links:
        return []
    url_of = {n: BASE + quote(n) for n, _ in links}
    link_names = [n for n, _ in links]
    old_map = {r['技能名称']: r for r in (old_skills or [])}
    if old_skills:
        need = [n for n in link_names
                if n not in old_map or not old_map[n].get('技能属性')]
        base = {n: old_map[n] for n in link_names
                if n in old_map and old_map[n].get('技能属性')}
    else:
        need = link_names
        base = {}
    print(f"[增量] 技能需抓取 {len(need)}（新增+属性为空），复用 {len(base)}")
    got = Spider('技能详情').run([(n, url_of[n]) for n in need], skill_ok, parse_skill_detail)
    # 限流增量重试：批量并发后，限流页（属性为空）通常已解除，单独轻量补抓
    missing = [n for n in need if not (got.get(n) or {}).get('技能属性')]
    if missing:
        print(f"[技能] {len(missing)} 个属性为空，限流冷却后单独补抓...")
        time.sleep(SERIAL_GAP)
        for i, n in enumerate(missing, 1):
            page = fetch_with_retry(url_of[n], skill_ok, retries=4)
            if page is not None:
                got[n] = parse_skill_detail(page, n)
            if i % 20 == 0 or i == len(missing):
                print(f"  补抓 {i}/{len(missing)}")
    rows = list(base.values())
    # 本次没抓到（限流）的技能沿用旧记录，避免被误判成「wiki 已删除」
    rows += [d for n in need for d in [got.get(n) or old_map.get(n)] if d]
    print(f"技能档案共 {len(rows)} 条")
    return rows


# ====================== 精灵详情页技能表（获取条件）======================
def parse_quests(page, pet_name):
    """解析精灵页的「图鉴课题」面板 —— 获取条件的唯一正确来源。

    .roco-quest-content > .roco-quest-row 每行一个课题：
        .roco-quest-text   条件文本「使用2次」+ 嵌套 <a title="水炮">
        .roco-quest-rew    奖励（技能石 / 魔法经验 / 课题研究点…）

    只保留「使用N次XX」这类课题；技能名取 .roco-quest-text 里的 a[title]。

    注意：不能用 .roco-sk-dcard 的 .roco-sk-source-task，那是挂在整个技能上的，
    归属精灵见 .roco-sk-source-pet，通常是别的精灵（水灵的「突袭」→ 流浪鼠）。
    """
    out = []
    for row in page.css('.roco-quest-row'):
        txt = ''.join(t.strip() for t in row.css('.roco-quest-text::text').getall())
        m = re.match(r'使用(\d+)次$', txt)
        if not m:
            continue
        skill = ''
        a = row.css('.roco-quest-text a[title]')
        if a:
            skill = (a[0].attrib.get('title') or '').strip()
        if not skill:
            generic = {'课题研究点', '魔法经验', '灵魂环印', '精灵果实'}
            for alt in row.css('.roco-quest-rew img::attr(alt)').getall():
                if alt and alt not in generic:
                    skill = alt
                    break
        if not skill:
            continue
        out.append({'精灵名称': pet_name, '技能名称': skill,
                    '获取条件': f'使用{m.group(1)}次{skill}'})
    return out


def parse_pet_skills(page, pet_name):
    """解析精灵详情页技能表（技能 + 来源），并把图鉴课题的获取条件挂上去。

    兼容两套模板：
    - 旧模板：.roco-sk-card（.roco-sk-name / .roco-sk-lv / data-source）
    - 新模板：.sd-p-skill-card（data-name / data-level / data-source 等属性）
    """
    quests = {q['技能名称']: q['获取条件'] for q in parse_quests(page, pet_name)}
    rows = []
    # 新模板：技能卡以 data-* 属性携带全部字段
    for c in page.css('.sd-p-skill-card'):
        name = (c.attrib.get('data-name') or '').strip()
        if not name:
            continue
        source = (c.attrib.get('data-source') or '').strip()
        rows.append({
            '精灵名称': pet_name,
            '技能名称': name,
            '来源类型': SOURCE_MAP.get(source, source),
            '等级': (c.attrib.get('data-level') or '').strip(),
            '获取条件': quests.get(name, ''),
        })
    if rows:
        return rows
    # 旧模板
    cards = page.css('.roco-sk-card')
    if not cards:
        return []
    for c in cards:
        name = text_of(c, '.roco-sk-name')
        if not name:
            continue
        source = (c.attrib.get('data-source', '') or '').strip()
        rows.append({
            '精灵名称': pet_name,
            '技能名称': name,
            '来源类型': SOURCE_MAP.get(source, source),
            '等级': text_of(c, '.roco-sk-lv'),
            '获取条件': quests.get(name, ''),
        })
    return rows


def parse_evo_chain(page, pet_name):
    """解析精灵页「进化链」面板，返回该页家族的所有进化步骤。

    兼容两套模板：
    - 新模板：.sd-p-evo-group（data-sd-p-evo-branch 标分支）> .sd-p-evo-chain，
      节点 .sd-p-evo-node 与连接 .sd-p-evo-connector 交替；
      名字在 .sd-p-evo-name-main a，阶位在 .sd-p-evo-name-sub，
      条件在 .sd-p-evo-condition。单分支也有 group，隐藏分支仍在 HTML 中。
    - 旧模板：.roco-evo-wrap > .roco-evo-timeline（data-branch 标分支）>
      .roco-evo-node / .roco-evo-conn > .roco-evo-conn-pill。
    一个家族的完整链在任意成员页都能拿到，按家族去重由 build_evo_chains 处理。
    """
    steps = []
    groups = page.css('.sd-p-evo-group') or page.css('.sd-p-evo-chain')
    for grp in groups:
        branch = (grp.attrib.get('data-sd-p-evo-branch') or '').strip() or '1'
        nodes = grp.css('.sd-p-evo-node')
        conns = grp.css('.sd-p-evo-connector')
        for i in range(min(len(nodes) - 1, len(conns))):
            src = _sd_evo_name(nodes[i])
            tgt = _sd_evo_name(nodes[i + 1])
            if not src or not tgt:
                continue
            steps.append({'分支': branch, '阶序': _sd_evo_sub(nodes[i]),
                          '源精灵': src, '目标精灵': tgt,
                          '进化条件': _sd_evo_cond(conns[i])})
    if steps:
        return steps
    # 旧模板
    for tl in page.css('.roco-evo-timeline'):
        branch = (tl.attrib.get('data-branch') or '').strip() or '1'
        nodes = tl.css('.roco-evo-node')
        conns = tl.css('.roco-evo-conn')
        for i in range(min(len(nodes) - 1, len(conns))):
            src = _evo_name(nodes[i])
            tgt = _evo_name(nodes[i + 1])
            if not src or not tgt:
                continue
            cond = ''.join(t.strip() for t in conns[i].css('.roco-evo-conn-pill *::text').getall())
            steps.append({'分支': branch, '阶序': _evo_sub(nodes[i]),
                          '源精灵': src, '目标精灵': tgt, '进化条件': cond})
    return steps


def _sd_evo_name(node):
    a = node.css('.sd-p-evo-name-main a')
    if a:
        return (a[0].attrib.get('title') or '').strip() or ''.join(
            t.strip() for t in a[0].css('::text').getall())
    img = node.css('.sd-p-evo-avatar img')
    return (img[0].attrib.get('alt') or '').strip() if img else ''


def _sd_evo_sub(node):
    return ''.join(t.strip() for t in node.css('.sd-p-evo-name-sub::text').getall())


def _sd_evo_cond(conn):
    cond = conn.css('.sd-p-evo-condition')
    return str((cond[0] if cond else conn).get_all_text(strip=True))


def _evo_name(node):
    a = node.css('.roco-evo-name-main a')
    if a:
        return (a[0].attrib.get('title') or '').strip() or ''.join(
            t.strip() for t in a[0].css('::text').getall())
    return ''.join(t.strip() for t in node.css('.roco-evo-name-main::text').getall())


def _evo_sub(node):
    return ''.join(t.strip() for t in node.css('.roco-evo-name-sub::text').getall())


# 旧模板把来源文本写进「等级」列（技能石 / 血脉·X / 传说），新模板 data-level 是纯数字。
# 统一约定：等级列只放数值——自学=真实等级，技能石/传说=999，血脉=15。
LEVEL_TEXT_MAP = {'技能石': '999', '传说': '999'}


def normalize_level(level):
    """把等级列归一化成纯数字；未知取值原样保留，不硬猜"""
    s = str(level or '').strip()
    if not s or s.isdigit():
        return s
    m = re.fullmatch(r'LV\s*(\d+)', s, re.IGNORECASE)
    if m:
        return m.group(1)
    if s.startswith('血脉·'):
        return '15'
    return LEVEL_TEXT_MAP.get(s, s)


def normalize_source_row(row):
    """统一来源表一行的「来源类型」「等级」写法"""
    row['来源类型'] = SOURCE_MAP.get(row.get('来源类型', ''), row.get('来源类型', ''))
    row['等级'] = normalize_level(row.get('等级', ''))
    return row


SKILL_CARD_ID_RE = re.compile(r'^[a-z]+-(.+)-\d+$')


def parse_skill_catalog(page, pet_name):
    """解析精灵详情页技能卡的完整信息，用于生成「精灵技能总表」。

    新模板 .sd-p-skill-card 以 data-* 携带全部字段：
        data-sd-p-skill = 来源-技能ID-顺序（如 blood-skill_000274-27）
        data-source / data-type / data-category / data-level / data-power / data-energy
        .sd-p-skill-description 效果文本
        .sd-p-learning 学习标签；血脉形如「普通血脉 · Lv.15」，前半段即血脉属性
    旧模板没有这些字段，本表只收录新模板页面（当前 wiki 已全量换版）。
    """
    rows = []
    for c in page.css('.sd-p-skill-card'):
        a = c.attrib
        name = (a.get('data-name') or '').strip()
        if not name or is_placeholder(name):
            continue
        source = SOURCE_MAP.get(a.get('data-source', ''), a.get('data-source', ''))
        m = SKILL_CARD_ID_RE.match(a.get('data-sd-p-skill') or '')
        learning = text_of(c, '.sd-p-learning')
        power = (a.get('data-power') or '').strip()
        rows.append({
            '精灵名称': pet_name,
            '技能ID': m.group(1) if m else '',
            '技能名称': name,
            '来源类型': source,
            '学习等级': normalize_level(a.get('data-level', '')),
            '血脉属性': learning.split('·')[0].strip() if source == '血脉' else '',
            '技能系别': (a.get('data-type') or '').strip(),
            '技能类别': (a.get('data-category') or '').strip(),
            '威力': '' if power in ('', '-1') else power,
            '能耗': (a.get('data-energy') or '').strip(),
            '技能描述': text_of(c, '.sd-p-skill-description'),
            '顺序': (a.get('data-order') or '').strip(),
        })
    return rows


def parse_pet_full(page, pet_name):
    """精灵详情页：同时解析技能来源与进化链（复用同一次抓取）"""
    skills = [normalize_source_row(r) for r in parse_pet_skills(page, pet_name)
              if not is_placeholder(r['精灵名称']) and not is_placeholder(r['技能名称'])]
    return {'skills': skills, 'evo': parse_evo_chain(page, pet_name),
            'catalog': parse_skill_catalog(page, pet_name)}


def crawl_sources(names):
    """精灵 × 技能 来源表 + 进化链（自学/技能石/血脉/传说，含获取条件）。

    复用精灵详情页的同一次抓取，同时取出技能来源、进化链与技能总表，避免二次请求。
    返回 (skills_rows, evo_by_spirit, catalog_rows)。
    """
    print(f"\n===== 抓取精灵详情页（技能表 + 进化链）=====")
    targets = [(n, BASE + quote(n)) for n in names]
    got = Spider('精灵技能表').run(
        targets, lambda p: bool(p.css('.roco-sk-card') or p.css('.sd-p-skill-card')), parse_pet_full)
    # 增量重试：技能与进化链均为空的精灵页，多半是被限流，冷却后单独补抓
    url_of = {n: u for n, u in targets}
    missing = [n for n in names
               if not ((got.get(n) or {}).get('skills') or (got.get(n) or {}).get('evo'))]
    if missing:
        print(f"[精灵] {len(missing)} 个技能/进化链为空，限流冷却后单独补抓...")
        time.sleep(SERIAL_GAP)
        for i, n in enumerate(missing, 1):
            page = fetch_with_retry(url_of[n],
                                    lambda p: bool(p.css('.roco-sk-card') or p.css('.sd-p-skill-card')),
                                    retries=4)
            if page is not None:
                got[n] = parse_pet_full(page, n)
            if i % 20 == 0 or i == len(missing):
                print(f"  补抓 {i}/{len(missing)}")
    skills, evo_by, catalog = [], {}, []
    for n in names:
        d = got.get(n)
        if d:
            skills.extend(d['skills'])
            evo_by[n] = d['evo']
            catalog.extend(d.get('catalog', []))
    print(f"技能来源共 {len(skills)} 条，技能总表 {len(catalog)} 条"
          f"（{len(got)}/{len(names)} 个精灵）；进化链待合并")
    return skills, evo_by, catalog


# ====================== 读写 ======================
def read_csv(path, fields=None):
    try:
        with open(path, encoding='utf-8-sig') as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        return []
    if fields:
        rows = [{k: (r.get(k) or '').strip() for k in fields} for r in rows]
    return rows


def save_csv(data, path, fields):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(data)
    print(f"已保存 {path}（{len(data)} 条）")


def save_xlsx(data, path, fields, sheet='精灵图鉴'):
    if not data:
        return
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet(sheet)
    ws.title = sheet
    ws.append(fields)
    for r in data:
        ws.append([r.get(k, '') for k in fields])
    last_col = get_column_letter(len(fields))
    ws.auto_filter.ref = f"A1:{last_col}{len(data) + 1}"
    ws.freeze_panes = 'A2'
    for idx, field in enumerate(fields, 1):
        width = len(field)
        for r in data:
            cell = str(r.get(field, ''))
            width = max(width, sum(2 if ord(ch) > 127 else 1 for ch in cell))
        ws.column_dimensions[get_column_letter(idx)].width = min(width + 4, 30)
    wb.save(path)
    print(f"已保存 {path}（{len(data)} 条）")


# ====================== 合并（最近两次快照 + 覆盖更新 + 缺失淘汰）======================
def snapshot_files():
    """按文件名（含时间戳）排序的快照列表"""
    if not os.path.isdir(SNAP_DIR):
        return []
    return sorted(f for f in os.listdir(SNAP_DIR)
                  if f.startswith('snapshot_') and f.endswith('.json'))


def load_prev_snapshot():
    """上一份快照 -> {表名: 键集合}；不存在返回 None（首次运行）"""
    files = snapshot_files()
    if not files:
        return None
    try:
        with open(os.path.join(SNAP_DIR, files[-1]), encoding='utf-8') as f:
            data = json.load(f)
        return {t: set(keys) for t, keys in data.get('tables', {}).items()}
    except (OSError, ValueError) as e:
        print(f"[警告] 快照读取失败，本次按首次运行处理：{e}")
        return None


def save_snapshot(keys_by_table):
    """保存本次抓取到的键集合，只保留最近 SNAP_KEEP 份"""
    os.makedirs(SNAP_DIR, exist_ok=True)
    # 用纳秒时间戳命名：字典序即时间序，且同一秒内多次运行也不会互相覆盖
    path = os.path.join(SNAP_DIR, f'snapshot_{time.time_ns():020d}.json')
    payload = {'time': time.strftime('%Y-%m-%d %H:%M:%S'),
               'tables': {t: sorted(keys) for t, keys in keys_by_table.items()}}
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False)
    files = snapshot_files()
    for old in files[:-SNAP_KEEP]:
        try:
            os.remove(os.path.join(SNAP_DIR, old))
        except OSError:
            pass
    print(f"快照已保存: {path}（保留最近 {SNAP_KEEP} 份）")


def overlay_row(old, new):
    """本次值非空则覆盖；本次为空时保留旧值，防止限流简化页把已有数据洗成空"""
    out = dict(old)
    for k, v in new.items():
        if str(v or '').strip():
            out[k] = v
    return out


def merge_table(table, old, new, prev_keys=None):
    """以本次抓取为准合并一张表。

    - 本次抓到：逐字段用本次值覆盖（空值保留旧值，即字段完整性保护）
    - 本次没抓到、上一份快照里有：疑似限流，暂留观察
    - 本次与上一份快照都没有：连续两次缺失，从主表移除
    - 首次运行（无上一份快照）：只更新 / 追加，不删除任何记录

    返回 (合并后的行, 本次键集合)。
    """
    key = TABLE_KEYS[table]
    old_map = {key(r): r for r in old}
    new_map = {key(r): r for r in new}
    prev = old_map.keys() if prev_keys is None else prev_keys
    merged, added, updated, pending, dropped = {}, [], 0, 0, []
    for k, row in old_map.items():
        if k in new_map:
            merged[k] = overlay_row(row, new_map[k])
            if merged[k] != row:
                updated += 1
        elif k in prev:
            merged[k] = row
            pending += 1
        else:
            dropped.append(k)
    for k, row in new_map.items():
        if k not in merged:
            merged[k] = row
            added.append(row)
    print(f"[{table}] 更新 {updated} / 新增 {len(added)} / 本次缺失待观察 {pending} / "
          f"连续两次缺失淘汰 {len(dropped)}")
    if dropped:
        print(f"    淘汰示例: {dropped[:5]}")
    return list(merged.values()), set(new_map)


def sort_spirits(rows):
    return sorted(rows, key=lambda r: (
        int(r['序号']) if str(r['序号']).isdigit() else 99999, r['精灵名称']))


# ====================== 生成合并表 ======================
def quest_index(sources):
    """基础精灵名 -> [(技能名, 获取条件)]，顺序同技能表

    来源只有精灵页自己的「图鉴课题」：每只精灵的课题奖励若干技能石，
    条件形如「使用2次水炮」。技能石在别的精灵身上也能用，但条件归属这只精灵。
    """
    idx, seen = defaultdict(list), defaultdict(set)
    for r in sources:
        pet, skill = r.get('精灵名称', ''), r.get('技能名称', '')
        cond = (r.get('获取条件') or '').strip()
        if not pet or not skill or not re.search(r'使用\d+次', cond):
            continue
        base = base_name(pet)
        if skill in seen[base]:
            continue
        seen[base].add(skill)
        idx[base].append((skill, cond))
    return idx


def _build_combined_from_stones(spirits, stones, max_skills=MAX_SKILLS):
    """从「图鉴课题」聚合页（skill_stones）构建合并表：按序号（物种）一行，
    技能石归属与官方页完全一致——分支形态不再各自成行，无重复。"""
    attr_by_no = {}
    for s in spirits:
        no = s.get('序号', '')
        if no and no not in attr_by_no:
            attr_by_no[no] = s.get('主属性', '')
    by_no, name_by_no = defaultdict(list), {}
    for r in stones:
        no = r.get('序号', '')
        if not str(no).isdigit():
            continue
        by_no[no].append(r)
        name_by_no.setdefault(no, r.get('精灵名称', ''))
    rows = []
    for no, lst in by_no.items():
        picked = []
        for r in lst:
            cond = (r.get('条件') or '').strip()
            m = re.match(r'^使用\d+次(.+)$', cond)
            skill = m.group(1).strip() if m else ''
            picked.append((skill, cond))
        picked = picked[:max_skills]
        rows.append({
            '序号': no,
            '精灵名称': name_by_no.get(no, ''),
            '属性': attr_by_no.get(no, ''),
            **{f'技能名称{i + 1}': picked[i][0] if i < len(picked) else ''
               for i in range(max_skills)},
            '条件': '; '.join(c for _, c in picked),
        })
    rows.sort(key=lambda x: (int(x['序号']) if str(x['序号']).isdigit() else 99999, x['精灵名称']))
    return rows


def build_combined(spirits, sources, stones=None, max_skills=MAX_SKILLS):
    """合并表：精灵 → 最多 N 个技能石 + 获取条件（视图）。

    优先以官方「图鉴课题」聚合页（skill_stones）为准：按序号（物种）一行，
    技能石归属与官方页完全一致，避免分支形态（如 191 棋契陛下 / 棋祈督）
    被拆成重复行、或技能被错归到其他序号。
    无聚合页数据时回退到逐精灵来源表（按基础名归组）。
    """
    if stones is None and os.path.exists(F_STONES):
        stones = read_csv(F_STONES, STONE_FIELDS)
    if stones:
        return _build_combined_from_stones(spirits, stones, max_skills)
    # 回退路径：基于来源表按基础名归组（与旧逻辑一致）
    by_base = defaultdict(list)
    for s in spirits:
        by_base[base_name(s.get('精灵名称', ''))].append(s)
    idx = quest_index(sources)
    rows = []
    for base, spirit_list in by_base.items():
        picked = idx.get(base, [])[:max_skills]
        if not picked:
            continue
        info = min(spirit_list, key=lambda r: (0 if r.get('精灵阶数') else 1,
                                               len(r.get('精灵名称', '')),
                                               r.get('精灵名称', '')))
        rows.append((info.get('序号', ''), info.get('主属性', ''),
                     info.get('精灵阶数', ''), base, tuple(picked)))
    groups = defaultdict(list)
    for no, attr, stage, base, picked in rows:
        groups[(no, picked)].append((0 if stage else 1, len(base), base, attr))
    combined = []
    for (no, picked), cands in groups.items():
        _, _, name, attr = min(cands)
        combined.append({
            '序号': no,
            '精灵名称': name,
            '属性': attr,
            **{f'技能名称{i + 1}': picked[i][0] if i < len(picked) else ''
               for i in range(max_skills)},
            '条件': '; '.join(c for _, c in picked),
        })
    combined.sort(key=lambda x: (int(x['序号']) if str(x['序号']).isdigit() else 99999, x['精灵名称']))
    return combined


# 普通进化（非特殊）= 纯等级进化：仅 LVx，或 LVx 进化（不带突破/天气/性别等附加条件）
STD_EVO = re.compile(r'^LV\d+\s*(进化)?\s*$', re.IGNORECASE)


def _is_special_family(steps):
    """家族是否特殊：存在分支（data-branch>1）或任一进化步骤不是纯等级进化。"""
    branches = {s['分支'] for s in steps}
    any_special = any(not STD_EVO.match(s['进化条件']) for s in steps)
    return len(branches) > 1 or any_special


def build_evo_chains(evo_by):
    """按家族去重，生成进化链步骤表，并标记特殊链。

    特殊 = 存在分支（data-branch>1）或任一进化步骤的条件不是纯等级进化
    （突破、天气/时间、性别、技能使用、道具、特定条件等都算特殊）。
    只有「LVx」或「LVx 进化」这类纯等级进化才算普通。
    """
    families = {}
    for spirit, steps in evo_by.items():
        if not steps:
            continue
        srcs = {s['源精灵'] for s in steps}
        tgts = {s['目标精灵'] for s in steps}
        roots = srcs - tgts
        root = (sorted(roots)[0] if roots else (sorted(srcs)[0] if srcs else spirit))
        if root in families:        # 同一家族只在首次出现的精灵页取一次
            continue
        special = '是' if _is_special_family(steps) else '否'
        families[root] = (steps, special)

    rows = []
    for root in sorted(families):
        steps, special = families[root]
        for s in steps:
            rows.append({'家族': root, '分支': s['分支'], '阶序': s['阶序'],
                         '源精灵': s['源精灵'], '目标精灵': s['目标精灵'],
                         '进化条件': s['进化条件'], '是否特殊': special})
    return rows


def flag_evo_special(rows):
    """按家族重算「是否特殊」标记（不重新联网，供 rebuild 复用）。"""
    from collections import defaultdict
    groups = defaultdict(list)
    for r in rows:
        groups[r['家族']].append(r)
    for items in groups.values():
        special = '是' if _is_special_family(items) else '否'
        for r in items:
            r['是否特殊'] = special
    return rows


def build_spirit_no_map(spirits):
    """基础名 -> 序号（忽略「（…）」形态后缀），供进化链按序号比对。"""
    m = {}
    for r in spirits:
        base = base_name(r.get('精灵名称', ''))
        if base and base not in m:
            m[base] = r.get('序号', '')
    return m


def enrich_evo_no(evo_rows, spirits):
    """为进化链每行补上源/目标精灵的 序号（按图鉴主 key 比对，规避形态后缀差异）。"""
    m = build_spirit_no_map(spirits)
    for r in evo_rows:
        r['源精灵序号'] = m.get(base_name(r.get('源精灵', '')), '')
        r['目标精灵序号'] = m.get(base_name(r.get('目标精灵', '')), '')
    return evo_rows


def sort_evo(rows):
    """按源精灵序号升序（0→x）排列，家族内步骤自然保持顺序。"""
    def key(r):
        no = r.get('源精灵序号', '')
        return (int(no) if str(no).isdigit() else 10 ** 9,
                r.get('家族', ''), str(r.get('分支', '')))
    return sorted(rows, key=key)


# ====================== 特性图鉴 ======================
def parse_feature_catalog(page):
    """解析特性图鉴页，返回 (特性表行, 精灵名 -> (特性名, 特性介绍)) 映射。

    .nrc-feature-catalog-card 上有：
        .nrc-skill-catalog-card-name     特性名
        .nrc-feature-card-description   特性介绍
        .nrc-feature-card-owner-avatar  拥有该特性的精灵头像（每个一）
        data-feature-primary-type / secondary-type  特性分类
        data-skill-catalog-text = 「特性名 介绍 特性编号 拥有者… 属性…」
    拥有者名字从 text 尾部解析：按介绍切分后去掉 6 位编号与末尾属性词；
    解析出的个数必须与头像数一致，否则不猜、置空以免错配。
    """
    rows, owners_map = [], {}
    for c in page.css('.nrc-feature-catalog-card'):
        name = text_of(c, '.nrc-skill-catalog-card-name')
        desc = text_of(c, '.nrc-feature-card-description')
        if not name or is_placeholder(name):
            continue
        avatars = c.css('.nrc-feature-card-owner-avatar')
        text = c.attrib.get('data-skill-catalog-text') or ''
        tail = text.split(desc, 1)[1].strip() if desc and desc in text else ''
        m = re.match(r'^(\d{6})\s+(.*)$', tail)
        number = m.group(1) if m else (c.attrib.get('data-skill-catalog-id') or '')
        tokens = m.group(2).split() if m else []
        attrs = set()
        for k in ('data-feature-primary-type', 'data-feature-secondary-type',
                  'data-skill-catalog-type'):
            for a in (c.attrib.get(k) or '').split('|'):
                if a.strip():
                    attrs.add(a.strip())
        while tokens and tokens[-1] in attrs:
            tokens.pop()
        owners = [t for t in tokens if not is_placeholder(t)]
        if len(owners) != len(avatars):
            owners = []
        rows.append({
            '特性编号': str(number),
            '特性名称': name,
            '特性介绍': desc,
            '主属性': (c.attrib.get('data-feature-primary-type') or '').strip(),
            '副属性': (c.attrib.get('data-feature-secondary-type') or '').strip(),
            '拥有精灵数': str(len(avatars)),
            '拥有精灵': '; '.join(owners),
        })
        for o in owners:
            owners_map[o] = (name, desc)
    return rows, owners_map


def owners_map_from_features(features):
    """从特性表的「拥有精灵」列还原 精灵名 -> (特性名, 特性介绍) 映射"""
    m = {}
    for r in features or []:
        for name in (r.get('拥有精灵') or '').split(';'):
            name = name.strip()
            if name:
                m[name] = (r.get('特性名称', ''), r.get('特性介绍', ''))
    return m


def mark_traits(spirits, owners_map):
    """给精灵档案补「特性」「特性介绍」两列：先全名精确，再基础名（去形态后缀）兜底"""
    for r in spirits:
        name = r.get('精灵名称', '')
        hit = owners_map.get(name) or owners_map.get(base_name(name))
        r['特性'] = hit[0] if hit else ''
        r['特性介绍'] = hit[1] if hit else ''
    return spirits


def crawl_features():
    """抓取特性图鉴页（一次静态请求即可覆盖全部特性与拥有精灵）"""
    print("\n===== 抓取特性图鉴 =====")
    page = fetch_with_retry(FEATURE_URL, lambda p: bool(p.css('.nrc-feature-catalog-card')))
    if page is None:
        print("特性图鉴页抓取失败")
        return [], {}
    rows, owners_map = parse_feature_catalog(page)
    print(f"解析到 {len(rows)} 个特性，覆盖 {len(owners_map)} 个精灵名")
    return rows, owners_map


# ====================== 体检 ======================
def mark_quest(spirits, sources, stones=None):
    """给精灵档案补「技能石课题」列（有/无）。有聚合页数据时按序号（物种）判定，否则回退基础名。"""
    if stones is None and os.path.exists(F_STONES):
        stones = read_csv(F_STONES, STONE_FIELDS)
    if stones:
        has_no = {r.get('序号') for r in stones if str(r.get('序号', '')).isdigit()}
        for r in spirits:
            r['技能石课题'] = '有' if r.get('序号') in has_no else '无'
        return spirits
    has = {base_name(r['精灵名称']) for r in sources
           if re.search(r'使用\d+次', r.get('获取条件') or '')}
    for r in spirits:
        r['技能石课题'] = '有' if base_name(r.get('精灵名称', '')) in has else '无'
    return spirits


def check_data(spirits, skills, sources, combined, evo, stones=None, label='',
               features=None, catalog=None):
    """检查数据完整性，返回 (统计 dict, 问题 list)"""
    stats, issues = {}, []

    # ---- 精灵图鉴 ----
    stats['精灵数'] = len(spirits)
    no_no = [r for r in spirits if not r.get('序号')]
    no_attr = [r for r in spirits if not r.get('主属性')]
    no_stage = [r for r in spirits
                if not r.get('精灵阶数') and '首领' not in (r.get('形态') or '')]
    if no_no:
        issues.append(f'精灵图鉴：{len(no_no)} 条缺序号')
    if no_attr:
        issues.append(f'精灵图鉴：{len(no_attr)} 条缺主属性，如 {[r["精灵名称"] for r in no_attr[:5]]}')
    if no_stage:
        issues.append(f'精灵图鉴：{len(no_stage)} 条缺阶数，如 {[r["精灵名称"] for r in no_stage[:5]]}')

    nums = sorted({int(r['序号']) for r in spirits if str(r['序号']).isdigit()})
    if nums:
        missing = [n for n in range(nums[0], nums[-1] + 1) if n not in set(nums)]
        stats['序号范围'] = f'{nums[0]}-{nums[-1]}'
        if missing:
            issues.append(f'精灵图鉴：序号断号 {len(missing)} 个，如 {missing[:20]}')

    # ---- 技能档案 ----
    stats['技能数'] = len(skills)
    empty_attr = {r['技能名称'] for r in skills if r.get('技能名称') and not r.get('技能属性')}
    if empty_attr:
        issues.append(f'技能档案：{len(empty_attr)} 个技能缺属性（疑似限流空页），如 {list(empty_attr)[:5]}')

    # ---- 来源表 ----
    stats['来源记录'] = len(sources)
    stats['覆盖精灵'] = len({r['精灵名称'] for r in sources})
    lack = [r['精灵名称'] for r in spirits
            if r['精灵名称'] not in {s['精灵名称'] for s in sources}]
    if lack:
        issues.append(f'来源表：{len(lack)} 个精灵没有来源记录，如 {lack[:8]}')

    quest_n = sum(1 for r in sources if re.search(r'使用\d+次', r.get('获取条件') or ''))
    stats['有获取条件'] = quest_n
    stats['有课题的形态行'] = sum(1 for r in spirits if r.get('技能石课题') == '有')
    if not quest_n:
        issues.append('来源表：一条获取条件都没解析到，检查 .roco-quest-row 结构')

    # ---- 合并表 ----
    stats['合并表行数'] = len(combined)
    empty_cond = [r for r in combined if not re.search(r'使用\d+次', r.get('条件') or '')]
    if empty_cond:
        issues.append(f'合并表：{len(empty_cond)} 行条件不完整，如 {[r["精灵名称"] for r in empty_cond[:5]]}')
    # 合并表按「序号 + 条件」去重（形态、同链成员会合并），所以比序号数
    want = len({r['序号'] for r in spirits if r.get('技能石课题') == '有'})
    stats['有课题的序号'] = want
    if combined and len(combined) < want:
        issues.append(f'合并表：有课题的序号 {want} 个，只进了 {len(combined)} 行')

    # ---- 技能石课题（与图鉴课题聚合页比对）----
    skill_count = sum(1 for r in combined for i in (1, 2, 3) if r.get(f'技能名称{i}'))
    stats['技能石课题数'] = skill_count
    if stones is None and os.path.exists(F_STONES):
        stones = read_csv(F_STONES, STONE_FIELDS)
    if stones:
        agg_total = len(stones)
        agg_species = len({r['序号'] for r in stones if str(r['序号']).isdigit()})
        stats['聚合页课题数'] = agg_total
        stats['聚合页物种数'] = agg_species
        if skill_count == agg_total and want == agg_species:
            stats['与聚合页比对'] = '一致'
        else:
            stats['与聚合页比对'] = '不一致'
            issues.append(f'技能石：合并表 {skill_count} 项 / {want} 物种，'
                          f'与聚合页 {agg_total} 项 / {agg_species} 物种 不一致')
    else:
        stats['与聚合页比对'] = '跳过（无聚合页数据）'

    # ---- 进化链 ----
    stats['进化家族'] = len(evo)
    if evo:
        stats['特殊家族'] = sum(1 for r in evo if r.get('是否特殊') == '是')
        no_map = build_spirit_no_map(spirits)      # 基础名 -> 序号（按序列比对）
        missing = sorted({(r['源精灵'], r['源精灵序号']) for r in evo
                          if base_name(r['源精灵']) not in no_map}
                         | {(r['目标精灵'], r['目标精灵序号']) for r in evo
                            if base_name(r['目标精灵']) not in no_map})
        if missing:
            issues.append(f'进化链：{len(missing)} 个名称在精灵表（按序号）找不到，'
                          f'如 {[m[0] for m in missing[:8]]}')

    # ---- 特性 ----
    if features:
        stats['特性数'] = len(features)
        stats['特性覆盖精灵名'] = len(owners_map_from_features(features))
        stats['有特性的精灵'] = sum(1 for r in spirits if r.get('特性'))
        no_trait = [r['精灵名称'] for r in spirits if not r.get('特性')]
        if no_trait:
            issues.append(f'特性：{len(no_trait)} 个精灵没匹配到特性，如 {no_trait[:5]}')

    # ---- 技能总表 ----
    if catalog:
        stats['技能总表行数'] = len(catalog)
        stats['技能总表精灵'] = len({r.get('精灵名称', '') for r in catalog})
        stats['带血脉属性'] = sum(1 for r in catalog if r.get('血脉属性'))
        # wiki 已下架的精灵（页面 404）本来就没有技能卡，只在成片缺失时报警
        no_cat = {r['精灵名称'] for r in spirits} - {r.get('精灵名称', '') for r in catalog}
        stats['技能总表缺失精灵'] = len(no_cat)
        if len(no_cat) > 5:
            issues.append(f'技能总表：{len(no_cat)} 个精灵没有技能卡记录（疑似限流），'
                          f'如 {sorted(no_cat)[:5]}')

    # ---- 输出 ----
    print(f"\n===== 体检{'：' + label if label else ''} =====")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    if issues:
        print(f"  发现 {len(issues)} 个问题：")
        for i, s in enumerate(issues, 1):
            print(f"    {i}. {s}")
    else:
        print("  未发现问题")
    return stats, issues


def diff_report(old_spirits, new_spirits, old_skills, new_skills):
    """对比上次数据与本次抓取，报告差异"""
    print(f"\n===== 与上次数据对比 =====")
    os_n = {(r['序号'], r['精灵名称']) for r in old_spirits}
    ns_n = {(r['序号'], r['精灵名称']) for r in new_spirits}
    added, gone = ns_n - os_n, os_n - ns_n
    print(f"  精灵：上次 {len(os_n)} → 本次 {len(ns_n)}；新增 {len(added)}，本次缺失 {len(gone)}")
    for r in sorted(added)[:20]:
        print(f"    + {r[0]} {r[1]}")
    if gone:
        print(f"    [本次缺失] {sorted(gone)[:20]}（连续两次缺失才淘汰）")

    os_k = {r['技能名称'] for r in old_skills}
    ns_k = {r['技能名称'] for r in new_skills}
    print(f"  技能：上次 {len(os_k)} → 本次 {len(ns_k)}；新增 {len(ns_k - os_k)}，本次缺失 {len(os_k - ns_k)}")
    for n in sorted(ns_k - os_k)[:20]:
        print(f"    + {n}")
    if os_k - ns_k:
        print(f"    [本次缺失] {sorted(os_k - ns_k)[:20]}（连续两次缺失才淘汰）")
    print("  说明：本次缺失的条目会与上一份快照比对，连续两次都读不到才从表中移除")


# ====================== 主流程 ======================
def write_all(spirits, skills, sources, combined, evo, stones, label,
              features=None, catalog=None):
    save_csv(spirits, F_SPIRITS, SPIRIT_FIELDS)
    save_xlsx(spirits, F_SPIRITS_XLSX, SPIRIT_FIELDS)
    save_csv(skills, F_SKILLS, SKILL_FIELDS)
    save_csv(sources, F_SOURCES, SOURCE_FIELDS)
    save_csv(combined, F_COMBINED, COMBINED_FIELDS)
    if stones is not None:
        save_csv(stones, F_STONES, STONE_FIELDS)
    save_csv(evo, F_EVO, EVO_FIELDS)
    if features is not None:
        save_csv(features, F_FEATURES, FEATURE_FIELDS)
    if catalog is not None:
        save_csv(catalog, F_CATALOG, CATALOG_FIELDS)
    check_data(spirits, skills, sources, combined, evo, stones, label, features, catalog)


def run_update():
    """抓取 → 体检 → 合并 → 复检（以本次抓取为准，连续两次缺失才淘汰）"""
    print("=" * 60)
    print("运行：体检 → 抓取 → 合并 → 复检")
    print("=" * 60)
    prev = load_prev_snapshot()
    if prev is None:
        print("[快照] 未发现上一份基底，本次只更新 / 追加，不淘汰任何记录")
    else:
        print(f"[快照] 上一份基底：{', '.join(f'{t}={len(k)}' for t, k in prev.items())} "
              f"（连续两次缺失才淘汰）")

    # 历史数据里若残留「???」等占位条目，先清掉
    old_spirits = [r for r in read_csv(F_SPIRITS, SPIRIT_FIELDS)
                   if not is_placeholder(r['精灵名称'])]
    old_skills = [r for r in read_csv(F_SKILLS, SKILL_FIELDS)
                  if not is_placeholder(r['技能名称'])]
    old_sources = [r for r in read_csv(F_SOURCES, SOURCE_FIELDS)
                   if not is_placeholder(r['精灵名称']) and not is_placeholder(r['技能名称'])]
    old_combined = read_csv(F_COMBINED, COMBINED_FIELDS)
    old_evo = read_csv(F_EVO, EVO_FIELDS)
    old_stones = read_csv(F_STONES, STONE_FIELDS)
    old_features = read_csv(F_FEATURES, FEATURE_FIELDS)
    old_catalog = read_csv(F_CATALOG, CATALOG_FIELDS)

    if old_spirits and old_skills:
        print(f"已有数据：精灵 {len(old_spirits)} / 技能 {len(old_skills)} / "
              f"来源 {len(old_sources)} / 合并表 {len(old_combined)}")
        check_data(old_spirits, old_skills, old_sources, old_combined, old_evo, label='上次数据')
    else:
        print("未发现已有数据，本次将全量建立")

    # 技能详情页最先生效：趁 wiki 限流滑动窗口干净，其内置的限流增量重试才能成功；
    # 否则在精灵/来源高压抓取之后，窗口持续发热，技能属性会大量缺失。
    print("\n[冷却] 等待 wiki 限流窗口重置...")
    time.sleep(SERIAL_GAP * 5)
    new_skills = crawl_skills(old_skills)
    new_spirits = sort_spirits(crawl_spirits())
    if not new_spirits or not new_skills:
        print("本次抓取为空，未改动数据文件")
        return
    new_sources, new_evo_by, new_catalog = crawl_sources([r['精灵名称'] for r in new_spirits])
    new_features, _ = crawl_features()

    if old_spirits:
        diff_report(old_spirits, new_spirits, old_skills, new_skills)

    pk = prev or {}
    print("\n===== 合并（覆盖更新 + 两次缺失淘汰）=====")
    spirits, k_spirits = merge_table('spirits', old_spirits, new_spirits, pk.get('spirits'))
    skills, k_skills = merge_table('skills', old_skills, new_skills, pk.get('skills'))
    sources, k_sources = merge_table('sources', old_sources, new_sources, pk.get('sources'))
    spirits = sort_spirits(spirits)
    print(f"合并结果：精灵 {len(spirits)} / 技能 {len(skills)} / 来源 {len(sources)}")

    # 技能石：一次静态请求覆盖全量；失败时按空结果处理，由 merge_table 暂留旧记录
    try:
        new_stones = crawl_skill_stones()
    except Exception as e:
        print(f"[警告] 聚合页抓取失败，本次按空结果处理：{e}")
        new_stones = []
    stones, k_stones = merge_table('stones', old_stones, new_stones, pk.get('stones'))
    stones.sort(key=lambda r: (int(r['序号']) if str(r['序号']).isdigit() else 10 ** 9,
                               r.get('精灵名称', '')))
    mark_quest(spirits, sources, stones)
    features, k_features = merge_table('features', old_features, new_features, pk.get('features'))
    mark_traits(spirits, owners_map_from_features(features))
    catalog, k_catalog = merge_table('catalog', old_catalog, new_catalog, pk.get('catalog'))

    # 进化链：按本次精灵页重建；整轮为空时保留旧数据，且不推进淘汰计数
    new_evo = sort_evo(enrich_evo_no(build_evo_chains(new_evo_by), spirits))
    if not new_evo and old_evo:
        print(f"[警告] 本次未解析到任何进化链，保留上次 {len(old_evo)} 行并跳过淘汰判定")
        evo = old_evo
        k_evo = {TABLE_KEYS['evo'](r) for r in old_evo}
    else:
        evo, k_evo = merge_table('evo', old_evo, new_evo, pk.get('evo'))
        evo = sort_evo(evo)

    write_all(spirits, skills, sources, build_combined(spirits, sources, stones),
              evo, stones, '合并后复检', features, catalog)
    save_snapshot({'spirits': k_spirits, 'skills': k_skills, 'sources': k_sources,
                   'stones': k_stones, 'evo': k_evo, 'features': k_features,
                   'catalog': k_catalog})
    print("\n更新完成")


def run_check():
    spirits = read_csv(F_SPIRITS, SPIRIT_FIELDS)
    skills = read_csv(F_SKILLS, SKILL_FIELDS)
    sources = read_csv(F_SOURCES, SOURCE_FIELDS)
    combined = read_csv(F_COMBINED, COMBINED_FIELDS)
    evo = read_csv(F_EVO, EVO_FIELDS)
    if not spirits and not skills:
        print("没有可体检的数据文件")
        return
    stones = read_csv(F_STONES, STONE_FIELDS) if os.path.exists(F_STONES) else None
    features = read_csv(F_FEATURES, FEATURE_FIELDS) if os.path.exists(F_FEATURES) else None
    catalog = read_csv(F_CATALOG, CATALOG_FIELDS) if os.path.exists(F_CATALOG) else None
    check_data(spirits, skills, sources, combined, evo, stones, '当前数据', features, catalog)


def run_rebuild():
    """不联网，用已有 CSV 重算合并表与「技能石课题」列，并据当前规则重算进化链特殊标记"""
    spirits = read_csv(F_SPIRITS, SPIRIT_FIELDS)
    skills = read_csv(F_SKILLS, SKILL_FIELDS)
    sources = read_csv(F_SOURCES, SOURCE_FIELDS)
    evo = sort_evo(enrich_evo_no(flag_evo_special(read_csv(F_EVO, EVO_FIELDS)), spirits))
    if not spirits or not sources:
        print("缺少数据文件，请先运行 nrc_scraper.py")
        return
    stones = read_csv(F_STONES, STONE_FIELDS) if os.path.exists(F_STONES) else None
    mark_quest(spirits, sources, stones)
    features = read_csv(F_FEATURES, FEATURE_FIELDS)
    if features:
        mark_traits(spirits, owners_map_from_features(features))
    write_all(spirits, skills, sources, build_combined(spirits, sources, stones), evo, stones,
              '重建后复检', features or None)


def run_features():
    """只抓特性图鉴：写出特性表，并给精灵档案补「特性」「特性介绍」两列。

    只发一次静态请求，不重抓 662 个精灵页；不写快照，避免影响其它表的淘汰判定。
    """
    features, owners_map = crawl_features()
    if not features:
        print("未解析到特性，未改动数据文件")
        return
    save_csv(features, F_FEATURES, FEATURE_FIELDS)
    spirits = [r for r in read_csv(F_SPIRITS, SPIRIT_FIELDS)
               if not is_placeholder(r.get('精灵名称', ''))]
    if not spirits:
        print("未找到精灵图鉴数据，已单独写出特性表")
        return
    mark_traits(sort_spirits(spirits), owners_map)
    save_csv(spirits, F_SPIRITS, SPIRIT_FIELDS)
    save_xlsx(spirits, F_SPIRITS_XLSX, SPIRIT_FIELDS)
    covered = sum(1 for r in spirits if r.get('特性'))
    print(f"特性表 {len(features)} 行已保存；精灵表 {len(spirits)} 行，"
          f"其中 {covered} 行已填特性 / 特性介绍")
    check_data(spirits, read_csv(F_SKILLS, SKILL_FIELDS),
               read_csv(F_SOURCES, SOURCE_FIELDS),
               read_csv(F_COMBINED, COMBINED_FIELDS),
               read_csv(F_EVO, EVO_FIELDS),
               read_csv(F_STONES, STONE_FIELDS), '特性更新后复检', features)


def run_normalize():
    """离线统一 skill_sources.csv 的「来源类型」「等级」写法（不联网、不动其它表）"""
    rows = read_csv(F_SOURCES, SOURCE_FIELDS)
    if not rows:
        print("没有来源表数据")
        return
    changed = 0
    for r in rows:
        before = (r.get('来源类型', ''), r.get('等级', ''))
        normalize_source_row(r)
        if (r.get('来源类型', ''), r.get('等级', '')) != before:
            changed += 1
    save_csv(rows, F_SOURCES, SOURCE_FIELDS)
    types = sorted({r['来源类型'] for r in rows})
    levels = sorted({r['等级'] for r in rows if r['等级']})
    print(f"来源表 {len(rows)} 行，已统一 {changed} 行；"
          f"来源类型 {types}；等级取值 {len(levels)} 种：{levels[:12]}")


def run_skillcatalog():
    """抓取全部精灵页，生成 / 刷新「精灵技能总表」（不动其它表、不写快照）。

    以精灵为单位整体替换：本次抓到技能卡的精灵用最新列表（版本改动后旧技能自然消失），
    抓取失败的精灵沿用旧记录，避免限流把整只精灵的技能清空。
    """
    names = [r['精灵名称'] for r in read_csv(F_SPIRITS, SPIRIT_FIELDS)
             if not is_placeholder(r.get('精灵名称', ''))]
    if not names:
        print("未找到精灵图鉴数据，请先运行 nrc_scraper.py")
        return
    _, _, catalog = crawl_sources(names)
    if not catalog:
        print("未解析到技能卡，未改动数据文件")
        return
    old = read_csv(F_CATALOG, CATALOG_FIELDS)
    new_by, old_by = defaultdict(list), defaultdict(list)
    for r in catalog:
        new_by[r['精灵名称']].append(r)
    for r in old:
        old_by[r['精灵名称']].append(r)
    merged, kept = [], []
    for name in sorted(set(old_by) | set(new_by)):
        if new_by.get(name):
            merged.extend(new_by[name])
        else:
            merged.extend(old_by.get(name, []))
            if old_by.get(name):
                kept.append(name)
    save_csv(merged, F_CATALOG, CATALOG_FIELDS)
    blood = sum(1 for r in merged if r.get('血脉属性'))
    print(f"技能总表 {len(merged)} 行 / {len({r['精灵名称'] for r in merged})} 个精灵，"
          f"其中 {blood} 行带血脉属性；本次抓取缺失、沿用旧数据的精灵 {len(kept)} 个")
    check_data(read_csv(F_SPIRITS, SPIRIT_FIELDS), read_csv(F_SKILLS, SKILL_FIELDS),
               read_csv(F_SOURCES, SOURCE_FIELDS), read_csv(F_COMBINED, COMBINED_FIELDS),
               read_csv(F_EVO, EVO_FIELDS), read_csv(F_STONES, STONE_FIELDS),
               '技能总表复检', read_csv(F_FEATURES, FEATURE_FIELDS) or None, merged)


# ====================== 图鉴课题（技能石）聚合页 ======================
def extract_skill_stones(html):
    """从图鉴课题聚合页解析技能石课题。

    数据内嵌在静态 HTML 的 book="..." 属性里（HTML 实体转义 JSON），
    一次静态请求即可覆盖全部 466 物种，无需分页 / 渲染。
    每个 handbook：number(序号) / name(物种名) / topics[]，
    其中 category=='skill' 的 topic 即技能石课题。
    """
    books = re.findall(r'book="([^"]+)"', html)
    rows = []
    for raw in books:
        try:
            o = json.loads(_html.unescape(raw))
        except Exception:
            continue
        no = str(o.get('number', '')).strip()
        name = (o.get('name') or '').strip()
        if is_placeholder(name):
            continue
        for t in o.get('topics', []):
            if t.get('category') != 'skill':
                continue
            rows.append({
                '序号': no,
                '精灵名称': name,
                '类别': '技能石',
                '条件': (t.get('text') or '').strip(),
                '奖励ID': ','.join(str(x) for x in t.get('rewards', [])),
            })
    return rows


def crawl_skill_stones():
    """抓取图鉴课题聚合页并解析全部技能石课题（静态请求即可，覆盖全量）"""
    url = "https://wiki.biligame.com/nrc/%E5%9B%BE%E9%89%B4%E8%AF%BE%E9%A2%98"
    print("抓取图鉴课题聚合页（技能石）...")
    page = fetch(url, stealth=False)            # 静态包已含全量数据，无需渲染
    body = getattr(page, 'body', b'') or b''
    if isinstance(body, str):
        body = body.encode('utf-8')
    html = body.decode('utf-8', 'ignore')
    rows = extract_skill_stones(html)
    rows.sort(key=lambda r: (int(r['序号']) if str(r['序号']).isdigit() else 10 ** 9, r['精灵名称']))
    return rows


def run_skillstones():
    rows = crawl_skill_stones()
    if not rows:
        print("未解析到任何技能石课题")
        return
    save_csv(rows, F_STONES, STONE_FIELDS)
    species = len({r['序号'] for r in rows if str(r['序号']).isdigit()})
    print(f"已写出 {F_STONES}：{len(rows)} 行（{species} 个物种）")


def validate_skill_stones():
    """交叉验证：聚合页技能石课题 vs 逐精灵页主爬取的课题（按序号比对）"""
    spirits = read_csv(F_SPIRITS, SPIRIT_FIELDS)
    sources = read_csv(F_SOURCES, SOURCE_FIELDS)
    stones = read_csv(F_STONES, STONE_FIELDS)
    if not stones:
        print("请先运行：python nrc_scraper.py skillstones")
        return
    if not spirits or not sources:
        print("缺少精灵 / 来源数据，无法验证")
        return

    no_map = build_spirit_no_map(spirits)          # 基础名 -> 序号
    # 全名精确匹配：分支形态（如「棋契陛下（白棋棋祈督分支）」）必须按全名定位序号，
    # 否则会被基础名归并到同一序号，造成验证误报。
    full_map = {r.get('精灵名称'): r.get('序号') for r in spirits if r.get('精灵名称')}

    # 主爬取：来源表里含「使用N次」获取条件的课题，按序号归集
    main = defaultdict(set)
    for r in sources:
        cond = r.get('获取条件') or ''
        if not re.search(r'使用\d+次', cond):
            continue
        nm = r.get('精灵名称', '')
        no = full_map.get(nm) or no_map.get(base_name(nm))   # 先全名，再基础名兜底
        if not no:
            continue
        for s in cond.split(';'):
            s = s.strip()
            if re.search(r'使用\d+次', s):
                main[no].add(s)

    # 聚合页：按序号归集条件
    agg = defaultdict(set)
    for r in stones:
        no = r.get('序号', '')
        if str(no).isdigit():
            agg[no].add(r.get('条件', '').strip())

    diffs = []
    for no in sorted(set(main) | set(agg), key=lambda x: int(x)):
        mset, aset = main.get(no, set()), agg.get(no, set())
        extra_main = mset - aset
        extra_agg = aset - mset
        if extra_main or extra_agg:
            diffs.append((no, extra_main, extra_agg))

    compared = len(set(main) | set(agg))
    print(f"\n===== 技能石课题交叉验证 =====")
    print(f"  对比物种数：{compared}")
    print(f"  一致物种数：{compared - len(diffs)}")
    print(f"  有差异物种数：{len(diffs)}")
    diff_rows = []
    for no, extra_main, extra_agg in diffs[:60]:
        name = next((r['精灵名称'] for r in stones if r['序号'] == no), '')
        print(f"  [{no} {name}]")
        for c in sorted(extra_main):
            print(f"      仅主爬取有: {c}")
            diff_rows.append({'序号': no, '精灵名称': name, '差异类型': '仅主爬取有', '条件': c})
        for c in sorted(extra_agg):
            print(f"      仅聚合页有: {c}")
            diff_rows.append({'序号': no, '精灵名称': name, '差异类型': '仅聚合页有', '条件': c})
    if diff_rows:
        save_csv(diff_rows, 'skill_stones_diff.csv', ['序号', '精灵名称', '差异类型', '条件'])
        print(f"  差异明细已写出 skill_stones_diff.csv（{len(diff_rows)} 行）")
    else:
        print("  未发现差异")
    print("  注：课题归属按全名精确映射序号（分支形态不再被基础名归并），"
          "故主爬取与聚合页可逐条对齐；若有差异会在此列出。")


def main():
    setup_logging()
    args = sys.argv[1:]
    mode = args[0] if args else 'update'
    if mode == 'update':
        run_update()
    elif mode == 'check':
        run_check()
    elif mode == 'rebuild':
        run_rebuild()
    elif mode == 'skillstones':
        run_skillstones()
    elif mode == 'validate':
        validate_skill_stones()
    elif mode == 'features':
        run_features()
    elif mode == 'normalize':
        run_normalize()
    elif mode == 'skillcatalog':
        run_skillcatalog()
    else:
        print(__doc__)
        return
    print("\n执行完毕")


if __name__ == "__main__":
    main()
