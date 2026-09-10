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

# Windows 下输出默认 GBK 会导致中文乱码，统一为 UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

from scrapling.fetchers import Fetcher, StealthyFetcher

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ====================== 配置 ======================
BASE = "https://wiki.biligame.com/nrc/"
SPIRIT_URL = BASE + quote("精灵图鉴")
SKILL_URL = BASE + quote("技能图鉴")

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

# 输出三张表 + 一个视图：
#   spirits_data_all.csv  精灵档案（622）
#   skills_data.csv       技能档案（579，技能自身的属性，与精灵无关）
#   skill_sources.csv     精灵 × 技能 来源（谁在哪、怎么获得）
#   combined_data.csv     精灵 → 最多 3 个技能石 + 获取条件（视图，由前两张算出）
SPIRIT_FIELDS = ['序号', '精灵名称', '精灵阶数', '进化位置', '主属性',
                 '副属性', '形态', '主形态/地区形态', '异色', '赛季', '技能石课题']
SKILL_FIELDS = ['技能名称', '技能属性', '技能类型', '威力', '耗能',
                '技能描述', '风味文本']
SOURCE_FIELDS = ['精灵名称', '技能名称', '来源类型', '等级', '获取条件']
COMBINED_FIELDS = ['序号', '精灵名称', '属性', '技能名称1', '技能名称2', '技能名称3', '条件']

F_SPIRITS = 'spirits_data_all.csv'
F_SPIRITS_XLSX = 'spirits_data_all.xlsx'
F_SKILLS = 'skills_data.csv'
F_SOURCES = 'skill_sources.csv'
F_COMBINED = 'combined_data.csv'


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
        if not name or not number:
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
    rows = parse_spirit_list(page)
    print(f"解析到 {len(rows)} 条精灵数据")
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
        if not name or not href or 'redlink=1' in href:
            continue
        if not href.startswith('http'):
            href = 'https://wiki.biligame.com' + href
        links.setdefault(name, href)
    print(f"技能图鉴去重后共 {len(links)} 个技能")
    return [(n, u) for n, u in links.items()]


def parse_skill_detail(page, skill_name):
    """解析技能详情页：属性 / 类型 / 威力 / 耗能 / 描述 / 风味"""
    detail = {'技能名称': skill_name, '技能属性': '', '技能类型': '', '威力': '',
              '耗能': '', '技能描述': '', '风味文本': ''}

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
    """限流的「简化页」没有属性标签"""
    return bool(page.css('.nrc-skill-tags'))


def crawl_skills():
    """技能档案：只取技能自身的属性。

    「谁能学这个技能」由精灵页产出（skill_sources.csv），覆盖面更全
    （技能页 30733 对 ⊂ 精灵页 30750 对），所以这里不解析来源。
    """
    print(f"\n===== 抓取技能图鉴 =====")
    links = crawl_skill_links()
    if not links:
        return []
    got = Spider('技能详情').run(links, skill_ok, parse_skill_detail)
    rows = [{k: d[k] for k in SKILL_FIELDS}
            for name, _ in links for d in [got.get(name)] if d]
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
    """解析精灵详情页技能表（技能 + 来源），并把图鉴课题的获取条件挂上去"""
    cards = page.css('.roco-sk-card')
    if not cards:
        return []
    quests = {q['技能名称']: q['获取条件'] for q in parse_quests(page, pet_name)}
    rows = []
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


def crawl_sources(names):
    """精灵 × 技能 来源表：自学(Lv.x) / 技能石 / 血脉 / 传说，并带获取条件"""
    print(f"\n===== 抓取精灵详情页技能表 =====")
    targets = [(n, BASE + quote(n)) for n in names]
    got = Spider('精灵技能表').run(
        targets, lambda p: bool(p.css('.roco-sk-card')), parse_pet_skills)
    rows = []
    for n in names:
        rows.extend(got.get(n, []))
    print(f"技能来源共 {len(rows)} 条（{len(got)}/{len(names)} 个精灵）")
    return rows


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


# ====================== 合并（保留原有 + 追加新增）======================
def merge_spirits(old, new):
    seen = {(r['序号'], r['精灵名称']) for r in old}
    merged, added = list(old), []
    for r in new:
        key = (r['序号'], r['精灵名称'])
        if key in seen:
            continue
        seen.add(key)
        merged.append(r)
        added.append(r)
    return sort_spirits(merged), added


def merge_skills(old, new):
    """技能档案按技能名去重（原有优先，只补新增技能）"""
    merged, added = list(old), []
    seen = {r['技能名称'] for r in old}
    for r in new:
        if r['技能名称'] in seen:
            continue
        seen.add(r['技能名称'])
        merged.append(r)
        added.append(r)
    return merged, added


def merge_sources(old, new):
    """来源表按 (精灵, 技能) 去重；新数据带条件时覆盖旧的空条件"""
    idx = {(r['精灵名称'], r['技能名称']): r for r in old}
    added = []
    for r in new:
        key = (r['精灵名称'], r['技能名称'])
        if key not in idx:
            idx[key] = r
            added.append(r)
        elif r.get('获取条件') and not idx[key].get('获取条件'):
            idx[key] = r
    return list(idx.values()), added


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


def build_combined(spirits, sources, max_skills=MAX_SKILLS):
    """精灵 -> 最多 N 个技能石 + 获取条件（视图）

    数据来自来源表里的「图鉴课题」：每只精灵的课题奖励若干技能石，
    条件形如「使用2次水炮」。没有课题的精灵不出现在表里。

    同一精灵的各形态（「鸭吉吉（燃了鸭）」等）条件完全相同，
    合并成一行，名称取括号前的基础名。

    同一序号的进化链成员（迪莫 / 圣光迪莫 / 圣水迪莫…）课题也一样，
    再合并成一行，保留本体形态（有「精灵阶数」的那个）。
    条件不同的仍各自成行。
    """
    by_base = defaultdict(list)
    for s in spirits:
        by_base[base_name(s.get('精灵名称', ''))].append(s)

    idx = quest_index(sources)

    rows = []
    for base, spirit_list in by_base.items():
        picked = idx.get(base, [])[:max_skills]
        if not picked:
            continue
        # 代表形态：优先有阶数的（本体），再取名字最短的
        info = min(spirit_list, key=lambda r: (0 if r.get('精灵阶数') else 1,
                                               len(r.get('精灵名称', '')),
                                               r.get('精灵名称', '')))
        rows.append((info.get('序号', ''), info.get('主属性', ''),
                     info.get('精灵阶数', ''), base, tuple(picked)))

    # 同一序号 + 相同条件 = 同一条进化链，只留一行
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
    combined.sort(key=lambda x: (
        int(x['序号']) if str(x['序号']).isdigit() else 99999, x['精灵名称']))
    return combined


# ====================== 体检 ======================
def mark_quest(spirits, sources):
    """给精灵档案补「技能石课题」列（有/无），替代单独的清单文件"""
    has = {base_name(r['精灵名称']) for r in sources
           if re.search(r'使用\d+次', r.get('获取条件') or '')}
    for r in spirits:
        r['技能石课题'] = '有' if base_name(r.get('精灵名称', '')) in has else '无'
    return spirits


def check_data(spirits, skills, sources, combined, label=''):
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
    print(f"  精灵：上次 {len(os_n)} → 本次 {len(ns_n)}；新增 {len(added)}，消失 {len(gone)}")
    for r in sorted(added)[:20]:
        print(f"    + {r[0]} {r[1]}")
    if gone:
        print(f"    [注意] 本次未抓到：{sorted(gone)[:20]}")

    os_k = {r['技能名称'] for r in old_skills}
    ns_k = {r['技能名称'] for r in new_skills}
    print(f"  技能：上次 {len(os_k)} → 本次 {len(ns_k)}；新增 {len(ns_k - os_k)}，消失 {len(os_k - ns_k)}")
    for n in sorted(ns_k - os_k)[:20]:
        print(f"    + {n}")
    if os_k - ns_k:
        print(f"    [注意] 本次未抓到：{sorted(os_k - ns_k)[:20]}")
    print(f"  技能来源记录：上次 {len(old_skills)} → 本次 {len(new_skills)}"
          f"（+{len(new_skills) - len(old_skills)}）")


# ====================== 主流程 ======================
def write_all(spirits, skills, sources, combined, label):
    save_csv(spirits, F_SPIRITS, SPIRIT_FIELDS)
    save_xlsx(spirits, F_SPIRITS_XLSX, SPIRIT_FIELDS)
    save_csv(skills, F_SKILLS, SKILL_FIELDS)
    save_csv(sources, F_SOURCES, SOURCE_FIELDS)
    save_csv(combined, F_COMBINED, COMBINED_FIELDS)
    check_data(spirits, skills, sources, combined, label)


def run_update():
    """抓取 → 体检 → 合并 → 复检（已有数据则先体检并报告差异）"""
    print("=" * 60)
    print("运行：体检 → 抓取 → 合并 → 复检")
    print("=" * 60)
    old_spirits = read_csv(F_SPIRITS, SPIRIT_FIELDS)
    old_skills = read_csv(F_SKILLS, SKILL_FIELDS)
    old_sources = read_csv(F_SOURCES, SOURCE_FIELDS)
    old_combined = read_csv(F_COMBINED, COMBINED_FIELDS)

    if old_spirits and old_skills:
        print(f"已有数据：精灵 {len(old_spirits)} / 技能 {len(old_skills)} / "
              f"来源 {len(old_sources)} / 合并表 {len(old_combined)}")
        check_data(old_spirits, old_skills, old_sources, old_combined, '上次数据')
    else:
        print("未发现已有数据，本次将全量建立")

    new_spirits = sort_spirits(crawl_spirits())
    new_skills = crawl_skills()
    if not new_spirits or not new_skills:
        print("本次抓取为空，未改动数据文件")
        return
    new_sources = crawl_sources([r['精灵名称'] for r in new_spirits])

    if old_spirits:
        diff_report(old_spirits, new_spirits, old_skills, new_skills)

    spirits, add_s = merge_spirits(old_spirits, new_spirits)
    skills, add_k = merge_skills(old_skills, new_skills)
    sources, add_r = merge_sources(old_sources, new_sources)
    print(f"\n合并结果：精灵 +{len(add_s)} → {len(spirits)}；"
          f"技能 +{len(add_k)} → {len(skills)}；来源 +{len(add_r)} → {len(sources)}")

    mark_quest(spirits, sources)
    write_all(spirits, skills, sources, build_combined(spirits, sources), '合并后复检')
    print("\n更新完成")


def run_check():
    spirits = read_csv(F_SPIRITS, SPIRIT_FIELDS)
    skills = read_csv(F_SKILLS, SKILL_FIELDS)
    sources = read_csv(F_SOURCES, SOURCE_FIELDS)
    combined = read_csv(F_COMBINED, COMBINED_FIELDS)
    if not spirits and not skills:
        print("没有可体检的数据文件")
        return
    check_data(spirits, skills, sources, combined, '当前数据')


def run_rebuild():
    """不联网，用已有 CSV 重算合并表与「技能石课题」列"""
    spirits = read_csv(F_SPIRITS, SPIRIT_FIELDS)
    skills = read_csv(F_SKILLS, SKILL_FIELDS)
    sources = read_csv(F_SOURCES, SOURCE_FIELDS)
    if not spirits or not sources:
        print("缺少数据文件，请先运行 nrc_scraper.py")
        return
    mark_quest(spirits, sources)
    write_all(spirits, skills, sources, build_combined(spirits, sources), '重建后复检')


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
    else:
        print(__doc__)
        return
    print("\n执行完毕")


if __name__ == "__main__":
    main()
