# -*- coding: utf-8 -*-
"""拉取 USDA FAS ESR 美豆、美玉米、美豆粕、美豆油、美小麦周度出口销售。

五个品种合成同一个网页，并各写一份 xlsx。
密钥只读环境变量 USDA_FAS_API_KEY。脚本和输出里都不写密钥。

更新一周只需要在仓库根目录运行：

    python excel_s/us_weekly_export/fetch_soy_esr.py

不带参数运行时，每次都会：
- 刷新参考表（commodities、countries、unitsOfMeasure、datareleasedates）；
- 对每个品种重新拉取「当前营销年」及其后已列入的营销年，覆盖 cache/ 里的旧文件；
- 历史营销年已有缓存就复用，缺哪年补哪年。
当前营销年优先取 datareleasedates 里该品种的 marketYear（营销年结束年）；
拿不到时按品种营销年起始月份（PRODUCTS 的 start_month，6、9 或 10 月）和今天日期推算。
API_YEARS 没覆盖到当前营销年时自动扩展。
接口失败则退出，不编造序列，也不拿旧缓存冒充新数据。
--from-cache 只用 cache/ 里已有的文件重画出表，不联网。
"""
from __future__ import annotations

import json
import os
import sys
import time
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
EXCEL_S = ROOT.parent
sys.path.insert(0, str(EXCEL_S))
from chart_watermark import html_watermark_css, html_watermark_div, seasonal_year_styles  # noqa: E402

BASE = "https://api.fas.usda.gov/api/esr"
# 接口 marketYear 是营销年结束年。2019 = 18/19，2027 = 26/27。
# 下限固定 2019；上限在 main() 里按当前营销年自动扩展（extend_api_years）。
FIRST_API_YEAR = 2019
API_YEARS = range(FIRST_API_YEAR, 2028)
PRODUCTS = (
    {"id": "soy", "code": 801, "en": "Soybeans", "zh": "美豆", "cache": "soy_801", "xlsx": "美豆周度出口销售.xlsx", "start_month": 9},
    {"id": "corn", "code": 401, "en": "Corn", "zh": "美玉米", "cache": "corn_401", "xlsx": "美玉米周度出口销售.xlsx", "start_month": 9},
    {"id": "meal", "code": 901, "en": "Soybean cake & meal", "zh": "美豆粕", "cache": "meal_901", "xlsx": "美豆粕周度出口销售.xlsx", "start_month": 10},
    {"id": "oil", "code": 902, "en": "Soybean Oil", "zh": "美豆油", "cache": "oil_902", "xlsx": "美豆油周度出口销售.xlsx", "start_month": 10},
    # 小麦用 All Wheat 107（各类小麦合计）。ESR 小麦营销年 6 月 1 日–5 月 31 日。
    # 出口目标.xlsx 没有小麦块时 target_optional 生效：进度留空，不编目标。
    {"id": "wheat", "code": 107, "en": "All Wheat", "zh": "美小麦", "cache": "wheat_107", "xlsx": "美小麦周度出口销售.xlsx", "start_month": 6, "target_optional": True},
)
MY_NOTE = {
    "jun": "ESR 官方营销年 6 月 1 日–5 月 31 日，接口 marketYear 是结束年。季节性图从 6 月的第 1 个报告周起。",
    "sep": "ESR 官方营销年 9 月 1 日–8 月 31 日，接口 marketYear 是结束年。",
    "oct": "ESR 官方营销年 10 月 1 日–9 月 30 日，接口 marketYear 是结束年。季节性图从 10 月的第 1 个报告周起。",
}
# 营销年起始月份与 MY_NOTE 口径的对应；index_rows 从数据识别出的口径必须与 PRODUCTS 的 start_month 一致。
CALENDAR_START = {"jun": 6, "sep": 9, "oct": 10}
TARGET_XLSX = ROOT / "出口目标.xlsx"
CACHE = ROOT / "cache"
OUT_HTML = ROOT / "美豆周度出口销售.html"
OUT_MD = ROOT / "口径说明.md"

# 接口字段。页面表头 currentGrossNewSales 对应接口 grossNewSales。
API_FIELDS = (
    "weeklyExports",
    "accumulatedExports",
    "outstandingSales",
    "grossNewSales",
    "currentMYNetSales",
    "currentMYTotalCommitment",
    "nextMYOutstandingSales",
    "nextMYNetSales",
)
# 展示名，顺序与需求一致
METRICS = (
    ("weeklyExports", "weeklyExports", "周度装运", "weeklyExports"),
    ("accumulatedExports", "accumulatedExports", "累计装运", "accumulatedExports"),
    ("outstandingSales", "outstandingSales", "本年度未装运", "outstandingSales"),
    ("currentGrossNewSales", "currentGrossNewSales", "本年度总新销售", "grossNewSales"),
    ("currentMYNetSales", "currentMYNetSales", "本年度净销售", "currentMYNetSales"),
    ("currentMYTotalCommitment", "currentMYTotalCommitment", "本年度总销售", "currentMYTotalCommitment"),
    ("nextMYOutstandingSales", "nextMYOutstandingSales", "下年度未装运", "nextMYOutstandingSales"),
    ("nextMYNetSales", "nextMYNetSales", "下年度净销售", "nextMYNetSales"),
    ("grossMinusNet", "grossMinusNet", "总新销售减净销售", None),
)
PROGRESS = (
    ("shipProg", "shipProg", "装运进度", "accumulatedExports", "cur"),
    ("saleProg", "saleProg", "本年度销售进度", "currentMYTotalCommitment", "cur"),
    ("nextProg", "nextProg", "下年度销售进度", "nextMYOutstandingSales", "next"),
)

# 欧盟：FAS regionId=1 的成员国，去掉空的集团行 1 和英国 4120；
# 比卢是一个代码；克罗地亚在 FAS 里记在 OTHER EUROPE，按成员国并入。
EU_CODES = (
    4010, 4050, 4090, 4190, 4210, 4230, 4270, 4280, 4330, 4351, 4359,
    4370, 4470, 4490, 4510, 4550, 4700, 4710, 4730, 4750, 4791, 4792,
    4840, 4850, 4870, 4910,
)
UK_CODE = 4120
CHINA = 5700
UNKNOWN = 9990
VIETNAM_OLD = 5510  # 若样本内出现南越旧码，并入越南

ONE = {
    "cn": (CHINA,),
    "jp": (5880,),
    "kr": (5800,),
    "tw": (5830,),
    "idn": (5600,),
    "th": (5490,),
    "vn": (5520,),
    "bd": (5380,),
    "pk": (5350,),
    "tr": (4890,),
    "uk": (UK_CODE,),
    "eg": (7290,),
    "dz": (7210,),
    "mx": (2010,),
    "unk": (UNKNOWN,),
}

# 分组, 名称, id, 是否加总行
ROWS = (
    ("汇总", "Total", "total", True),
    ("汇总", "中国", "cn", False),
    ("汇总", "除中国外", "excn", True),
    ("汇总", "Unknown", "unk", False),
    ("东亚", "东亚", "ea", True),
    ("东亚", "中国", "cn", False),
    ("东亚", "日本", "jp", False),
    ("东亚", "韩国", "kr", False),
    ("东亚", "中国台湾", "tw", False),
    ("东南亚", "东南亚", "sea", True),
    ("东南亚", "印尼", "idn", False),
    ("东南亚", "泰国", "th", False),
    ("东南亚", "越南", "vn", False),
    ("东南亚", "孟加拉", "bd", False),
    ("中亚", "中亚", "ca", True),
    ("中亚", "巴基斯坦", "pk", False),
    ("中亚", "土耳其", "tr", False),
    ("欧洲", "欧盟", "eu", False),
    ("欧洲", "英国", "uk", False),
    ("欧洲", "欧盟+英国", "eupuk", True),
    ("非洲", "非洲", "af", True),
    ("非洲", "埃及", "eg", False),
    ("非洲", "阿尔及利亚", "dz", False),
    ("美洲", "墨西哥", "mx", False),
)

# 对照表只列地区里的目的地。Total、除中国外、Unknown 不进表。
# currentMYNetSales 表与季节性图同一口径：最新一周的单周净销售。
TABLE_ROWS = (
    ("东亚", "中国", "cn"),
    ("东亚", "日本", "jp"),
    ("东亚", "韩国", "kr"),
    ("东亚", "中国台湾", "tw"),
    ("东南亚", "印尼", "idn"),
    ("东南亚", "泰国", "th"),
    ("东南亚", "越南", "vn"),
    ("东南亚", "孟加拉", "bd"),
    ("中亚", "巴基斯坦", "pk"),
    ("中亚", "土耳其", "tr"),
    ("欧洲", "欧盟", "eu"),
    ("欧洲", "英国", "uk"),
    ("欧洲", "欧盟+英国", "eupuk"),
    ("非洲", "埃及", "eg"),
    ("非洲", "阿尔及利亚", "dz"),
    ("美洲", "墨西哥", "mx"),
)
TABLE_METRICS = (
    (
        "currentMYTotalCommitment",
        "本年度总销售",
        "currentMYTotalCommitment",
        "本年度总销售 currentMYTotalCommitment：本营销年累计装运与本年度未装运之和。",
    ),
    (
        "currentMYNetSales",
        "本年度净销售",
        "currentMYNetSales",
        "本年度净销售 currentMYNetSales 指本周单周净销售，不是本营销年累计。",
    ),
    (
        "nextMYOutstandingSales",
        "下一年度未装运",
        "nextMYOutstandingSales",
        "下一年度未装运 nextMYOutstandingSales：下一营销年已经销售、尚未装运的数量。",
    ),
)
SHOT_ROWS = (("Total", "total"), ("中国", "cn"), ("除中国外", "excn"))
SHOT_FIELDS = (
    ("currentMYNetSales", "本年度净销售", "万吨"),
    ("currentMYTotalCommitment", "本年度总销售", "万吨"),
    ("shipProg", "装运进度", "%"),
    ("saleProg", "本年度销售进度", "%"),
)

YEAR_ORDER_NOTE = ("18/19", "19/20", "20/21", "21/22", "22/23", "23/24", "24/25", "25/26", "26/27")

CN_FONT = "微软雅黑"
FILL_HEAD = PatternFill("solid", fgColor="1F3A5F")
FILL_AGG = PatternFill("solid", fgColor="F4F7FB")
FILL_NOTE = PatternFill("solid", fgColor="FFF8E8")
FONT_HEAD = Font(name=CN_FONT, color="FFFFFF", bold=True, size=10)
FONT_BASE = Font(name=CN_FONT, size=10, color="222222")
FONT_AGG = Font(name=CN_FONT, size=10, color="222222", bold=True)
FONT_POS = Font(name=CN_FONT, size=10, color="D00000", bold=True)
FONT_NEG = Font(name=CN_FONT, size=10, color="17803D", bold=True)
FONT_ZERO = Font(name=CN_FONT, size=10, color="222222")
THIN = Border(
    left=Side(style="thin", color="E1E4E8"),
    right=Side(style="thin", color="E1E4E8"),
    top=Side(style="thin", color="E1E4E8"),
    bottom=Side(style="thin", color="E1E4E8"),
)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(vertical="center", horizontal="left")


def fail(msg: str) -> None:
    print(msg)
    raise SystemExit(1)


def redact(text: str) -> str:
    secret = os.environ.get("USDA_FAS_API_KEY") or ""
    if secret:
        text = text.replace(secret, "***")
    return text


def tenths_div(numer: int, denom: int) -> int:
    """Round numer/denom to an integer, half away from zero."""
    if denom == 0:
        fail("除数为 0")
    sign = -1 if (numer < 0) ^ (denom < 0) else 1
    n, d = abs(numer), abs(denom)
    q, r = divmod(n, d)
    if r * 2 >= d:
        q += 1
    return sign * q


def wan_tenths(tons: int) -> int:
    """吨 → 0.1 万吨的整数。1 万吨 = 10000 吨，十分位看千吨的余数。"""
    return tenths_div(int(tons), 1000)


def fmt_tenths(t) -> str:
    if t is None:
        return "—"
    sign = "-" if t < 0 else ""
    a = abs(int(t))
    return f"{sign}{a // 10}.{a % 10}"


def pp_tenths(diff_percent: Decimal) -> int:
    q = diff_percent.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return int(q * 10)


def label_of_api(api_year: int) -> str:
    return f"{(api_year - 1) % 100:02d}/{api_year % 100:02d}"


def next_my_label(label: str) -> str:
    end = int(label.split("/")[1])
    return f"{end:02d}/{(end + 1) % 100:02d}"


def load_key() -> str:
    key = (os.environ.get("USDA_FAS_API_KEY") or "").strip()
    if not key:
        fail("未设置环境变量 USDA_FAS_API_KEY。密钥只从环境变量读取，不写在脚本里。")
    return key


def esr_get(session: requests.Session, path: str, allow_empty: bool = False):
    url = BASE + path
    last = "请求失败"
    for attempt in range(3):
        try:
            resp = session.get(url, timeout=180)
        except requests.RequestException as exc:
            last = redact(str(exc))
            time.sleep(2)
            continue
        if resp.status_code == 429 and attempt < 2:
            time.sleep(3)
            continue
        if resp.status_code != 200:
            body = redact(resp.text[:300])
            fail(f"ESR {path} 返回 {resp.status_code}。{body}")
        try:
            data = resp.json()
        except ValueError:
            fail(f"ESR {path} 不是 JSON。{redact(resp.text[:200])}")
        if allow_empty and data == []:
            return None
        return data
    fail(f"ESR {path} 连续失败。{last}")


def _parse_my_label(lab) -> str:
    text = str(lab).strip()
    if text.endswith("年度"):
        text = text[:-2]
    if "/" not in text:
        fail(f"无法识别出口目标列名：{lab}。列对不上，不算进度。")
    return text


def load_targets() -> dict:
    """四个品种各一块。没有名为本年度/下年度的列，年份列必须彼此一致。"""
    if not TARGET_XLSX.exists():
        fail(f"找不到出口目标：{TARGET_XLSX}")
    wb = load_workbook(TARGET_XLSX, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    titles = {p["zh"]: p["id"] for p in PRODUCTS}
    found = {}
    label_sig = None
    i = 0
    while i < len(rows):
        head = rows[i][0] if rows[i] and rows[i][0] is not None else None
        if head not in titles:
            i += 1
            continue
        if i + 2 >= len(rows):
            fail(f"{head} 后面缺少累计出口目标和日期两行。列对不上，不算进度。")
        vals, labs = rows[i + 1], rows[i + 2]
        if vals[0] is None or str(vals[0]).strip() != "累计出口目标(t)":
            fail(f"{head} 的下一行不是「累计出口目标(t)」，而是 {vals[0]!r}。列对不上，不算进度。")
        if labs[0] is None or str(labs[0]).strip() != "日期":
            fail(f"{head} 的日期行不是「日期」，而是 {labs[0]!r}。列对不上，不算进度。")
        out = {}
        order = []
        for val, lab in zip(vals[1:], labs[1:]):
            if lab is None and val is None:
                continue
            if lab is None or val is None:
                fail(f"{head} 有目标和年份对不齐：{lab!r} {val!r}。不算进度。")
            key = _parse_my_label(lab)
            if key in out:
                fail(f"{head} 年份重复：{key}。不算进度。")
            if isinstance(val, bool) or not isinstance(val, (int, float)) or (
                isinstance(val, float) and not float(val).is_integer()
            ):
                fail(f"{head} {key} 的目标不是整数吨：{val!r}。不算进度。")
            out[key] = int(val)
            order.append(key)
        for lab in YEAR_ORDER_NOTE:
            if lab not in out:
                fail(f"{head} 出口目标缺少 {lab}。不算进度。")
        if label_sig is None:
            label_sig = tuple(order)
        elif tuple(order) != label_sig:
            fail(
                f"{head} 的年份列是 {order}，与前面的 {list(label_sig)} 不一致。不算进度。"
            )
        pid = titles[head]
        if pid in found:
            fail(f"{head} 在出口目标里出现两次。不算进度。")
        found[pid] = out
        i += 3
    for p in PRODUCTS:
        if p["id"] not in found and p.get("target_optional"):
            found[p["id"]] = {}
    missing = [p["zh"] for p in PRODUCTS if p["id"] not in found]
    if missing:
        fail("出口目标.xlsx 缺少：" + "、".join(missing) + "。不算进度。")
    return found


def parse_day(iso: str):
    return datetime.strptime(iso[:10], "%Y-%m-%d").date()


def as_int(value, field: str, where: str) -> int:
    if value is None:
        fail(f"{where} 的 {field} 为空，停止，不补数。")
    if isinstance(value, bool):
        fail(f"{where} 的 {field} 类型异常")
    if isinstance(value, float):
        if not value.is_integer():
            fail(f"{where} 的 {field} 不是整数吨")
        return int(value)
    return int(value)


def build_session(key: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({"X-Api-Key": key, "Accept": "application/json"})
    return s


def _unit_name(units, unit_id) -> str:
    for u in units:
        if u.get("unitId") == unit_id:
            return str(u.get("unitNames", "")).strip()
    return ""


def _release_day(release) -> str | None:
    if release and release.get("releaseTimeStamp"):
        return str(release["releaseTimeStamp"])[:10]
    return None


def fetch_reference(session: requests.Session):
    commodities = esr_get(session, "/commodities")
    countries = esr_get(session, "/countries")
    units = esr_get(session, "/unitsOfMeasure")
    releases = esr_get(session, "/datareleasedates")
    CACHE.mkdir(parents=True, exist_ok=True)
    for name, obj in (
        ("commodities.json", commodities),
        ("countries.json", countries),
        ("unitsOfMeasure.json", units),
        ("datareleasedates.json", releases),
    ):
        (CACHE / name).write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")
    by_code = {c.get("commodityCode"): c for c in commodities}
    for spec in PRODUCTS:
        row = by_code.get(spec["code"])
        if row is None:
            fail(f"商品表里没有 {spec['code']} {spec['en']}")
        got = str(row.get("commodityName", "")).strip().lower()
        if got != spec["en"].lower():
            fail(f"代码 {spec['code']} 不是 {spec['en']}：{row.get('commodityName')}")
        unit_name = _unit_name(units, row.get("unitId"))
        if unit_name.lower() != "metric tons":
            fail(f"{spec['zh']} 单位不是 Metric Tons：{unit_name or row.get('unitId')}")
    names = {c["countryCode"]: c["countryDescription"].strip() for c in countries}
    region1 = {c["countryCode"] for c in countries if c.get("regionId") == 1}
    release_map = {}
    for spec in PRODUCTS:
        release = next((r for r in releases if r.get("commodityCode") == spec["code"]), None)
        release_map[spec["code"]] = release
    return names, region1, release_map


def cache_path(spec, year: int) -> Path:
    return CACHE / f"{spec['cache']}_my{year}.json"


def fetch_years(session: requests.Session, spec, years, allow_empty_years=(), tag=None) -> dict:
    """拉取并覆盖写缓存。allow_empty_years 里的年份若返回空列表，视作尚未发布；
    此时若磁盘上已有该年缓存，说明接口与旧数据矛盾，停止，不拿旧缓存冒充。"""
    out = {}
    tag = tag or {}
    for year in list(years):
        path = f"/exports/commodityCode/{spec['code']}/allCountries/marketYear/{year}"
        allow_empty = year in allow_empty_years
        data = esr_get(session, path, allow_empty=allow_empty)
        if data is None:
            if cache_path(spec, year).exists():
                fail(
                    f"{spec['zh']} marketYear {year} 接口返回空，但缓存里有该年旧数据。"
                    "停止，不用旧缓存冒充新数据。"
                )
            print(f"{spec['zh']} marketYear {year} 没有数据，视作该营销年尚未发布")
            continue
        if not isinstance(data, list) or not data:
            fail(f"{spec['zh']} 营销年参数 {year} 没有数据，停止，不补数。")
        cache_path(spec, year).write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        out[year] = data
        dates = sorted({str(r.get("weekEndingDate", ""))[:10] for r in data})
        print(
            f"{spec['zh']} my {year} {label_of_api(year)} rows={len(data)} "
            f"weeks={len(dates)} {dates[0]} → {dates[-1]}{tag.get(year, '')}"
        )
    return out


def current_market_year(spec, release, today: date | None = None) -> tuple[int, str]:
    """当前营销年（接口口径＝营销年结束年）与判定来源。
    优先 datareleasedates 里该品种的 marketYear；拿不到就按 start_month 和今天推算：
    今天月份 ≥ 起始月份，则营销年结束于明年，否则结束于今年。"""
    if release:
        my = release.get("marketYear")
        try:
            my = int(my)
        except (TypeError, ValueError):
            my = None
        if my and my >= FIRST_API_YEAR:
            return my, "datareleasedates.marketYear"
    today = today or date.today()
    start = spec["start_month"]
    my = today.year + 1 if today.month >= start else today.year
    return my, f"按 {start} 月起始与今天 {today.isoformat()} 推算"


def extend_api_years(latest_year: int) -> None:
    """API_YEARS 没覆盖到当前营销年时扩展到当前年。只扩不缩。"""
    global API_YEARS
    if latest_year > API_YEARS[-1]:
        API_YEARS = range(FIRST_API_YEAR, latest_year + 1)
        print(f"API_YEARS 扩展到 {FIRST_API_YEAR}–{latest_year}")


def index_rows(raw_years: dict, names: dict, region1: set, commodity_code: int):
    """raw[(api_year, date, code)] = metric tons. Also per-year week lists."""
    raw = {}
    weeks = {}
    present_codes = set()
    old_vietnam = False
    for year, rows in raw_years.items():
        bucket = defaultdict(list)
        for rec in rows:
            if rec.get("commodityCode") != commodity_code:
                fail(f"{year} 混入了其他商品 {rec.get('commodityCode')}")
            if rec.get("unitId") != 1:
                fail(f"{year} 出现非公吨 unitId={rec.get('unitId')}")
            code = int(rec["countryCode"])
            day = str(rec["weekEndingDate"])[:10]
            where = f"{year} {day} {code}"
            metrics = {f: as_int(rec.get(f), f, where) for f in API_FIELDS}
            key = (year, day, code)
            if key in raw:
                fail(f"重复记录 {where}")
            raw[key] = metrics
            bucket[day].append(code)
            present_codes.add(code)
            if code == VIETNAM_OLD:
                old_vietnam = True
        ordered = sorted(bucket)
        for i in range(1, len(ordered)):
            delta = (parse_day(ordered[i]) - parse_day(ordered[i - 1])).days
            if delta != 7:
                fail(
                    f"代码 {commodity_code} {label_of_api(year)} 周序列不连续："
                    f"{ordered[i - 1]} → {ordered[i]} 间隔 {delta} 天。停止，不对齐猜测。"
                )
        if parse_day(ordered[0]) < parse_day("2018-06-01"):
            fail(f"代码 {commodity_code} 首周 {ordered[0]} 早于 2018-06-01")
        weeks[year] = ordered
        # 每个周末至少有一个国家
        if any(len(bucket[d]) == 0 for d in ordered):
            fail(f"{year} 有空周")
    # 相邻营销年只允许共享一个边界周末
    years = sorted(weeks)
    for a, b in zip(years, years[1:]):
        overlap = set(weeks[a]) & set(weeks[b])
        if len(overlap) > 1:
            fail(f"{label_of_api(a)} 与 {label_of_api(b)} 重叠周末超过 1 个：{sorted(overlap)}")
    unknown_eu = (present_codes & region1) - set(EU_CODES) - {1, UK_CODE}
    if unknown_eu:
        shown = ", ".join(f"{c} {names.get(c, '')}" for c in sorted(unknown_eu))
        fail(f"欧盟地区出现未列入的国家，停止以免少加：{shown}")
    bloc_nonzero = False
    for (year, day, code), metrics in raw.items():
        if code == 1 and any(v != 0 for v in metrics.values()):
            bloc_nonzero = True
            break
    if bloc_nonzero:
        fail("接口同时给出欧盟集团行和成员国行，停止以免双重计算。")
    months = {parse_day(ordered[0]).month for ordered in weeks.values()}
    if months == {6}:
        calendar = "jun"
    elif months == {9}:
        calendar = "sep"
    elif months == {10}:
        calendar = "oct"
    else:
        fail(f"代码 {commodity_code} 各营销年首周月份不一致：{sorted(months)}。先核对官方营销年，不算进度。")
    start_month = CALENDAR_START[calendar]
    for year, ordered in weeks.items():
        first = parse_day(ordered[0])
        last = parse_day(ordered[-1])
        # 进行中的营销年，末周可以仍在起始年；上限是结束年该月 14 日。
        end_cap = parse_day(f"{year}-{start_month:02d}-14")
        if not (first.year == year - 1 and first.month == start_month and first.day <= 14):
            fail(
                f"代码 {commodity_code} {label_of_api(year)} 首周 {ordered[0]} "
                f"不在该营销年 {start_month} 月上旬"
            )
        if last > end_cap or last < first:
            fail(f"代码 {commodity_code} {label_of_api(year)} 末周 {ordered[-1]} 晚于 {end_cap}")
    return raw, weeks, present_codes, old_vietnam, calendar


def sum_codes(raw, year, day, codes) -> dict:
    acc = {f: 0 for f in API_FIELDS}
    for code in codes:
        rec = raw.get((year, day, code))
        if rec is None:
            continue
        for f in API_FIELDS:
            acc[f] += rec[f]
    acc["grossMinusNet"] = acc["grossNewSales"] - acc["currentMYNetSales"]
    return acc


def subtract(a: dict, b: dict) -> dict:
    out = {f: a[f] - b[f] for f in list(API_FIELDS) + ["grossMinusNet"]}
    return out


def add(a: dict, b: dict) -> dict:
    return {f: a[f] + b[f] for f in list(API_FIELDS) + ["grossMinusNet"]}


def build_series(raw, weeks, present_codes, old_vietnam):
    vn_codes = (5520, VIETNAM_OLD) if old_vietnam else (5520,)
    code_map = dict(ONE)
    code_map["vn"] = vn_codes
    code_map["eu"] = EU_CODES
    code_map["eupuk"] = EU_CODES + (UK_CODE,)
    code_map["ea"] = code_map["cn"] + code_map["jp"] + code_map["kr"] + code_map["tw"]
    code_map["sea"] = code_map["idn"] + code_map["th"] + code_map["vn"] + code_map["bd"]
    code_map["ca"] = code_map["pk"] + code_map["tr"]
    code_map["af"] = code_map["eg"] + code_map["dz"]
    all_codes = tuple(sorted(present_codes))
    ids = []
    for _g, _n, i, _a in ROWS:
        if i not in ids:
            ids.append(i)
    values = {i: {} for i in ids}
    for year, days in weeks.items():
        lab = label_of_api(year)
        for i in ids:
            values[i][lab] = []
        for day in days:
            parts = {i: sum_codes(raw, year, day, code_map[i]) for i in code_map}
            total = sum_codes(raw, year, day, all_codes)
            parts["total"] = total
            parts["excn"] = subtract(total, parts["cn"])
            # 恒等
            for f in list(API_FIELDS) + ["grossMinusNet"]:
                if parts["excn"][f] + parts["cn"][f] != total[f]:
                    fail(f"{lab} {day} 除中国外 + 中国 ≠ Total")
                if parts["ea"][f] != parts["cn"][f] + parts["jp"][f] + parts["kr"][f] + parts["tw"][f]:
                    fail(f"{lab} {day} 东亚加总不一致")
                if parts["eupuk"][f] != parts["eu"][f] + parts["uk"][f]:
                    fail(f"{lab} {day} 欧盟+英国加总不一致")
                if parts["sea"][f] != parts["idn"][f] + parts["th"][f] + parts["vn"][f] + parts["bd"][f]:
                    fail(f"{lab} {day} 东南亚加总不一致")
            for i in ids:
                values[i][lab].append(parts[i])
    return values, code_map, all_codes


def year_labels():
    return [label_of_api(y) for y in API_YEARS]


def week_value(values, ident, label, index, field):
    series = values[ident].get(label) or []
    if index >= len(series):
        return None
    return series[index][field]


def progress_tenths(tons, target):
    if target is None or tons is None:
        return None
    return tenths_div(int(tons) * 1000, int(target))


def delta_wan_tenths(cur, prev):
    if cur is None or prev is None:
        return None
    return wan_tenths(int(cur) - int(prev))


def delta_wan_vs_mean(cur, prevs):
    if cur is None or len(prevs) != 5 or any(p is None for p in prevs):
        return None
    # round((cur - mean)/10000, 1) = round((5*cur - sum)/5000)
    return tenths_div(5 * int(cur) - sum(int(p) for p in prevs), 5000)


def delta_pp(cur_tons, cur_tgt, prev_tons, prev_tgt):
    if cur_tons is None or prev_tons is None or not cur_tgt or not prev_tgt:
        return None
    cur = Decimal(int(cur_tons)) / Decimal(int(cur_tgt)) * 100
    prev = Decimal(int(prev_tons)) / Decimal(int(prev_tgt)) * 100
    return pp_tenths(cur - prev)


def delta_pp_mean(cur_tons, cur_tgt, pairs):
    if cur_tons is None or not cur_tgt or len(pairs) != 5:
        return None
    if any(t is None or not tgt for t, tgt in pairs):
        return None
    cur = Decimal(int(cur_tons)) / Decimal(int(cur_tgt)) * 100
    avg = sum(Decimal(int(t)) / Decimal(int(tgt)) * 100 for t, tgt in pairs) / Decimal(5)
    return pp_tenths(cur - avg)


def prior_labels(label: str, n: int):
    labels = year_labels()
    i = labels.index(label)
    if i < n:
        return None
    return labels[i - n : i]


def compare_at(values, targets, ident, label, index, field, kind):
    """kind cur/next selects which target a progress field uses. field is tons field or grossMinusNet."""
    cur = week_value(values, ident, label, index, field)
    prev_lab_list = prior_labels(label, 1)
    prev = None
    if prev_lab_list:
        prev = week_value(values, ident, prev_lab_list[0], index, field)
    five = prior_labels(label, 5)
    prev5 = None
    if five:
        prev5 = [week_value(values, ident, lab, index, field) for lab in five]
    if kind == "tons":
        return wan_tenths(cur) if cur is not None else None, delta_wan_tenths(cur, prev), (
            delta_wan_vs_mean(cur, prev5) if prev5 is not None else None
        )
    # progress
    tgt_key = label if kind == "cur" else next_my_label(label)
    tgt = targets.get(tgt_key)
    level = progress_tenths(cur, tgt)
    yoy = None
    y5 = None
    if prev_lab_list:
        pl = prev_lab_list[0]
        pt = targets.get(pl if kind == "cur" else next_my_label(pl))
        yoy = delta_pp(cur, tgt, prev, pt)
    if five:
        pairs = []
        for lab in five:
            tt = targets.get(lab if kind == "cur" else next_my_label(lab))
            pairs.append((week_value(values, ident, lab, index, field), tt))
        y5 = delta_pp_mean(cur, tgt, pairs)
    return level, yoy, y5


def series_payload(values, targets, weeks):
    labels = [label_of_api(y) for y in sorted(weeks)]
    week_map = {label_of_api(y): weeks[y] for y in weeks}
    ids = []
    for _g, _n, i, _a in ROWS:
        if i not in ids:
            ids.append(i)
    series = {}
    for ident in ids:
        series[ident] = {}
        for key, _en, _zh, src in METRICS:
            series[ident][key] = {}
            for lab in labels:
                arr = []
                for i in range(len(values[ident][lab])):
                    tons = values[ident][lab][i]["grossMinusNet" if src is None else src]
                    arr.append(wan_tenths(tons))
                series[ident][key][lab] = arr
        for key, _en, _zh, src, kind in PROGRESS:
            series[ident][key] = {}
            for lab in labels:
                tgt = targets.get(lab if kind == "cur" else next_my_label(lab))
                arr = []
                for i in range(len(values[ident][lab])):
                    tons = values[ident][lab][i][src]
                    arr.append(progress_tenths(tons, tgt) if tgt else None)
                series[ident][key][lab] = arr
    latest_year = labels[-1]
    latest_idx = len(week_map[latest_year]) - 1
    latest_date = week_map[latest_year][latest_idx]
    deltas = {}
    for ident in ids:
        deltas[ident] = {}
        for key, _en, _zh, src in METRICS:
            field = "grossMinusNet" if src is None else src
            _lv, yoy, y5 = compare_at(values, targets, ident, latest_year, latest_idx, field, "tons")
            deltas[ident][key] = {"yoy": yoy, "y5": y5}
        for key, _en, _zh, src, kind in PROGRESS:
            _lv, yoy, y5 = compare_at(values, targets, ident, latest_year, latest_idx, src, kind)
            deltas[ident][key] = {"yoy": yoy, "y5": y5}
    return {
        "years": labels,
        "weeks": week_map,
        "latest": {
            "year": latest_year,
            "weekIndex": latest_idx + 1,
            "date": latest_date,
        },
        "rows": [
            {"group": g, "name": n, "id": i, "agg": a} for g, n, i, a in ROWS
        ],
        "series": series,
        "deltas": deltas,
        "targetsWan": {k: wan_tenths(v) for k, v in targets.items()},
    }


def apply_header(ws, headers, row):
    for col, text in enumerate(headers, 1):
        cell = ws.cell(row, col, text)
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
        cell.alignment = WRAP
        cell.border = THIN
    ws.row_dimensions[row].height = 36
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{row}"
    ws.freeze_panes = f"A{row + 1}"
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{ws.max_row or row}"


def _put_delta(cell, tenths, font):
    cell.border = THIN
    cell.alignment = Alignment(vertical="center", horizontal="right")
    if tenths is None:
        cell.font = font
        return
    cell.value = int(tenths) / 10
    cell.number_format = "0.0"
    if tenths > 0:
        cell.font = FONT_POS
    elif tenths < 0:
        cell.font = FONT_NEG
    else:
        cell.font = FONT_ZERO


def write_snapshot_sheet(wb, payload, release_day):
    ws = wb.active
    ws.title = "本周快照"
    latest = payload["latest"]
    note = (
        f"{payload.get('zh', '')}（{payload.get('code', '')}）截至 {latest['date']}　"
        f"{latest['year']} 第 {latest['weekIndex']} 周"
        f"　发布 {release_day or '接口未给发布时点'}　数量为万吨，进度为百分数，均 1 位小数"
    )
    ws.cell(1, 1, note).font = Font(name=CN_FONT, size=11, bold=True)
    ws.merge_cells("A1:E1")
    headers = ["类别", "本年度净销售", "本年度总销售", "装运进度", "本年度销售进度"]
    for col, text in enumerate(headers, 1):
        cell = ws.cell(3, col, text)
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
        cell.alignment = WRAP
        cell.border = THIN
    lab = latest["year"]
    wi = latest["weekIndex"] - 1
    for r_i, (name, ident) in enumerate(SHOT_ROWS):
        excel_row = 4 + r_i
        cell = ws.cell(excel_row, 1, name)
        cell.font = FONT_AGG
        cell.border = THIN
        cell.alignment = LEFT
        for col, (key, _label, _unit) in enumerate(SHOT_FIELDS, 2):
            val = payload["series"][ident][key][lab][wi]
            c = ws.cell(excel_row, col, None if val is None else int(val) / 10)
            c.font = FONT_BASE
            c.border = THIN
            c.alignment = Alignment(vertical="center", horizontal="right")
            if val is not None:
                c.number_format = "0.0"
    ws.column_dimensions["A"].width = 14
    for col in ("B", "C", "D", "E"):
        ws.column_dimensions[col].width = 18
    ws.row_dimensions[3].height = 22
    ws.oddFooter.center.text = "shengbanghengchang_carina"


def write_metric_sheet(wb, payload, key, zh, en, definition):
    ws = wb.create_sheet(zh)
    latest = payload["latest"]
    lab = latest["year"]
    wi = latest["weekIndex"] - 1
    ws.cell(1, 1, f"{zh} {en}").font = Font(name=CN_FONT, size=14, bold=True)
    ws.merge_cells("A1:E1")
    if key == "currentMYNetSales":
        note = (
            f"{definition} 截至 {latest['date']}，{lab} 第 {latest['weekIndex']} 周。"
            "单位万吨，1 位小数。正数红色，负数绿色。不含 Total 和未知目的地。"
        )
    else:
        note = (
            f"{definition} 截至 {latest['date']}，{lab} 第 {latest['weekIndex']} 周。"
            f"同比去年是与上一营销年同一周序之差，都是该市场年的第 {latest['weekIndex']} 周。"
            "同比5年均值是当前值减去此前五个营销年同一周序的算术平均，不含当年；不满五个营销年留空。"
            "单位万吨，1 位小数。正数红色，负数绿色。不含 Total 和未知目的地。"
        )
    ws.cell(2, 1, note).font = Font(name=CN_FONT, size=10)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 48
    headers = ["地区", "目的地", zh, "同比去年", "同比5年均值"]
    for col, text in enumerate(headers, 1):
        cell = ws.cell(4, col, text)
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
        cell.alignment = Alignment(vertical="center", horizontal="center")
        cell.border = THIN
    start = 5
    i = 0
    rows = TABLE_ROWS
    while i < len(rows):
        group = rows[i][0]
        j = i + 1
        while j < len(rows) and rows[j][0] == group:
            j += 1
        span = j - i
        top = start + i
        if span > 1:
            ws.merge_cells(start_row=top, start_column=1, end_row=top + span - 1, end_column=1)
        region = ws.cell(top, 1, group)
        region.font = FONT_AGG
        region.alignment = Alignment(vertical="center", horizontal="center")
        region.fill = FILL_AGG
        for k in range(i, j):
            excel_row = start + k
            _g, name, ident = rows[k]
            level = payload["series"][ident][key][lab][wi]
            yoy = payload["deltas"][ident][key]["yoy"]
            y5 = payload["deltas"][ident][key]["y5"]
            rc = ws.cell(excel_row, 1)
            rc.border = THIN
            rc.fill = FILL_AGG
            dest = ws.cell(excel_row, 2, name)
            dest.font = FONT_BASE
            dest.border = THIN
            dest.alignment = LEFT
            lv = ws.cell(excel_row, 3, None if level is None else int(level) / 10)
            lv.font = FONT_BASE
            lv.border = THIN
            lv.alignment = Alignment(vertical="center", horizontal="right")
            if level is not None:
                lv.number_format = "0.0"
            _put_delta(ws.cell(excel_row, 4), yoy, FONT_BASE)
            _put_delta(ws.cell(excel_row, 5), y5, FONT_BASE)
            ws.row_dimensions[excel_row].height = 18
        i = j
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = "shengbanghengchang_carina"
    ws.auto_filter.ref = f"A4:E{4 + len(TABLE_ROWS)}"



def write_detail_sheet(wb, values, targets, weeks):
    ws = wb.create_sheet("周度明细")
    headers = ["营销年", "周序", "周末", "分组", "目的地"]
    fields = []
    for key, en, zh, src in METRICS:
        headers.append(f"{en}_{zh}_吨")
        fields.append((key, src, "ton"))
    for key, en, zh, src in METRICS:
        headers.append(f"{en}_{zh}_万吨")
        fields.append((key, src, "wan"))
    for key, en, zh, _src, _kind in PROGRESS:
        headers.append(f"{en}_{zh}_%")
        fields.append((key, None, "pct"))
    ws.append(headers)
    for cell in ws[1]:
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
        cell.alignment = WRAP
    ws.row_dimensions[1].height = 32
    labels_weeks = [(label_of_api(y), weeks[y], y) for y in API_YEARS if y in weeks]
    for lab, days, _year in labels_weeks:
        for idx in range(len(days)):
            for group, name, ident, _agg in ROWS:
                rec = values[ident][lab][idx]
                row = [lab, idx + 1, days[idx], group, name]
                ton_vals = []
                for _key, src, kind in fields:
                    if kind != "ton":
                        continue
                    field = "grossMinusNet" if src is None else src
                    ton_vals.append(rec[field])
                row.extend(ton_vals)
                wan_vals = [wan_tenths(v) / 10 for v in ton_vals]
                row.extend(wan_vals)
                for _key, _en, _zh, src, kind in PROGRESS:
                    tgt = targets.get(lab if kind == "cur" else next_my_label(lab))
                    tenths = progress_tenths(rec[src], tgt) if tgt else None
                    row.append(None if tenths is None else tenths / 10)
                ws.append(row)
    # 万吨与进度列格式
    n_metrics = len(METRICS)
    wan_start = 6 + n_metrics
    wan_end = wan_start + n_metrics - 1
    pct_end = wan_end + len(PROGRESS)
    for row in ws.iter_rows(min_row=2, min_col=wan_start, max_col=pct_end):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0.0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    for i, w in enumerate([10, 8, 14, 10, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def write_code_sheet(wb, names, code_map, old_vietnam, present_codes):
    ws = wb.create_sheet("目的地代码")
    ws.append(["id", "页面名称", "组成", "countryCode", "FAS 英文名", "说明"])
    for cell in ws[1]:
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
    desc = {
        "total": "该周全部目的地行之和，含 Unknown，含中国",
        "excn": "Total − 中国（5700），不含中国台湾和香港",
        "ea": "中国+日本+韩国+中国台湾",
        "sea": "印尼+泰国+越南+孟加拉",
        "ca": "巴基斯坦+土耳其。按用户分组，不是 FAS 的地区名",
        "eu": "欧盟成员国目的地之和，不含英国。比卢为一个代码。克罗地亚 4791 并入",
        "eupuk": "欧盟 + 英国",
        "af": "埃及+阿尔及利亚",
        "unk": "FAS UNKNOWN，已包含在 Total 里，这里单列",
    }
    name_by_id = {}
    for _g, n, i, _a in ROWS:
        name_by_id.setdefault(i, n)
    for ident, codes in code_map.items():
        for code in codes:
            ws.append([
                ident,
                name_by_id.get(ident, ident),
                "、".join(name_by_id.get(ident, ident) for _ in [0]) if False else desc.get(ident, "单一目的地"),
                code,
                names.get(code, ""),
                "样本内有记录" if code in present_codes else "样本内无记录，按 0",
            ])
    ws.append(["total", "Total", desc["total"], "—", "全部 countryCode", "不把地区行再加总"])
    ws.append(["excn", "除中国外", desc["excn"], "—", "Total − 5700", ""])
    if old_vietnam:
        ws.append(["vn", "越南", "5520+5510", VIETNAM_OLD, names.get(VIETNAM_OLD, ""), "旧码并入越南"])
    ws.append(["—", "不计入欧盟", "英国虽在 FAS regionId=1，但按用户要求单列", UK_CODE, names.get(UK_CODE, ""), ""])
    ws.append(["—", "不使用", "EUROPEAN UNION - 27 集团行。样本内无非零记录，故用成员国相加", 1, names.get(1, ""), "避免与成员国双重计算"])
    for col, w in enumerate([12, 16, 42, 14, 36, 28], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    for row in ws.iter_rows(min_row=2, max_col=6):
        for cell in row:
            cell.font = FONT_BASE
            cell.alignment = Alignment(wrap_text=True, vertical="center")


def write_target_sheet(wb, targets):
    ws = wb.create_sheet("出口目标")
    ws.append(["营销年", "累计出口目标_吨", "万吨", "用途"])
    for cell in ws[1]:
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
    labels = []
    # 保持工作簿原列顺序：把 targets 的键按起始年排序，包括 16/17
    def sort_key(lab):
        return int(lab.split("/")[0])
    for lab in sorted(targets, key=sort_key):
        labels.append(lab)
        start = int(lab.split("/")[0])
        if start < 18:
            use = "样本自 18/19 起，此列不参与页面"
        elif lab == "26/27":
            use = "26/27 本年度目标；25/26 的下年度目标。27/28 没有，故 26/27 下年度销售进度留空"
        else:
            use = "本年度目标，也作为上一营销年的下年度目标"
        ws.append([lab, targets[lab], wan_tenths(targets[lab]) / 10, use])
    ws.append(["27/28", None, None, "出口目标.xlsx 没有这一列，下年度销售进度留空"])
    for row in ws.iter_rows(min_row=2, max_col=4):
        for cell in row:
            cell.font = FONT_BASE
            cell.alignment = LEFT
        if isinstance(row[1].value, int):
            row[1].number_format = "#,##0"
        if isinstance(row[2].value, float):
            row[2].number_format = "0.0"
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 78
    note = ws.cell(ws.max_row + 2, 1, "来源为同目录 出口目标.xlsx，本脚本不改那个文件。")
    note.font = FONT_BASE


def _cls(tenths):
    if tenths is None:
        return ""
    if tenths > 0:
        return "pos"
    if tenths < 0:
        return "neg"
    return "zero"


# 截图区里只有 Total 有出口目标。中国、除中国外不除总目标。
SHOT_PROGRESS_TOTAL_ONLY = frozenset({"shipProg", "saleProg"})


def shot_block(payload) -> str:
    latest = payload["latest"]
    notes = {
        "currentMYNetSales": "currentMYNetSales",
        "currentMYTotalCommitment": "currentMYTotalCommitment",
        "shipProg": "累计装运 Total / 出口目标",
        "saleProg": "本年度总销售 Total / 出口目标",
    }
    heads = ['<div class="shot-corner"></div>']
    for key, label, unit in SHOT_FIELDS:
        heads.append(
            f'<div class="shot-colhead">{label}<small>{notes[key]} · {unit}</small></div>'
        )
    body = ["".join(heads)]
    for name, ident in SHOT_ROWS:
        row = [f'<div class="shot-rowhead">{name}</div>']
        for key, _label, unit in SHOT_FIELDS:
            if key in SHOT_PROGRESS_TOTAL_ONLY and ident != "total":
                row.append('<div class="shot-na"></div>')
            else:
                row.append(
                    f'<figure class="shot-cell" data-id="{ident}" data-key="{key}" data-unit="{unit}">'
                    "<canvas></canvas></figure>"
                )
        body.append("".join(row))
    pid = payload["id"]
    return (
        '<p class="kicker">截图区</p>'
        f'<section id="shot-{pid}" class="shot">'
        f'<div class="shot-top"><h2>{payload["zh"]}周度出口销售</h2>'
        f'<span>{latest["year"]} 第 {latest["weekIndex"]} 周 · 截至 {latest["date"]}</span></div>'
        f'<div class="legend" id="shot-legend-{pid}"></div>'
        '<div class="shot-grid">' + "".join(body) + "</div></section>"
    )


def tables_block(payload) -> str:
    latest = payload["latest"]
    lab = latest["year"]
    wi = latest["weekIndex"] - 1
    parts = []
    for key, zh, en, definition in TABLE_METRICS:
        body = []
        i = 0
        while i < len(TABLE_ROWS):
            group = TABLE_ROWS[i][0]
            j = i + 1
            while j < len(TABLE_ROWS) and TABLE_ROWS[j][0] == group:
                j += 1
            span = j - i
            for k in range(i, j):
                _g, name, ident = TABLE_ROWS[k]
                level = payload["series"][ident][key][lab][wi]
                yoy = payload["deltas"][ident][key]["yoy"]
                y5 = payload["deltas"][ident][key]["y5"]
                region = f'<td class="region" rowspan="{span}">{group}</td>' if k == i else ""
                if key == "currentMYNetSales":
                    level_txt = "数据缺失" if level is None else fmt_tenths(level)
                    yoy_txt = "" if yoy is None else fmt_tenths(yoy)
                    y5_txt = "" if y5 is None else fmt_tenths(y5)
                    lv_cls = _cls(level)
                    lv_td = f'<td class="{lv_cls}">{level_txt}</td>' if lv_cls else f"<td>{level_txt}</td>"
                else:
                    level_txt = fmt_tenths(level)
                    yoy_txt = fmt_tenths(yoy)
                    y5_txt = fmt_tenths(y5)
                    lv_td = f"<td>{level_txt}</td>"
                body.append(
                    "<tr>"
                    + region
                    + f'<td class="l">{name}</td>'
                    + lv_td
                    + f'<td class="{_cls(yoy)}">{yoy_txt}</td>'
                    + f'<td class="{_cls(y5)}">{y5_txt}</td></tr>'
                )
            i = j
        if key == "currentMYNetSales":
            defn = (
                f"{definition} 截至 {latest['date']}，{lab} 第 {latest['weekIndex']} 周。"
                "单位万吨，1 位小数。正数红色，负数绿色。不含 Total 和未知目的地。"
            )
        else:
            defn = (
                f"{definition} 截至 {latest['date']}，{lab} 第 {latest['weekIndex']} 周。"
                f"同比去年是与上一营销年同一周序之差，都是该市场年的第 {latest['weekIndex']} 周。"
                "同比5年均值是当前值减去此前五个营销年同一周序的算术平均，不含当年；不满五个营销年留空。"
                "单位万吨，1 位小数。正数红色，负数绿色。本表不含 Total 和未知目的地。"
            )
        parts.append(
            f'<section class="block" id="t-{payload["id"]}-{key}">'
            f"<h2>{zh} {en}</h2>"
            f'<p class="def">{defn}</p>'
            '<table class="metric"><thead><tr>'
            '<th class="l">地区</th><th class="l">目的地</th>'
            f"<th>{zh} {en}</th><th>同比去年</th><th>同比5年均值</th>"
            "</tr></thead><tbody>"
            + "".join(body)
            + "</tbody></table></section>"
        )
    return "\n".join(parts)


def target_sentence(latest_year: str, targets: dict) -> str:
    cur = targets.get(latest_year)
    nxt_lab = next_my_label(latest_year)
    nxt = targets.get(nxt_lab)
    cur_s = f"{fmt_tenths(wan_tenths(cur))} 万吨" if cur else "没有这一列，进度留空"
    if nxt is None:
        nxt_s = f"{nxt_lab} 不在出口目标表里，下年度销售进度留空"
    else:
        nxt_s = f"{nxt_lab} 为 {fmt_tenths(wan_tenths(nxt))} 万吨"
    return f"本年度出口目标取「{latest_year}年度」列，{cur_s}。下年度取 {nxt_s}。"


def foot_text(payload) -> str:
    axis = {"jun": "6 月", "sep": "9 月", "oct": "10 月"}[payload["calendar"]]
    return (
        f"横轴从{axis}的第 1 个报告周起，按各营销年的周序对齐。月份刻度用周数最完整的一年。"
        "颜色从这张图的最新营销年往回数：红、紫、黄、绿、蓝，更早的年份为浅色。"
        "同比 5 年均值是当前值减去此前五个营销年同周序的平均值；不满五个营销年就留空。"
        f"{payload['targetNote']}"
        "万吨按行四舍五入，分行相加可能与加总行相差 0.1。"
    )


def other_block(payload) -> str:
    pid = payload["id"]
    return (
        '<h2 class="sec">其他指标</h2>'
        '<p class="def">截图区：净销售和总销售是 Total、中国、除中国外。装运进度和销售进度只画 Total，中国和除中国外注明无分项目标。这里可再按目的地看其余指标，横轴从该营销年第 1 周起，最新一周标在端点旁。</p>'
        '<div class="controls">'
        f'<label for="dest-{pid}">目的地</label>'
        f'<select id="dest-{pid}"></select>'
        "</div>"
        f'<div class="legend" id="legend-{pid}"></div>'
        f'<div class="grid" id="grid-{pid}"></div>'
    )


def panel_html(payload) -> str:
    latest = payload["latest"]
    pid = payload["id"]
    on = " on" if pid == "soy" else ""
    sub = (
        f'{payload["zh"]} commodityCode {payload["code"]} {payload["en"]}，公吨。'
        f'{payload["myNote"]} 最新一周 {latest["date"]}（{latest["year"]} 第 {latest["weekIndex"]} 周）。'
        f'发布 {payload.get("releaseDay") or "接口未返回"}。'
        f'{payload["targetNote"]} 数量为万吨，进度为百分数，均 1 位小数。'
    )
    return (
        f'<section class="panel{on}" id="panel-{pid}" data-pid="{pid}">'
        f'<p class="sub">{sub}</p>'
        + shot_block(payload)
        + other_block(payload)
        + tables_block(payload)
        + f'<p class="foot">{foot_text(payload)}</p>'
        + "</section>"
    )


def write_html(products):
    data = json.dumps(products, ensure_ascii=False).replace("<", "\\u003c")
    html = HTML_TEMPLATE
    html = html.replace("@@WM_CSS@@", html_watermark_css())
    html = html.replace("@@WM_DIV@@", html_watermark_div())
    html = html.replace("@@PRODUCTS@@", data)
    html = html.replace("@@PANELS@@", "\n".join(panel_html(p) for p in products))
    OUT_HTML.write_text(html, encoding="utf-8")


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>美豆周度出口销售</title>
<style>
  body { margin: 0; background: #f3f5f7; color: #1c1c1c;
    font-family: "Microsoft YaHei", "PingFang SC", sans-serif; }
  .wrap { position: relative; z-index: 1; max-width: 1480px; margin: 0 auto;
    padding: 18px 16px 72px; }
  h1 { font-size: 26px; margin: 0 0 8px; }
  .sub { color: #4d5966; font-size: 13px; line-height: 1.6; margin: 0; }
  .tabs { display: flex; gap: 8px; margin: 14px 0 4px; flex-wrap: wrap; }
  .tab { font: inherit; background: #fff; color: #1c1c1c; border: 1px solid #d7dee6;
    border-radius: 6px; padding: 6px 16px; cursor: pointer; }
  .tab.on { background: #1f3a5f; color: #fff; border-color: #1f3a5f; }
  .panel { display: none; }
  .panel.on { display: block; }
  .controls { display: flex; gap: 10px; align-items: center; margin: 16px 0 10px; }
  select { font: inherit; padding: 6px 8px; background: #fff; }
  .legend { display: flex; flex-wrap: wrap; gap: 10px 14px; font-size: 12px; margin: 4px 0 12px; }
  .legend b { font-weight: 600; }
  .swatch { display: inline-block; width: 22px; height: 3px; margin-right: 4px; vertical-align: middle; }
  .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .card { background: #fff; border: 1px solid #e3e6ea; border-radius: 8px; padding: 8px 10px 10px; }
  .card h2 { font-size: 14px; margin: 0; font-weight: 650; }
  .card p { margin: 2px 0 6px; color: #66707a; font-size: 12px; min-height: 16px; }
  canvas { width: 100%; height: 280px; cursor: crosshair; display: block; }
  .kicker { margin: 18px 0 6px; font-size: 13px; letter-spacing: 0.08em; color: #5c6770; }
  .sec { font-size: 18px; margin: 26px 0 4px; }
  .shot { background: #fff; border: 1px solid #e6eaee; border-radius: 8px; padding: 10px 12px 8px; }
  .shot-top { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 2px; }
  .shot-top h2 { margin: 0; font-size: 18px; }
  .shot-top span { color: #5c6770; font-size: 12px; }
  .shot .legend { margin: 2px 0 6px; gap: 8px 12px; }
  .shot-grid { display: grid; grid-template-columns: 78px repeat(4, minmax(0, 1fr)); gap: 0 6px; align-items: end; }
  .shot-colhead { text-align: center; font-size: 13px; line-height: 1.25; font-weight: 700; padding-bottom: 2px; }
  .shot-colhead small { display: block; font-weight: 500; color: #5c6770; font-size: 10px; }
  .shot-rowhead { font-size: 14px; font-weight: 700; display: flex; align-items: center; justify-content: flex-end; white-space: nowrap; padding: 0 6px 16px 0; }
  .shot-cell { margin: 0; }
  .shot canvas { width: 100%; height: 148px; cursor: default; }
  .shot-na { height: 148px; display: flex; align-items: center; justify-content: center; }
  .shot-na p { margin: 0; color: #5c6770; font-size: 13px; font-weight: 650; }
  .block { margin: 22px 0 8px; }
  .block h2 { font-size: 16px; margin: 0 0 6px; }
  .def { color: #3d4852; font-size: 13px; line-height: 1.6; margin: 0 0 8px; }
  table.metric { border-collapse: collapse; font-size: 13px; background: #fff; }
  table.metric th, table.metric td { border: 1px solid #d7dee6; padding: 5px 10px; white-space: nowrap; text-align: right; }
  table.metric th { background: #1f3a5f; color: #fff; font-weight: 600; }
  table.metric td.l, table.metric th.l { text-align: left; }
  table.metric td.region { text-align: center; vertical-align: middle; background: #f4f7fb; font-weight: 650; }
  .pos { color: #d00000; font-weight: 700; }
  .neg { color: #17803d; font-weight: 700; }
  .zero { color: #222; }
  .foot { color: #5c6770; font-size: 12px; line-height: 1.55; margin-top: 10px; }
  @media (max-width: 900px) {
    .grid { grid-template-columns: 1fr; }
  }
</style>
<style>@@WM_CSS@@</style>
</head>
<body>
@@WM_DIV@@
<div class="wrap">
  <h1>美国周度出口销售</h1>
  <p class="sub">USDA FAS ESR，公吨。美豆、美玉米、美豆粕、美豆油、美小麦同一套目的地和表。出口目标读同目录工作簿各品种「累计出口目标(t)」，列是 16/17年度到 26/27年度，没有单独的本年度、下年度列，也没有 27/28。正数红、负数绿。默认显示美豆。</p>
  <nav class="tabs" id="tabs">
    <button type="button" class="tab on" data-id="soy">美豆</button>
    <button type="button" class="tab" data-id="corn">美玉米</button>
    <button type="button" class="tab" data-id="meal">美豆粕</button>
    <button type="button" class="tab" data-id="oil">美豆油</button>
    <button type="button" class="tab" data-id="wheat">美小麦</button>
  </nav>
  @@PANELS@@
</div>
<script>
const PRODUCTS = @@PRODUCTS@@;
const CHARTS = [
  ["weeklyExports","周度装运 weeklyExports","万吨",""],
  ["accumulatedExports","累计装运 accumulatedExports","万吨",""],
  ["outstandingSales","本年度未装运 outstandingSales","万吨",""],
  ["currentGrossNewSales","本年度总新销售 currentGrossNewSales","万吨","接口字段是 grossNewSales"],
  ["currentMYNetSales","本年度净销售 currentMYNetSales","万吨",""],
  ["grossMinusNet","grossMinusNet 总新销售 − 净销售","万吨",""],
  ["currentMYTotalCommitment","本年度总销售 currentMYTotalCommitment","万吨",""],
  ["nextMYOutstandingSales","下年度未装运 nextMYOutstandingSales","万吨",""],
  ["nextMYNetSales","下年度净销售 nextMYNetSales","万吨",""],
  ["shipProg","装运进度","%", "累计装运 / 本年度出口目标"],
  ["saleProg","本年度销售进度","%", "本年度总销售 / 本年度出口目标"],
  ["nextProg","下年度销售进度","%", "下年度未装运 / 下一营销年出口目标。目标表没有下一列时留空"]
];

function t1(tenths) {
  if (tenths == null) return "—";
  var sign = tenths < 0 ? "-" : "";
  var a = Math.abs(tenths);
  return sign + Math.floor(a / 10) + "." + (a % 10);
}
function cls(tenths) {
  if (tenths == null) return "";
  if (tenths > 0) return "pos";
  if (tenths < 0) return "neg";
  return "zero";
}
function num(tenths) { return tenths == null ? null : tenths / 10; }

function styleOf(product, y) {
  return (product.yearStyle && product.yearStyle[y]) || {color:"#ccc", width:1};
}
function fillLegend(el, product) {
  if (!el) return;
  product.years.slice().reverse().forEach(function (y) {
    var s = styleOf(product, y);
    var span = document.createElement("span");
    var sw = document.createElement("i");
    sw.className = "swatch";
    sw.style.background = s.color;
    sw.style.height = Math.max(2, s.width) + "px";
    span.appendChild(sw);
    var b = document.createElement("b");
    b.textContent = y;
    b.style.color = s.color;
    span.appendChild(b);
    el.appendChild(span);
  });
}
function refWeeks(product) {
  var best = [];
  product.years.forEach(function (y) {
    var w = product.weeks[y] || [];
    if (w.length > best.length) best = w;
  });
  return best;
}
function tickMonths(product) {
  if (product.calendar === "oct") return {"10月":1, "1月":1, "4月":1, "7月":1};
  if (product.calendar === "jun") return {"6月":1, "9月":1, "12月":1, "3月":1};
  return {"9月":1, "12月":1, "3月":1, "6月":1};
}
function niceStep(span, n) {
  var step0 = span / n;
  if (!(step0 > 0)) return 1;
  var mag = Math.pow(10, Math.floor(Math.log10(step0)));
  var err = step0 / mag;
  if (err >= 7.5) return 10 * mag;
  if (err >= 3.5) return 5 * mag;
  if (err >= 1.5) return 2 * mag;
  return mag;
}
function fmtTick(v) {
  var av = Math.abs(v);
  if (av >= 100) return String(Math.round(v));
  if (av >= 10) return (Math.round(v * 10) / 10).toFixed(1);
  return (Math.round(v * 10) / 10).toFixed(1);
}
function drawChart(card, product, ident, compact) {
  var pack = product.series[ident][card.key];
  var canvas = card.canvas;
  var dpr = window.devicePixelRatio || 1;
  var w = canvas.clientWidth || (compact ? 280 : 560);
  var h = canvas.clientHeight || (compact ? 148 : 270);
  canvas.width = Math.round(w * dpr);
  canvas.height = Math.round(h * dpr);
  var ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.textBaseline = "alphabetic";
  ctx.textAlign = "left";
  ctx.clearRect(0, 0, w, h);
  var maxWeek = 0;
  product.years.forEach(function (y) { maxWeek = Math.max(maxWeek, (pack[y] || []).length); });
  var lo = 0, hi = 0, any = false;
  product.years.forEach(function (y) {
    (pack[y] || []).forEach(function (t) {
      if (t == null) return;
      any = true;
      var v = t / 10;
      if (v < lo) lo = v;
      if (v > hi) hi = v;
    });
  });
  var percent = card.unit === "%";
  var padL = compact ? 42 : 52, padR = compact ? 36 : 52, padT = compact ? 12 : 18, padB = compact ? 16 : 28;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, w, h);
  if (!any) {
    ctx.fillStyle = "#889";
    ctx.font = "12px Microsoft YaHei, sans-serif";
    ctx.fillText("这一项没有可画的数", padL, 28);
    card.geom = null;
    if (card.cap) card.cap.textContent = card.note || "";
    return;
  }
  var span = hi - lo || 1;
  var step = niceStep(span, 5);
  var y0 = Math.floor(lo / step) * step;
  var y1 = Math.ceil(hi / step) * step;
  if (y1 === y0) y1 = y0 + step;
  var plotW = w - padL - padR;
  var plotH = h - padT - padB;
  function xAt(i) { return padL + (maxWeek <= 1 ? plotW / 2 : (i / (maxWeek - 1)) * plotW); }
  function yAt(v) { return padT + (y1 - v) / (y1 - y0) * plotH; }
  ctx.strokeStyle = "#e6eaee";
  ctx.fillStyle = "#66707a";
  ctx.font = (compact ? "10px " : "11px ") + "Microsoft YaHei, sans-serif";
  ctx.lineWidth = 1;
  for (var v = y0; v <= y1 + step * 0.01; v += step) {
    var yy = yAt(v);
    ctx.beginPath();
    ctx.moveTo(padL, yy);
    ctx.lineTo(w - padR, yy);
    ctx.stroke();
    ctx.fillStyle = "#66707a";
    ctx.fillText(percent ? fmtTick(v) + "%" : fmtTick(v), 2, yy + 3);
  }
  var ref = refWeeks(product);
  var showM = tickMonths(product);
  var prevM = "";
  ctx.fillStyle = "#8b95a1";
  ref.forEach(function (iso, i) {
    var m = iso.slice(5, 7).replace(/^0/, "") + "月";
    if (m !== prevM) {
      prevM = m;
      if (compact && !showM[m]) {
        ctx.strokeStyle = "#f0f2f4";
        ctx.beginPath();
        ctx.moveTo(xAt(i), padT);
        ctx.lineTo(xAt(i), padT + plotH);
        ctx.stroke();
        return;
      }
      ctx.fillStyle = "#8b95a1";
      ctx.fillText(m, xAt(i) - 6, h - 3);
      ctx.strokeStyle = "#f0f2f4";
      ctx.beginPath();
      ctx.moveTo(xAt(i), padT);
      ctx.lineTo(xAt(i), padT + plotH);
      ctx.stroke();
    }
  });
  product.years.forEach(function (y) {
    var arr = pack[y] || [];
    var st = styleOf(product, y);
    ctx.strokeStyle = st.color;
    ctx.lineWidth = st.width;
    ctx.beginPath();
    var pen = false;
    arr.forEach(function (t, i) {
      if (t == null) { pen = false; return; }
      var x = xAt(i), yy = yAt(t / 10);
      if (!pen) { ctx.moveTo(x, yy); pen = true; }
      else ctx.lineTo(x, yy);
    });
    ctx.stroke();
    if (y === product.latest.year) {
      var last = arr.length - 1;
      if (last >= 0 && arr[last] != null) {
        var lx = xAt(last);
        var ly = yAt(arr[last] / 10);
        ctx.fillStyle = st.color;
        ctx.beginPath();
        ctx.arc(lx, ly, 3.4, 0, Math.PI * 2);
        ctx.fill();
        var label = t1(arr[last]) + (percent ? "%" : "");
        ctx.font = (compact ? "bold 11px " : "bold 12px ") + "Microsoft YaHei, sans-serif";
        ctx.textBaseline = "middle";
        var tw = ctx.measureText(label).width;
        var tx = lx + 8;
        var ty = ly - 11;
        if (tx + tw > w - 4) tx = lx - tw - 8;
        if (ty < 12) ty = ly + 12;
        ctx.lineWidth = 3;
        ctx.strokeStyle = "#ffffff";
        ctx.strokeText(label, tx, ty);
        ctx.fillStyle = st.color;
        ctx.fillText(label, tx, ty);
      }
    }
  });
  card.geom = {padL:padL, plotW:plotW, maxWeek:maxWeek, y0:y0, y1:y1, padT:padT, plotH:plotH};
  if (!card.cap) return;
  var d = product.deltas[ident][card.key];
  var lv = pack[product.latest.year][product.latest.weekIndex - 1];
  card.cap.textContent = "本周 " + t1(lv) + " " + card.unit
    + "　同比去年 " + t1(d.yoy)
    + "　较5年均值 " + t1(d.y5)
    + (card.note ? "　" + card.note : "");
}
var runtime = {};
PRODUCTS.forEach(function (product) {
  var pid = product.id;
  fillLegend(document.getElementById("shot-legend-" + pid), product);
  fillLegend(document.getElementById("legend-" + pid), product);
  var destSel = document.getElementById("dest-" + pid);
  product.rows.forEach(function (r, i) {
    var opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = r.group + " · " + r.name;
    destSel.appendChild(opt);
  });
  var cards = [];
  var grid = document.getElementById("grid-" + pid);
  CHARTS.forEach(function (c) {
    var card = document.createElement("div");
    card.className = "card";
    var h = document.createElement("h2");
    h.textContent = c[1];
    var p = document.createElement("p");
    var canvas = document.createElement("canvas");
    card.appendChild(h);
    card.appendChild(p);
    card.appendChild(canvas);
    grid.appendChild(card);
    cards.push({key:c[0], unit:c[2], note:c[3], canvas:canvas, cap:p, geom:null});
  });
  var shotCards = [];
  document.querySelectorAll("#shot-" + pid + " .shot-cell").forEach(function (fig) {
    shotCards.push({
      key: fig.getAttribute("data-key"),
      unit: fig.getAttribute("data-unit"),
      ident: fig.getAttribute("data-id"),
      canvas: fig.querySelector("canvas"),
      cap: null,
      geom: null
    });
  });
  function drawOne(card) {
    drawChart(card, product, product.rows[Number(destSel.value)].id, false);
  }
  cards.forEach(function (card) {
    card.canvas.addEventListener("mousemove", function (ev) {
      if (!card.geom) return;
      var g = card.geom;
      var i = Math.round((ev.offsetX - g.padL) / (g.plotW / Math.max(1, g.maxWeek - 1)));
      if (i < 0 || i >= g.maxWeek) return;
      var row = product.rows[Number(destSel.value)];
      var bits = ["第" + (i + 1) + "周"];
      product.years.forEach(function (y) {
        var arr = product.series[row.id][card.key][y] || [];
        if (i < arr.length && arr[i] != null) {
          var dt = (product.weeks[y] || [])[i] || "";
          bits.push(y + " " + dt.slice(5) + " " + t1(arr[i]));
        }
      });
      card.cap.textContent = bits.join("　");
    });
    card.canvas.addEventListener("mouseleave", function () { drawOne(card); });
  });
  destSel.addEventListener("change", function () { drawProduct(pid); });
  runtime[pid] = {product:product, cards:cards, shotCards:shotCards, drawOne:drawOne};
});
function drawProduct(pid) {
  var rt = runtime[pid];
  if (!rt) return;
  rt.shotCards.forEach(function (c) { drawChart(c, rt.product, c.ident, true); });
  rt.cards.forEach(rt.drawOne);
}
function activate(pid) {
  if (!runtime[pid]) pid = "soy";
  document.querySelectorAll(".panel").forEach(function (el) {
    el.classList.toggle("on", el.getAttribute("data-pid") === pid);
  });
  document.querySelectorAll(".tab").forEach(function (el) {
    el.classList.toggle("on", el.getAttribute("data-id") === pid);
  });
  drawProduct(pid);
}
document.querySelectorAll(".tab").forEach(function (btn) {
  btn.addEventListener("click", function () { activate(btn.getAttribute("data-id")); });
});
window.addEventListener("resize", function () {
  var on = document.querySelector(".panel.on");
  if (on) drawProduct(on.getAttribute("data-pid"));
});
var startId = (location.hash || "").replace("#", "");
activate(startId);
</script>
</body>
</html>
"""


def boundary_example(raw, weeks):
    years = [y for y in API_YEARS if y in weeks]
    for a, b in zip(years, years[1:]):
        overlap = set(weeks[a]) & set(weeks[b])
        if not overlap:
            continue
        day = sorted(overlap)[0]
        old = raw.get((a, day, CHINA))
        new = raw.get((b, day, CHINA))
        if old and new:
            return (
                f"例如中国在 {day}：{label_of_api(a)} 文件累计装运 {old['accumulatedExports']} 吨，"
                f"{label_of_api(b)} 文件累计装运 {new['accumulatedExports']} 吨、"
                f"未装运 {new['outstandingSales']} 吨"
                f"（与旧文件该周 nextMYOutstandingSales {old['nextMYOutstandingSales']} 吨相同）。"
            )
    return "相邻营销年没有共享的中国边界周。"


def write_md(records):
    lines = [
        "# 美国周度出口销售口径",
        "",
        "五个品种来自 USDA FAS Export Sales Reporting，单位都是 Metric Tons。"
        "网页是 `美豆周度出口销售.html`，顶部切换美豆、美玉米、美豆粕、美豆油、美小麦，默认美豆。"
        "密钥只从环境变量 `USDA_FAS_API_KEY` 读取，不写入 html、xlsx、md 或缓存。",
        "",
        "## 出口目标",
        "",
        "`出口目标.xlsx` 有四块，顺序是美豆、美玉米、美豆粕、美豆油。每块三行：品种名、累计出口目标(t)、日期。"
        "日期列四个品种一致，都是 16/17年度到 26/27年度。表里没有名为「本年度」或「下年度」的列，也没有 27/28。"
        "表里没有美小麦块，美小麦的装运进度、销售进度留空，同比进度也留空；在表里按同样三行格式加上「美小麦」块后会自动计算。"
        "本年度出口目标取当前营销年那一列。下年度出口目标取下一列。当前营销年若是 26/27，下一列 27/28 不存在，下年度销售进度留空。"
        "16/17 与 17/18 早于样本，页面不用。脚本不改这个文件。",
        "",
        "装运进度 = accumulatedExports / 本年度出口目标。"
        "本年度销售进度 = currentMYTotalCommitment / 本年度出口目标。"
        "下年度销售进度 = nextMYOutstandingSales / 下年度出口目标。",
        "",
        "## 接口与营销年",
        "",
        "基址 `https://api.fas.usda.gov/api/esr`，路径 "
        "`/exports/commodityCode/{代码}/allCountries/marketYear/{年}`。"
        f"`{{年}}` 是营销年结束年。脚本请求 {API_YEARS[0]} 至 {API_YEARS[-1]}，上限随当前营销年自动扩展。"
        "当前营销年优先取 datareleasedates 里该品种的 marketYear；拿不到时按该品种营销年起始月份（美豆、美玉米 9 月，美豆粕、美豆油 10 月，美小麦 6 月）和运行当天日期推算。"
        "当前营销年之后的年份若返回空列表，视作尚未发布，不补数。"
        "周序按周末从早到晚，相邻周末必须相隔 7 天，否则停止。"
        "首周月份用来确认官方营销年：都在 9 月上旬则是 9 月 1 日–8 月 31 日；都在 10 月上旬则是 10 月 1 日–9 月 30 日；都在 6 月上旬则是 6 月 1 日–5 月 31 日（美小麦）。"
        "季节性图从该营销年第 1 个报告周画起。月份刻度用周数最完整的一年。",
        "颜色从这一张图的最新营销年往回数：红、紫、黄、绿、蓝，更早浅色。红色不固定在某一年。",
        "网页水印用 `excel_s/chart_watermark.py`，文字 shengbanghengchang_carina。",
        "",
        "某国某周没有记录时按 0。营销年第 1 周的未装运里通常已包含上一营销年结转的下年度销售，这是接口原数。",
        "",
        "## 字段",
        "",
        "页面表头 currentGrossNewSales（本年度总新销售）用接口字段 grossNewSales。",
        "grossMinusNet = grossNewSales − currentMYNetSales，先按吨相减再换算。",
        "其余字段与接口同名。展示单位万吨，1 位小数，一半向上进位。同比先在吨上相减再换算。"
        "万吨按行四舍五入，分行相加可能与加总行相差 0.1。",
        "",
        "## 目的地",
        "",
        "五个品种用同一套分组。Total 是该周全部目的地之和，含中国和 Unknown。除中国外 = Total − 中国 5700。",
        "东亚 = 中国 5700、日本 5880、韩国 5800、中国台湾 5830。",
        "东南亚 = 印尼 5600、泰国 5490、越南 5520、孟加拉 5380。",
        "中亚 = 巴基斯坦 5350、土耳其 4890。这是指定分组，不是 FAS 地区名。",
        "欧盟是成员国相加，不含英国 4120。比卢是代码 4230。克罗地亚 4791 在 FAS 里属 OTHER EUROPE，仍并入欧盟。",
        "集团行代码 1 若出现非零记录，脚本停止，避免和成员国加两次。",
        "欧盟+英国 = 欧盟 + 英国。非洲 = 埃及 7290 + 阿尔及利亚 7210。美洲只列墨西哥 2010。",
        "Unknown 是 9990，只在其他指标的下拉里，不进三张表。表内也不含 Total。地区行不能再加总去对 Total。",
        "",
        "## 同比",
        "",
        "同比去年是同一周序的上一营销年。同比 5 年是当前值减去此前五个营销年同一周序的算术平均，不含当年。"
        "不满五年留空，不做 3 年均值。数量差是万吨，进度差是百分点，1 位小数。正数红，负数绿，零为黑色。",
        "",
        "## 页面",
        "",
        "每个品种一页：白底截图区仍是 4 列。本年度净销售、本年度总销售是 Total、中国、除中国外三行。"
        "装运进度 = 累计装运 Total / 出口目标，本年度销售进度 = 本年度总销售 Total / 出口目标，这两列只画 Total。"
        "中国和除中国外不除总目标，格内注明无分项目标。"
        "没有表格，没有目的地下拉。进度图纵轴是百分比，最新一周有数据标签。",
        "截图区下面可以用下拉看其余指标。再下面三张表：本年度总销售、本年度净销售、下一年度未装运。"
        "本年度净销售表是最新一周的单周净销售，与季节性图最新点同一口径。列是地区、目的地、指标、同比去年、同比5年均值。",
        "",
        "## 分品种",
        "",
    ]
    for rec in records:
        spec = rec["spec"]
        payload = rec["payload"]
        latest = payload["latest"]
        lab = latest["year"]
        wi = latest["weekIndex"] - 1
        series = payload["series"]

        def shown(ident, key, _series=series, _lab=lab, _wi=wi):
            return fmt_tenths(_series[ident][key][_lab][_wi])

        years = sorted(rec["weeks"])
        first = rec["weeks"][years[0]][0]
        last = rec["weeks"][years[-1]][-1]
        vn = "样本内出现南越旧码 5510，已并入越南。" if rec["old_vietnam"] else "样本内没有南越旧码 5510。"
        span = ""
        if payload.get("marketYearStart") and payload.get("marketYearEnd"):
            span = (
                f"发布记录当前营销年 marketYear {payload.get('releaseMarketYear')}，"
                f"{payload['marketYearStart']} 至 {payload['marketYearEnd']}。"
            )
        lines.append(f"### {spec['zh']} {spec['code']} {spec['en']}")
        lines.append("")
        lines.append(
            f"{payload['myNote']} 样本首周 {first}，最新一周 {last}，"
            f"{lab} 第 {latest['weekIndex']} 周。发布 {payload.get('releaseDay') or '接口未给'}。{span}{vn}"
        )
        lines.append(payload["targetNote"])
        lines.append(
            f"Total 周度装运 {shown('total','weeklyExports')} 万吨，累计装运 {shown('total','accumulatedExports')} 万吨，"
            f"本年度净销售 {shown('total','currentMYNetSales')} 万吨，本年度总销售 {shown('total','currentMYTotalCommitment')} 万吨，"
            + (
                f"装运进度 {shown('total','shipProg')}%，本年度销售进度 {shown('total','saleProg')}%。"
                if series["total"]["shipProg"][lab][wi] is not None
                else "出口目标表里没有该品种的本年度目标，装运进度、本年度销售进度留空。"
            )
        )
        lines.append(
            f"中国周度装运 {shown('cn','weeklyExports')} 万吨，累计装运 {shown('cn','accumulatedExports')} 万吨，"
            f"本年度净销售 {shown('cn','currentMYNetSales')} 万吨，本年度总销售 {shown('cn','currentMYTotalCommitment')} 万吨。"
            "中国、除中国外无分项目标，截图区不计算这两行的装运进度和销售进度。"
        )
        lines.append(" ".join(overlap_bits(rec["weeks"])))
        lines.append("")
    lines.extend(
        [
            "## 文件",
            "",
            "- `fetch_soy_esr.py`：拉取与出表。更新一周只需要在仓库根目录运行 "
            "`python excel_s/us_weekly_export/fetch_soy_esr.py`（需设好环境变量 `USDA_FAS_API_KEY`）。",
            "- 缓存写在 `cache/`，是接口原文。不带参数运行时，每次都刷新参考表（commodities、countries、unitsOfMeasure、datareleasedates），"
            "并对每个品种重新拉取当前营销年；历史营销年已有缓存就复用，缺哪年补哪年。接口失败就退出，不用旧缓存冒充新数据。",
            "- `python excel_s/us_weekly_export/fetch_soy_esr.py --from-cache`：不联网，只用 `cache/` 里已有的文件重画出表；缺历史营销年缓存时停止。",
            "- `美豆周度出口销售.html`：五个品种合成的网页。",
            "- `美豆周度出口销售.xlsx`、`美玉米周度出口销售.xlsx`、`美豆粕周度出口销售.xlsx`、`美豆油周度出口销售.xlsx`、`美小麦周度出口销售.xlsx`。",
            "- `出口目标.xlsx`：只读。",
            "",
        ]
    )
    OUT_MD.write_text("\n".join(lines), encoding="utf-8")



def assert_no_secret(paths):
    secret = os.environ.get("USDA_FAS_API_KEY") or ""
    if not secret:
        return
    for path in paths:
        blob = path.read_bytes()
        if secret.encode("utf-8") in blob:
            path.unlink(missing_ok=True)
            fail(f"密钥出现在输出中，已删除 {path.name}")


def years_on_disk(spec) -> dict:
    out = {}
    for year in API_YEARS:
        path = cache_path(spec, year)
        if path.exists():
            out[year] = json.loads(path.read_text(encoding="utf-8"))
    return out


def load_reference_cache():
    countries = json.loads((CACHE / "countries.json").read_text(encoding="utf-8"))
    names = {c["countryCode"]: c["countryDescription"].strip() for c in countries}
    region1 = {c["countryCode"] for c in countries if c.get("regionId") == 1}
    releases = json.loads((CACHE / "datareleasedates.json").read_text(encoding="utf-8"))
    release_map = {}
    for spec in PRODUCTS:
        release_map[spec["code"]] = next(
            (r for r in releases if r.get("commodityCode") == spec["code"]), None
        )
    return names, region1, release_map


def check_china(raw, values, latest):
    api_year = 2000 + int(latest["year"].split("/")[1])
    day = latest["date"]
    got = values["cn"][latest["year"]][latest["weekIndex"] - 1]
    src = raw.get((api_year, day, CHINA))
    if src is None:
        for field in API_FIELDS:
            if got[field] != 0:
                fail(f"中国该周没有原始行，但汇总值不是 0：{field}={got[field]}")
        return
    for field in API_FIELDS:
        if got[field] != src[field]:
            fail(f"中国加总与原始行不一致：{field}")


def attach_meta(payload, spec, calendar, release, targets):
    latest = payload["latest"]
    payload["id"] = spec["id"]
    payload["zh"] = spec["zh"]
    payload["code"] = spec["code"]
    payload["en"] = spec["en"]
    payload["calendar"] = calendar
    payload["myNote"] = MY_NOTE[calendar]
    payload["yearStyle"] = seasonal_year_styles(payload["years"])
    payload["releaseDay"] = _release_day(release)
    payload["targetNote"] = target_sentence(latest["year"], targets)
    if release:
        payload["marketYearStart"] = str(release.get("marketYearStart") or "")[:10]
        payload["marketYearEnd"] = str(release.get("marketYearEnd") or "")[:10]
        payload["releaseMarketYear"] = release.get("marketYear")
    return payload


def overlap_bits(weeks) -> list[str]:
    bits = []
    for a, b in zip(sorted(weeks), sorted(weeks)[1:]):
        overlap = sorted(set(weeks[a]) & set(weeks[b]))
        if overlap:
            bits.append(
                f"{label_of_api(a)} 与 {label_of_api(b)} 共用周末 {overlap[0]}，按两个文件分计。"
            )
        else:
            bits.append(
                f"{label_of_api(a)} 末周 {weeks[a][-1]}，{label_of_api(b)} 首周 {weeks[b][0]}，中间不重复。"
            )
    return bits


def main():
    from_cache = "--from-cache" in sys.argv
    all_targets = load_targets()
    print("出口目标四品种年份列一致：16/17 至 26/27。没有本年度/下年度专列，没有 27/28。")
    if from_cache:
        print("--from-cache：只用 cache/ 里的文件重画，不联网")
        names, region1, release_map = load_reference_cache()
    else:
        key = load_key()
        session = build_session(key)
        print("ESR 参考表（每次刷新）")
        names, region1, release_map = fetch_reference(session)
    current = {}
    current_how = {}
    for spec in PRODUCTS:
        my, how = current_market_year(spec, release_map.get(spec["code"]))
        current[spec["id"]] = my
        current_how[spec["id"]] = how
        print(f"{spec['zh']} 当前营销年 marketYear {my}（{label_of_api(my)}），依据：{how}")
    extend_api_years(max(current.values()))
    cached = {}
    for spec in PRODUCTS:
        got = years_on_disk(spec)
        cur = current[spec["id"]]
        history = [year for year in API_YEARS if year < cur]
        if from_cache:
            missing = [year for year in history if year not in got]
            if missing:
                fail(
                    f"{spec['zh']} 缺少历史营销年缓存 {missing}，不能离线出表。"
                    "请设置 USDA_FAS_API_KEY 后不带 --from-cache 运行。"
                )
            if cur not in got:
                print(f"{spec['zh']} 缓存里没有当前营销年 {cur}，只画到已有年份")
        else:
            missing = [year for year in history if year not in got]
            if missing:
                print(f"ESR {spec['zh']} 补拉历史营销年 {missing}")
                got.update(fetch_years(session, spec, missing))
            elif history:
                print(f"ESR {spec['zh']} 历史营销年 {history[0]}–{history[-1]} 复用缓存")
            fresh_years = [year for year in API_YEARS if year >= cur]
            # 当前营销年及其后已列入的年份：丢掉旧缓存，每次重拉。
            for year in fresh_years:
                got.pop(year, None)
            print(f"ESR {spec['zh']} 重拉当前营销年 {cur}" + (
                f"，以及其后 {fresh_years[1:]}（空则视作未发布）" if len(fresh_years) > 1 else ""
            ))
            # 当前营销年若来自 datareleasedates，接口必须有数；按日期推算时允许刚开年尚无数据。
            from_release = current_how[spec["id"]] == "datareleasedates.marketYear"
            allow_empty = [year for year in fresh_years if year > cur or not from_release]
            tag = {cur: "  ← 当前营销年，本次重拉"}
            got.update(fetch_years(session, spec, fresh_years, allow_empty_years=allow_empty, tag=tag))
            if any(year not in got for year in history):
                fail(f"{spec['zh']} 仍缺历史营销年数据，停止，不补数。")
        cached[spec["id"]] = got
    products = []
    records = []
    paths = [OUT_HTML, OUT_MD, Path(__file__)]
    for spec in PRODUCTS:
        targets = all_targets[spec["id"]]
        raw_years = cached[spec["id"]]
        raw, weeks, present_codes, old_vietnam, calendar = index_rows(
            raw_years, names, region1, spec["code"]
        )
        if CALENDAR_START[calendar] != spec["start_month"]:
            fail(
                f"{spec['zh']} 数据识别出的营销年口径 {calendar} 与设定起始月 {spec['start_month']} 不一致，"
                "先核对官方营销年。"
            )
        values, code_map, _all = build_series(raw, weeks, present_codes, old_vietnam)
        payload = series_payload(values, targets, weeks)
        release = release_map.get(spec["code"])
        attach_meta(payload, spec, calendar, release, targets)
        check_china(raw, values, payload["latest"])
        wb = Workbook()
        write_snapshot_sheet(wb, payload, payload.get("releaseDay"))
        for key, zh, en, definition in TABLE_METRICS:
            write_metric_sheet(wb, payload, key, zh, en, definition)
        write_detail_sheet(wb, values, targets, weeks)
        write_code_sheet(wb, names, code_map, old_vietnam, present_codes)
        write_target_sheet(wb, targets)
        out_xlsx = ROOT / spec["xlsx"]
        wb.save(out_xlsx)
        paths.append(out_xlsx)
        products.append(payload)
        records.append(
            {"spec": spec, "payload": payload, "weeks": weeks, "old_vietnam": old_vietnam}
        )
        latest = payload["latest"]
        week_tons = payload["series"]["total"]["weeklyExports"][latest["year"]][latest["weekIndex"] - 1]
        cn_tons = payload["series"]["cn"]["weeklyExports"][latest["year"]][latest["weekIndex"] - 1]
        print(
            f"{spec['zh']} 最新 {latest['date']} {latest['year']} 第{latest['weekIndex']}周 "
            f"Total周度 {fmt_tenths(week_tons)} 万吨 中国周度 {fmt_tenths(cn_tons)} 万吨 "
            f"营销年口径 {calendar}"
        )
    write_html(products)
    write_md(records)
    assert_no_secret(paths)
    print(OUT_HTML)
    for spec in PRODUCTS:
        print(ROOT / spec["xlsx"])
    print(OUT_MD)


if __name__ == "__main__":
    main()
