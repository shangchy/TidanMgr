"""字段定义、表头、地域字典与纯文本工具（无 Qt 依赖）。"""
from __future__ import annotations

import re
from typing import Any

FIELDS = [
    "task_name",
    "type_code",
    "operator_code",
    "industry_code",
    "allow_print_url",
    "url",
    "quantity",
    "duration",
    "age_max",
    "age_min",
    "pv",
    "province",
    "exclude_province",
    "city",
    "exclude_city",
]
DISPLAY_FIELDS = ["checked"] + FIELDS + ["created_at", "print_count", "last_printed_at", "action"]
HEADERS = {
    "checked": "☐",
    "task_name": "任务名",
    "type_code": "类型",
    "operator_code": "运营商",
    "industry_code": "行业编码",
    "allow_print_url": "允许",
    "url": "URL",
    "quantity": "数量",
    "duration": "时长",
    "age_max": "年龄上限",
    "age_min": "年龄下限",
    "pv": "pv",
    "province": "省份",
    "exclude_province": "排除省份",
    "city": "地市",
    "exclude_city": "排除地市",
    "print_count": "打印次数",
    "last_printed_at": "打印时间",
    "created_at": "提单时间",
    "action": "操作",
}


def coerce_allow_print_url(raw: Any) -> bool:
    """是否允许导出 URL；缺省为 True（兼容旧数据）。"""
    if raw is None or raw == "":
        return True
    if isinstance(raw, bool):
        return raw
    s = str(raw).strip().lower()
    if s in ("0", "false", "no", "否"):
        return False
    if s in ("1", "true", "yes", "是"):
        return True
    return True


FROZEN_COLUMNS = 2
# 「允许」列在右侧滚动表中的列索引；逻辑列号 = 滚动列 + 冻结列数
ALLOW_PRINT_SCROLL_COL = FIELDS.index("allow_print_url") - 1
ALLOW_PRINT_URL_DISPLAY_COL = ALLOW_PRINT_SCROLL_COL + FROZEN_COLUMNS
MAIN_SCROLL_COLUMNS = len(DISPLAY_FIELDS) - FROZEN_COLUMNS
# 历史滚动区：与主表相同字段列 + 打印次数/时间 + 提单时间 + 删除时间，不含「操作」列
HISTORY_SCROLL_FIELDS = FIELDS[1:] + ["created_at", "print_count", "last_printed_at", "deleted_at"]
HISTORY_SCROLL_COLUMNS = len(HISTORY_SCROLL_FIELDS)

PRINT_LOG_DATA_FIELDS = [
    "filename",
    "printed_at",
    "row_count",
    "include_print_url",
    "file_exists",
    "path",
]
PRINT_LOG_HEADERS = {
    "filename": "文件名",
    "printed_at": "打印时间",
    "row_count": "数据条数",
    "include_print_url": "打印URL列",
    "file_exists": "文件状态",
    "path": "保存路径",
}
PRINT_LOG_COL_COUNT = 1 + len(PRINT_LOG_DATA_FIELDS) + 1

TYPE_MAP = {"DB": "dpi-白", "DJ": "106", "XC": "小程序", "DH": "dpi-灰", "DY": "抖音"}
OP_MAP = {"YD": "移动", "LT": "联通", "DX": "电信", "YX": "移动|电信"}
DURATIONS = ["一天", "本周", "长期"]
PROVINCES = [
    "北京", "天津", "上海", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江",
    "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南",
    "广东", "海南", "四川", "贵州", "云南", "陕西", "甘肃", "青海", "台湾",
    "内蒙古", "广西", "西藏", "宁夏", "新疆", "香港", "澳门",
]
CITIES = [
    "石家庄", "唐山", "秦皇岛", "邯郸", "邢台", "保定", "张家口", "承德", "沧州", "廊坊", "衡水",
    "太原", "大同", "阳泉", "长治", "晋城", "朔州", "晋中", "运城", "忻州", "临汾", "吕梁",
    "沈阳", "大连", "鞍山", "抚顺", "本溪", "丹东", "锦州", "营口", "阜新", "辽阳", "盘锦", "铁岭", "朝阳", "葫芦岛",
    "长春", "吉林", "四平", "辽源", "通化", "白山", "松原", "白城",
    "哈尔滨", "齐齐哈尔", "鸡西", "鹤岗", "双鸭山", "大庆", "伊春", "佳木斯", "七台河", "牡丹江", "黑河", "绥化",
    "南京", "无锡", "徐州", "常州", "苏州", "南通", "连云港", "淮安", "盐城", "扬州", "镇江", "泰州", "宿迁",
    "杭州", "宁波", "温州", "嘉兴", "湖州", "绍兴", "金华", "衢州", "舟山", "台州", "丽水",
    "合肥", "芜湖", "蚌埠", "淮南", "马鞍山", "淮北", "铜陵", "安庆", "黄山", "滁州", "阜阳", "宿州", "六安", "亳州", "池州", "宣城",
    "福州", "厦门", "莆田", "三明", "泉州", "漳州", "南平", "龙岩", "宁德",
    "南昌", "景德镇", "萍乡", "九江", "新余", "鹰潭", "赣州", "吉安", "宜春", "抚州", "上饶",
    "济南", "青岛", "淄博", "枣庄", "东营", "烟台", "潍坊", "济宁", "泰安", "威海", "日照", "临沂", "德州", "聊城", "滨州", "菏泽",
]

# CITIES 按省级行政区连续分段，与 PROVINCES 中「河北」起至「山东」顺序一一对应（其余省暂无地市表数据）
_PROVINCE_CITY_BLOCK_NAMES = PROVINCES[4:15]
_PROVINCE_CITY_BLOCK_SIZES = [11, 11, 14, 8, 12, 13, 11, 16, 9, 11, 16]


def _build_province_to_cities() -> dict[str, tuple[str, ...]]:
    d: dict[str, tuple[str, ...]] = {}
    i = 0
    for name, n in zip(_PROVINCE_CITY_BLOCK_NAMES, _PROVINCE_CITY_BLOCK_SIZES):
        d[name] = tuple(CITIES[i : i + n])
        i += n
    assert i == len(CITIES), "PROVINCE_CITY_BLOCK_SIZES 与 CITIES 总长度不一致"
    return d


PROVINCE_TO_CITIES = _build_province_to_cities()


def cities_under_provinces(provinces: list[str]) -> list[str]:
    """已选省份下的地市并集，顺序与 CITIES 全局顺序一致。"""
    union: set[str] = set()
    for p in provinces:
        union.update(PROVINCE_TO_CITIES.get(p, ()))
    return [c for c in CITIES if c in union]


def split_multi(value: str) -> list[str]:
    return [x.strip() for x in str(value or "").split("|") if x.strip()]


# 任务名前半段：从前到后第一个「地域」中文词 + 紧随其后的 6 位字母数字（运营商2+类型2+行业2）
_REGION_NAME_TOKENS = sorted(set(PROVINCES + CITIES + ["全国"]), key=len, reverse=True)
_REGION_NUM_RE = re.compile(r"^([一二两三四五六七八九十]+省|[一二两三四五六七八九十]+市|\d+省|\d+市)")


def find_earliest_region_in_left(left: str) -> tuple[int, int, str] | None:
    """返回 (起始下标, 结束下标不含, 地域词)。取最靠前起点；同起点取长匹配。"""
    n = len(left)
    best: tuple[int, int, str] | None = None
    for start in range(n):
        ch = left[start]
        if not ("\u4e00" <= ch <= "\u9fff"):
            continue
        cand: str | None = None
        for tok in _REGION_NAME_TOKENS:
            if start + len(tok) <= n and left[start : start + len(tok)] == tok:
                cand = tok
                break
        if cand is None:
            m = _REGION_NUM_RE.match(left[start:])
            if m:
                cand = m.group(1)
        if not cand:
            continue
        end = start + len(cand)
        if best is None or start < best[0] or (start == best[0] and len(cand) > len(best[2])):
            best = (start, end, cand)
    return best


def first_url(value: str) -> str:
    if not value:
        return ""
    for line in str(value).splitlines():
        clean = line.strip()
        if clean:
            return clean
    return ""


def url_lines_for_filter(raw: Any) -> list[str]:
    """多行 URL 拆成去重后的非空行列表；无有效行时为 [\"\"] 供筛选空单元格。"""
    text = str(raw or "").replace("\r\n", "\n").replace("\r", "\n")
    out: list[str] = []
    seen: set[str] = set()
    for line in text.split("\n"):
        t = line.strip()
        if not t or t in seen:
            continue
        seen.add(t)
        out.append(t)
    return out if out else [""]


def sanitize_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(text or "")).strip() or "客户"


def int_to_cn(n: int) -> str:
    return {1: "一", 2: "两", 3: "三", 4: "四", 5: "五", 6: "六", 7: "七", 8: "八", 9: "九", 10: "十"}.get(n, str(n))
