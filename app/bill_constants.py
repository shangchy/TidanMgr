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
    "source",
    "printed_at",
    "row_count",
    "include_print_url",
    "file_exists",
    "path",
]
PRINT_LOG_HEADERS = {
    "filename": "文件名",
    "source": "来源",
    "printed_at": "创建时间",
    "row_count": "数据条数",
    "include_print_url": "打印URL",
    "file_exists": "状态",
    "path": "路径",
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
PROVINCE_TO_CITIES: dict[str, tuple[str, ...]] = {
    "北京": ("北京",),
    "天津": ("天津",),
    "上海": ("上海",),
    "重庆": ("重庆",),
    "河北": ("石家庄", "唐山", "秦皇岛", "邯郸", "邢台", "保定", "张家口", "承德", "沧州", "廊坊", "衡水"),
    "山西": ("太原", "大同", "阳泉", "长治", "晋城", "朔州", "晋中", "运城", "忻州", "临汾", "吕梁"),
    "辽宁": ("沈阳", "大连", "鞍山", "抚顺", "本溪", "丹东", "锦州", "营口", "阜新", "辽阳", "盘锦", "铁岭", "朝阳", "葫芦岛"),
    "吉林": ("长春", "吉林", "四平", "辽源", "通化", "白山", "松原", "白城", "延边朝鲜族自治州"),
    "黑龙江": ("哈尔滨", "齐齐哈尔", "鸡西", "鹤岗", "双鸭山", "大庆", "伊春", "佳木斯", "七台河", "牡丹江", "黑河", "绥化", "大兴安岭地区"),
    "江苏": ("南京", "无锡", "徐州", "常州", "苏州", "南通", "连云港", "淮安", "盐城", "扬州", "镇江", "泰州", "宿迁"),
    "浙江": ("杭州", "宁波", "温州", "嘉兴", "湖州", "绍兴", "金华", "衢州", "舟山", "台州", "丽水"),
    "安徽": ("合肥", "芜湖", "蚌埠", "淮南", "马鞍山", "淮北", "铜陵", "安庆", "黄山", "滁州", "阜阳", "宿州", "六安", "亳州", "池州", "宣城"),
    "福建": ("福州", "厦门", "莆田", "三明", "泉州", "漳州", "南平", "龙岩", "宁德"),
    "江西": ("南昌", "景德镇", "萍乡", "九江", "新余", "鹰潭", "赣州", "吉安", "宜春", "抚州", "上饶"),
    "山东": ("济南", "青岛", "淄博", "枣庄", "东营", "烟台", "潍坊", "济宁", "泰安", "威海", "日照", "临沂", "德州", "聊城", "滨州", "菏泽"),
    "河南": ("郑州", "开封", "洛阳", "平顶山", "安阳", "鹤壁", "新乡", "焦作", "濮阳", "许昌", "漯河", "三门峡", "南阳", "商丘", "信阳", "周口", "驻马店", "济源"),
    "湖北": ("武汉", "黄石", "十堰", "宜昌", "襄阳", "鄂州", "荆门", "孝感", "荆州", "黄冈", "咸宁", "随州", "恩施土家族苗族自治州", "仙桃", "潜江", "天门", "神农架"),
    "湖南": ("长沙", "株洲", "湘潭", "衡阳", "邵阳", "岳阳", "常德", "张家界", "益阳", "郴州", "永州", "怀化", "娄底", "湘西土家族苗族自治州"),
    "广东": ("广州", "深圳", "珠海", "汕头", "佛山", "韶关", "湛江", "肇庆", "江门", "茂名", "惠州", "梅州", "汕尾", "河源", "阳江", "清远", "东莞", "中山", "潮州", "揭阳", "云浮"),
    "海南": ("海口", "三亚", "三沙", "儋州", "五指山", "琼海", "文昌", "万宁", "东方", "定安", "屯昌", "澄迈", "临高", "白沙黎族自治县", "昌江黎族自治县", "乐东黎族自治县", "陵水黎族自治县", "保亭黎族苗族自治县", "琼中黎族苗族自治县"),
    "四川": ("成都", "自贡", "攀枝花", "泸州", "德阳", "绵阳", "广元", "遂宁", "内江", "乐山", "南充", "眉山", "宜宾", "广安", "达州", "雅安", "巴中", "资阳", "阿坝藏族羌族自治州", "甘孜藏族自治州", "凉山彝族自治州"),
    "贵州": ("贵阳", "六盘水", "遵义", "安顺", "毕节", "铜仁", "黔西南布依族苗族自治州", "黔东南苗族侗族自治州", "黔南布依族苗族自治州"),
    "云南": ("昆明", "曲靖", "玉溪", "保山", "昭通", "丽江", "普洱", "临沧", "楚雄彝族自治州", "红河哈尼族彝族自治州", "文山壮族苗族自治州", "西双版纳傣族自治州", "大理白族自治州", "德宏傣族景颇族自治州", "怒江傈僳族自治州", "迪庆藏族自治州"),
    "陕西": ("西安", "铜川", "宝鸡", "咸阳", "渭南", "延安", "汉中", "榆林", "安康", "商洛"),
    "甘肃": ("兰州", "嘉峪关", "金昌", "白银", "天水", "武威", "张掖", "平凉", "酒泉", "庆阳", "定西", "陇南", "临夏回族自治州", "甘南藏族自治州"),
    "青海": ("西宁", "海东", "海北藏族自治州", "黄南藏族自治州", "海南藏族自治州", "果洛藏族自治州", "玉树藏族自治州", "海西蒙古族藏族自治州"),
    "台湾": ("台北", "新北", "桃园", "台中", "台南", "高雄", "基隆", "新竹", "嘉义"),
    "内蒙古": ("呼和浩特", "包头", "乌海", "赤峰", "通辽", "鄂尔多斯", "呼伦贝尔", "巴彦淖尔", "乌兰察布", "兴安盟", "锡林郭勒盟", "阿拉善盟"),
    "广西": ("南宁", "柳州", "桂林", "梧州", "北海", "防城港", "钦州", "贵港", "玉林", "百色", "贺州", "河池", "来宾", "崇左"),
    "西藏": ("拉萨", "日喀则", "昌都", "林芝", "山南", "那曲", "阿里地区"),
    "宁夏": ("银川", "石嘴山", "吴忠", "固原", "中卫"),
    "新疆": ("乌鲁木齐", "克拉玛依", "吐鲁番", "哈密", "昌吉回族自治州", "博尔塔拉蒙古自治州", "巴音郭楞蒙古自治州", "阿克苏地区", "克孜勒苏柯尔克孜自治州", "喀什地区", "和田地区", "伊犁哈萨克自治州", "塔城地区", "阿勒泰地区", "石河子", "阿拉尔", "图木舒克", "五家渠", "北屯", "铁门关", "双河", "可克达拉", "昆玉", "胡杨河", "新星"),
    "香港": ("香港",),
    "澳门": ("澳门",),
}

CITIES = [c for p in PROVINCES for c in PROVINCE_TO_CITIES.get(p, ())]


def _build_city_alias_to_full() -> dict[str, str]:
    """构建地市简称 -> 全称映射；仅保留唯一映射，避免歧义。"""
    cand: dict[str, set[str]] = {}
    suffixes = ("特别行政区", "自治州", "自治县", "自治旗", "地区", "盟", "市", "县")
    for full in CITIES:
        aliases = {full}
        for suf in suffixes:
            if full.endswith(suf) and len(full) > len(suf):
                aliases.add(full[: -len(suf)])
        if "自治州" in full:
            aliases.add(full.split("自治州", 1)[0])
        if "自治县" in full:
            aliases.add(full.split("自治县", 1)[0])
        if len(full) >= 2:
            aliases.add(full[:2])
        if len(full) >= 3:
            aliases.add(full[:3])
        for a in aliases:
            a = a.strip()
            if not a:
                continue
            cand.setdefault(a, set()).add(full)
    out: dict[str, str] = {}
    for alias, targets in cand.items():
        if len(targets) == 1:
            out[alias] = next(iter(targets))
    return out


CITY_ALIAS_TO_FULL = _build_city_alias_to_full()


def normalize_city_name(raw: str, allowed: set[str] | None = None) -> str | None:
    """地市简称/全称规范化；未命中返回 None。"""
    t = str(raw or "").strip()
    if not t:
        return ""
    if allowed is not None and t in allowed:
        return t
    if t in set(CITIES):
        if allowed is None or t in allowed:
            return t
        return None
    full = CITY_ALIAS_TO_FULL.get(t)
    if not full:
        return None
    if allowed is not None and full not in allowed:
        return None
    return full


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


def region_in_task_name_requires_customer_prefix(region: str) -> bool:
    """
    任务名左段中若出现省/市名或「几省/几市」等地域词，则其前须填写任务编号（客户段），全国除外。
    """
    r = str(region or "").strip()
    if not r or r == "全国":
        return False
    if r in PROVINCES or r in CITIES:
        return True
    if _REGION_NUM_RE.fullmatch(r):
        return True
    return False


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


def sanitize_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(text or "")).strip() or "客户"


def int_to_cn(n: int) -> str:
    return {1: "一", 2: "两", 3: "三", 4: "四", 5: "五", 6: "六", 7: "七", 8: "八", 9: "九", 10: "十"}.get(n, str(n))
