#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
地址坐标转换工具
━━━━━━━━━━━━━━━━━━━━━━━━━
模块化重构版 | 支持在线更新
"""

import sys, os, json, math, re, time, threading, ssl, urllib.request, urllib.parse, difflib, datetime
import io, platform, subprocess, tempfile
from tkinter import font as tkfont
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
# ═══════════════════════════════════════════════════════════
# 0. 延迟加载第三方库
# ═══════════════════════════════════════════════════════════
_openpyxl = _PIL = None

def _get_openpyxl():
    global _openpyxl
    if _openpyxl is None:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment
        from openpyxl.utils import get_column_letter
        _openpyxl = (openpyxl, PatternFill, Font, Alignment, get_column_letter)
    return _openpyxl

def _get_pil():
    global _PIL
    if _PIL is None:
        from PIL import Image, ImageTk
        _PIL = (Image, ImageTk)
    return _PIL


# ═══════════════════════════════════════════════════════════
# 1. 常量与配置模块
# ═══════════════════════════════════════════════════════════
class Config:
    APP_NAME      = "地址坐标转换工具"
    VERSION       = "2.8.37"
    CONFIG_FILE   = "addr_config.json"
    ABNORMAL_CLR  = "FFFF00"

    # 更新服务器 —— 请替换为您的 Cloudflare Worker 自定义域名
    UPDATE_URL    = "https://addressmap.mypdftool1.top"

    # 更多工具推广配置
    TOOLS = [
        {
            "name": "PDF 工具箱",
            "icon": "📄",
            "desc": "PDF 转 Word/Excel/PPT、合并拆分、OCR 识别、批量签章、表单填写，一站式办公解决方案",
            "version": "v10.0.39",
            "size": "128 MB",
            "category": "办公效率",
            "download_url": "https://pub-13f79ef5b83f4d82bb6b3678d818810e.r2.dev/pdf_tool_setup.exe",
            "color": "#dc2626",
            "badge": "🔥 热门"
        },
        {
            "name": "智能 C 盘清理",
            "icon": "🧹",
            "desc": "深度清理系统垃圾、注册表优化、内存加速、大文件分析、网络优化，让电脑重获新生",
            "version": "v2.1.0",
            "size": "28 MB",
            "category": "系统优化",
            "download_url": "https://addressmap.mypdftool1.top/download/smart_cleaner_setup.exe",
            "color": "#059669",
            "badge": ""
        },
        {
            "name": "印章制作大师",
            "icon": "🔴",
            "desc": "圆形章、椭圆章、防伪纹仿真生成，支持字体弧形排布、断点防伪，效果逼真",
            "version": "v3.5.2",
            "size": "12 MB",
            "category": "设计工具",
            "download_url": "https://pub-f682efdabe984dd48f0a151c7b77ce77.r2.dev/sealmaker/%E5%8D%B0%E7%AB%A0%E5%A4%A7%E5%B8%88_%E5%AE%89%E8%A3%85%E5%8C%85.exe",
            "color": "#d97706",
            "badge": ""
        },
        {
            "name": "地址标准化工具箱",
            "icon": "📍",
            "desc": "基于高德/百度地图 API 的地址智能解析、坐标转换、批量编码，支持正逆向地理编码",
            "version": "v2.8.13",
            "size": "13 MB",
            "category": "数据处理",
            "download_url": "https://addressmap.mypdftool1.top/download/setup_addressmap_tool_v2.8.6.exe",
            "color": "#2563eb",
            "badge": "✅ 当前"
        },
    ]

    # API 端点
    GAODE_GEO     = "https://restapi.amap.com/v3/geocode/geo"
    GAODE_REGO    = "https://restapi.amap.com/v3/geocode/regeo"
    GAODE_STATIC  = "https://restapi.amap.com/v3/staticmap"
    BAIDU_GEO     = "https://api.map.baidu.com/geocoding/v3/"
    BAIDU_REGO    = "https://api.map.baidu.com/reverse_geocoding/v3/"
    BAIDU_STATIC  = "https://api.map.baidu.com/staticimage/v2"

    # POI 检索 API 端点
    GAODE_POI_TEXT   = "https://restapi.amap.com/v3/place/text"
    GAODE_POI_AROUND = "https://restapi.amap.com/v3/place/around"
    GAODE_DIST       = "https://restapi.amap.com/v3/config/district"
    BAIDU_POI        = "https://api.map.baidu.com/place/v2/search"

    # 颜色主题
    C_PRIMARY   = "#1e293b"
    C_SECONDARY = "#334155"
    C_ACCENT    = "#d97706"
    C_BG        = "#f1f5f9"
    C_CARD      = "#ffffff"
    C_TEXT      = "#0f172a"
    C_TEXT_M    = "#475569"
    C_TEXT_L    = "#64748b"
    C_BORDER    = "#e2e8f0"
    C_SUCCESS   = "#059669"
    C_DANGER    = "#dc2626"
    C_WARNING   = "#d97706"
    C_INFO      = "#2563eb"

    # 坐标常量
    PI = 3.14159265358979
    X_PI = 52.3598775598299
    EARTH_A = 6378245.0
    EARTH_EE = 0.006693421622965943

    # 列映射
    COLS = {
        "serial": 1, "province": 2, "city": 3, "district": 4,
        "addr1": 5, "addr2": 6, "gcj_lng": 7, "gcj_lat": 8,
        "wgs_lng": 9, "wgs_lat": 10, "std_addr": 11,
        "rego_full": 12, "rego_div": 13, "rego_street": 14
    }

    # 错误码映射
    GAODE_ERR = {
        "10001": "高德 Key 无效或已过期", "10003": "高德日调用额度已用完",
        "10004": "高德 Key 类型不正确（需Web服务Key）", "10005": "高德 Key IP白名单限制",
        "10007": "高德数字签名验证失败", "20003": "高德静态地图权限不足或服务端异常",
    }
    # POI 关键词扩展映射（用于突破500条限制）
    POI_KEYWORD_EXPAND = {
        "网吧": ["网咖", "电竞馆", "网络会所", "电竞网吧"],
        "网咖": ["网吧", "电竞馆", "网络会所"],
        "便利店": ["超市", "小卖部", "杂货店", "士多店", "社区超市"],
        "超市": ["便利店", "生鲜超市", "社区超市", "小卖部"],
        "商铺": ["店面", "门店", "铺面", "商业店铺"],
        "火锅店": ["火锅", "串串香", "麻辣烫", "涮锅"],
        "加油站": ["加油", " petrol station", "油站"],
        "幼儿园": ["托儿所", "学前班", "幼托", "早教中心"],
        "银行": ["支行", "ATM", "信用社", "储蓄所"],
        "药店": ["药房", "大药房", "连锁药店", "医保药店"],
        "酒店": ["宾馆", "旅馆", "招待所", "客栈", "民宿"],
        "宾馆": ["酒店", "旅馆", "招待所", "客栈"],
        "旅馆": ["酒店", "宾馆", "招待所"],
        "景区": ["风景区", "旅游景点", "公园", "度假区", "名胜区"],
        "公园": ["景区", "风景区", "游乐场", "主题公园"],
        "餐厅": ["饭店", "餐馆", "美食城", "酒楼", "大排档"],
        "饭店": ["餐厅", "餐馆", "酒楼", "大排档"],
        "学校": ["中学", "小学", "培训机构", "教育中心"],
        "医院": ["诊所", "卫生院", "门诊部", "卫生服务中心"],
        "诊所": ["医院", "门诊部", "卫生所"],
    }

    BAIDU_ERR = {
        "0": "请求正常",
        "1": "服务器内部错误（百度服务端临时故障，请稍后重试）",
        "2": "请求参数非法（请检查地址格式）",
        "3": "百度权限校验失败（AK无地理编码权限，请检查AK配置）",
        "4": "百度配额校验失败（当日/当月配额已用完）",
        "5": "百度 AK 不存在或非法（请检查AK是否正确）",
        "6": "百度服务器繁忙（请降低请求频率后重试）",
        "7": "百度服务不可用（该服务已下线或维护中）",
        "10": "百度IP白名单校验失败（当前IP不在白名单内）",
        "101": "百度SN校验失败（数字签名验证失败）",
        "102": "百度SK不存在或非法",
        "200": "百度无相关数据（该地址无匹配结果）",
        "301": "百度永久配额超限",
        "302": "百度天配额超限（今日额度已用完，请明日再试）",
        "401": "百度天配额超限（今日额度已用完，请明日再试）",
        "402": "百度月配额超限",
    }

    # 精度等级
    LEVEL_SCORE = {
        "门址": 100, "门牌号": 100, "门址点": 100, "门牌": 100,
        "POI": 90, "兴趣点": 90, "楼宇": 90, "大厦": 90, "建筑物": 90, "购物": 90,
        "住宅区": 85, "住宅小区": 85, "小区": 85, "社区": 80, 
        "村庄": 90, "村": 90, "自然村": 90, "村委会": 90,
        "道路": 70, "街道": 70, "路段": 70, "道路交叉口": 70, "商圈": 65, "开发区": 65,
        "乡镇": 60, "乡镇级": 60, "街道级": 60, "镇": 60, "乡": 60,
        "区县": 50, "区县级": 50, "区县级别": 50, "区": 50, "县": 50,
        "城市": 30, "市级": 30, "地市级": 30, "市": 30, "省": 10, "省级": 10, "国家": 5
    }

    @classmethod
    def load(cls):
        default = {
            "gaode_key": "", "baidu_ak": "",
            "enable_baidu_fallback": True, "baidu_trigger_level": "乡镇",
            "max_retry": 3, "base_interval": 0.2, "skip_existing": True,
            "check_drift": True, "enable_cache": True, "baidu_debug": True,
            "auto_open_excel": False
        }
        if os.path.exists(cls.CONFIG_FILE):
            try:
                with open(cls.CONFIG_FILE, "r", encoding="utf-8") as f:
                    default.update(json.load(f))
            except: pass
        return default

    @classmethod
    def save(cls, cfg):
        with open(cls.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════════════════
# 2. 工具函数模块
# ═══════════════════════════════════════════════════════════
def url_enc(s): return urllib.parse.quote(str(s), safe="")

def http_get(url, max_retry=3, timeout=15):
    ctx = ssl.create_default_context()
    ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
    for retry in range(1, max_retry + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                data = resp.read().decode("utf-8")
                if data: return True, data
        except Exception as e:
            if retry == max_retry: return False, str(e)
            time.sleep(0.3 * retry)
    return False, ""

def json_val(js, key):
    """提取 JSON 字段值：优先 json.loads，失败则正则兜底"""
    try:
        d = json.loads(js)
        if isinstance(d, dict):
            v = d.get(key)
            return str(v) if v is not None else ""
    except: pass
    m = re.search(r'"' + re.escape(key) + r'"\s*:\s*("?)([^"\,}\s]+)\1', js)
    return m.group(2) if m else ""

def geo_field(js, field):
    """从 geocodes[0] 中提取字段：优先 json.loads"""
    try:
        d = json.loads(js)
        gc = d.get("geocodes", [])
        if gc and isinstance(gc, list) and isinstance(gc[0], dict):
            v = gc[0].get(field)
            return str(v) if v is not None else ""
    except: pass
    m = re.search(r'"geocodes"\s*:\s*\[\s*\{.*?"' + re.escape(field) + r'"\s*:\s*"([^"]*)"', js, re.DOTALL)
    return m.group(1) if m else ""

def regeo_field(js, field):
    """从 regeocode 中提取字段：优先 json.loads"""
    try:
        d = json.loads(js)
        rg = d.get("regeocode", {})
        if isinstance(rg, dict):
            v = rg.get(field)
            return str(v) if v is not None else ""
    except: pass
    m = re.search(r'"regeocode"\s*:\s*\{.*?"' + re.escape(field) + r'"\s*:\s*"([^"]*)"', js, re.DOTALL)
    return m.group(1) if m else ""

def regeo_comp(js, field):
    """从 addressComponent 中提取字段：优先 json.loads"""
    try:
        d = json.loads(js)
        rg = d.get("regeocode", {})
        if isinstance(rg, dict):
            ac = rg.get("addressComponent", {})
            if isinstance(ac, dict):
                v = ac.get(field)
                return str(v) if v is not None else ""
    except: pass
    m = re.search(r'"addressComponent"\s*:\s*\{.*?"' + re.escape(field) + r'"\s*:\s*"([^"]*)"', js, re.DOTALL)
    return m.group(1) if m else ""
def baidu_fmt(js):
    try:
        d = json.loads(js)
        if d.get("status") in (0, "0") and isinstance(d.get("result"), dict):
            r = d["result"]
            for k in ["formatted_address", "sematic_description"]:
                v = r.get(k)
                if v and str(v).strip(): return str(v).strip()
            ac = r.get("addressComponent", {})
            parts = [str(ac.get(k, "")).strip() for k in ["province", "city", "district", "town", "street", "street_number"] if ac.get(k)]
            if parts: return "".join(parts)
    except: pass
    return ""

def clean_addr(addr):
    if not addr: return ""
    addr = str(addr)
    for k, v in {"　": " ", "，": ",", "。": ".", "；": ";", "：": ":", "（": "(", "）": ")", "\r\n": "", "\n": "", "\t": ""}.items():
        addr = addr.replace(k, v)
    while "  " in addr: addr = addr.replace("  ", " ")
    return addr.strip()

def rm_paren(s):
    s = re.sub(r"[(（][^)）]*[)）]", "", s)
    while "  " in s: s = s.replace("  ", " ")
    return s.strip()

def addr_levels(addr):
    levels = [addr]
    a2 = rm_paren(addr)
    if a2 != addr and len(a2) > 4: levels.append(a2)

    # 去掉省级前缀（如"云南省曲靖市..." → "曲靖市..."）
    prov_match = re.match(r'^(北京|天津|上海|重庆|河北|山西|辽宁|吉林|黑龙江|江苏|浙江|安徽|福建|江西|山东|河南|湖北|湖南|广东|海南|四川|贵州|云南|陕西|甘肃|青海|台湾|内蒙古|广西|西藏|宁夏|新疆|香港|澳门)[省市自治区]+', a2)
    if prov_match:
        rest = a2[prov_match.end():].strip()
        if rest and len(rest) > 3 and rest not in levels:
            levels.append(rest)

    # 去掉市级前缀（如"曲靖市麒麟区..." → "麒麟区..."）
    search_base = a2[prov_match.end():].strip() if prov_match else a2
    city_match = re.match(r'^([^省]+?[市州盟])', search_base)
    if city_match:
        rest = search_base[city_match.end():].strip()
        if rest and len(rest) > 3 and rest not in levels:
            levels.append(rest)

    # 按分隔符截断
    for sep in [",", " ", "-"]:
        pos = a2.rfind(sep)
        if pos > 3: 
            truncated = a2[:pos].strip()
            if truncated and truncated not in levels:
                levels.append(truncated)
            break

    # 原来的正则截断（兜底）
    m = re.search(r"(.+?(省|市))", addr)
    if m and m.group(1) not in levels: 
        levels.append(m.group(1))

    return levels


def extract_city(addr):
    """从地址字符串中提取市级行政区划名称（用于高德/百度API city参数）

    v2.7.14 优化：若未提取到市级名，则退而求其次提取县/区级名。
    百度地理编码的 city 参数对区县名同样具有限定作用，
    可显著提升「XX县XX镇」这类纯县镇地址的匹配率。
    """
    if not addr:
        return ""
    addr = str(addr).strip()
    # 直辖市
    for direct in ["北京市", "上海市", "天津市", "重庆市"]:
        if direct in addr:
            return direct
    # 方向1：匹配 XX市、XX州、XX盟
    m = re.search(r'([^省]+?[市州盟])', addr)
    if m:
        city = m.group(1)
        if len(city) >= 2 and "省" not in city:
            return city
    # 方向2（v2.7.14新增）：未提取到市级时，提取县/区名作为备选
    # 百度 API 的 city 参数对区县名同样有效，可缩小搜索范围
    m = re.search(r'([^省市]+?[县区])', addr)
    if m:
        district = m.group(1)
        if len(district) >= 2:
            return district
    return ""

def build_addr(prov, city, dist, addr):
    result = ""
    for part in [prov, city, dist]:
        if part and part not in addr and part not in result: result += part
    return result + addr

def has_detail(addr):
    return bool(addr) and any(k in addr for k in ["路", "街", "号", "村", "小区", "大厦", "道", "巷", "弄", "镇", "乡", "社区"])

def extract_district(addr):
    """提取地址文本中的区县级名称（如“麒麟区”“陆良县”“赛罕区”）。
    排除“自治区”等省级后缀；对“曲靖市麒麟区”这类含市级前缀的长词，
    用零宽前瞻取内部纯区县名，避免与结果地址因写法差异而误判漂移。"""
    if not addr:
        return set()
    out = set()
    for d in re.findall(r'(?=([\u4e00-\u9fa5]{2,6}?(?:区|县|旗)))', str(addr)):
        if "自治区" in d:
            continue
        if any(x in d for x in ["市", "州", "盟", "省"]):
            for s in re.findall(r'(?=([\u4e00-\u9fa5]{2,4}?(?:区|县|旗)))', d):
                if "自治区" not in s and not any(x in s for x in ["市", "州", "盟", "省"]):
                    out.add(s)
        else:
            out.add(d)
    return out


def check_drift(std_addr, prov, city, dist, raw_addr=""):
    """漂移检测（修复版：补上“地址文本区县”比对，A县→B县 不再漏检）
    ① 省/市/区县列中的任一行政区划不在结果地址中 → 漂移
    ② 输入地址（省市区列 + 地址文本）中的区县名与结果区县名不一致 → 漂移
    当输入未提供任何区县信息时，不做②判定，避免误伤。"""
    if not std_addr: return True
    if any(p and p not in std_addr for p in [dist, city, prov]):
        return True
    src = extract_district("".join(filter(None, [prov, city, dist, raw_addr])))
    dst = extract_district(std_addr)
    if src and dst and not (src & dst):
        return True
    return False

def extract_core(raw_addr, prov, city, dist):
    s = raw_addr
    for p in [prov, city, dist]:
        if p: s = s.replace(p, "")
    return s.strip()

def dedup_addr(addr):
    if not addr: return addr
    for _ in range(3): addr = re.sub(r'([^\s]+?)\1+', r'\1', addr)
    addr = re.sub(r'([^\s]{2,6})\s*\1', r'\1', addr)
    return addr.strip()

def build_query_addr(a1, a2, city, dist, do_clean, do_merge):
    """清洗 + 智能合并 + 市/区县拼接，生成最终查询地址（修复版：补入区县列，
    保证清洗预览、去重预处理、批量处理三者逻辑完全一致）"""
    if do_clean:
        a1 = _clean_single_addr(a1)
    if do_merge and a2:
        a1 = _clean_single_addr(a1)
        if len(a1) <= 6 and any(k in a1 for k in ["区", "县", "镇", "乡"]) and any(k in a2 for k in ["路", "街", "号", "村", "小区", "大厦", "道", "巷"]):
            a1 = a1 + a2
    if dist and dist not in a1:
        a1 = dist + a1
    if city and city not in a1:
        a1 = city + a1
    return a1


def get_poi_district_units(gkey, city):
    """通过高德行政区划查询获取城市下属区县列表（POI 区县下钻检索用）；
    查询失败或无下属区县时回退为城市本身，兼容直辖市、县级市等场景"""
    try:
        url = f"{Config.GAODE_DIST}?keywords={url_enc(city)}&subdistrict=1&level=city&extensions=base&key={gkey}"
        ok, resp = http_get(url, 2)
        if not ok:
            return [city]
        d = json.loads(resp)
        if d.get("status") != "1":
            return [city]
        dists = d.get("districts") or []
        if not dists:
            return [city]
        subs = dists[0].get("districts") or []
        units = [s.get("name", "") for s in subs if s.get("name")]
        return units if units else [city]
    except Exception:
        return [city]


def _clean_single_addr(addr):
    """单条地址清洗：去空格、统一符号、去重前缀"""
    if not addr: return ""
    addr = str(addr).strip()
    for k, v in {"　": "", " ": "", "\r\n": "", "\n": "", "\t": ""}.items():
        addr = addr.replace(k, v)
    addr = addr.replace("（", "(").replace("）", ")")
    addr = addr.replace("【", "[").replace("】", "]")
    prov_match = re.match(r'^(北京|天津|上海|重庆|河北|山西|辽宁|吉林|黑龙江|江苏|浙江|安徽|福建|江西|山东|河南|湖北|湖南|广东|海南|四川|贵州|云南|陕西|甘肃|青海|台湾|内蒙古|广西|西藏|宁夏|新疆|香港|澳门)[省市自治区]+(.+?[市州盟].+)', addr)
    if prov_match:
        addr = prov_match.group(2)
    city_match = re.match(r'^([^省]+?[市州盟])([^市]+?[区县镇乡街道].+)', addr)
    if city_match and len(city_match.group(2)) > 3:
        addr = city_match.group(2)
    return addr.strip()


def _sim_addr(a, b):
    """地址相似度：序列比对+尾部惩罚+长度检查，≥0.85视为重复"""
    if not a or not b: return 0.0
    a, b = str(a).strip(), str(b).strip()
    if a == b: return 1.0
    # 长度差异超过25% → 直接不合并（防止"村"和"村小组"误判）
    max_len = max(len(a), len(b))
    if max_len > 0 and abs(len(a) - len(b)) / max_len > 0.25:
        return 0.0
    # 序列相似度（关注顺序和位置，比字符集合更准确）
    seq_sim = difflib.SequenceMatcher(None, a, b).ratio()
    # 尾部差异惩罚：最后4个字不同 → 相似度打8折（减少自然村名误判）
    tail = min(4, len(a), len(b))
    if tail > 0 and a[-tail:] != b[-tail:]:
        seq_sim *= 0.8
    return seq_sim


def _xml_esc(s):
    """XML特殊字符转义"""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _gen_kml(proc, srow, lrow, addr_idx, out_path):
    """从Excel生成KML文件，使用WGS-84坐标"""
    kml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        '  <Document>',
        '    <name>地址坐标转换结果</name>',
        '    <Style id="redPin">',
        '      <IconStyle>',
        '        <Icon><href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href></Icon>',
        '      </IconStyle>',
        '    </Style>'
    ]
    count = 0
    for r in range(srow, lrow + 1):
        std = str(proc.ws.cell(row=r, column=Config.COLS["std_addr"]).value or "").strip()
        if not std or "失败" in std or "错误" in std or "无匹配" in std:
            continue
        name = std.split(" [精度:")[0].strip()
        # 读取原始地址（如果addr_idx有效）
        orig = ""
        if addr_idx and addr_idx > 0:
            orig = str(proc.ws.cell(row=r, column=addr_idx).value or "").strip()
        gl = proc.ws.cell(row=r, column=Config.COLS["gcj_lng"]).value
        gt = proc.ws.cell(row=r, column=Config.COLS["gcj_lat"]).value
        try:
            gl, gt = float(gl), float(gt)
            wl, wt = gcj02_to_wgs84(gl, gt)
        except:
            continue
        level = "未知"; source = "高德"
        m = re.search(r'\[精度:([^\]]+)\]', std)
        if m: level = m.group(1)
        m = re.search(r'\[来源:([^\]]+)\]', std)
        if m: source = m.group(1)
        kml_lines.append('    <Placemark>')
        kml_lines.append(f'      <name>{_xml_esc(name)}</name>')
        desc = f"原始地址：{_xml_esc(orig)}\nGCJ-02：{gl},{gt}\nWGS-84：{round(wl,6)},{round(wt,6)}\n精度：{level} | 来源：{source}\nExcel行号：第{r}行"
        kml_lines.append(f'      <description><![CDATA[{desc}]]></description>')
        kml_lines.append('      <styleUrl>#redPin</styleUrl>')
        kml_lines.append(f'      <Point><coordinates>{round(wl,6)},{round(wt,6)},0</coordinates></Point>')
        kml_lines.append('    </Placemark>')
        count += 1
    kml_lines += ['  </Document>', '</kml>']
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(kml_lines))
    return count

def level_score(level): return Config.LEVEL_SCORE.get(level, 20)

def gaode_err_msg(code, default=""):
    return Config.GAODE_ERR.get(str(code), default or f"高德未知错误码: {code}")

def baidu_err_msg(code, default=""):
    return Config.BAIDU_ERR.get(str(code), default or f"百度未知错误码: {code}")


# ═══════════════════════════════════════════════════════════
# 3. 坐标转换模块
# ═══════════════════════════════════════════════════════════
def _delta(lng, lat):
    dlat = -100.0 + 2.0*lng + 3.0*lat + 0.2*lat*lat + 0.1*lng*lat + 0.2*math.sqrt(abs(lng))
    dlat += (20.0*math.sin(6.0*lng*Config.PI) + 20.0*math.sin(2.0*lng*Config.PI)) * 2.0/3.0
    dlat += (20.0*math.sin(lat*Config.PI) + 40.0*math.sin(lat/3.0*Config.PI)) * 2.0/3.0
    dlat += (160.0*math.sin(lat/12.0*Config.PI) + 320.0*math.sin(lat*Config.PI/30.0)) * 2.0/3.0
    dlng = 300.0 + lng + 2.0*lat + 0.1*lng*lng + 0.1*lng*lat + 0.1*math.sqrt(abs(lng))
    dlng += (20.0*math.sin(6.0*lng*Config.PI) + 20.0*math.sin(2.0*lng*Config.PI)) * 2.0/3.0
    dlng += (20.0*math.sin(lng*Config.PI) + 40.0*math.sin(lng/3.0*Config.PI)) * 2.0/3.0
    dlng += (150.0*math.sin(lng/12.0*Config.PI) + 300.0*math.sin(lng*Config.PI/30.0)) * 2.0/3.0
    return dlng, dlat

def _iter_transform(src_lng, src_lat, direction):
    lng, lat = src_lng, src_lat
    out_lng, out_lat = src_lng, src_lat
    for _ in range(3):
        dlng, dlat = _delta(lng - 105, lat - 35)
        radlat = lat / 180.0 * Config.PI
        magic = 1 - Config.EARTH_EE * math.sin(radlat) ** 2
        sqrtmagic = math.sqrt(magic)
        dlat = (dlat * 180.0) / ((Config.EARTH_A * (1 - Config.EARTH_EE)) / (magic * sqrtmagic) * Config.PI)
        dlng = (dlng * 180.0) / (Config.EARTH_A / sqrtmagic * math.cos(radlat) * Config.PI)
        out_lat = src_lat + dlat * direction
        out_lng = src_lng + dlng * direction
        lat, lng = out_lat, out_lng
    return out_lng, out_lat

def wgs84_to_gcj02(lng, lat): return _iter_transform(lng, lat, 1)
def gcj02_to_wgs84(lng, lat): return _iter_transform(lng, lat, -1)

def bd09_to_gcj02(bd_lng, bd_lat):
    x, y = bd_lng - 0.0065, bd_lat - 0.006
    z = math.sqrt(x*x + y*y) - 0.00002 * math.sin(y * Config.X_PI)
    theta = math.atan2(y, x) - 0.000003 * math.cos(x * Config.X_PI)
    return z * math.cos(theta), z * math.sin(theta)

def gcj02_to_bd09(lng, lat):
    z = math.sqrt(lng*lng + lat*lat) + 0.00002 * math.sin(lat * Config.X_PI)
    theta = math.atan2(lat, lng) + 0.000003 * math.cos(lng * Config.X_PI)
    return z * math.cos(theta) + 0.0065, z * math.sin(theta) + 0.006

def convert_coords(lng, lat, src_sys, dst_sys):
    if src_sys == dst_sys: return lng, lat
    valid = {"wgs84", "bd09", "gcj02"}
    if src_sys not in valid or dst_sys not in valid:
        raise ValueError(f"非法坐标系: {src_sys} -> {dst_sys}，仅支持 wgs84/bd09/gcj02")
    gcj = {"wgs84": wgs84_to_gcj02, "bd09": bd09_to_gcj02, "gcj02": lambda x,y:(x,y)}[src_sys](lng, lat)
    return {"wgs84": gcj02_to_wgs84, "bd09": gcj02_to_bd09, "gcj02": lambda x,y:(x,y)}[dst_sys](*gcj)


# ═══════════════════════════════════════════════════════════
# 4. 自动更新模块（零依赖 urllib）
# ═══════════════════════════════════════════════════════════
def get_latest_version():
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
        req = urllib.request.Request(f"{Config.UPDATE_URL}/version", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10, context=ctx) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except: return None

def cmp_ver(local, remote):
    try:
        lp = list(map(int, local.split(".")))
        rp = list(map(int, remote.split(".")))
        ml = max(len(lp), len(rp))
        lp += [0]*(ml - len(lp))
        rp += [0]*(ml - len(rp))
        return rp > lp
    except: return False

def _verify_exe(path, min_mb=1):
    if not os.path.exists(path): return False
    if os.path.getsize(path) < min_mb * 1024 * 1024: return False
    try:
        with open(path, 'rb') as f:
            return f.read(2) == b'MZ'
    except: return False

def download_update(download_url, new_name, status_cb=None, root=None, progress_cb=None):
    """
    后台线程下载 Inno Setup 安装包，下载完成后在主线程弹出安装确认
    progress_cb: 可选的进度回调，接收 (downloaded_bytes, total_bytes)
    """
    if status_cb:
        status_cb("正在连接更新服务器...")

    temp_dir = tempfile.gettempdir()
    setup_path = os.path.normpath(os.path.join(temp_dir, new_name))

    def _do_download():
        """在后台线程执行下载"""
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
            req = urllib.request.Request(download_url, headers={"User-Agent": "Mozilla/5.0"})

            with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
                total = int(resp.headers.get('Content-Length', 0))
                downloaded = 0
                with open(setup_path, "wb") as f:
                    while True:
                        chunk = resp.read(65536)  # 64KB 块
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_cb:
                            try:
                                progress_cb(downloaded, total)
                            except:
                                pass
                        if total > 0 and status_cb:
                            pct = min(100, int(downloaded * 100 / total))
                            msg = f"正在下载... {pct}% ({downloaded//1024//1024}MB / {total//1024//1024}MB)"
                            if root and root.winfo_exists():
                                root.after(0, lambda m=msg: status_cb(m))
                            else:
                                status_cb(msg)

            # 校验安装包完整性（EXE 魔数校验）
            if not _verify_exe(setup_path, min_mb=5):
                try: os.remove(setup_path)
                except: pass
                _show_err("下载的安装包不完整或已损坏，请稍后重试。")
                return

            # 下载完成 → 回到主线程执行安装确认
            if root and root.winfo_exists():
                root.after(0, lambda: _run_installer(setup_path, status_cb, root))
            else:
                _run_installer(setup_path, status_cb, root)

        except Exception as e:
            _show_err(f"更新过程出错：{str(e)}")
            try: os.remove(setup_path)
            except: pass

    def _show_err(msg):
        """在主线程显示错误"""
        def _show():
            if status_cb:
                status_cb(f"更新失败：{msg}")
            messagebox.showerror("更新失败", msg)
        if root and root.winfo_exists():
            root.after(0, _show)
        else:
            _show()

    def _run_installer(path, scb, rt):
        """在主线程询问并启动安装程序"""
        if scb:
            scb("下载完成，准备安装...")

        if not messagebox.askyesno(
            "准备安装",
            f"新版本安装包已下载完成。\n\n"
            f"点击「是」立即运行安装程序（推荐先关闭当前窗口中的未保存工作）。\n"
            f"点击「否」稍后手动运行：\n{path}"
        ):
            messagebox.showinfo(
                "稍后安装",
                f"安装包已保存到临时目录，您可以稍后手动运行：\n{path}"
            )
            if scb:
                scb("就绪")
            return

        args = [path, "/VERYSILENT", "/CLOSEAPPLICATIONS", "/NOCANCEL", "/SUPPRESSMSGBOXES"]
        try:
            subprocess.Popen(args, shell=False)
        except Exception as e:
            messagebox.showerror("启动失败", f"无法启动安装程序：{str(e)}\n\n请手动运行：\n{path}")
            if scb:
                scb("就绪")
            return

        if scb:
            scb("安装程序已启动，本程序即将退出...")
        # 延迟 1.5 秒后销毁窗口
        if rt and rt.winfo_exists():
            rt.after(1500, rt.destroy)

    # 启动后台下载线程
    threading.Thread(target=_do_download, daemon=True).start()
# ═══════════════════════════════════════════════════════════
# 5. 地理编码核心模块
# ═══════════════════════════════════════════════════════════
class GeoCoder:
    def __init__(self, cfg):
        self.cfg = cfg
        self.gkey = cfg.get("gaode_key", "")
        self.bak = cfg.get("baidu_ak", "")
        self.use_baidu = cfg.get("enable_baidu_fallback", True)
        self.trigger = cfg.get("baidu_trigger_level", "乡镇")
        self.retry = cfg.get("max_retry", 3)
        self.interval = cfg.get("base_interval", 0.2)
        self.gcnt = self.bcnt = 0
        self.cache, self.rcache, self.ccache = {}, {}, {}
        self.stopped = False
        self.gok = self.bok = False
        self._checked = False

    def stop(self): self.stopped = True

    def check(self):
        if self._checked:
            return (True, "") if self.gok else (False, "高德验证失败")
        if not self.gkey or len(self.gkey) < 10: return False, "高德密钥为空"
        self.gok = self._chk_gaode()
        if self.use_baidu and self.bak: self.bok = self._chk_baidu()
        self._checked = True
        if not self.gok:
            return False, "高德验证失败"
        return True, ""

    def _chk_gaode(self):
        ok, resp = http_get(f"{Config.GAODE_GEO}?address={url_enc('北京市天安门')}&output=json&key={self.gkey}", 2)
        return ok and json_val(resp, "status") == "1"

    def _chk_baidu(self):
        ok, resp = http_get(f"{Config.BAIDU_GEO}?address={url_enc('北京市天安门')}&output=json&ak={self.bak}", 2)
        return ok and json_val(resp, "status") == "0"

    def gaode_geo(self, address, city=""):
        url = f"{Config.GAODE_GEO}?address={url_enc(address)}"
        if city: url += f"&city={url_enc(city)}"
        url += f"&output=json&key={self.gkey}"
        ok, resp = http_get(url, self.retry)
        if not ok: return None, resp
        self.gcnt += 1
        if json_val(resp, "status") != "1":
            c = json_val(resp, "infocode")
            return None, f"高德错误[{c}]: {gaode_err_msg(c, json_val(resp, 'info'))}"
        return {"location": geo_field(resp, "location"), "std_addr": geo_field(resp, "formatted_address"),
                "level": geo_field(resp, "level"), "source": "高德"}, ""

    def gaode_regeo(self, lng, lat):
        loc = f"{lng},{lat}"
        if self.cfg.get("enable_cache") and loc in self.rcache: return self.rcache[loc], ""
        url = f"{Config.GAODE_REGO}?location={loc}&output=json&extensions=all&key={self.gkey}"
        ok, resp = http_get(url, self.retry)
        if not ok: return None, resp
        self.gcnt += 1
        if json_val(resp, "status") != "1":
            c = json_val(resp, "infocode")
            return None, f"高德逆地理错误[{c}]: {gaode_err_msg(c, json_val(resp, 'info'))}"
        r = {"full": regeo_field(resp, "formatted_address"),
             "division": f"{regeo_comp(resp, 'province')}{regeo_comp(resp, 'city')}{regeo_comp(resp, 'district')}",
             "street": f"{regeo_comp(resp, 'street')}{regeo_comp(resp, 'number')}",
             "province": regeo_comp(resp, "province"), "city": regeo_comp(resp, "city"),
             "district": regeo_comp(resp, "district"), "township": regeo_comp(resp, "township")}
        if self.cfg.get("enable_cache"): self.rcache[loc] = r
        return r, ""

    def baidu_geo(self, address, city=""):
        c = clean_addr(address)
        for ch in ["&", "#", "=", "?", "%"]: c = c.replace(ch, "")
        if not c: return None, "地址为空"
        # 百度API：city参数传城市名可提升匹配精度
        # 若外部未传入city，自动从address中提取（含县级兜底）
        if not city:
            city = extract_city(c)
        url = f"{Config.BAIDU_GEO}?address={url_enc(c)}&output=json&ak={self.bak}"
        if city and city not in c:
            url += f"&city={url_enc(city)}"
        ok, resp = http_get(url, 2)
        if not ok: return None, f"百度请求失败: {resp[:80]}"
        self.bcnt += 1
        try: d = json.loads(resp)
        except: return None, f"百度JSON解析失败: {resp[:200]}"
        status = d.get("status")
        if status not in (0, "0"):
            # 返回更详细的错误信息，包含原始响应用于诊断
            msg = d.get("msg", "") or d.get("message", "")
            err_info = baidu_err_msg(str(status), msg or "未知错误")
            # 附加诊断建议
            advice = ""
            if str(status) == "1": advice = "【诊断】百度返回服务端错误，可能原因：①该地址无匹配结果 ②请求参数异常 ③百度服务端瞬时故障"
            elif str(status) == "3": advice = "【诊断】AK无地理编码权限，请前往 https://lbsyun.baidu.com 开通"
            elif str(status) in ("4", "302", "401", "402"): advice = "【诊断】配额已用完，请购买或更换AK"
            elif str(status) == "5": advice = "【诊断】AK不存在或非法，请检查AK配置"
            elif str(status) == "6": advice = "【诊断】请求过于频繁，请降低并发"
            elif str(status) == "10": advice = "【诊断】IP白名单限制，请在百度控制台添加当前IP"
            # 调试信息：显示请求URL和原始响应前200字符
            debug_info = f"[调试] URL参数city={city[:20] if city else '无'} | 响应预览: {resp[:200]}"
            return None, f"百度错误[{status}]: {err_info} {advice}\n{debug_info}"
        r = d.get("result", {})
        if not isinstance(r, dict): return None, "百度返回无result字段"
        loc = r.get("location", {})
        try: blng, blat = float(loc.get("lng", 0)), float(loc.get("lat", 0))
        except: return None, "百度坐标格式错误"
        if blng == 0 or blat == 0: return None, "百度返回坐标为0"
        glng, glat = bd09_to_gcj02(blng, blat)
        level = r.get("level", "未知")
        ck = (round(blng, 6), round(blat, 6))
        if not baidu_fmt(resp) and ck in self.ccache: std = self.ccache[ck]
        elif not baidu_fmt(resp):
            try:
                ru = f"{Config.BAIDU_REGO}?ak={self.bak}&output=json&coordtype=bd09ll&location={blat},{blng}"
                ok2, r2 = http_get(ru, 2)
                if ok2:
                    self.bcnt += 1
                    dr = json.loads(r2)
                    if dr.get("status") in (0, "0"):
                        fa = baidu_fmt(r2)
                        if fa: self.ccache[ck] = fa; std = fa
            except: pass
        # 村庄级别但标准地址不完整（缺村庄名），从原始查询地址提取村庄名补全
        if level in ("村庄", "村", "自然村") and not any(k in std for k in ["村", "寨", "屯"]):
            # 优先从原始查询地址 c 中提取村庄名（比百度逆编码更可靠）
            vm = re.findall(r'[^\s,，]{2,6}[村寨屯]', c)
            vm = [v for v in vm if not any(b in v for b in ["民委员会", "委员会", "居民"])]
            if vm:
                village = vm[-1]
                if village not in std:
                    std = std + village
            else:
                # 原始地址提取失败，尝试百度逆编码 addressComponent 拼接
                try:
                    ru = f"{Config.BAIDU_REGO}?ak={self.bak}&output=json&coordtype=bd09ll&location={blat},{blng}"
                    ok2, r2 = http_get(ru, 2)
                    if ok2:
                        self.bcnt += 1
                        dr = json.loads(r2)
                        if dr.get("status") in (0, "0"):
                            r = dr.get("result", {})
                            ac = r.get("addressComponent", {})
                            parts = [str(ac.get(k, "")).strip() for k in ["province", "city", "district", "town", "street", "street_number"] if ac.get(k)]
                            if parts:
                                full = "".join(parts)
                                if len(full) > len(std):
                                    std = full
                            # street 为空时，尝试 sematic_description 或 poi 补充
                            if not any(k in std for k in ["村", "寨", "屯"]):
                                sd = str(r.get("sematic_description", "")).strip()
                                if sd and any(k in sd for k in ["村", "寨", "屯"]):
                                    vm2 = re.findall(r'[^\s,，]{2,6}[村寨屯]', sd)
                                    vm2 = [v for v in vm2 if not any(b in v for b in ["民委员会", "委员会"])]
                                    if vm2:
                                        village = vm2[-1]
                                        if village not in std:
                                            std = std + village
                                else:
                                    poi = r.get("poiRegions", [])
                                    if isinstance(poi, list) and poi:
                                        pn = str(poi[0].get("name", "")).strip()
                                        if pn and any(k in pn for k in ["村", "寨", "屯"]) and pn not in std:
                                            std = std + pn
                except: pass
        std = dedup_addr(std) or c
        return {"location": f"{glng},{glat}", "std_addr": std, "level": level, "source": "百度"}, ""

    def resolve(self, addr_str, prov_str, city_str, dist_str):
        blog = ""
        if self.stopped: return None, False, "用户取消", False, blog
        if not addr_str: return None, True, "详细地址为空", False, blog
        api_city = "".join(filter(None, [prov_str, city_str, dist_str]))
        # 方向 C：自动提取城市名，提升高德 POI 匹配精度
        auto_city = extract_city(addr_str) if not city_str else ""
        gcity = city_str if "市" in city_str else auto_city
        query = build_addr(prov_str, city_str, dist_str, addr_str)
        cq = clean_addr(query)
        ckey = f"{api_city}|{addr_str}".lower().replace(" ", "").replace("，", "").replace("。", "")
        if self.cfg.get("enable_cache") and ckey in self.cache:
            c = self.cache[ckey]
            return c, c.get("is_abnormal", False), "", False, blog
        has_d = has_detail(addr_str)
        levels = addr_levels(cq)
        gcj_lng = gcj_lat = 0.0
        std_addr = level = source = ""
        is_drift = is_abnormal = False
        err_msg = location = ""
        # 方向 A：在所有查询结果中保留精度最高的
        best_res = None
        best_score = -1
        best_err = ""
        for la in levels:
            if self.stopped: return None, False, "用户取消", False, blog
            res, err = self.gaode_geo(la, gcity)
            if res:
                sc = level_score(res["level"])
                # 保留精度最高的结果
                if sc > best_score:
                    best_score = sc
                    best_res = res
                # 精度足够时提前退出
                if res.get("location") and "," in res["location"] and sc >= level_score("街道"):
                    break
            if err:
                for fc in ["10001", "10003", "10004", "10005", "10007"]:
                    if fc in err: return None, True, f"高德全局错误：{err}", True, blog
                if not best_err: best_err = err
            time.sleep(self.interval)
        # 使用最优结果
        if best_res:
            location = best_res["location"]
            std_addr = best_res["std_addr"]
            level = best_res["level"]
            source = best_res["source"]
            try:
                p = location.split(",")
                gcj_lng, gcj_lat = float(p[0]), float(p[1])
            except: gcj_lng = gcj_lat = 0.0
        else:
            err_msg, is_abnormal = best_err or "高德无匹配结果", True
        if self.cfg.get("check_drift") and std_addr:
            is_drift = check_drift(std_addr, prov_str, city_str, dist_str, addr_str)
            if is_drift: is_abnormal = True
        # 百度Fallback：精度不足或漂移时自动切换
        if self.use_baidu and location and self.bak and self.bok:
            gs = level_score(level); ts = level_score(self.trigger)
            should_fallback = is_drift or (gs <= ts)   # 去掉 has_d 限制
            if is_drift:
                blog = f"[百度Fallback] 地址疑似漂移，尝试百度...\n"
            elif gs <= ts:
                blog = f"[百度Fallback] 高德精度不足({level})，尝试百度...\n"
            else:
                blog = f"[百度Fallback] 高德精度满足要求({level})，未触发百度Fallback\n"
                bd = None
            if should_fallback:
                blog += f"[调试] 自动提取城市名: '{gcity}' | 原始地址: '{addr_str[:30]}'\n"
                bd, be = self.baidu_geo(query, gcity)   # 只调用一次
                if bd:
                    bs = level_score(bd["level"])
                    bdr = check_drift(bd["std_addr"], prov_str, city_str, dist_str, addr_str)
                    if bs > gs and not bdr:
                        p = bd["location"].split(",")
                        gcj_lng, gcj_lat = float(p[0]), float(p[1])
                        std_addr, level, source = bd["std_addr"], bd["level"], "百度"
                        is_drift = is_abnormal = False
                        blog += f"[百度Fallback] 百度提升成功: {std_addr} [精度:{level}]\n"
                    else:
                        blog += f"[百度Fallback] 百度未提升({bd.get('level','')}/{bs} vs {level}/{gs})\n"
                        # 百度未提升：保留高德结果，不污染地址，不标记异常
                else:
                    blog += f"[百度Fallback] 百度失败: {be}\n"
                    # 百度失败：保留高德结果，不污染地址，不标记异常
        if gcj_lng and gcj_lat:
            if self.cfg.get("enable_cache") and not is_abnormal and level_score(level) > level_score("区县"):
                self.cache[ckey] = {"gcj_lng": gcj_lng, "gcj_lat": gcj_lat, "std_addr": std_addr,
                                    "level": level, "source": source, "is_drift": is_drift, "is_abnormal": is_abnormal}
            return {"gcj_lng": gcj_lng, "gcj_lat": gcj_lat, "std_addr": std_addr,
                    "level": level, "source": source, "is_drift": is_drift, "is_abnormal": is_abnormal}, is_abnormal, err_msg or ("解析异常" if is_abnormal else ""), False, blog
        return None, True, err_msg or "双平台均无法解析", False, blog


# ═══════════════════════════════════════════════════════════
# 6. Excel 处理模块
# ═══════════════════════════════════════════════════════════
class ExcelProcessor:
    def __init__(self, filepath):
        openpyxl, PatternFill, Font, Alignment, get_column_letter = _get_openpyxl()
        self._ox, self._pf, self._ft, self._al = openpyxl, PatternFill, Font, Alignment
        self._gcl = get_column_letter
        self.fp = filepath
        self.wb = openpyxl.load_workbook(filepath)
        if "Sheet1" not in self.wb.sheetnames: raise ValueError("Excel 中未找到 Sheet1")
        self.ws = self.wb["Sheet1"]
        self.yf = PatternFill(start_color=Config.ABNORMAL_CLR, end_color=Config.ABNORMAL_CLR, fill_type="solid")
        self.nf = PatternFill(fill_type=None)

    def last_row(self, col=Config.COLS["addr1"]):
        for r in range(self.ws.max_row, 1, -1):
            if self.ws.cell(row=r, column=col).value is not None and str(self.ws.cell(row=r, column=col).value).strip():
                return r
        return 1

    def is_normal(self, row):
        lng = self.ws.cell(row=row, column=Config.COLS["gcj_lng"]).value
        addr = str(self.ws.cell(row=row, column=Config.COLS["std_addr"]).value or "")
        if not lng or not isinstance(lng, (int, float)) or lng == 0: return False
        return not any(k in addr for k in ["失败", "错误", "漂移", "无匹配", "空地址", "无效", "超出范围"])

    def is_abnormal(self, row):
        if self.is_normal(row): return False
        addr = str(self.ws.cell(row=row, column=Config.COLS["std_addr"]).value or "")
        return bool(addr) and any(k in addr for k in ["失败", "错误", "无匹配", "空地址", "无效", "超出范围"])

    def get_abnormal(self, start_row=2):
        abnormal, headers = [], [self.ws.cell(row=1, column=c).value for c in range(1, self.ws.max_column + 1)]
        for r in range(start_row, self.last_row() + 1):
            if self.is_abnormal(r):
                abnormal.append({headers[c - 1] if headers[c - 1] else f"列{c}": self.ws.cell(row=r, column=c).value
                                 for c in range(1, self.ws.max_column + 1)})
                abnormal[-1]["_原始行号"] = r
        return abnormal, headers

    def export_abnormal(self, out_path, start_row=2):
        abnormal, headers = self.get_abnormal(start_row)
        if not abnormal: return False, "没有异常行"
        wb = self._ox.Workbook(); ws = wb.active; ws.title = "异常行"
        outh = ["原始行号"] + [h for h in headers if h]
        for col, h in enumerate(outh, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = self._ft(bold=True, color="FFFFFF")
            c.fill = self._pf(start_color="dc3545", end_color="dc3545", fill_type="solid")
            c.alignment = self._al(horizontal="center", vertical="center")
        for ri, rd in enumerate(abnormal, 2):
            ws.cell(row=ri, column=1, value=rd.get("_原始行号", ""))
            for ci, h in enumerate(outh[1:], 2):
                ws.cell(row=ri, column=ci, value=rd.get(h, ""))
        for i, h in enumerate(outh, 1):
            ws.column_dimensions[self._gcl(i)].width = min(max(len(str(h)) * 1.5 + 2, 10), 40)
        wb.save(out_path); return True, f"已导出 {len(abnormal)} 条异常行"

    def clear_forward(self, row):
        for c in range(Config.COLS["gcj_lng"], Config.COLS["std_addr"] + 1):
            self.ws.cell(row=row, column=c).value = None
            self.ws.cell(row=row, column=c).fill = self.nf

    def write_forward(self, row, result):
        if not result:
            for c in range(Config.COLS["gcj_lng"], Config.COLS["std_addr"] + 1):
                self.ws.cell(row=row, column=c).fill = self.yf
            self.ws.cell(row=row, column=Config.COLS["gcj_lng"]).value = "无匹配"
            self.ws.cell(row=row, column=Config.COLS["gcj_lat"]).value = "无匹配"
            self.ws.cell(row=row, column=Config.COLS["wgs_lng"]).value = "无匹配"
            self.ws.cell(row=row, column=Config.COLS["wgs_lat"]).value = "无匹配"
            self.ws.cell(row=row, column=Config.COLS["std_addr"]).value = "解析失败"; return
        gl, gt = result["gcj_lng"], result["gcj_lat"]
        self.ws.cell(row=row, column=Config.COLS["gcj_lng"]).value = gl
        self.ws.cell(row=row, column=Config.COLS["gcj_lat"]).value = gt
        if 73 <= gl <= 135 and 3 <= gt <= 53:
            wl, wt = gcj02_to_wgs84(gl, gt)
            self.ws.cell(row=row, column=Config.COLS["wgs_lng"]).value = round(wl, 6)
            self.ws.cell(row=row, column=Config.COLS["wgs_lat"]).value = round(wt, 6)
        else:
            self.ws.cell(row=row, column=Config.COLS["wgs_lng"]).value = "超出范围"
            self.ws.cell(row=row, column=Config.COLS["wgs_lat"]).value = "超出范围"
            result["is_abnormal"] = True
        final = f"{result['std_addr']} [精度:{result['level']}][来源:{result['source']}]"
        if result.get("is_drift"): final += " 疑似漂移"
        self.ws.cell(row=row, column=Config.COLS["std_addr"]).value = final
        fill = self.yf if result.get("is_abnormal") else self.nf
        for c in range(Config.COLS["gcj_lng"], Config.COLS["std_addr"] + 1):
            self.ws.cell(row=row, column=c).fill = fill

    def clear_reverse(self, row):
        for c in range(Config.COLS["rego_full"], Config.COLS["rego_street"] + 1):
            self.ws.cell(row=row, column=c).value = None

    def write_reverse(self, row, result):
        if not result:
            self.ws.cell(row=row, column=Config.COLS["rego_full"]).value = "解析失败"; return
        self.ws.cell(row=row, column=Config.COLS["rego_full"]).value = result.get("full", "")
        self.ws.cell(row=row, column=Config.COLS["rego_div"]).value = result.get("division", "")
        self.ws.cell(row=row, column=Config.COLS["rego_street"]).value = result.get("street", "")

    def save(self): self.wb.save(self.fp)

    def close(self):
        if self.wb:
            self.wb.close()
            self.wb = None


# ═══════════════════════════════════════════════════════════
# 7. 批量处理模块
# ═══════════════════════════════════════════════════════════
class BatchProcessor:
    def __init__(self): self.geo = None; self.stop_flag = False
    def stop(self): self.stop_flag = True; self.geo.stop() if self.geo else None
    def init(self, cfg):
        self.geo = GeoCoder(cfg)
        self.stop_flag = False
        ok, msg = self.geo.check()
        if not ok or msg:
            return False, msg or "密钥验证失败"
        return True, ""

    def forward(self, proc, srow, lrow, aidx, a2idx, cidx, didx, skip, only_ab, 
                do_clean, do_merge, do_dedup, log_cb, prog_cb):
        total = lrow - srow + 1; success = fail = skipc = a2u = drift = bdup = 0
        # 地址去重预处理
        # 地址去重预处理（清洗逻辑与实际处理完全一致）
        dedup_map = {}
        result_cache = {}
        if do_dedup:
            log_cb("[预处理] 正在分析地址去重...\n", "info")
            vals = []
            for i in range(srow, lrow + 1):
                a1 = str(proc.ws.cell(row=i, column=aidx).value or "").strip()
                a2 = str(proc.ws.cell(row=i, column=a2idx).value or "").strip()
                city = str(proc.ws.cell(row=i, column=cidx).value or "").strip()
                dist = str(proc.ws.cell(row=i, column=didx).value or "").strip()
                # 与实际处理一致的清洗流程（含区县拼接）
                a1 = build_query_addr(a1, a2, city, dist, do_clean, do_merge)
                vals.append(a1)
            for i in range(len(vals)):
                if (srow + i) in dedup_map: continue
                dedup_map[srow + i] = 0
                for j in range(i + 1, len(vals)):
                    if (srow + j) in dedup_map: continue
                    sim = _sim_addr(vals[i], vals[j])
                    if sim >= 0.85:
                        dedup_map[srow + j] = srow + i
            dup_count = sum(1 for v in dedup_map.values() if v != 0)
            log_cb(f"[预处理] 发现 {dup_count} 个重复地址，将复用查询结果\n", "info")
        for i in range(srow, lrow + 1):
            try:
                if self.stop_flag: log_cb("\n⚠ 用户取消\n", "warning"); break
                pv = i - srow + 1; prog_cb(pv, total)
                if skip and proc.is_normal(i): skipc += 1; log_cb(f"第 {i} 行：跳过\n", "skip"); continue
                if only_ab and not proc.is_abnormal(i): skipc += 1; continue
                proc.clear_forward(i)
                # 去重复用
                rep = dedup_map.get(i, 0)
                if rep != 0 and rep in result_cache:
                    res = result_cache[rep]
                    proc.write_forward(i, res)
                    success += 1
                    log_cb(f"第 {i} 行：复用第 {rep} 行结果（去重）\n", "skip")
                    continue
                city = str(proc.ws.cell(row=i, column=cidx).value or "").strip()
                a1 = str(proc.ws.cell(row=i, column=aidx).value or "").strip()
                a2 = str(proc.ws.cell(row=i, column=a2idx).value or "").strip()
                dist = str(proc.ws.cell(row=i, column=didx).value or "").strip()
                # 清洗 + 合并 + 市/区县拼接（与去重预处理、清洗预览完全一致）
                a1 = build_query_addr(a1, a2, city, dist, do_clean, do_merge)
                if not a1: proc.write_forward(i, None); fail += 1; log_cb(f"第 {i} 行：地址为空\n", "error"); continue
                res, ab, err, fatal, blog = self.geo.resolve(a1, "", city, dist)
                if fatal: log_cb(f"致命错误：{err}，已终止\n", "error"); break
                # 检查主地址结果：精度不足时才做内容匹配校验
                # 避免高精度结果（门址/POI/村庄/街道）因字符串不完全匹配而误触发地址2
                if res and not ab and res.get("source") == "高德" and a2:
                    if level_score(res.get("level", "未知")) < level_score("街道"):
                        std_lower = res["std_addr"].lower()
                        a1_clean = _clean_single_addr(a1)
                        a2_clean = _clean_single_addr(a2)
                        match_a1 = a1_clean and (a1_clean.lower() in std_lower or std_lower in a1_clean.lower())
                        match_a2 = a2_clean and (a2_clean.lower() in std_lower or std_lower in a2_clean.lower())
                        if not match_a1 and not match_a2:
                            ab = True; res["is_abnormal"] = True; res["std_addr"] += " [内容不匹配]"
                # 地址2备用策略：主地址失败 或 标记异常 时，尝试地址2
                orig_res = res
                if (not res) or ab:
                    if a2:
                        r2, a22, e2, f2, blog2 = self.geo.resolve(a2, "", city, dist)
                        if f2: log_cb(f"致命错误：{e2}，已终止\n", "error"); break
                        if r2 and not a22:
                            # 地址2备用策略：
                            # ① 地址1完全失败 → 无条件采用地址2
                            # ② 地址2精度严格高于地址1 → 直接采用地址2（不比较距离）
                            # ③ 地址2精度与地址1相同 → 比较坐标距离（≤200米才采用）
                            # ④ 地址2精度低于地址1 → 保留地址1
                            r2_score = level_score(r2.get("level", "未知"))
                            res_score = level_score(res.get("level", "未知")) if res else 0
                            should_replace = False
                            if not res:
                                should_replace = True  # 地址1完全失败，无条件使用地址2
                                log_cb(f"[地址2备用] 地址1完全失败，采用地址2结果\n", "info")
                            elif r2_score > res_score:
                                # 地址2精度严格更高 → 直接采用，不比较距离
                                should_replace = True
                                log_cb(f"[地址2备用] 地址2精度更高({r2.get('level','')}/{r2_score} > {res.get('level','')}/{res_score})，直接采用地址2结果\n", "info")
                            elif r2_score == res_score:
                                # 同精度 → 比较坐标距离（200米内才采用）
                                if res.get("gcj_lng") and res.get("gcj_lat") and r2.get("gcj_lng") and r2.get("gcj_lat"):
                                    try:
                                        dx = abs(r2["gcj_lng"] - res["gcj_lng"]) * 111000 * math.cos(math.radians(res["gcj_lat"]))
                                        dy = abs(r2["gcj_lat"] - res["gcj_lat"]) * 111000
                                        dist = math.sqrt(dx*dx + dy*dy)
                                        if dist <= 200:
                                            should_replace = True
                                            log_cb(f"[地址2备用] 同精度({r2.get('level','')}/{r2_score})，距离{int(dist)}米≤200米，采用地址2结果\n", "info")
                                        else:
                                            log_cb(f"[地址2备用] 同精度({r2.get('level','')}/{r2_score})，距离{int(dist)}米>200米，保留地址1结果\n", "info")
                                    except:
                                        should_replace = True  # 距离计算异常，回退到采用地址2
                                        log_cb(f"[地址2备用] 同精度距离计算异常，回退采用地址2结果\n", "warning")
                                else:
                                    # 无坐标无法比较距离，同精度时默认采用地址2
                                    should_replace = True
                                    log_cb(f"[地址2备用] 同精度({r2.get('level','')}/{r2_score})，无坐标对比距离，采用地址2结果\n", "info")
                            else:
                                # 地址2精度低于地址1 → 保留地址1
                                log_cb(f"[地址2备用] 地址2精度({r2.get('level','')}/{r2_score})低于地址1({res.get('level','')}/{res_score})，保留地址1结果\n", "info")

                            if should_replace:
                                res = r2; a2u += 1
                                drift += 1 if r2.get("is_drift") else 0
                                bdup += 1 if r2.get("source") == "百度" else 0
                                if blog2: log_cb(blog2, "baidu")
                                log_cb(f"[地址2备用] 地址2查询成功: {r2.get('std_addr','')} [精度:{r2.get('level','')}]\n", "success")
                        else:
                            res = None
                    else:
                        res = None
                    # 修复：主地址仅“疑似漂移”且坐标有效时，保留结果并标黄展示，
                    # 避免被覆盖为“解析失败”导致坐标与地址丢失
                    if res is None and orig_res and orig_res.get("is_drift"):
                        res = orig_res
                        drift += 1
                        log_cb(f"[漂移保留] 结果疑似漂移（保留并标黄）：{orig_res.get('std_addr','')} [精度:{orig_res.get('level','')}]\n", "warning")
                proc.write_forward(i, res)
                if do_dedup and res:
                    result_cache[i] = res
                if res: success += 1; tag = "success"
                else: fail += 1; tag = "error"
                if blog: log_cb(blog, "baidu")
                log_cb(f"第 {i} 行：{'成功' if res else '失败'}\n", tag)
                time.sleep(self.geo.interval)
            except Exception as e:
                log_cb(f"第 {i} 行处理异常: {str(e)[:80]}\n", "error")
                fail += 1
                continue
        return {"total": total, "success": success, "fail": fail, "skip": skipc,
                "addr2_used": a2u, "drift_cnt": drift, "baidu_up": bdup,
                "gaode_count": self.geo.gcnt, "baidu_count": self.geo.bcnt}

    def reverse(self, proc, srow, lrow, glidx, gtidx, wlidx, wtidx, skip, log_cb, prog_cb):
        total = lrow - srow + 1; success = fail = skipc = 0; gerr = False
        for i in range(srow, lrow + 1):
            try:
                if self.stop_flag: log_cb("\n⚠ 用户取消\n", "warning"); break
                pv = i - srow + 1; prog_cb(pv, total)
                if skip:
                    ex = str(proc.ws.cell(row=i, column=Config.COLS["rego_full"]).value or "").strip()
                    if ex and "失败" not in ex and "错误" not in ex: skipc += 1; log_cb(f"第 {i} 行：跳过\n", "skip"); continue
                proc.clear_reverse(i)
                gl = proc.ws.cell(row=i, column=glidx).value; gt = proc.ws.cell(row=i, column=gtidx).value
                try: gl = float(gl) if gl else 0; gt = float(gt) if gt else 0
                except: gl = gt = 0
                if not (gl and gt and 73 <= gl <= 135 and 3 <= gt <= 53):
                    wl = proc.ws.cell(row=i, column=wlidx).value if wlidx > 0 else 0
                    wt = proc.ws.cell(row=i, column=wtidx).value if wtidx > 0 else 0
                    try: wl = float(wl) if wl else 0; wt = float(wt) if wt else 0
                    except: wl = wt = 0
                    if wl and wt and 73 <= wl <= 135 and 3 <= wt <= 53: gl, gt = wgs84_to_gcj02(wl, wt)
                    else: proc.write_reverse(i, None); fail += 1; log_cb(f"第 {i} 行：无有效坐标\n", "error"); continue
                if not (gl and gt):
                    proc.ws.cell(row=i, column=Config.COLS["rego_full"]).value = "无有效坐标"; fail += 1; continue
                res, err = self.geo.gaode_regeo(gl, gt)
                if not res:
                    proc.ws.cell(row=i, column=Config.COLS["rego_full"]).value = f"失败：{err}"
                    if "10001" in err or "10003" in err: gerr = True; fail += 1; log_cb(f"第 {i} 行：全局错误 {err}\n"); break
                    fail += 1; log_cb(f"第 {i} 行：失败 - {err}\n")
                else: proc.write_reverse(i, res); success += 1; log_cb(f"第 {i} 行：成功\n", "success")
                time.sleep(self.geo.interval)
            except Exception as e:
                log_cb(f"第 {i} 行处理异常: {str(e)[:80]}\n", "error")
                fail += 1
                continue
        return {"total": total, "success": success, "fail": fail, "skip": skipc,
                "global_error": gerr, "gaode_count": self.geo.gcnt}


# ═══════════════════════════════════════════════════════════
# 8. UI 组件模块
# ═══════════════════════════════════════════════════════════
class SButton(tk.Button):
    def __init__(self, master, text, command, btn_type="primary", **kw):
        colors = {"primary": (Config.C_ACCENT, "#fff", "#b45309"), "secondary": (Config.C_INFO, "#fff", "#1d4ed8"),
                  "success": (Config.C_SUCCESS, "#fff", "#047857"), "danger": (Config.C_DANGER, "#fff", "#b91c1c"),
                  "warning": ("#f59e0b", "#fff", "#d97706")}
        bg, fg, act = colors.get(btn_type, colors["primary"])
        super().__init__(master, text=text, command=command, font=("微软雅黑", 10, "bold"), bg=bg, fg=fg, bd=0,
                         padx=18, pady=8, cursor="hand2", relief=tk.FLAT, activebackground=act, activeforeground="#fff", **kw)
        self._nb = bg
        self.bind("<Enter>", lambda e: self.config(bg=act))
        self.bind("<Leave>", lambda e: self.config(bg=self._nb))

class Card(tk.Frame):
    def __init__(self, master, title=None, accent=Config.C_ACCENT, **kw):
        super().__init__(master, bg=Config.C_CARD, highlightbackground=Config.C_BORDER, highlightthickness=1, **kw)
        if accent: tk.Frame(self, bg=accent, height=3).pack(fill=tk.X)
        if title:
            tf = tk.Frame(self, bg=Config.C_CARD); tf.pack(fill=tk.X, padx=10, pady=(6, 0))
            tk.Label(tf, text=title, font=("微软雅黑", 11, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
            tk.Frame(self, bg=Config.C_BORDER, height=1).pack(fill=tk.X, padx=10, pady=6)

class Collapsible(tk.Frame):
    def __init__(self, master, title, expanded=False, tf=None, tfg=Config.C_ACCENT, **kw):
        super().__init__(master, bg=Config.C_CARD, **kw)
        self.expanded = expanded
        self.header = tk.Frame(self, bg=Config.C_CARD, cursor="hand2")
        self.header.pack(fill=tk.X); self.header.bind("<Button-1>", self._toggle)
        self.arrow = tk.Label(self.header, text="▼" if expanded else "▶", font=("微软雅黑", 9, "bold"),
                              bg=Config.C_CARD, fg=tfg, width=2)
        self.arrow.pack(side=tk.LEFT, padx=(10, 4)); self.arrow.bind("<Button-1>", self._toggle)
        self.tl = tk.Label(self.header, text=title, font=tf or ("微软雅黑", 12, "bold"), bg=Config.C_CARD, fg=tfg)
        self.tl.pack(side=tk.LEFT, padx=(0, 8), pady=8); self.tl.bind("<Button-1>", self._toggle)
        tk.Frame(self, bg=Config.C_BORDER, height=1).pack(fill=tk.X, padx=10)
        self.content = tk.Frame(self, bg=Config.C_CARD)
        if expanded: self.content.pack(fill=tk.X, padx=10, pady=(6, 10))
        for w in (self.header, self.arrow, self.tl):
            w.bind("<Enter>", lambda e: self.header.config(bg="#f8fafc"))
            w.bind("<Leave>", lambda e: self.header.config(bg=Config.C_CARD))
    def _toggle(self, e=None):
        self.expanded = not self.expanded; self.arrow.config(text="▼" if self.expanded else "▶")
        self.content.pack(fill=tk.X, padx=10, pady=(6, 10)) if self.expanded else self.content.pack_forget()
    def cf(self): return self.content

class PreviewManager:
    def __init__(self, parent, notebook, tab_text="📋 数据预览", refresh_cmd=None):
        self.notebook = notebook; self.page_size = 20; self.data = []; self.headers = []; self.page = 0; self.tp = 0
        self.refresh_cmd = refresh_cmd
        self.frame = tk.Frame(notebook, bg=Config.C_CARD); notebook.add(self.frame, text=f"  {tab_text}  ")
        tb = tk.Frame(self.frame, bg=Config.C_CARD); tb.pack(fill=tk.X, padx=10, pady=(8, 4))
        self.info = tk.Label(tb, text="请先导入 Excel 文件", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_L)
        self.info.pack(side=tk.LEFT)
        if refresh_cmd:
            rb = tk.Button(tb, text="🔄 刷新", font=("微软雅黑", 9), bg=Config.C_BORDER, fg=Config.C_TEXT,
                           bd=0, padx=12, pady=4, cursor="hand2", relief=tk.FLAT, command=refresh_cmd)
            rb.pack(side=tk.RIGHT)
            rb.bind("<Enter>", lambda e: rb.config(bg="#cbd5e1"))
            rb.bind("<Leave>", lambda e: rb.config(bg=Config.C_BORDER))
        tf = tk.Frame(self.frame, bg=Config.C_CARD); tf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))
        self.tree = ttk.Treeview(tf, show="headings", selectmode="browse")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview); vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview); hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.bind("<MouseWheel>", lambda e: self.tree.yview_scroll(int(-1*(e.delta/120)), "units"))
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(int(-1*(e.delta/120)), "units"))
        pf = tk.Frame(self.frame, bg=Config.C_CARD); pf.pack(fill=tk.X, padx=10, pady=(0, 8))
        self.pl = tk.Label(pf, text="第 0 / 0 页", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT)
        self.pl.pack(side=tk.LEFT)
        tk.Button(pf, text="◀ 上一页", font=("微软雅黑", 9), bg=Config.C_BORDER, fg=Config.C_TEXT, bd=0, padx=12, pady=4,
                  cursor="hand2", relief=tk.FLAT, command=self.prev).pack(side=tk.LEFT, padx=(12, 4))
        tk.Button(pf, text="下一页 ▶", font=("微软雅黑", 9), bg=Config.C_BORDER, fg=Config.C_TEXT, bd=0, padx=12, pady=4,
                  cursor="hand2", relief=tk.FLAT, command=self.next).pack(side=tk.LEFT, padx=4)

    def load(self, headers, data):
        self.headers, self.data = headers, data; self.page = 0
        self.tp = max(1, (len(data)+self.page_size-1)//self.page_size) if data else 0
        self.refresh(); self.info.config(text=f"共 {len(data)} 行 | {len(headers)} 列（预览前 {min(len(data),200)} 行）")

    def refresh(self):
        for c in self.tree["columns"]: self.tree.heading(c, text="")
        self.tree.delete(*self.tree.get_children())
        if not self.data:
            self.tree["columns"] = ("提示",); self.tree.heading("提示", text="暂无数据")
            self.tree.column("提示", width=400); self.pl.config(text="第 0 / 0 页"); return
        dh = ["序号"] + self.headers; self.tree["columns"] = dh
        self.tree.column("序号", width=55, anchor="center"); self.tree.heading("序号", text="序号")
        for h in dh[1:]:
            self.tree.column(h, width=120, anchor="w", minwidth=80, stretch=False); self.tree.heading(h, text=h)
        si = self.page * self.page_size; ei = min(si + self.page_size, len(self.data))
        for i in range(si, ei):
            vals = [i+1] + [str(self.data[i].get(h, ""))[:27]+"..." if len(str(self.data[i].get(h,"")))>30 else str(self.data[i].get(h,"")) for h in dh[1:]]
            self.tree.insert("", tk.END, values=vals)
        self.pl.config(text=f"第 {self.page+1} / {self.tp} 页（共 {len(self.data)} 行）")
        self.tree.update_idletasks()

    def next(self):
        if self.page < self.tp - 1: self.page += 1; self.refresh()
    def prev(self):
        if self.page > 0: self.page -= 1; self.refresh()


# ═══════════════════════════════════════════════════════════
# 9. 主应用模块
# ═══════════════════════════════════════════════════════════
class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{Config.APP_NAME} v{Config.VERSION}")
        self.root.geometry("1400x900"); self.root.minsize(1320, 820); self.root.state("zoomed")
        self.cfg = Config.load(); self.geo = GeoCoder(self.cfg)
        self.running = False; self.stop_flag = False; self.is_updating = False; self._map_loading = False
        self.poi_running = False
        self._last = {}; self.bp = BatchProcessor()
        self.df = None  # 延迟创建字体
        # 使用期限检查（2028-01-01）
        self.is_expired = datetime.datetime.now() > datetime.datetime(2028, 1, 1)
        self._build(); self.show("home")
        self.root.after(1000, lambda: threading.Thread(target=self._auto_check, daemon=True).start())
        if self.is_expired:
            self.root.after(1500, self._show_expire_notice)

    # ---------- 导航与页面框架 ----------
    def _build(self):
        # 延迟初始化全局字体
        if self.df is None:
            self.df = tkfont.Font(family="微软雅黑", size=10)
            self.root.option_add("*Font", self.df)
        main = tk.Frame(self.root, bg=Config.C_BG); main.pack(fill=tk.BOTH, expand=True)
        nav = tk.Frame(main, bg=Config.C_PRIMARY, width=210); nav.pack(side=tk.LEFT, fill=tk.Y); nav.pack_propagate(False)
        # Logo
        lf = tk.Frame(nav, bg=Config.C_PRIMARY, height=110); lf.pack(fill=tk.X, pady=(20, 15))
        tk.Label(lf, text="◉", font=("微软雅黑", 28), bg=Config.C_PRIMARY, fg=Config.C_ACCENT).pack(pady=(6, 0))
        tk.Label(lf, text="地址转换", font=("微软雅黑", 18, "bold"), bg=Config.C_PRIMARY, fg=Config.C_ACCENT).pack(pady=(2, 0))
        tk.Label(lf, text="标准化工具", font=("微软雅黑", 10), bg=Config.C_PRIMARY, fg="#8892b0").pack()
        # 导航按钮
        self.navs = {}
        items = [("home", "首页", "🏠"), ("single", "单条查询", "📍"), ("batch_fwd", "批量正向编码", "📊"),
                 ("batch_rev", "批量逆向编码", "🔄"), ("converter", "坐标转换", "🧭"), ("poi", "POI 采集", "📍"), ("tools", "更多工具", "🧰"), ("settings", "API 设置", "⚙️")]
        for key, text, icon in items:
            btn = tk.Button(nav, text=f"{icon}   {text}", font=("微软雅黑", 11), bg=Config.C_PRIMARY, fg="#94a3b8",
                           activebackground=Config.C_SECONDARY, activeforeground=Config.C_ACCENT, bd=0, padx=20, pady=10,
                           anchor="w", cursor="hand2", relief=tk.FLAT, command=lambda k=key: self.show(k))
            btn.pack(fill=tk.X, pady=(3, 3), padx=(10, 10))
            self.navs[key] = btn
        tk.Frame(nav, bg=Config.C_PRIMARY).pack(fill=tk.BOTH, expand=True)
        # 检查更新按钮
        self.ubtn = tk.Button(nav, text="🔍 检查更新", font=("微软雅黑", 10), bg=Config.C_SECONDARY, fg="#e2e8f0",
                              activebackground=Config.C_PRIMARY, activeforeground=Config.C_ACCENT, bd=0, padx=18, pady=6,
                              cursor="hand2", relief=tk.FLAT, command=self._check_update)
        self.ubtn.pack(fill=tk.X, padx=10, pady=(0, 4))
        # 状态 + 进度条（整合到左侧栏）
        self.status = tk.StringVar(value="就绪")
        sf = tk.Frame(nav, bg=Config.C_SECONDARY, height=42); sf.pack(fill=tk.X, padx=10, pady=(0, 4))
        tk.Label(sf, textvariable=self.status, font=("微软雅黑", 9), bg=Config.C_SECONDARY, fg="#94a3b8").pack(side=tk.LEFT, padx=8, pady=4)
        self.upd_prog = ttk.Progressbar(sf, orient=tk.HORIZONTAL, mode="determinate", length=80)
        self.upd_prog.pack(side=tk.RIGHT, padx=8, pady=4)
        # 版本号
        tk.Label(nav, text=f"v{Config.VERSION}", font=("微软雅黑", 9), bg=Config.C_PRIMARY, fg="#64748b").pack(pady=(0, 6))
        # 内容区 — 去掉 padx/pady，让内容真正填满
        self.content = tk.Frame(main, bg=Config.C_BG); self.content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # 内容区框架（页面延迟构建）
        self.pages = {}
        self._page_builders = {
            "home": self._build_home,
            "single": self._build_single,
            "batch_fwd": lambda: self._build_batch("fwd"),
            "batch_rev": lambda: self._build_batch("rev"),
            "converter": self._build_converter,
            "poi": self._build_poi,
            "tools": self._build_tools,
            "settings": self._build_settings,
        }

    def _nav_style(self, active):
        for k, btn in self.navs.items():
            if k == active: btn.config(bg=Config.C_SECONDARY, fg=Config.C_ACCENT, font=("微软雅黑", 11, "bold"))
            else: btn.config(bg=Config.C_PRIMARY, fg="#94a3b8", font=("微软雅黑", 11))

    def show(self, key):
        if self._check_expired(): return
        # 延迟构建：首次切换到某页面时才创建
        if key not in self.pages and key in self._page_builders:
            self._page_builders[key]()
        self._nav_style(key)
        for p in self.pages.values(): p.pack_forget()
        self.pages[key].pack(fill=tk.BOTH, expand=True)
        self.cur = key

    def _page(self): return tk.Frame(self.content, bg=Config.C_BG)

    def _get_start_row(self, direction):
        try: return int(getattr(self, f"{direction}_start").get())
        except: return 2

    def _bind_hover(self, btn, bg, hover_bg):
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))

    def _col_idx(self, headers, col_name, default_col):
        if not col_name: return default_col
        try: return headers.index(col_name) + 1
        except ValueError: return default_col

    def _pick_header(self, headers, chosen, *keywords):
        """列名识别：优先用户选择，其次按关键词模糊匹配表头（预览与批量处理共用，保证一致）"""
        if chosen and chosen in headers: return chosen
        if chosen:
            for h in headers:
                if chosen in h: return h
        for kw in keywords:
            for h in headers:
                if kw in h: return h
        return ""

    def _log_cfg(self, w):
        for tag, color in [("success", Config.C_SUCCESS), ("error", Config.C_DANGER), ("warning", Config.C_WARNING),
                           ("info", Config.C_INFO), ("skip", Config.C_TEXT_L), ("baidu", "#ca8a04")]:
            w.tag_config(tag, foreground=color)

    def _log(self, direction, msg, tag="normal"):
        getattr(self, f"{direction}_log").insert(tk.END, msg, tag); getattr(self, f"{direction}_log").see(tk.END)

    # ---------- 首页 ----------
    def _build_home(self):
        f = self._page(); self.pages["home"] = f
        tf = tk.Frame(f, bg=Config.C_BG); tf.pack(fill=tk.X, pady=(0, 4))
        tk.Label(tf, text="欢迎使用地址转换工具", font=("微软雅黑", 24, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w")
        tk.Label(tf, text="基于高德/百度地图 API 的地址智能解析与坐标转换", font=("微软雅黑", 12), bg=Config.C_BG, fg=Config.C_TEXT_L).pack(anchor="w", pady=(4, 0))
        cf = tk.Frame(f, bg=Config.C_BG); cf.pack(fill=tk.X, pady=(8, 6))
        items = [("📍 单条查询", "快速查询单个地址的 GCJ-02 / WGS-84 坐标\n支持地址查询与经纬度逆编码", Config.C_INFO),
                 ("📊 批量正向编码", "从 Excel 读取地址，批量获取\n标准化结果与坐标", Config.C_SUCCESS),
                 ("🔄 批量逆向编码", "从 Excel 读取坐标，批量解析\n为标准地址信息", Config.C_WARNING),
                 ("🧭 坐标转换", "GCJ-02 / WGS-84 / BD-09\n三系坐标互转，支持单点与批量", "#7c3aed"),
                 ("⚙️ 双平台配置", "支持高德 + 百度 API\n智能 Fallback 自动切换", Config.C_DANGER)]
        for i, (t, d, c) in enumerate(items):
            cd = tk.Frame(cf, bg=Config.C_CARD, highlightbackground=Config.C_BORDER, highlightthickness=1)
            cd.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0 if i==0 else 8, 0))
            tk.Frame(cd, bg=c, height=4).pack(fill=tk.X)
            tk.Label(cd, text=t, font=("微软雅黑", 12, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w", padx=12, pady=(10, 4))
            tk.Label(cd, text=d, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_L, justify=tk.LEFT).pack(anchor="w", padx=12, pady=(0, 10))
        steps = Collapsible(f, title="▎使用步骤", expanded=True, tf=("微软雅黑", 13, "bold"), tfg=Config.C_ACCENT)
        steps.pack(fill=tk.X, pady=(6, 6))
        for s in ["1. 点击左侧「API 设置」，配置高德/百度 API Key（自动保存到 addr_config.json）",
                  "2. 准备 Excel 文件，确保包含 Sheet1，按下方表格格式填写数据",
                  "3. 选择「批量正向编码」或「批量逆向编码」，选择文件后开始处理",
                  "4. 结果自动写入 Excel，异常行黄色高亮；支持断点续传与失败重试"]:
            tk.Label(steps.cf(), text=s, font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT_M).pack(anchor="w", padx=6, pady=3)
        fmt = Collapsible(f, title="▎Excel 数据格式说明", expanded=True, tf=("微软雅黑", 13, "bold"), tfg=Config.C_SUCCESS)
        fmt.pack(fill=tk.X, pady=(0, 6))
        db = tk.Button(fmt.header, text="下载模板", font=("微软雅黑", 10, "bold"), bg=Config.C_ACCENT, fg="#fff",
                       bd=0, padx=14, pady=4, cursor="hand2", relief=tk.FLAT, command=self._dl_template)
        db.pack(side=tk.RIGHT, padx=14); db.bind("<Enter>", lambda e: db.config(bg="#b45309")); db.bind("<Leave>", lambda e: db.config(bg=Config.C_ACCENT))
        tk.Label(fmt.cf(), text="▎正向编码（地址 → 坐标）", font=("微软雅黑", 11, "bold"), bg=Config.C_CARD, fg=Config.C_SUCCESS).pack(anchor="w", padx=6, pady=(6, 5))
        self._mk_table(fmt.cf(), [["序号","省(选填)","市(选填)","县/区(选填)","地址1(必填)","地址2(选填)","GCJ-02\n经度","GCJ-02\n纬度","WGS-84\n经度","WGS-84\n纬度","标准地址"],
                                  ["输入","输入","输入","输入","输入","输入","输出","输出","输出","输出","输出"]], [6,8,8,8,16,14,10,10,10,10,24])
        tk.Label(fmt.cf(), text="▎逆向编码（坐标 → 地址）", font=("微软雅黑", 11, "bold"), bg=Config.C_CARD, fg=Config.C_WARNING).pack(anchor="w", padx=6, pady=(14, 4))
        self._mk_table(fmt.cf(), [["GCJ-02\n经度","GCJ-02\n纬度","WGS-84\n经度","WGS-84\n纬度","逆地理\n完整地址","省 - 市 - 区","街道\n门牌号"],
                                  ["输入","输入","输入(备选)","输入(备选)","输出","输出","输出"]], [11,11,11,11,18,18,12])

    def _mk_table(self, parent, rows, widths):
        tf = tk.Frame(parent, bg=Config.C_CARD); tf.pack(fill=tk.X, padx=4, pady=4)
        for ri, row in enumerate(rows):
            if ri > 0: tk.Frame(tf, bg=Config.C_BORDER, height=1).pack(fill=tk.X)
            bgc, fgc, font = (Config.C_SECONDARY, "white", ("微软雅黑", 10, "bold")) if ri == 0 else (("#f8fafc" if ri%2==0 else Config.C_CARD), Config.C_TEXT_M, ("微软雅黑", 10))
            rc = tk.Frame(tf, bg=bgc); rc.pack(fill=tk.X)
            for cell, w in zip(row, widths):
                tk.Label(rc, text=cell, font=font, bg=bgc, fg=fgc, width=w, anchor="center").pack(side=tk.LEFT, padx=1, pady=3)

    def _dl_template(self):
        if self._check_expired(): return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="地址标准化模板.xlsx")
        if not path: return
        try:
            ox, pf, ft, al, gcl = _get_openpyxl()
            wb = ox.Workbook(); ws = wb.active; ws.title = "Sheet1"
            hs = ["序号","省","市","区县","地址1(必填)","地址2(选填)","GCJ-02经度","GCJ-02纬度","WGS-84经度","WGS-84纬度","标准地址","完整地址","省市区","街道门牌号"]
            for col, h in enumerate(hs, 1):
                c = ws.cell(row=1, column=col, value=h); c.font = ft(bold=True, color="FFFFFF")
                c.fill = pf(start_color="1e293b", end_color="1e293b", fill_type="solid"); c.alignment = al(horizontal="center", vertical="center")
            for ri, row in enumerate([[1,"四川省","成都市","","西南电子科技大学","",""]*2+[""]*4,
                                       [2,"北京市","北京市","","天安门","人民英雄纪念碑",""]*2+[""]*4,
                                       [3,"云南省","曲靖市","麒麟区","曲靖市农村信用合作社联合社","寥廓北路中段271-1号",""]*2+[""]*4], 2):
                for ci, v in enumerate(row, 1): ws.cell(row=ri, column=ci, value=v)
            for i, w in enumerate([8,12,12,12,25,20,15,15,15,15,40,30,20,20], 1): ws.column_dimensions[gcl(i)].width = w
            wb.save(path); messagebox.showinfo("成功", f"模板已保存到：\n{path}")
        except Exception as e: messagebox.showerror("错误", f"保存失败：{e}")

    # ---------- 单条查询 ----------
    def _build_single(self):
        f = self._page(); self.pages["single"] = f
        sp = tk.Frame(f, bg=Config.C_BG); sp.pack(fill=tk.BOTH, expand=True)
        sp.columnconfigure(0, weight=35, minsize=400); sp.columnconfigure(1, weight=65, minsize=700); sp.rowconfigure(0, weight=1)
        left = tk.Frame(sp, bg=Config.C_BG); left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        # 参数卡片
        param = Card(left, title=None, accent=None); param.pack(fill=tk.X, pady=(0, 8))
        inp = tk.Frame(param, bg=Config.C_CARD); inp.pack(fill=tk.X, padx=14, pady=2)
        # 地址输入
        self.af = tk.Frame(inp, bg=Config.C_CARD); self.af.pack(fill=tk.X)
        self.albl = tk.Label(self.af, text="* 地址（必填）", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT); self.albl.pack(anchor="w")
        aef = tk.Frame(self.af, bg=Config.C_CARD); aef.pack(fill=tk.X, pady=(4, 0))
        self.saddr = tk.Entry(aef, font=("微软雅黑", 11), relief=tk.SOLID, bd=1); self.saddr.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=0)
        self.saddr.bind("<Return>", lambda e: self._do_single())
        self.saddr.insert(0, "麒麟区雄业金都")
        tk.Button(aef, text="×", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT_L, bd=0, cursor="hand2",
                  command=lambda: self.saddr.delete(0, tk.END)).pack(side=tk.RIGHT, padx=(4, 0))
        # 坐标输入
        self.cff = tk.Frame(inp, bg=Config.C_CARD)
        tk.Label(self.cff, text="经纬度", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
        tk.Label(self.cff, text="经度在前,纬度在后，支持中英文逗号、空格、顿号分割,如：103.791246,25.491221", font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_TEXT_L).pack(anchor="w")
        self.scoord = tk.Entry(self.cff, font=("微软雅黑", 11), relief=tk.SOLID, bd=1); self.scoord.pack(fill=tk.X, pady=(4, 0), ipady=0)
        self.scoord.bind("<Return>", lambda e: self._do_single())
        self.scoord.insert(0, "103.791246,25.491221")
        sf = tk.Frame(self.cff, bg=Config.C_CARD); sf.pack(fill=tk.X, pady=(6, 0))
        tk.Label(sf, text="坐标系：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.csys = tk.StringVar(value="GCJ-02")
        for v, t in [("GCJ-02","GCJ-02"),("WGS-84","WGS-84"),("BD-09","BD-09")]:
            tk.Radiobutton(sf, text=t, variable=self.csys, value=v, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(side=tk.LEFT, padx=(0, 12))
        # 按钮
        bf = tk.Frame(param, bg=Config.C_CARD); bf.pack(fill=tk.X, padx=14, pady=(0, 12))
        self.squery_btn = SButton(bf, "查询", self._do_single, "primary")
        self.squery_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
        mf = tk.Frame(bf, bg=Config.C_CARD); mf.pack(side=tk.RIGHT, padx=(10, 0))
        self.smode = tk.StringVar(value="address")
        tk.Radiobutton(mf, text="根据地址", variable=self.smode, value="address", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT,
                       selectcolor=Config.C_CARD, command=self._switch_mode).pack(side=tk.LEFT, padx=(0, 10))
        tk.Radiobutton(mf, text="根据经纬度", variable=self.smode, value="coordinate", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT,
                       selectcolor=Config.C_CARD, command=self._switch_mode).pack(side=tk.LEFT)
        # 结果卡片
        rc = Card(left, title="标准化地址", accent=Config.C_ACCENT); rc.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        rh = tk.Frame(rc, bg=Config.C_CARD); rh.pack(fill=tk.X, padx=14, pady=(10, 6))
        tk.Label(rh, text="查询结果", font=("微软雅黑", 13, "bold"), bg=Config.C_CARD, fg=Config.C_ACCENT).pack(side=tk.LEFT)
        SButton(rh, "📋 复制全部", self._copy_results, "secondary").pack(side=tk.RIGHT)
        tk.Frame(rc, bg=Config.C_BORDER, height=1).pack(fill=tk.X, padx=14, pady=(0, 6))
        self.sres = {}
        for key, label in [("std_addr","标准化地址"),("gcj_lng","GCJ-02经度"),("gcj_lat","GCJ-02纬度"),("wgs_loc","WGS-84经纬度"),
                           ("province","省份"),("city_r","城市"),("district","区县"),("township","乡镇/街道"),("street","道路名称")]:
            r = tk.Frame(rc, bg=Config.C_CARD); r.pack(fill=tk.X, padx=10, pady=2)
            tk.Label(r, text=label, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_M, width=11, anchor="e").pack(side=tk.LEFT)
            e = tk.Entry(r, font=("微软雅黑", 10), relief=tk.SOLID, bd=1, state="readonly", bg="#f8fafc"); e.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, ipady=0)
            self.sres[key] = e
        self.serr = tk.Label(rc, text="", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_DANGER, wraplength=360)
        self.serr.pack(fill=tk.X, padx=10, pady=(2, 2))
        tk.Label(rc, text="查询日志", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w", padx=10, pady=(4, 2))
        tk.Frame(rc, bg=Config.C_BORDER, height=1).pack(fill=tk.X, padx=10, pady=(0, 3))
        rf = tk.Frame(rc, bg=Config.C_CARD); rf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 4))
        self.raw = scrolledtext.ScrolledText(rf, font=("Consolas", 10), wrap=tk.WORD, height=10, bg="#f8fafc", fg=Config.C_TEXT, relief=tk.SOLID, bd=1)
        self.raw.pack(fill=tk.BOTH, expand=True)
        # 右侧地图
        right = tk.Frame(sp, bg=Config.C_BG); right.grid(row=0, column=1, sticky="nsew"); right.rowconfigure(0, weight=1); right.columnconfigure(0, weight=1)
        mc = tk.Frame(right, bg="#e2e8f0", highlightbackground=Config.C_BORDER, highlightthickness=1); mc.pack(fill=tk.BOTH, expand=True)
        self.mlbl = tk.Label(mc, bg="#e2e8f0"); self.mlbl.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        self.mlbl.image = None
        self.mlbl.bind("<MouseWheel>", self._map_scroll)
        self.mlbl.bind("<Button-4>", lambda e: self._map_zoom(1))
        self.mlbl.bind("<Button-5>", lambda e: self._map_zoom(-1))
        self.zlvl = 16; self.mlng = 116.397428; self.mlat = 39.90923
        # 地图控件
        zf = tk.Frame(mc, bg="#ffffff", highlightthickness=0); zf.place(x=12, y=12)
        for txt, cmd in [("+", self._map_zoom_in), ("−", self._map_zoom_out)]:
            tk.Button(zf, text=txt, font=("微软雅黑", 12, "bold"), bg="#ffffff", fg="#475569", bd=0, width=2, height=1,
                      relief=tk.FLAT, activebackground="#f1f5f9", command=cmd).pack(pady=(0, 1))
        pf = tk.Frame(mc, bg="#ffffff", highlightthickness=0); pf.place(relx=1.0, y=12, anchor="ne", x=-12)
        tk.Button(pf, text="▲", font=("微软雅黑", 9), bg="#ffffff", fg="#475569", bd=0, width=3, height=1, relief=tk.FLAT,
                  activebackground="#f1f5f9", command=lambda: self._pan("up")).pack()
        mr = tk.Frame(pf, bg="#ffffff"); mr.pack(fill=tk.X)
        tk.Button(mr, text="◀", font=("微软雅黑", 9), bg="#ffffff", fg="#475569", bd=0, width=3, height=1, relief=tk.FLAT,
                  activebackground="#f1f5f9", command=lambda: self._pan("left")).pack(side=tk.LEFT)
        tk.Label(mr, text="", font=("微软雅黑", 8), bg="#ffffff", width=1).pack(side=tk.LEFT)
        tk.Button(mr, text="▶", font=("微软雅黑", 9), bg="#ffffff", fg="#475569", bd=0, width=3, height=1, relief=tk.FLAT,
                  activebackground="#f1f5f9", command=lambda: self._pan("right")).pack(side=tk.LEFT)
        tk.Button(pf, text="▼", font=("微软雅黑", 9), bg="#ffffff", fg="#475569", bd=0, width=3, height=1, relief=tk.FLAT,
                  activebackground="#f1f5f9", command=lambda: self._pan("down")).pack()
        self.scl = tk.Label(mc, text="", font=("微软雅黑", 9), bg="#ffffff", fg="#475569", bd=0, relief=tk.FLAT, padx=6, pady=3)
        self.scl.place(relx=1.0, rely=1.0, anchor="se", x=-12, y=-12)
        self.root.after(1500, lambda: self._load_map(self.mlng, self.mlat))

    def _map_scroll(self, e):
        if e.delta > 0: self._map_zoom_in()
        else: self._map_zoom_out()
    def _map_zoom(self, d):
        self._map_zoom_in() if d > 0 else self._map_zoom_out()
    def _map_zoom_in(self):
        if self.zlvl < 19: self.zlvl += 1; self._load_map(self.mlng, self.mlat)
    def _map_zoom_out(self):
        if self.zlvl > 3: self.zlvl -= 1; self._load_map(self.mlng, self.mlat)
    def _pan(self, d):
        w = max(self.mlbl.winfo_width(), 600); h = max(self.mlbl.winfo_height(), 500)
        sx, sy = w // 4, h // 4; dx = {"left": -sx, "right": sx}.get(d, 0); dy = {"up": -sy, "down": sy}.get(d, 0)
        mw = 256.0 * (2 ** self.zlvl); lpp = 360.0 / mw
        self.mlng = max(-180, min(180, self.mlng + dx * lpp))
        self.mlat = max(-85, min(85, self.mlat - dy * lpp * math.cos(math.radians(self.mlat))))
        self._load_map(self.mlng, self.mlat)

    def _switch_mode(self):
        if self.smode.get() == "address":
            self.albl.config(text="* 地址（必填）"); self.af.pack(fill=tk.X); self.cff.pack_forget()
        else:
            self.albl.config(text="经纬度"); self.af.pack_forget(); self.cff.pack(fill=tk.X)
        self._clear_single()

    def _clear_single(self):
        self.saddr.delete(0, tk.END); self.saddr.insert(0, "麒麟区雄业金都")
        self.scoord.delete(0, tk.END); self.scoord.insert(0, "103.791246,25.491221")
        for e in self.sres.values(): e.config(state="normal"); e.delete(0, tk.END); e.config(state="readonly")
        self.serr.config(text=""); self.raw.delete("1.0", tk.END)
        self.mlng, self.mlat = 116.397428, 39.90923; self.zlvl = 16; self._load_map(self.mlng, self.mlat); self._last = {}

    def _copy_results(self):
        if self._check_expired(): return
        if not self._last: messagebox.showinfo("提示", "暂无查询结果"); return
        lines = ["="*45, "地址坐标转换结果", "="*45]
        for k, l in [("std_addr","标准化地址"),("gcj_lng","GCJ-02经度"),("gcj_lat","GCJ-02纬度"),("wgs_loc","WGS-84经纬度"),
                     ("province","省份"),("city_r","城市"),("district","区县"),("township","乡镇/街道"),("street","道路名称")]:
            v = self._last.get(k, ""); 
            if v: lines.append(f"{l}: {v}")
        lines.append("="*45)
        self.root.clipboard_clear(); self.root.clipboard_append("\n".join(lines))
        self.status.set("结果已复制"); self.root.after(2000, lambda: self.status.set("就绪"))

    def _parse_coord(self, s):
        """解析坐标字符串，支持英文逗号、中文逗号、空格、顿号等分隔符"""
        import re
        parts = re.split(r'[，,、\s]+', s.strip())
        if len(parts) != 2:
            raise ValueError("坐标格式错误，请使用「经度,纬度」格式（支持中文逗号、空格分隔）")
        try:
            lng = float(parts[0])
            lat = float(parts[1])
            return lng, lat
        except ValueError:
            raise ValueError("坐标必须为数字")

    def _do_single(self):
        if self._check_expired(): return
        mode = self.smode.get()
        if not self.cfg.get("gaode_key"): self.serr.config(text="请先配置高德 API Key"); return
        # 防并发：禁用查询按钮
        if hasattr(self, 'squery_btn'):
            self.squery_btn.config(state=tk.DISABLED)
        self.geo = GeoCoder(self.cfg); self.status.set("正在查询..."); self.serr.config(text="")
        for e in self.sres.values(): e.config(state="normal"); e.delete(0, tk.END); e.config(state="readonly")
        self.raw.delete("1.0", tk.END); self._last = {}
        def task():
            try:
                ok, msg = self.geo.check()
                if not ok:
                    self.root.after(0, lambda: self.serr.config(text=f"密钥错误：{msg}"))
                    self.root.after(0, lambda: self.status.set("就绪"))
                    self.root.after(0, lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
                    return
                if mode == "address":
                    addr = self.saddr.get().strip()
                    if not addr:
                        self.root.after(0, lambda: self.serr.config(text="请输入地址"))
                        self.root.after(0, lambda: self.status.set("就绪"))
                        self.root.after(0, lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
                        return
                    res, ab, err, fatal, blog = self.geo.resolve(addr, "", "", "")
                    self.root.after(0, lambda: self._show_addr(res, ab, err, fatal, blog))
                else:
                    try:
                        lng, lat = self._parse_coord(self.scoord.get().strip())
                    except Exception as e:
                        self.root.after(0, lambda: self.serr.config(text=f"坐标格式错误：{e}"))
                        self.root.after(0, lambda: self.status.set("就绪"))
                        self.root.after(0, lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
                        return
                    gl, gt = {"WGS-84": wgs84_to_gcj02, "BD-09": bd09_to_gcj02, "GCJ-02": lambda x,y:(x,y)}[self.csys.get()](lng, lat)
                    self._do_regeo(gl, gt)
            except Exception as e:
                self.root.after(0, lambda: self.serr.config(text=f"查询异常：{e}"))
                self.root.after(0, lambda: self.status.set("就绪"))
                self.root.after(0, lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
        threading.Thread(target=task, daemon=True).start()
    def _show_addr(self, res, ab, err, fatal, blog):
        self.status.set("就绪")
        # 恢复查询按钮状态
        if hasattr(self, 'squery_btn'):
            self.squery_btn.config(state=tk.NORMAL)
        if fatal: self.serr.config(text=f"查询失败（致命错误）：{err}"); self._build_log(None, ab, err, fatal, blog); return
        if not res: self.serr.config(text=f"未找到结果：{err}"); self._build_log(None, ab, err, fatal, blog); return
        gl, gt = res["gcj_lng"], res["gcj_lat"]; self.mlng, self.mlat = gl, gt
        wl, wt = gcj02_to_wgs84(gl, gt)
        fields = {"gcj_lng": str(gl), "gcj_lat": str(gt), "wgs_loc": f"{round(wl,6)},{round(wt,6)}", "std_addr": res["std_addr"]}
        self._last.update(fields)
        for k, v in fields.items():
            if k in self.sres: self.sres[k].config(state="normal"); self.sres[k].delete(0, tk.END); self.sres[k].insert(0, str(v)); self.sres[k].config(state="readonly")
        self._build_log(res, ab, err, fatal, blog); self._load_map(gl, gt)
        def regeo_task():
            rr, _ = self.geo.gaode_regeo(gl, gt)
            if rr: self.root.after(0, lambda: self._fill_regeo(rr))
        threading.Thread(target=regeo_task, daemon=True).start()

    def _build_log(self, res, ab, err, fatal, blog):
        self.raw.delete("1.0", tk.END); lines = ["="*50, "查询日志", "="*50]
        if fatal: lines.append(f"[致命错误] {err}"); self.raw.insert(tk.END, "\n".join(lines)); return
        if not res: lines.append(f"[查询失败] {err or '未找到结果'}"); self.raw.insert(tk.END, "\n".join(lines)); return
        lines += [f"[高德查询] 地址: {res.get('std_addr','')}", f"[高德查询] 精度: {res.get('level','未知')}",
                  f"[高德查询] 坐标: {res.get('gcj_lng','')}, {res.get('gcj_lat','')}"]
        if res.get('is_drift'): lines.append("[高德查询] 警告: 疑似地址漂移")
        if blog: lines.append("-"*50); lines += [l.strip() for l in blog.strip().split("\n") if l.strip()]
        lines.append("-"*50); lines.append(f"[最终结果] 状态: {'异常' if ab else '成功'}")
        lines += [f"[最终结果] 来源: {res.get('source','高德')}", f"[最终结果] 精度: {res.get('level','未知')}", f"[最终结果] 标准化地址: {res.get('std_addr','')}"]
        lines.append("="*50); self.raw.insert(tk.END, "\n".join(lines))

    def _fill_regeo(self, rr):
        for k, v in [("province", rr.get("province","")), ("city_r", rr.get("city","")), ("district", rr.get("district","")),
                     ("township", rr.get("township","")), ("street", rr.get("street",""))]:
            if k in self.sres: self.sres[k].config(state="normal"); self.sres[k].delete(0, tk.END); self.sres[k].insert(0, str(v)); self.sres[k].config(state="readonly"); self._last[k] = str(v)

    def _do_regeo(self, gl, gt):
        self.mlng, self.mlat = gl, gt
        def task():
            rr, err = self.geo.gaode_regeo(gl, gt)
            self.root.after(0, lambda: self._show_regeo(gl, gt, rr, err))
        threading.Thread(target=task, daemon=True).start()

    def _show_regeo(self, gl, gt, rr, err):
        self.status.set("就绪")
        # 恢复查询按钮状态
        if hasattr(self, 'squery_btn'):
            self.squery_btn.config(state=tk.NORMAL)
        if not rr: self.serr.config(text=f"逆地理编码失败：{err}"); return
        wl, wt = gcj02_to_wgs84(gl, gt)
        fields = {"gcj_lng": str(gl), "gcj_lat": str(gt), "wgs_loc": f"{round(wl,6)},{round(wt,6)}", "province": rr.get("province",""),
                  "city_r": rr.get("city",""), "district": rr.get("district",""), "township": rr.get("township",""),
                  "street": rr.get("street",""), "std_addr": rr.get("full","")}
        self._last.update(fields)
        for k, v in fields.items():
            if k in self.sres: self.sres[k].config(state="normal"); self.sres[k].delete(0, tk.END); self.sres[k].insert(0, str(v)); self.sres[k].config(state="readonly")
        self.raw.delete("1.0", tk.END); self.raw.insert(tk.END, json.dumps(rr, ensure_ascii=False, indent=2)); self._load_map(gl, gt)

    def _load_map(self, lng, lat):
        # 防并发锁
        if self._map_loading:
            return
        self._map_loading = True
        def task():
            try:
                key = self.cfg.get("gaode_key", ""); zoom = self.zlvl; Image, ImageTk = _get_pil()
                rw = self.mlbl.winfo_width(); rh = self.mlbl.winfo_height()
                if rw < 100 or rh < 100: rw = max(self.root.winfo_width() - 640, 600); rh = max(self.root.winfo_height() - 120, 500)
                w, h = max(rw, 600), max(rh, 500)
                w = min(w, 1024); h = min(h, 1024)  # 高德静态地图最大尺寸限制
                if key:
                    try:
                        url = f"{Config.GAODE_STATIC}?location={lng},{lat}&zoom={zoom}&size={w}*{h}&markers=mid,0xFF0000,{lng},{lat}&key={key}"
                        ok, data = self._fetch_img(url)
                        if ok: img = Image.open(io.BytesIO(data)); self.root.after(0, lambda img=img: self._set_img(img)); self.root.after(0, self._upd_scale); return
                    except: pass
                bak = self.cfg.get("baidu_ak", "")
                if bak:
                    try:
                        bl, bt = gcj02_to_bd09(lng, lat)
                        url = f"{Config.BAIDU_STATIC}?ak={bak}&center={bl},{bt}&width={w}&height={h}&zoom={zoom}&markers={bl},{bt}"
                        ok, data = self._fetch_img(url)
                        if ok: img = Image.open(io.BytesIO(data)); self.root.after(0, lambda img=img: self._set_img(img)); self.root.after(0, self._upd_scale); return
                    except: pass
                em = "地图加载失败\n"
                if not key: em += "未配置高德 API Key\n"
                else: em += "高德静态地图服务异常\n"
                if not bak: em += "未配置百度 AK\n"
                em += "\n可能原因：\n1. 高德Key类型不正确\n2. 静态地图API未开通或配额超限\n3. 网络连接问题"
                self.root.after(0, lambda: self.mlbl.config(text=em))
            finally:
                self._map_loading = False
        threading.Thread(target=task, daemon=True).start()
    def _fetch_img(self, url):
        ctx = ssl.create_default_context(); ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            ct = resp.headers.get('Content-Type', ''); data = resp.read()
            if b'application/json' in ct.encode() or data.startswith(b'{'): return False, None
            return True, data

    def _set_img(self, img):
        try:
            Image, ImageTk = _get_pil(); w = max(self.mlbl.winfo_width(), 600); h = max(self.mlbl.winfo_height(), 500)
            # 兼容新旧 Pillow
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.LANCZOS
            img = img.resize((w, h), resample)
            photo = ImageTk.PhotoImage(img)
            self.mlbl.config(image=photo, text=""); self.mlbl.image = photo
        except Exception as e: self.mlbl.config(text=f"图片处理失败: {str(e)[:80]}")

    def _upd_scale(self):
        sm = {3:"1000 km",4:"500 km",5:"200 km",6:"100 km",7:"50 km",8:"20 km",9:"10 km",10:"5 km",11:"2 km",12:"1 km",
              13:"500 m",14:"200 m",15:"100 m",16:"50 m",17:"25 m",18:"10 m",19:"5 m"}
        self.scl.config(text=sm.get(self.zlvl, f"zoom {self.zlvl}"))

    # ---------- 批量编码（通用） ----------
    def _build_batch(self, direction):
        is_fwd = direction == "fwd"
        f = self._page(); self.pages[f"batch_{direction}"] = f
        tk.Label(f, text="批量正向地理编码（地址 -> 坐标）" if is_fwd else "批量逆向地理编码（坐标 -> 地址）",
                 font=("微软雅黑", 16, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w", pady=(0, 6))
        sp = tk.Frame(f, bg=Config.C_BG); sp.pack(fill=tk.BOTH, expand=True)
        sp.columnconfigure(0, weight=0, minsize=300); sp.columnconfigure(1, weight=1, minsize=700); sp.rowconfigure(0, weight=1)
        # 左侧滚动面板
        lo = tk.Frame(sp, bg=Config.C_CARD, highlightbackground=Config.C_BORDER, highlightthickness=1)
        lo.grid(row=0, column=0, sticky="nsew", padx=(0, 6)); lo.grid_propagate(False); lo.config(width=320)
        lo.grid_rowconfigure(0, weight=1); lo.grid_columnconfigure(0, weight=1)
        canvas = tk.Canvas(lo, bg=Config.C_CARD, highlightthickness=0); canvas.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(lo, orient="vertical", command=canvas.yview); vsb.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=vsb.set)
        outer_li = tk.Frame(canvas, bg=Config.C_CARD)
        cw = canvas.create_window((0, 0), window=outer_li, anchor="nw", width=300)
        outer_li.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        def _mw(e): canvas.yview_scroll(int(-1*(e.delta/120)), "units"); return "break"
        canvas.bind("<MouseWheel>", _mw); outer_li.bind("<MouseWheel>", _mw)
        inner_li = tk.Frame(outer_li, bg=Config.C_CARD); inner_li.pack(fill=tk.BOTH, expand=True, padx=8, pady=0)
        # 内容
        li = inner_li
        tk.Label(li, text="▎第一步：导入文件", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_ACCENT).pack(anchor="w", pady=(0, 2))
        pv = tk.StringVar()
        tk.Entry(li, textvariable=pv, font=("微软雅黑", 10), relief=tk.SOLID, bd=1, state="readonly", bg="#f8fafc").pack(fill=tk.X, pady=(0, 3))
        setattr(self, f"{direction}_path", pv)
        ib = tk.Button(li, text="📂 导入 Excel 文件", font=("微软雅黑", 11, "bold"), bg=Config.C_ACCENT, fg="#fff", bd=0, padx=18, pady=8,
                       cursor="hand2", relief=tk.FLAT, command=lambda: self._import(direction))
        ib.pack(fill=tk.X, pady=(0, 4)); self._bind_hover(ib, Config.C_ACCENT, "#b45309")
        tk.Frame(li, bg=Config.C_BORDER, height=1).pack(fill=tk.X, pady=(0, 3))
        # ═══ 数据预处理（仅正向编码需要）═══
        if is_fwd:
            pre = Collapsible(li, title="▎数据预处理", expanded=False, tf=("微软雅黑", 11, "bold"), tfg=Config.C_SUCCESS)
            pre.pack(fill=tk.X, pady=(0, 6))
            pcf = pre.cf()
            # 使用卡片式子区域
            pinner = tk.Frame(pcf, bg="#f8fafc", highlightbackground=Config.C_BORDER, highlightthickness=1)
            pinner.pack(fill=tk.X, pady=(3, 4), padx=2)
            self.fwd_clean = tk.BooleanVar(value=False)
            tk.Checkbutton(pinner, text="自动清洗（去空格、统一符号、去重前缀）", variable=self.fwd_clean,
                           font=("微软雅黑", 9), bg="#f8fafc", fg=Config.C_TEXT, selectcolor="#f8fafc").pack(anchor="w", padx=8, pady=(6, 2))
            self.fwd_merge = tk.BooleanVar(value=False)
            tk.Checkbutton(pinner, text="智能合并（短区名 + 地址2 详细地址）", variable=self.fwd_merge,
                           font=("微软雅黑", 9), bg="#f8fafc", fg=Config.C_TEXT, selectcolor="#f8fafc").pack(anchor="w", padx=8, pady=0)
            self.fwd_dedup = tk.BooleanVar(value=False)
            tk.Checkbutton(pinner, text="地址去重（相似度≥85% 复用结果，节省配额）", variable=self.fwd_dedup,
                           font=("微软雅黑", 9), bg="#f8fafc", fg=Config.C_TEXT, selectcolor="#f8fafc").pack(anchor="w", padx=8, pady=(2, 6))
            pbf = tk.Frame(pcf, bg=Config.C_CARD)
            pbf.pack(fill=tk.X, pady=(2, 2))
            pb = tk.Button(pbf, text="👁 预览清洗效果", font=("微软雅黑", 9, "bold"), bg=Config.C_INFO, fg="#fff", bd=0, padx=12, pady=5,
                           cursor="hand2", relief=tk.FLAT, command=self._preview_clean)
            pb.pack(side=tk.LEFT); pb.bind("<Enter>", lambda e: pb.config(bg="#1d4ed8")); pb.bind("<Leave>", lambda e: pb.config(bg=Config.C_INFO))
            self.fwd_dedup_lbl = tk.Label(pbf, text="", font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_SUCCESS)
            self.fwd_dedup_lbl.pack(side=tk.LEFT, padx=(10, 0))
            tk.Frame(li, bg=Config.C_BORDER, height=1).pack(fill=tk.X, pady=(0, 3))
        tk.Label(li, text="▎第二步：指定数据列", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_ACCENT).pack(anchor="w", pady=(0, 2))
        cf = tk.Frame(li, bg=Config.C_CARD); cf.pack(fill=tk.X, pady=(0, 4))
        combos = [("* 地址1列（必填）", "fwd_addr_col"), ("* 地址2列（必填，备用）", "fwd_addr2_col"),
                  ("  城市列（可选）", "fwd_city_col"), ("  区县列（可选）", "fwd_dist_col")] if is_fwd else \
                 [("* GCJ-02 经度列", "rev_gcj_lng_col"), ("* GCJ-02 纬度列", "rev_gcj_lat_col"),
                  ("  WGS-84 经度列（备选）", "rev_wgs_lng_col"), ("  WGS-84 纬度列（备选）", "rev_wgs_lat_col")]
        for lbl, attr in combos:
            fg = Config.C_TEXT if "*" in lbl else Config.C_TEXT_L
            tk.Label(cf, text=lbl, font=("微软雅黑", 10), bg=Config.C_CARD, fg=fg).pack(anchor="w")
            cb = ttk.Combobox(cf, values=[], font=("微软雅黑", 10), width=30, state="readonly"); cb.pack(fill=tk.X, pady=(2, 4))
            setattr(self, attr, cb)
        tk.Frame(li, bg=Config.C_BORDER, height=1).pack(fill=tk.X, pady=(0, 3))
        tk.Label(li, text="▎第三步：处理参数", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_ACCENT).pack(anchor="w", pady=(0, 2))
        pg = tk.Frame(li, bg=Config.C_CARD); pg.pack(fill=tk.X, pady=(0, 4))
        tk.Label(pg, text="起始行：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).grid(row=0, column=0, sticky="w")
        ss = tk.Spinbox(pg, from_=2, to=100000, width=8, font=("微软雅黑", 10), relief=tk.SOLID, bd=1)
        ss.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=4); ss.delete(0, tk.END); ss.insert(0, "2")
        setattr(self, f"{direction}_start", ss)
        sv = tk.BooleanVar(value=True)
        tk.Checkbutton(li, text="跳过已有结果", variable=sv, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", pady=0)
        setattr(self, f"{direction}_skip", sv)
        if is_fwd:
            av = tk.BooleanVar(value=False)
            tk.Checkbutton(li, text="只重新处理异常行", variable=av, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", pady=0)
            setattr(self, "fwd_abnormal", av)
        av2 = tk.BooleanVar(value=self.cfg.get("auto_open_excel", False))
        tk.Checkbutton(li, text="完成后自动打开文件", variable=av2, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", pady=0)
        setattr(self, f"{direction}_auto_open", av2)
        tk.Frame(li, bg=Config.C_BORDER, height=1).pack(fill=tk.X, pady=(6, 5))
        tk.Label(li, text="▎第四步：开始转换", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_ACCENT).pack(anchor="w", pady=(0, 2))
        br1 = tk.Frame(li, bg=Config.C_CARD); br1.pack(fill=tk.X, pady=(0, 5))
        sb = tk.Button(br1, text="▶ 开始处理", font=("微软雅黑", 10, "bold"), bg=Config.C_ACCENT, fg=Config.C_PRIMARY, bd=0, padx=14, pady=6,
                       cursor="hand2", relief=tk.FLAT, command=lambda: self._start_batch(direction))
        sb.pack(side=tk.LEFT, fill=tk.X, expand=True); self._bind_hover(sb, Config.C_ACCENT, "#b45309")
        setattr(self, f"{direction}_start_btn", sb)
        xb = tk.Button(br1, text="⏹ 停止", font=("微软雅黑", 10, "bold"), bg=Config.C_DANGER, fg="white", bd=0, padx=14, pady=6,
                       cursor="hand2", relief=tk.FLAT, state=tk.DISABLED, command=self._stop_batch)
        xb.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True); setattr(self, f"{direction}_stop_btn", xb)
        tk.Frame(li, bg=Config.C_BORDER, height=1).pack(fill=tk.X, pady=(10, 6))
        tk.Label(li, text="▎结果导出", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_ACCENT).pack(anchor="w", pady=(0, 2))
        br2 = tk.Frame(li, bg=Config.C_CARD); br2.pack(fill=tk.X, pady=(0, 4))
        # 导出异常行 — 与「开始处理」同规格纯按钮
        eb = tk.Button(br2, text="📤  导出异常行", font=("微软雅黑", 10, "bold"), bg=Config.C_WARNING, fg="#fff",
                       bd=0, padx=14, pady=6, cursor="hand2", relief=tk.FLAT,
                       command=lambda: self._export_abnormal(direction))
        eb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6)); self._bind_hover(eb, Config.C_WARNING, "#d97706")
        # 导出KML — 与「开始处理」同规格纯按钮
        kb = tk.Button(br2, text="🌏  导出 KML", font=("微软雅黑", 10, "bold"), bg=Config.C_SUCCESS, fg="#fff",
                       bd=0, padx=14, pady=6, cursor="hand2", relief=tk.FLAT,
                       command=lambda: self._export_kml(direction))
        kb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0)); self._bind_hover(kb, Config.C_SUCCESS, "#047857")
        pf = tk.Frame(li, bg=Config.C_CARD); pf.pack(fill=tk.X, pady=(0, 3))
        prog = ttk.Progressbar(pf, orient=tk.HORIZONTAL, mode="determinate"); prog.pack(fill=tk.X)
        pl = tk.Label(pf, text="就绪", font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_TEXT_L); pl.pack(anchor="w", pady=(4, 0))
        setattr(self, f"{direction}_prog", prog); setattr(self, f"{direction}_prog_lbl", pl)
        # 右侧
        right = tk.Frame(sp, bg=Config.C_BG); right.grid(row=0, column=1, sticky="nsew"); right.rowconfigure(0, weight=1); right.columnconfigure(0, weight=1)
        nb = ttk.Notebook(right); nb.pack(fill=tk.BOTH, expand=True); setattr(self, f"{direction}_notebook", nb)
        pm = PreviewManager(right, nb, "📋 数据预览",
                              refresh_cmd=lambda d=direction: self._refresh_preview(d))
        setattr(self, f"{direction}_preview", pm)
        lf = tk.Frame(nb, bg=Config.C_CARD); nb.add(lf, text="  📊 进度预览  ")
        log = scrolledtext.ScrolledText(lf, font=("Consolas", 10), wrap=tk.WORD, height=12, bg="#f8fafc", fg=Config.C_TEXT, relief=tk.SOLID, bd=1, padx=6, pady=4)
        log.pack(fill=tk.BOTH, expand=True, padx=6, pady=4); self._log_cfg(log); setattr(self, f"{direction}_log", log)

    def _import(self, direction):
        if self._check_expired(): return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not path: return
        getattr(self, f"{direction}_path").set(path); self._load_preview(path, direction)

    def _load_preview(self, path, direction):
        try:
            ox = _get_openpyxl()[0]
            wb = ox.load_workbook(path, read_only=True, data_only=True)
            if "Sheet1" not in wb.sheetnames: messagebox.showwarning("提示", "Excel 中未找到 Sheet1"); wb.close(); return
            ws = wb["Sheet1"]
            hs = [str(c.value).strip() if c.value else f"列{c.column}" for c in ws[1]]
            data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(data) >= 200: break
                data.append({hs[i] if i < len(hs) else f"列{i+1}": (str(v) if v is not None else "") for i, v in enumerate(row)})
            wb.close()
            # 填充所有相关下拉框的选项
            if direction == "fwd":
                for attr in ["fwd_addr_col", "fwd_addr2_col", "fwd_city_col", "fwd_dist_col"]:
                    cb = getattr(self, attr, None)
                    if cb: cb["values"] = hs
            else:
                for attr in ["rev_gcj_lng_col", "rev_gcj_lat_col", "rev_wgs_lng_col", "rev_wgs_lat_col"]:
                    cb = getattr(self, attr, None)
                    if cb: cb["values"] = hs
            getattr(self, f"{direction}_preview").load(hs, data); self._auto_sel(hs, direction)
            getattr(self, f"{direction}_notebook").select(0); self.status.set(f"已加载：{os.path.basename(path)}")
        except Exception as e: messagebox.showerror("错误", f"读取失败：{e}")

    def _auto_sel(self, hs, direction):
        if direction == "fwd":
            am = a2m = cm = None
            for h in hs:
                hl = h.lower()
                if am is None and any(k in hl for k in ["地址1","地址一","addr1","address1","详细地址"]): am = h
                if a2m is None and any(k in hl for k in ["地址2","地址二","addr2","address2","备用地址"]): a2m = h
                if cm is None and any(k in hl for k in ["市","city","城市"]): cm = h
            if am: self.fwd_addr_col.set(am)
            if a2m: self.fwd_addr2_col.set(a2m)
            if cm: self.fwd_city_col.set(cm)
        else:
            glm = gtm = wlm = wtm = None
            for h in hs:
                hl = h.lower()
                if glm is None and any(k in hl for k in ["gcj-02经度","gcj02经度","gcj经度","火星经度"]): glm = h
                if gtm is None and any(k in hl for k in ["gcj-02纬度","gcj02纬度","gcj纬度","火星纬度"]): gtm = h
                if wlm is None and any(k in hl for k in ["wgs-84经度","wgs84经度","wgs经度","84经度"]): wlm = h
                if wtm is None and any(k in hl for k in ["wgs-84纬度","wgs84纬度","wgs纬度","84纬度"]): wtm = h
            if not glm:
                for h in hs:
                    if any(k in h.lower() for k in ["经度","lng","lon","x"]): glm = h; break
            if not gtm:
                for h in hs:
                    if any(k in h.lower() for k in ["纬度","lat","y"]): gtm = h; break
            if glm: self.rev_gcj_lng_col.set(glm)
            if gtm: self.rev_gcj_lat_col.set(gtm)
            if wlm: self.rev_wgs_lng_col.set(wlm)
            if wtm: self.rev_wgs_lat_col.set(wtm)

    def _show_expire_notice(self):
        """软件过期提示"""
        messagebox.showwarning("软件已过期",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "  本软件已达使用截止期限\n"
            "  截止日期：2028 年 1 月 1 日\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "继续使用或获取最新版本，请联系：\n"
            "📧  1178258589@qq.com\n\n"
            "感谢您的使用！")

    def _check_expired(self):
        """检查是否过期，过期则弹窗并返回 True"""
        if self.is_expired:
            self._show_expire_notice()
            return True
        return False

    def _refresh_preview(self, direction):
        """刷新数据预览：重新读取 Excel 文件并更新预览"""
        if self._check_expired(): return
        path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先导入 Excel 文件")
            return
        try:
            self._load_preview(path, direction)
            getattr(self, f"{direction}_notebook").select(0)
            self.status.set("预览已刷新")
        except Exception as e:
            messagebox.showerror("错误", f"刷新失败：{e}")

    def _export_abnormal(self, direction):
        if self._check_expired(): return
        path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请先选择有效的 Excel 文件"); return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                              initialfile=f"异常行_{'正向' if direction=='fwd' else '逆向'}编码.xlsx")
        if not out: return
        def task():
            try:
                proc = ExcelProcessor(path); ok, msg = proc.export_abnormal(out, self._get_start_row(direction))
                self.root.after(0, lambda: messagebox.showinfo("导出结果", msg) if ok else messagebox.showwarning("导出失败", msg))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", f"导出异常：{e}"))
            finally:
                try:
                    if 'proc' in locals() and proc is not None:
                        proc.close()
                except: pass
        threading.Thread(target=task, daemon=True).start()

    def _preview_clean(self):
        """预览地址清洗效果（严格遵循 do_clean / do_merge 开关）"""
        if self._check_expired(): return
        path = self.fwd_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先导入 Excel 文件"); return
        if not self.fwd_clean.get() and not self.fwd_merge.get():
            messagebox.showinfo("提示", "请先勾选「自动清洗」或「智能合并」后再预览"); return
        try:
            ox = _get_openpyxl()[0]
            wb = ox.load_workbook(path, read_only=True, data_only=True)
            if "Sheet1" not in wb.sheetnames: messagebox.showwarning("提示", "Excel 中未找到 Sheet1"); wb.close(); return
            ws = wb["Sheet1"]
            hs = [str(c.value).strip() if c.value else f"列{c.column}" for c in ws[1]]
            data = [{hs[i] if i < len(hs) else f"列{i+1}": (str(v) if v is not None else "") for i, v in enumerate(row)}
                    for row in list(ws.iter_rows(min_row=2, values_only=True))[:200]]
            wb.close()
            a1c = self._pick_header(hs, self.fwd_addr_col.get(), "地址1")
            a2c = self._pick_header(hs, self.fwd_addr2_col.get(), "地址2")
            ctc = self._pick_header(hs, self.fwd_city_col.get(), "市")
            dtc = self._pick_header(hs, self.fwd_dist_col.get(), "区县", "区/县", "县", "区")
            do_clean = self.fwd_clean.get()
            do_merge = self.fwd_merge.get()
            preview = []
            for i, row in enumerate(data[:50], 1):
                a1 = str(row.get(a1c, "") or "").strip()
                a2 = str(row.get(a2c, "") or "").strip()
                city = str(row.get(ctc, "") or "").strip()
                dist = str(row.get(dtc, "") or "").strip()
                orig = a1
                # 与实际批量处理完全一致的清洗+合并+市/区县拼接逻辑
                a1 = build_query_addr(a1, a2, city, dist, do_clean, do_merge)
                if a1 != orig: preview.append(f"第{i}行: '{orig}' → '{a1}'")
            tips = []
            if do_clean: tips.append("自动清洗")
            if do_merge: tips.append("智能合并")
            tips.append("行政区划拼接")
            header = f"【预览模式：{' + '.join(tips)}】\n"
            msg = header + "\n".join(preview[:30]) + (f"\n... 共 {len(preview)} 条修改" if len(preview) > 30 else "") if preview else header + "前50行无需清洗"
            messagebox.showinfo("清洗预览", msg)
        except Exception as e: messagebox.showerror("错误", f"预览失败：{e}")
    def _export_kml(self, direction):
        """导出KML文件"""
        if self._check_expired(): return
        path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请先选择有效的 Excel 文件"); return
        out = filedialog.asksaveasfilename(defaultextension=".kml", filetypes=[("KML文件", "*.kml")],
                                           initialfile=f"地址坐标结果_{time.strftime('%Y%m%d')}.kml")
        if not out: return
        def task():
            try:
                proc = ExcelProcessor(path)
                sr = self._get_start_row(direction)
                hs = [str(proc.ws.cell(row=1, column=c).value or "").strip() for c in range(1, proc.ws.max_column + 1)]
                if direction == "fwd":
                    addr_idx = self._col_idx(hs, self.fwd_addr_col.get(), Config.COLS["addr1"])
                    lr = proc.last_row(addr_idx)
                else:
                    # 逆向编码：用 GCJ-02 经度列找最后一行
                    lng_idx = self._col_idx(hs, self.rev_gcj_lng_col.get(), Config.COLS["gcj_lng"])
                    lat_idx = self._col_idx(hs, self.rev_gcj_lat_col.get(), Config.COLS["gcj_lat"])
                    lr = max(proc.last_row(lng_idx), proc.last_row(lat_idx))
                    addr_idx = 0  # 逆向编码不需要地址列作为KML名称
                count = _gen_kml(proc, sr, lr, addr_idx, out)
                self.root.after(0, lambda: messagebox.showinfo("导出成功", f"已导出 {count} 个坐标点到 KML\n文件：{out}"))
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("错误", f"导出失败：{e}"))
            finally:
                try:
                    if 'proc' in locals() and proc is not None:
                        proc.close()
                except: pass
        threading.Thread(target=task, daemon=True).start()


    def _start_batch(self, direction):
        if self._check_expired(): return
        is_fwd = direction == "fwd"; path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请选择有效的 Excel 文件"); return
        if not self.cfg.get("gaode_key"): messagebox.showwarning("提示", "请先配置高德 API Key"); self.show("settings"); return
        sr = self._get_start_row(direction)
        # 防御性检查：确保所有必要的 UI 组件已创建
        required_attrs = [f"{direction}_path", f"{direction}_start", f"{direction}_start_btn",
                          f"{direction}_stop_btn", f"{direction}_log", f"{direction}_prog",
                          f"{direction}_prog_lbl", f"{direction}_auto_open"]
        if is_fwd:
            required_attrs += ["fwd_addr_col", "fwd_addr2_col", "fwd_city_col", "fwd_dist_col",
                               "fwd_skip", "fwd_abnormal", "fwd_clean", "fwd_merge", "fwd_dedup"]
        else:
            required_attrs += ["rev_gcj_lng_col", "rev_gcj_lat_col", "rev_wgs_lng_col",
                               "rev_wgs_lat_col", "rev_skip"]
        missing = [a for a in required_attrs if not hasattr(self, a)]
        if missing:
            messagebox.showerror("内部错误",
                f"以下 UI 组件未初始化，请重新切换页面后重试：\n\n" +
                "\n".join(f"  • {a}" for a in missing))
            return
        self.running, self.stop_flag = True, False
        getattr(self, f"{direction}_start_btn").config(state=tk.DISABLED)
        getattr(self, f"{direction}_stop_btn").config(state=tk.NORMAL)
        getattr(self, f"{direction}_log").delete("1.0", tk.END)
        getattr(self, f"{direction}_log").insert(tk.END, f"[系统] 启动处理，文件: {os.path.basename(path)}\n")
        ao = getattr(self, f"{direction}_auto_open").get(); self.cfg["auto_open_excel"] = ao; Config.save(self.cfg)
        def task():
            # 先定义回调函数，确保后续日志可用
            def log_cb(msg, tag="normal"):
                self.root.after(0, lambda: self._log(direction, msg, tag))
            def prog_cb(val, total_val):
                self.root.after(0, lambda: (getattr(self, f"{direction}_prog_lbl").config(text=f"{val} / {total_val}"), getattr(self, f"{direction}_prog").config(value=val)))
            try:
                proc = ExcelProcessor(path)
                log_cb(f"[系统] Excel 读取成功，共 {proc.ws.max_row} 行 {proc.ws.max_column} 列\n", "info")
                hs = [str(proc.ws.cell(row=1, column=c).value or "").strip() for c in range(1, proc.ws.max_column + 1)]
                if is_fwd:
                    aidx = self._col_idx(hs, self._pick_header(hs, self.fwd_addr_col.get(), "地址1"), Config.COLS["addr1"])
                    a2idx = self._col_idx(hs, self._pick_header(hs, self.fwd_addr2_col.get(), "地址2"), Config.COLS["addr2"])
                    cidx = self._col_idx(hs, self._pick_header(hs, self.fwd_city_col.get(), "市"), Config.COLS["city"])
                    didx = self._col_idx(hs, self._pick_header(hs, self.fwd_dist_col.get(), "区县", "区/县", "县", "区"), Config.COLS["district"])
                    if not self.fwd_addr_col.get() or not self.fwd_addr2_col.get():
                        log_cb("[错误] 地址1列和地址2列必须同时指定，处理已终止\n", "error")
                        return
                    lr = proc.last_row(aidx)
                else:
                    lng_col = self.rev_gcj_lng_col.get()
                    lat_col = self.rev_gcj_lat_col.get()
                    if not lng_col or not lat_col:
                        log_cb("[错误] 请先选择 GCJ-02 经度列和纬度列，处理已终止\n", "error")
                        return
                    glidx = self._col_idx(hs, lng_col, Config.COLS["gcj_lng"])
                    gtidx = self._col_idx(hs, lat_col, Config.COLS["gcj_lat"])
                    # 校验列名是否真实存在于表头（防止回退到默认列导致读错数据）
                    if lng_col not in hs or lat_col not in hs:
                        log_cb(f"[错误] 所选列不存在于Excel表头，请重新选择（经度:{lng_col}, 纬度:{lat_col}）\n", "error")
                        return
                    wlidx = self._col_idx(hs, self.rev_wgs_lng_col.get(), 0) if self.rev_wgs_lng_col.get() else 0
                    wtidx = self._col_idx(hs, self.rev_wgs_lat_col.get(), 0) if self.rev_wgs_lat_col.get() else 0
                    lr = max(proc.last_row(glidx), proc.last_row(gtidx), proc.last_row(wlidx) if wlidx>0 else 1, proc.last_row(wtidx) if wtidx>0 else 1)
                    if lr < 2: lr = proc.ws.max_row
                total = lr - sr + 1
                if total <= 0: log_cb("提示：没有需要处理的数据\n", "info"); return
                log_cb("[系统] 正在验证 API 密钥...\n", "info")
                ok, msg = self.bp.init(self.cfg)
                if not ok:
                    log_cb(f"[密钥错误] {msg}\n", "error")
                    return
                log_cb("[系统] API 密钥验证通过\n", "info")
                self.root.after(0, lambda: getattr(self, f"{direction}_prog").config(maximum=total, value=0))
                if is_fwd:
                    do_clean = self.fwd_clean.get()
                    do_merge = self.fwd_merge.get()
                    do_dedup = self.fwd_dedup.get()
                    st = self.bp.forward(proc, sr, lr, aidx, a2idx, cidx, didx, 
                                         self.fwd_skip.get(), 
                                         (self.fwd_abnormal.get() if hasattr(self, "fwd_abnormal") else False),
                                         do_clean, do_merge, do_dedup,
                                         log_cb, prog_cb)
                    self.root.after(0, lambda: self._log(direction, f"\n{'='*45}\n处理完成统计\n{'='*45}\n总行数：{st['total']}\n成功：{st['success']}\n使用地址2：{st['addr2_used']}\n百度优化：{st['baidu_up']}\n疑似漂移：{st['drift_cnt']}\n失败：{st['fail']}\n跳过：{st['skip']}\n高德调用：{st['gaode_count']}\n百度调用：{st['baidu_count']}\n{'='*45}\n", "info"))
                else:
                    st = self.bp.reverse(proc, sr, lr, glidx, gtidx, wlidx, wtidx, self.rev_skip.get(), log_cb, prog_cb)
                    if st.get("global_error"): self.root.after(0, lambda: self._log(direction, "[全局错误] 已终止处理！\n", "error"))
                    else: self.root.after(0, lambda: self._log(direction, f"\n{'='*45}\n处理完成统计\n{'='*45}\n总行数：{st['total']}\n成功：{st['success']}\n失败：{st['fail']}\n跳过：{st['skip']}\n高德调用：{st['gaode_count']}\n{'='*45}\n", "info"))
                proc.save()
                if ao and os.path.exists(path): self._open(path)
            except Exception as e: self.root.after(0, lambda: self._log(direction, f"[处理异常] {str(e)}\n", "error"))
            finally:
                self.running = False
                try:
                    if "proc" in locals() and proc:
                        proc.close()
                except: pass
                self.root.after(0, lambda: (getattr(self, f"{direction}_start_btn").config(state=tk.NORMAL), getattr(self, f"{direction}_stop_btn").config(state=tk.DISABLED), self.status.set("就绪")))
        threading.Thread(target=task, daemon=True).start()

    def _stop_batch(self):
        self.stop_flag = True; self.bp.stop(); self.status.set("正在停止...")

    def _open(self, path):
        if self._check_expired(): return
        try:
            p = platform.system()
            if p == "Windows": os.startfile(path)
            elif p == "Darwin": subprocess.run(["open", path], check=False)
            else: subprocess.run(["xdg-open", path], check=False)
        except Exception as e: print(f"打开文件失败: {e}")

    # ---------- 坐标转换 ----------
    def _build_converter(self):
        f = self._page(); self.pages["converter"] = f
        tk.Label(f, text="坐标转换工具", font=("微软雅黑", 18, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w", pady=(0, 4))
        tk.Label(f, text="GCJ-02 / WGS-84 / BD-09 三系坐标互转，支持单点转换与批量Excel转换", font=("微软雅黑", 11), bg=Config.C_BG, fg=Config.C_TEXT_L).pack(anchor="w", pady=(0, 10))
        sc = Card(f, title="▎单点转换", accent=Config.C_INFO); sc.pack(fill=tk.X, pady=(0, 4))
        sf = tk.Frame(sc, bg=Config.C_CARD); sf.pack(fill=tk.X, padx=8, pady=5)
        rf = tk.Frame(sf, bg=Config.C_CARD); rf.pack(fill=tk.X, pady=(0, 4))
        tk.Label(rf, text="输入坐标：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_in = tk.Entry(rf, font=("微软雅黑", 11), width=30, relief=tk.SOLID, bd=1); self.cv_in.pack(side=tk.LEFT, padx=10, ipady=0); self.cv_in.insert(0, "103.791246,25.491221")
        tk.Label(rf, text="转换方式：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT, padx=(15, 0))
        self.cv_trans = ttk.Combobox(rf, values=["GCJ-02 → WGS-84","GCJ-02 → BD-09","WGS-84 → GCJ-02","WGS-84 → BD-09","BD-09 → GCJ-02","BD-09 → WGS-84"],
                                     font=("微软雅黑", 10), width=18, state="readonly")
        self.cv_trans.set("GCJ-02 → WGS-84"); self.cv_trans.pack(side=tk.LEFT, padx=10)
        SButton(rf, "  转换", self._do_cv_single, "primary").pack(side=tk.LEFT, padx=10)
        tk.Label(rf, text="格式：经度,纬度", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_L).pack(side=tk.LEFT, padx=(5, 0))
        of = tk.Frame(sf, bg=Config.C_CARD); of.pack(fill=tk.X, pady=(0, 5))
        tk.Label(of, text="转换结果：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_out = tk.Entry(of, font=("微软雅黑", 11), width=38, relief=tk.SOLID, bd=1, state="readonly", bg="#f8fafc"); self.cv_out.pack(side=tk.LEFT, padx=10, ipady=0)
        SButton(of, "📋 复制", self._copy_cv, "secondary").pack(side=tk.LEFT)
        self.cv_err = tk.Label(sf, text="", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_DANGER); self.cv_err.pack(anchor="w", pady=(6, 0))
        bc = Card(f, title="▎批量Excel转换", accent=Config.C_SUCCESS); bc.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        bf = tk.Frame(bc, bg=Config.C_CARD); bf.pack(fill=tk.X, padx=8, pady=5)
        fr = tk.Frame(bf, bg=Config.C_CARD); fr.pack(fill=tk.X, pady=(0, 4))
        tk.Label(fr, text="Excel 文件：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_bp = tk.StringVar()
        tk.Entry(fr, textvariable=self.cv_bp, font=("微软雅黑", 10), width=50, relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True, ipady=0)
        tk.Button(fr, text="浏览...", font=("微软雅黑", 10), bg=Config.C_BORDER, fg=Config.C_TEXT, bd=0, padx=15, pady=5, command=self._sel_cv_file).pack(side=tk.LEFT)
        tk.Button(fr, text="读取列", font=("微软雅黑", 10), bg=Config.C_INFO, fg="white", bd=0, padx=15, pady=5, command=self._load_cv_cols).pack(side=tk.LEFT, padx=6)
        self.cv_cf = tk.Frame(bf, bg=Config.C_CARD); self.cv_cf.pack(fill=tk.X, pady=(0, 10)); self.cv_cf.pack_forget()
        sg = tk.LabelFrame(self.cv_cf, text=" 源数据与转换方式 ", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT, bd=1, relief=tk.SOLID)
        sg.pack(fill=tk.X, pady=6)
        si = tk.Frame(sg, bg=Config.C_CARD); si.pack(fill=tk.X, padx=8, pady=6)
        tk.Label(si, text="源经度列：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_lng = ttk.Combobox(si, values=[], font=("微软雅黑", 10), width=16, state="readonly"); self.cv_lng.pack(side=tk.LEFT, padx=(0, 14))
        tk.Label(si, text="源纬度列：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_lat = ttk.Combobox(si, values=[], font=("微软雅黑", 10), width=16, state="readonly"); self.cv_lat.pack(side=tk.LEFT, padx=(0, 14))
        tk.Label(si, text="转换方式：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_mode = ttk.Combobox(si, values=["GCJ-02 → WGS-84","GCJ-02 → BD-09","WGS-84 → GCJ-02","WGS-84 → BD-09","BD-09 → GCJ-02","BD-09 → WGS-84"],
                                      font=("微软雅黑", 10), width=18, state="readonly")
        self.cv_mode.set("GCJ-02 → WGS-84"); self.cv_mode.pack(side=tk.LEFT)
        self.cv_mode.bind("<<ComboboxSelected>>", self._cv_mode_chg)
        og = tk.LabelFrame(self.cv_cf, text=" 输出设置（选择Excel中目标列） ", font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT, bd=1, relief=tk.SOLID)
        og.pack(fill=tk.X, pady=6)
        oi = tk.Frame(og, bg=Config.C_CARD); oi.pack(fill=tk.X, padx=8, pady=6)
        tk.Label(oi, text="输出经度列：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_olng = ttk.Combobox(oi, values=[], font=("微软雅黑", 10), width=20, state="readonly"); self.cv_olng.pack(side=tk.LEFT, padx=(0, 20))
        tk.Label(oi, text="输出纬度列：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
        self.cv_olat = ttk.Combobox(oi, values=[], font=("微软雅黑", 10), width=20, state="readonly"); self.cv_olat.pack(side=tk.LEFT)
        self.cv_sk = tk.BooleanVar(value=True)
        tk.Checkbutton(oi, text="跳过已有结果", variable=self.cv_sk, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(side=tk.LEFT, padx=(20, 10))
        self.cv_ao = tk.BooleanVar(value=self.cfg.get("auto_open_excel", False))
        tk.Checkbutton(oi, text="完成后自动打开文件", variable=self.cv_ao, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(side=tk.LEFT)
        bpr = tk.Frame(self.cv_cf, bg=Config.C_CARD); bpr.pack(fill=tk.X, pady=4)
        self.cv_sbtn = SButton(bpr, "  开始批量转换", self._start_cv_batch, "primary"); self.cv_sbtn.pack(side=tk.LEFT)
        self.cv_xbtn = SButton(bpr, "  停止", self._stop_cv_batch, "danger"); self.cv_xbtn.config(state=tk.DISABLED); self.cv_xbtn.pack(side=tk.LEFT, padx=10)
        self.cv_prog = ttk.Progressbar(bpr, orient=tk.HORIZONTAL, mode="determinate", length=300); self.cv_prog.pack(side=tk.LEFT, padx=(10, 0))
        self.cv_plbl = tk.Label(bpr, text="0 / 0", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT); self.cv_plbl.pack(side=tk.LEFT, padx=10)
        lgf = tk.Frame(bc, bg=Config.C_CARD); lgf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.cv_log = scrolledtext.ScrolledText(lgf, font=("Consolas", 10), wrap=tk.WORD, height=12, bg="#f8fafc", fg=Config.C_TEXT, relief=tk.SOLID, bd=1, padx=6, pady=4)
        self.cv_log.pack(fill=tk.BOTH, expand=True)

    def _sel_cv_file(self):
        if self._check_expired(): return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path: self.cv_bp.set(path); self._load_cv_cols()

    def _load_cv_cols(self):
        if self._check_expired(): return
        path = self.cv_bp.get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请选择有效的 Excel 文件"); return
        try:
            ox = _get_openpyxl()[0]
            wb = ox.load_workbook(path, read_only=True, data_only=True)
            if "Sheet1" not in wb.sheetnames: messagebox.showwarning("提示", "Excel 中未找到 Sheet1"); wb.close(); return
            ws = wb["Sheet1"]
            hs = [str(c.value).strip() if c.value else f"列{c.column}" for c in ws[1]]
            wb.close()
            for cb in [self.cv_lng, self.cv_lat, self.cv_olng, self.cv_olat]: cb["values"] = hs
            lm = lt = om = ot = None
            for h in hs:
                hl = h.lower()
                if lm is None and any(k in hl for k in ["经度","lng","lon","x"]): lm = h
                if lt is None and any(k in hl for k in ["纬度","lat","y"]): lt = h
                if om is None and any(k in hl for k in ["wgs84经度","wgs-84经度","84经度"]): om = h
                if ot is None and any(k in hl for k in ["wgs84纬度","wgs-84纬度","84纬度"]): ot = h
            if lm: self.cv_lng.set(lm)
            if lt: self.cv_lat.set(lt)
            if om: self.cv_olng.set(om)
            if ot: self.cv_olat.set(ot)
            self.cv_cf.pack(fill=tk.X, pady=(0, 10))
            self.cv_log.delete("1.0", tk.END); self.cv_log.insert(tk.END, f"已读取文件，共 {len(hs)} 列\n")
        except Exception as e: messagebox.showerror("错误", f"读取列失败：{e}")

    def _do_cv_single(self):
        if self._check_expired(): return
        self.cv_err.config(text=""); self.cv_out.config(state="normal"); self.cv_out.delete(0, tk.END); self.cv_out.config(state="readonly")
        try:
            lng, lat = self._parse_coord(self.cv_in.get().strip())
            trans = self.cv_trans.get()
            if "→" not in trans: raise ValueError("请选择有效的转换方式")
            sn, dn = [s.strip() for s in trans.split("→")]
            ss = sn.lower().replace("-","").replace("_",""); ds = dn.lower().replace("-","").replace("_","")
            ss = {"gcj02":"gcj02","wgs84":"wgs84","bd09":"bd09"}.get(ss, "gcj02")
            ds = {"gcj02":"gcj02","wgs84":"wgs84","bd09":"bd09"}.get(ds, "wgs84")
            if ss == ds: self.cv_err.config(text="源坐标系与目标坐标系相同"); return
            ol, ot = convert_coords(lng, lat, ss, ds)
            self.cv_out.config(state="normal"); self.cv_out.delete(0, tk.END); self.cv_out.insert(0, f"{round(ol,6)},{round(ot,6)}"); self.cv_out.config(state="readonly")
        except Exception as e: self.cv_err.config(text=f"转换失败：{e}")

    def _copy_cv(self):
        if self._check_expired(): return
        v = self.cv_out.get()
        if v: self.root.clipboard_clear(); self.root.clipboard_append(v); self.status.set("转换结果已复制"); self.root.after(2000, lambda: self.status.set("就绪"))

    def _cv_mode_chg(self, e=None):
        mode = self.cv_mode.get()
        if "→" in mode:
            dst = mode.split("→")[1].strip()
            lv = list(self.cv_olng["values"] or []); lt = list(self.cv_olat["values"] or [])
            if f"{dst}经度" in lv: self.cv_olng.set(f"{dst}经度")
            elif lv: self.cv_olng.set(lv[-1])
            if f"{dst}纬度" in lt: self.cv_olat.set(f"{dst}纬度")
            elif lt: self.cv_olat.set(lt[-1])

    def _stop_cv_batch(self): self.stop_flag = True; self.status.set("正在停止...")

    def _start_cv_batch(self):
        if self._check_expired(): return
        path = self.cv_bp.get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请选择有效的 Excel 文件"); return
        lc, lt = self.cv_lng.get(), self.cv_lat.get()
        if not lc or not lt: messagebox.showwarning("提示", "请选择经度列和纬度列"); return
        if lc == lt: messagebox.showwarning("提示", "经度列和纬度列不能相同"); return
        olc, olt = self.cv_olng.get().strip(), self.cv_olat.get().strip()
        if not olc or not olt: messagebox.showwarning("提示", "请选择输出经度列和输出纬度列"); return
        mode = self.cv_mode.get()
        if "→" not in mode: messagebox.showwarning("提示", "请选择有效的转换方式"); return
        sn, dn = [s.strip() for s in mode.split("→")]
        ss = sn.lower().replace("-","").replace("_",""); ds = dn.lower().replace("-","").replace("_","")
        ss = {"gcj02":"gcj02","wgs84":"wgs84","bd09":"bd09"}.get(ss, "gcj02")
        ds = {"gcj02":"gcj02","wgs84":"wgs84","bd09":"bd09"}.get(ds, "wgs84")
        if ss == ds: messagebox.showwarning("提示", "源坐标系与目标坐标系相同"); return
        self.stop_flag = False; self.cv_sbtn.config(state=tk.DISABLED); self.cv_xbtn.config(state=tk.NORMAL); self.cv_log.delete("1.0", tk.END)
        def task():
            try:
                ox, pf, ft, al, gcl = _get_openpyxl()
                wb = ox.load_workbook(path)
                if "Sheet1" not in wb.sheetnames:
                    try: wb.close()
                    except: pass
                    self.root.after(0, lambda: messagebox.showwarning("提示", "Excel 中未找到 Sheet1")); return
                ws = wb["Sheet1"]
                hs = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
                try: li = hs.index(lc) + 1; ti = hs.index(lt) + 1
                except ValueError:
                    try: wb.close()
                    except: pass
                    self.root.after(0, lambda: messagebox.showwarning("提示", "未找到指定的列名")); return
                if olc in hs: oli = hs.index(olc) + 1
                else: oli = len(hs) + 1; ws.cell(row=1, column=oli, value=olc); hs.append(olc)
                if olt in hs: oti = hs.index(olt) + 1
                else: oti = len(hs) + 1; ws.cell(row=1, column=oti, value=olt); hs.append(olt)
                lr = ws.max_row; self.root.after(0, lambda: self.cv_prog.config(maximum=lr-1, value=0))
                sc = fc = 0
                for r in range(2, lr + 1):
                    if self.stop_flag: self.root.after(0, lambda: self.cv_log.insert(tk.END, "\n用户取消\n")); break
                    pv = r - 1
                    self.root.after(0, lambda v=pv, t=lr-1: (self.cv_plbl.config(text=f"{v} / {t}"), self.cv_prog.config(value=v)))
                    try:
                        lv = ws.cell(row=r, column=li).value; tv = ws.cell(row=r, column=ti).value
                        if lv is None or tv is None: continue
                        ol, ot = convert_coords(float(lv), float(tv), ss, ds)
                        ws.cell(row=r, column=oli, value=round(ol, 6)); ws.cell(row=r, column=oti, value=round(ot, 6)); sc += 1
                    except Exception as e: fc += 1; ws.cell(row=r, column=oli, value=f"错误:{str(e)[:20]}"); ws.cell(row=r, column=oti, value=f"错误:{str(e)[:20]}")
                wb.save(path)
                self.root.after(0, lambda: self.cv_log.insert(tk.END, f"\n{'='*45}\n转换完成统计\n{'='*45}\n总行数：{lr-1}\n成功：{sc}\n失败：{fc}\n输出列：{olc}, {olt}\n文件已保存：{path}\n{'='*45}\n"))
            except Exception as e: self.root.after(0, lambda: self.cv_log.insert(tk.END, f"[处理异常] {str(e)}\n"))
            finally: self.root.after(0, lambda: (self.cv_sbtn.config(state=tk.NORMAL), self.cv_xbtn.config(state=tk.DISABLED), self.status.set("就绪")))
        threading.Thread(target=task, daemon=True).start()

    # ---------- 设置 ----------
    def _build_tools(self):
        """更多工具推广页面 — 精致卡片布局"""
        f = self._page(); self.pages["tools"] = f

        # ═══ 标题区 ═══
        tf = tk.Frame(f, bg=Config.C_BG); tf.pack(fill=tk.X, pady=(0, 10))
        title_row = tk.Frame(tf, bg=Config.C_BG); title_row.pack(fill=tk.X)
        tk.Label(title_row, text="🧰", font=("Segoe UI Emoji", 32), bg=Config.C_BG).pack(side=tk.LEFT)
        tcol = tk.Frame(title_row, bg=Config.C_BG); tcol.pack(side=tk.LEFT, padx=(12, 0))
        tk.Label(tcol, text="更多实用工具", font=("微软雅黑", 24, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w")
        tk.Label(tcol, text="基于同一技术栈打造的专业工具，提升您的办公效率", font=("微软雅黑", 11), 
                 bg=Config.C_BG, fg=Config.C_TEXT_L).pack(anchor="w", pady=(2, 0))

        # ═══ 联系信息横幅 ═══
        banner = tk.Frame(f, bg=Config.C_ACCENT, height=50); banner.pack(fill=tk.X, pady=(0, 12))
        banner.pack_propagate(False)
        binner = tk.Frame(banner, bg=Config.C_ACCENT); binner.pack(fill=tk.BOTH, expand=True, padx=20)
        tk.Label(binner, text="📢", font=("Segoe UI Emoji", 20), bg=Config.C_ACCENT, fg="white").pack(side=tk.LEFT)
        btext = tk.Frame(binner, bg=Config.C_ACCENT); btext.pack(side=tk.LEFT, padx=(12, 0))
        tk.Label(btext, text="商务合作与定制开发", font=("微软雅黑", 12, "bold"), bg=Config.C_ACCENT, fg="white").pack(anchor="w")
        tk.Label(btext, text="📧  1178258589@qq.com    |    支持企业内网部署 · 私有化定制 · API 接口集成", 
                 font=("微软雅黑", 10), bg=Config.C_ACCENT, fg="#fed7aa").pack(anchor="w")

        # ═══ 工具网格 ═══
        grid = tk.Frame(f, bg=Config.C_BG); grid.pack(fill=tk.BOTH, expand=True)
        # 不持久化下载状态，每次打开都显示"立即下载"
        self._downloaded = set(self.cfg.get("downloaded_tools", []))

        for i, tool in enumerate(Config.TOOLS):
            row, col = i // 2, i % 2

            # 外阴影层
            shadow = tk.Frame(grid, bg="#cbd5e1", highlightthickness=0)
            shadow.grid(row=row, column=col, sticky="nsew", padx=(0 if col==0 else 14, 2), pady=(0, 14))

            # 主卡片 — 固定高度 340
            card = tk.Frame(shadow, bg=Config.C_CARD, highlightbackground=Config.C_BORDER, highlightthickness=1, height=340)
            card.pack(fill=tk.BOTH, expand=True, padx=(0, 2), pady=(0, 4))
            card.pack_propagate(False)

            # 顶部色条
            bar = tk.Frame(card, bg=tool["color"], height=5); bar.pack(fill=tk.X)

            # 内容区
            inner = tk.Frame(card, bg=Config.C_CARD); inner.pack(fill=tk.BOTH, expand=True, padx=18, pady=(14, 16))

            # ═ 标题行 ═
            hf = tk.Frame(inner, bg=Config.C_CARD); hf.pack(fill=tk.X)

            # 图标圆形背景
            icon_bg = tk.Frame(hf, width=52, height=52, bg=tool["color"])
            icon_bg.pack(side=tk.LEFT); icon_bg.pack_propagate(False)
            tk.Label(icon_bg, text=tool["icon"], font=("Segoe UI Emoji", 26), bg=tool["color"], fg="white").place(relx=0.5, rely=0.5, anchor="center")

            # 标题和元信息
            tinfo = tk.Frame(hf, bg=Config.C_CARD); tinfo.pack(side=tk.LEFT, padx=(12, 0), fill=tk.Y)
            tk.Label(tinfo, text=tool["name"], font=("微软雅黑", 15, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")

            meta = tk.Frame(tinfo, bg=Config.C_CARD); meta.pack(anchor="w", pady=(3, 0))
            tk.Label(meta, text=tool["category"], font=("微软雅黑", 9), bg="#f1f5f9", fg=Config.C_TEXT_M,
                     padx=10, pady=3).pack(side=tk.LEFT)
            tk.Label(meta, text=tool["version"], font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_TEXT_L).pack(side=tk.LEFT, padx=(10, 0))
            tk.Label(meta, text="·", font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_TEXT_L).pack(side=tk.LEFT, padx=(6, 0))
            tk.Label(meta, text=tool["size"], font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_TEXT_L).pack(side=tk.LEFT, padx=(6, 0))

            # 角标
            if tool.get("badge"):
                badge = tk.Label(hf, text=tool["badge"], font=("微软雅黑", 9, "bold"), bg=tool["color"], fg="white",
                                 padx=10, pady=4)
                badge.pack(side=tk.RIGHT, anchor="n")

            # ═ 分隔线 ═
            tk.Frame(inner, bg=Config.C_BORDER, height=1).pack(fill=tk.X, pady=(10, 0))

            # ═ 描述 — 固定3行高度 ═
            desc_frame = tk.Frame(inner, bg=Config.C_CARD, height=58)
            desc_frame.pack(fill=tk.X, pady=(10, 0))
            desc_frame.pack_propagate(False)
            desc_text = tool["desc"]
            if len(desc_text) > 70:
                desc_text = desc_text[:67] + "..."
            tk.Label(desc_frame, text=desc_text, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_M,
                     wraplength=480, justify=tk.LEFT, anchor="nw").pack(fill=tk.BOTH, expand=True)



            # ═ 底部占位，把按钮推到底部 ═
            tk.Frame(inner, bg=Config.C_CARD).pack(fill=tk.BOTH, expand=True)

            # ═ 按钮区 — 固定在底部 ═
            bf = tk.Frame(inner, bg=Config.C_CARD); bf.pack(fill=tk.X, pady=(10, 0))
            is_current = tool["name"] == "地址标准化工具箱"
            is_downloaded = tool["name"] in self._downloaded

            if is_current:
                cur_frame = tk.Frame(bf, bg="#dcfce7", padx=12, pady=8)
                cur_frame.pack(side=tk.LEFT)
                tk.Label(cur_frame, text="✅ 当前正在使用", font=("微软雅黑", 10, "bold"), 
                         bg="#dcfce7", fg=Config.C_SUCCESS).pack()
            else:
                # 所有非当前工具统一显示"立即下载"
                btn = tk.Button(bf, text="⬇️  立即下载", font=("微软雅黑", 10, "bold"), 
                                bg=tool["color"], fg="white", bd=0, padx=24, pady=8, 
                                cursor="hand2", relief=tk.FLAT,
                                activebackground=self._darken(tool["color"]),
                                activeforeground="white",
                                command=lambda t=tool: self._download_tool(t))
                btn.pack(side=tk.LEFT)
                btn.bind("<Enter>", lambda e, b=btn, c=tool["color"]: b.config(bg=self._darken(c)))
                btn.bind("<Leave>", lambda e, b=btn, c=tool["color"]: b.config(bg=c))

            # 卡片悬停效果
            def _on_enter(e, c=card, color=tool["color"]):
                c.config(highlightbackground=color, highlightthickness=2)
            def _on_leave(e, c=card):
                c.config(highlightbackground=Config.C_BORDER, highlightthickness=1)
            card.bind("<Enter>", _on_enter)
            card.bind("<Leave>", _on_leave)

        # 配置网格权重
        grid.columnconfigure(0, weight=1); grid.columnconfigure(1, weight=1)
        for r in range((len(Config.TOOLS) + 1) // 2):
            grid.rowconfigure(r, weight=1)

    def _darken(self, hex_color, factor=0.85):
        """将颜色加深，用于悬停效果"""
        try:
            h = hex_color.lstrip('#')
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            r, g, b = int(r * factor), int(g * factor), int(b * factor)
            return f"#{r:02x}{g:02x}{b:02x}"
        except:
            return hex_color

    def _download_tool(self, tool):
        """下载工具软件"""
        if self._check_expired(): return
        url = tool.get("download_url", "")
        if not url:
            messagebox.showerror("错误", "下载链接未配置")
            return
        self._downloaded.add(tool["name"])
        self.cfg["downloaded_tools"] = list(self._downloaded)
        Config.save(self.cfg)
        try:
            import webbrowser
            webbrowser.open(url)
            messagebox.showinfo("下载启动", 
                f"『{tool['name']}』下载已启动\n\n"
                f"浏览器将打开下载页面，请保存安装包后运行。\n\n"
                f"下载地址：\n{url}")
            self.show("tools")
        except Exception as e:
            messagebox.showerror("错误", f"无法启动浏览器：{str(e)}\n\n请手动访问：\n{url}")

    # ---------- POI 采集助手 ----------
    def _build_poi(self):
        f = self._page(); self.pages["poi"] = f
        tk.Label(f, text="POI 采集助手", font=("微软雅黑", 18, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w", pady=(0, 4))
        tk.Label(f, text="基于高德/百度地图 API 的周边商户/设施/景点名录采集", font=("微软雅黑", 11), bg=Config.C_BG, fg=Config.C_TEXT_L).pack(anchor="w", pady=(0, 10))
        sp = tk.Frame(f, bg=Config.C_BG); sp.pack(fill=tk.BOTH, expand=True)
        sp.columnconfigure(0, weight=0, minsize=380); sp.columnconfigure(1, weight=1, minsize=650); sp.rowconfigure(0, weight=1)
        left = tk.Frame(sp, bg=Config.C_BG); left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        mode_card = Card(left, title="▎搜索模式", accent=Config.C_ACCENT); mode_card.pack(fill=tk.X, pady=(0, 8))
        mf = tk.Frame(mode_card, bg=Config.C_CARD); mf.pack(fill=tk.X, padx=14, pady=8)
        self.poi_mode = tk.StringVar(value="district")
        tk.Radiobutton(mf, text="行政区划检索", variable=self.poi_mode, value="district",
                       font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD,
                       command=self._poi_switch_mode).pack(anchor="w", pady=2)
        tk.Radiobutton(mf, text="周边范围检索", variable=self.poi_mode, value="around",
                       font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD,
                       command=self._poi_switch_mode).pack(anchor="w", pady=2)
        param_card = Card(left, title="▎检索参数", accent=Config.C_INFO); param_card.pack(fill=tk.X, pady=(0, 8))
        pf = tk.Frame(param_card, bg=Config.C_CARD); pf.pack(fill=tk.X, padx=14, pady=8)
        self.poi_dist_frame = tk.Frame(pf, bg=Config.C_CARD)
        self.poi_dist_frame.pack(fill=tk.X)
        tk.Label(self.poi_dist_frame, text="城市：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
        self.poi_city = tk.Entry(self.poi_dist_frame, font=("微软雅黑", 11), relief=tk.SOLID, bd=1); self.poi_city.pack(fill=tk.X, pady=(2, 4)); self.poi_city.insert(0, "曲靖市")
        tk.Label(self.poi_dist_frame, text="关键词：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
        # v2.8.5: 关键词改为 Combobox（下拉+手动输入）
        self.poi_kw_dist = ttk.Combobox(self.poi_dist_frame, values=["便利店","网吧","商铺","火锅店","加油站","幼儿园","银行","药店","酒店","景区"],
                                         font=("微软雅黑", 11), width=30, state="normal")
        self.poi_kw_dist.set("便利店"); self.poi_kw_dist.pack(fill=tk.X, pady=(2, 4))
        self.poi_around_frame = tk.Frame(pf, bg=Config.C_CARD)
        tk.Label(self.poi_around_frame, text="中心地址：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
        self.poi_center = tk.Entry(self.poi_around_frame, font=("微软雅黑", 11), relief=tk.SOLID, bd=1); self.poi_center.pack(fill=tk.X, pady=(2, 4)); self.poi_center.insert(0, "麒麟区雄业金都")
        tk.Label(self.poi_around_frame, text="半径（米）：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
        self.poi_radius = ttk.Combobox(self.poi_around_frame, values=["500", "1000", "2000", "3000", "5000", "10000"], font=("微软雅黑", 10), width=12, state="readonly")
        self.poi_radius.set("3000"); self.poi_radius.pack(anchor="w", pady=(2, 4))
        tk.Label(self.poi_around_frame, text="关键词：", font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w")
        self.poi_kw_around = ttk.Combobox(self.poi_around_frame, values=["便利店","网吧","商铺","火锅店","加油站","幼儿园","银行","药店","酒店","景区"],
                                           font=("微软雅黑", 11), width=30, state="normal")
        self.poi_kw_around.set("便利店"); self.poi_kw_around.pack(fill=tk.X, pady=(2, 4))
        opt_card = Card(left, title="▎高级选项", accent=Config.C_SUCCESS); opt_card.pack(fill=tk.X, pady=(0, 8))
        of = tk.Frame(opt_card, bg=Config.C_CARD); of.pack(fill=tk.X, padx=14, pady=8)
        self.poi_supplement_tel = tk.BooleanVar(value=False)
        tk.Checkbutton(of, text="补充电话（高德缺失时触发百度）", variable=self.poi_supplement_tel,
                       font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", pady=2)
        self.poi_break_limit = tk.BooleanVar(value=False)
        tk.Checkbutton(of, text="突破500条限制（关键词拆分检索）", variable=self.poi_break_limit,
                       font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", pady=2)
        quota_frame = tk.Frame(left, bg=Config.C_CARD, highlightbackground=Config.C_BORDER, highlightthickness=1)
        quota_frame.pack(fill=tk.X, pady=(0, 8))
        self.poi_quota_lbl = tk.Label(quota_frame, text="预计消耗：开始采集后动态计算\n剩余配额：请配置密钥后查看",
                                      font=("微软雅黑", 9), bg=Config.C_CARD, fg=Config.C_TEXT_M, justify=tk.LEFT)
        self.poi_quota_lbl.pack(anchor="w", padx=10, pady=8)
        # ═══ 统一按钮布局（2×2 grid，确保4个按钮大小完全一致）═══
        btn_frame = tk.Frame(left, bg=Config.C_BG); btn_frame.pack(fill=tk.X, pady=(0, 8))
        btn_frame.columnconfigure(0, weight=1); btn_frame.columnconfigure(1, weight=1)
        self.poi_start_btn = tk.Button(btn_frame, text="▶ 开始采集", font=("微软雅黑", 11, "bold"), bg=Config.C_ACCENT, fg=Config.C_PRIMARY,
                                       bd=0, padx=18, pady=8, cursor="hand2", relief=tk.FLAT, command=self._start_poi)
        self.poi_start_btn.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        self.poi_start_btn.bind("<Enter>", lambda e: self.poi_start_btn.config(bg="#b45309"))
        self.poi_start_btn.bind("<Leave>", lambda e: self.poi_start_btn.config(bg=Config.C_ACCENT))
        self.poi_stop_btn = tk.Button(btn_frame, text="⏹ 停止", font=("微软雅黑", 11, "bold"), bg=Config.C_DANGER, fg="white",
                                      bd=0, padx=18, pady=8, cursor="hand2", relief=tk.FLAT, state=tk.DISABLED, command=self._stop_poi)
        self.poi_stop_btn.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self.poi_excel_btn = tk.Button(btn_frame, text="📤 导出 Excel", font=("微软雅黑", 11, "bold"), bg=Config.C_SUCCESS, fg="#fff",
                                       bd=0, padx=18, pady=8, cursor="hand2", relief=tk.FLAT, state=tk.DISABLED, command=self._export_poi_excel)
        self.poi_excel_btn.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        self.poi_excel_btn.bind("<Enter>", lambda e: self.poi_excel_btn.config(bg="#047857"))
        self.poi_excel_btn.bind("<Leave>", lambda e: self.poi_excel_btn.config(bg=Config.C_SUCCESS))
        self.poi_kml_btn = tk.Button(btn_frame, text="🌏 导出 KML", font=("微软雅黑", 11, "bold"), bg=Config.C_INFO, fg="#fff",
                                     bd=0, padx=18, pady=8, cursor="hand2", relief=tk.FLAT, state=tk.DISABLED, command=self._export_poi_kml)
        self.poi_kml_btn.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))
        self.poi_kml_btn.bind("<Enter>", lambda e: self.poi_kml_btn.config(bg="#1d4ed8"))
        self.poi_kml_btn.bind("<Leave>", lambda e: self.poi_kml_btn.config(bg=Config.C_INFO))
        right = tk.Frame(sp, bg=Config.C_BG); right.grid(row=0, column=1, sticky="nsew"); right.rowconfigure(0, weight=1); right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1); right.columnconfigure(0, weight=1)
        nb = ttk.Notebook(right); nb.pack(fill=tk.BOTH, expand=True)
        preview_frame = tk.Frame(nb, bg=Config.C_CARD); nb.add(preview_frame, text="  📋 数据预览  ")
        tf = tk.Frame(preview_frame, bg=Config.C_CARD); tf.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self.poi_tree = ttk.Treeview(tf, show="headings", selectmode="browse")
        self.poi_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.poi_tree.yview); vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.poi_tree.xview); hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.poi_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        cols = ["序号", "名称", "地址", "电话", "GCJ-02经度", "GCJ-02纬度", "WGS-84经度", "WGS-84纬度", "分类", "数据来源"]
        self.poi_tree["columns"] = cols
        self.poi_tree.column("序号", width=50, anchor="center"); self.poi_tree.heading("序号", text="序号")
        self.poi_tree.column("名称", width=150, anchor="w"); self.poi_tree.heading("名称", text="名称")
        self.poi_tree.column("地址", width=200, anchor="w"); self.poi_tree.heading("地址", text="地址")
        self.poi_tree.column("电话", width=120, anchor="w"); self.poi_tree.heading("电话", text="电话")
        self.poi_tree.column("GCJ-02经度", width=100, anchor="w"); self.poi_tree.heading("GCJ-02经度", text="GCJ-02经度")
        self.poi_tree.column("GCJ-02纬度", width=100, anchor="w"); self.poi_tree.heading("GCJ-02纬度", text="GCJ-02纬度")
        self.poi_tree.column("WGS-84经度", width=100, anchor="w"); self.poi_tree.heading("WGS-84经度", text="WGS-84经度")
        self.poi_tree.column("WGS-84纬度", width=100, anchor="w"); self.poi_tree.heading("WGS-84纬度", text="WGS-84纬度")
        self.poi_tree.column("分类", width=150, anchor="w"); self.poi_tree.heading("分类", text="分类")
        self.poi_tree.column("数据来源", width=80, anchor="center"); self.poi_tree.heading("数据来源", text="数据来源")
        self.poi_tree.insert("", tk.END, values=("", "", "暂无数据", "点击「开始采集」后结果将显示在这里", "", "", "", "", "", ""))
        log_frame = tk.Frame(nb, bg=Config.C_CARD); nb.add(log_frame, text="  📊 采集日志  ")
        self.poi_log = scrolledtext.ScrolledText(log_frame, font=("Consolas", 10), wrap=tk.WORD, height=12,
                                                  bg="#f8fafc", fg=Config.C_TEXT, relief=tk.SOLID, bd=1)
        self.poi_log.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self._log_cfg(self.poi_log)
        self.poi_log.insert(tk.END, "[提示] 选择搜索模式并填写参数后，点击「开始采集」开始检索\n", "info")
        self.poi_results = []
        self.poi_stop_flag = False
        self._poi_switch_mode()

    def _poi_switch_mode(self):
        if self.poi_mode.get() == "district":
            self.poi_dist_frame.pack(fill=tk.X); self.poi_around_frame.pack_forget()
        else:
            self.poi_dist_frame.pack_forget(); self.poi_around_frame.pack(fill=tk.X)

    def _start_poi(self):
        if self._check_expired(): return
        if self.poi_running:
            messagebox.showinfo("提示", "正在采集中，请勿重复启动")
            return
        if not self.cfg.get("gaode_key"):
            messagebox.showwarning("提示", "请先配置高德 API Key")
            self.show("settings")
            return
        self.poi_running = True
        self.poi_stop_flag = False
        self.poi_start_btn.config(state=tk.DISABLED); self.poi_stop_btn.config(state=tk.NORMAL)
        self.poi_excel_btn.config(state=tk.DISABLED); self.poi_kml_btn.config(state=tk.DISABLED)
        self.poi_results = []
        self.poi_log.delete("1.0", tk.END)
        self.poi_log.insert(tk.END, "[系统] 启动 POI 采集...\n", "info")
        for item in self.poi_tree.get_children(): self.poi_tree.delete(item)
        threading.Thread(target=self._poi_task, daemon=True).start()

    def _stop_poi(self):
        self.poi_stop_flag = True; self.poi_log.insert(tk.END, "\n⚠ 用户请求停止...\n", "warning")

    def _poi_task(self):
        try:
            mode = self.poi_mode.get()
            gkey = self.cfg.get("gaode_key", "")
            bak = self.cfg.get("baidu_ak", "")
            geo = GeoCoder(self.cfg)
            ok, msg = geo.check()
            if not ok:
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[错误] 密钥验证失败: {msg}\n", "error"))
                return
            gaode_count = baidu_count = 0
            raw_results = []
            clng = clat = 0.0
            kw = city = radius = ""
            if mode == "district":
                city = self.poi_city.get().strip()
                kw = self.poi_kw_dist.get().strip()
                if not city or not kw:
                    self.root.after(0, lambda: self.poi_log.insert(tk.END, "[错误] 城市和关键词不能为空\n", "error"))
                    return
                # 修复：区县下钻 —— 高德同请求参数翻页最多200条，且市级检索仅覆盖市辖区；
                # 通过行政区划查询获取下属区县，逐区县检索并合并，突破数量上限
                units = get_poi_district_units(gkey, city)
                if len(units) > 1:
                    self.root.after(0, lambda u=units: self.poi_log.insert(tk.END, f"[高德] 已下钻 {len(u)} 个区县逐区检索: {'、'.join(u)}\n", "info"))
                for unit in units:
                    if self.poi_stop_flag: break
                    self.root.after(0, lambda un=unit: self.poi_log.insert(tk.END, f"[高德] 行政区划检索: {un} + {kw}\n", "info"))
                    page = 1
                    while not self.poi_stop_flag and page <= 40:
                        url = f"{Config.GAODE_POI_TEXT}?key={gkey}&keywords={url_enc(kw)}&city={url_enc(unit)}&offset=25&page={page}&extensions=all&output=json"
                        ok, resp = http_get(url, 2)
                        if not ok:
                            self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 请求失败: {resp[:80]}\n", "error"))
                            break
                        gaode_count += 1
                        try:
                            d = json.loads(resp)
                            if d.get("status") != "1":
                                info = d.get("info", "")
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 错误: {info}\n", "error"))
                                break
                            pois = d.get("pois", [])
                            if not pois:
                                self.root.after(0, lambda p=page: self.poi_log.insert(tk.END, f"[高德] 第{p}页无数据，采集结束\n", "info"))
                                break
                            for p in pois:
                                loc = p.get("location", "")
                                lng, lat = 0, 0
                                if loc and "," in loc:
                                    try: lng, lat = map(float, loc.split(","))
                                    except: pass
                                raw_results.append({
                                    "name": p.get("name", ""), "address": p.get("address", ""),
                                    "tel": p.get("tel", ""), "type": p.get("type", ""),
                                    "gcj_lng": lng, "gcj_lat": lat, "source": "高德"
                                })
                            self.root.after(0, lambda p=page, c=len(pois): self.poi_log.insert(tk.END, f"[高德] 第{p}页: {c}条\n", "info"))
                            if len(pois) < 25: break
                        except Exception as e:
                            self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 解析异常: {e}\n", "error"))
                            break
                        page += 1; time.sleep(0.3)
            else:
                center = self.poi_center.get().strip()
                kw = self.poi_kw_around.get().strip()
                radius = self.poi_radius.get()
                if not center or not kw:
                    self.root.after(0, lambda: self.poi_log.insert(tk.END, "[错误] 中心地址和关键词不能为空\n", "error"))
                    return
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 解析中心地址: {center}\n", "info"))
                res, ab, err, fatal, blog = geo.resolve(center, "", "", "")
                if not res or ab:
                    self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[错误] 中心地址解析失败: {err}\n", "error"))
                    return
                clng, clat = res["gcj_lng"], res["gcj_lat"]
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 中心坐标: {clng},{clat}\n", "info"))
                page = 1
                while not self.poi_stop_flag and page <= 40:
                    url = f"{Config.GAODE_POI_AROUND}?key={gkey}&location={clng},{clat}&keywords={url_enc(kw)}&radius={radius}&offset=25&page={page}&extensions=all&output=json"
                    ok, resp = http_get(url, 2)
                    if not ok:
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 请求失败: {resp[:80]}\n", "error"))
                        break
                    gaode_count += 1
                    try:
                        d = json.loads(resp)
                        if d.get("status") != "1":
                            info = d.get("info", "")
                            self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 错误: {info}\n", "error"))
                            break
                        pois = d.get("pois", [])
                        if not pois:
                            self.root.after(0, lambda p=page: self.poi_log.insert(tk.END, f"[高德] 第{p}页无数据，采集结束\n", "info"))
                            break
                        for p in pois:
                            loc = p.get("location", "")
                            lng, lat = 0, 0
                            if loc and "," in loc:
                                try: lng, lat = map(float, loc.split(","))
                                except: pass
                            raw_results.append({
                                "name": p.get("name", ""), "address": p.get("address", ""),
                                "tel": p.get("tel", ""), "type": p.get("type", ""),
                                "gcj_lng": lng, "gcj_lat": lat, "source": "高德"
                            })
                        self.root.after(0, lambda p=page, c=len(pois): self.poi_log.insert(tk.END, f"[高德] 第{p}页: {c}条\n", "info"))
                        if len(pois) < 25: break
                    except Exception as e:
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[高德] 解析异常: {e}\n", "error"))
                        break
                    page += 1; time.sleep(0.3)
            # ═══ 突破500条限制：关键词拆分检索 ═══
            expand_dup = 0   # 新增：初始化
            if self.poi_break_limit.get() and len(raw_results) >= 480:
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[突破限制] 高德已返回{len(raw_results)}条（接近上限），启动关键词拆分检索...\n", "info"))
                main_kw = kw
                expand_kws = Config.POI_KEYWORD_EXPAND.get(main_kw, [])
                # 若主关键词不在映射表中，尝试用包含关系匹配
                if not expand_kws:
                    for k, v in Config.POI_KEYWORD_EXPAND.items():
                        if k in main_kw or main_kw in k:
                            expand_kws = v
                            break
                # 去重：排除主关键词本身 + 扩展词之间互斥
                expand_kws = list(dict.fromkeys([ek for ek in expand_kws if ek != main_kw]))
                for ek in expand_kws:
                    if self.poi_stop_flag: break
                    self.root.after(0, lambda k=ek: self.poi_log.insert(tk.END, f"[突破限制] 扩展关键词: {k}\n", "info"))
                    for unit in (units if mode == "district" else [""]):
                        if self.poi_stop_flag: break
                        ep = 1
                        while not self.poi_stop_flag and ep <= 40:  # 每个扩展词+区县最多40页（1000条）
                            if mode == "district":
                                eurl = f"{Config.GAODE_POI_TEXT}?key={gkey}&keywords={url_enc(ek)}&city={url_enc(unit)}&offset=25&page={ep}&extensions=all&output=json"
                            else:
                                eurl = f"{Config.GAODE_POI_AROUND}?key={gkey}&location={clng},{clat}&keywords={url_enc(ek)}&radius={radius}&offset=25&page={ep}&extensions=all&output=json"
                            eok, eresp = http_get(eurl, 2)
                            if not eok:
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[突破限制] 扩展词请求失败\n", "error"))
                                break
                            gaode_count += 1
                            try:
                                ed = json.loads(eresp)
                                if ed.get("status") != "1": break
                                epois = ed.get("pois", [])
                                if not epois: break
                                for p in epois:
                                    loc = p.get("location", "")
                                    lng, lat = 0, 0
                                    if loc and "," in loc:
                                        try: lng, lat = map(float, loc.split(","))
                                        except: pass
                                    # 与已有结果坐标去重（距离 < 50 米视为重复）
                                    is_dup = False
                                    for existing in raw_results:
                                        if not existing.get("gcj_lng") or not existing.get("gcj_lat"):
                                            continue
                                        dx = abs(lng - existing["gcj_lng"]) * 111000 * math.cos(math.radians(lat))
                                        dy = abs(lat - existing["gcj_lat"]) * 111000
                                        if math.sqrt(dx*dx + dy*dy) < 50:
                                            is_dup = True
                                            expand_dup += 1
                                            break
                                    if not is_dup:
                                        raw_results.append({
                                            "name": p.get("name", ""), "address": p.get("address", ""),
                                            "tel": p.get("tel", ""), "type": p.get("type", ""),
                                            "gcj_lng": lng, "gcj_lat": lat, "source": "高德(扩展)"
                                        })
                                if len(epois) < 25: break
                            except Exception as ee:
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[突破限制] 扩展词解析异常: {ee}\n", "error"))
                                break
                            ep += 1; time.sleep(0.3)
                        if self.poi_stop_flag: break
                    if self.poi_stop_flag: break
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[突破限制] 关键词拆分检索完成，当前共{len(raw_results)}条（去重{expand_dup}条）\n", "info"))

            # 百度补缺逻辑
            need_baidu = False
            if bak and self.cfg.get("enable_baidu_fallback", True):
                if len(raw_results) < 20:
                    need_baidu = True
                    self.root.after(0, lambda c=len(raw_results): self.poi_log.insert(tk.END, f"[百度Fallback] 高德结果仅{c}条，触发百度补缺\n", "baidu"))
                elif self.poi_supplement_tel.get():
                    # 只要存在任何一条缺少电话的POI，就触发百度补电话
                    missing_tel_count = sum(1 for r in raw_results if not r.get("tel"))
                    if missing_tel_count > 0:
                        need_baidu = True
                        self.root.after(0, lambda mc=missing_tel_count, tr=len(raw_results): self.poi_log.insert(tk.END, f"[百度Fallback] {mc}/{tr}条POI缺少电话，触发百度补全\n", "baidu"))
            baidu_results = []
            if need_baidu:
                self.root.after(0, lambda: self.poi_log.insert(tk.END, "[百度] 开始补缺检索...\n", "baidu"))
                bn = 0
                while not self.poi_stop_flag and bn < 40:
                    if mode == "district":
                        url = f"{Config.BAIDU_POI}?query={url_enc(kw)}&region={url_enc(city)}&output=json&ak={bak}&page_size=10&page_num={bn}&scope=2"
                    else:
                        bd_lng, bd_lat = gcj02_to_bd09(clng, clat)
                        url = f"{Config.BAIDU_POI}?query={url_enc(kw)}&location={bd_lat},{bd_lng}&radius={radius}&output=json&ak={bak}&page_size=10&page_num={bn}&scope=2"
                    ok, resp = http_get(url, 2)
                    if not ok:
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[百度] 请求失败: {resp[:80]}\n", "error"))
                        break
                    baidu_count += 1
                    try:
                        d = json.loads(resp)
                        if d.get("status") != 0:
                            self.root.after(0, lambda s=d.get('status'): self.poi_log.insert(tk.END, f"[百度] 错误码: {s}\n", "error"))
                            break
                        results = d.get("results", [])
                        if not results: break
                        for r in results:
                            loc = r.get("location", {})
                            blng, blat = float(loc.get("lng", 0)), float(loc.get("lat", 0))
                            glng, glat = bd09_to_gcj02(blng, blat)
                            baidu_results.append({
                                "name": r.get("name", ""), "address": r.get("address", ""),
                                "tel": r.get("telephone", ""), "type": r.get("detail_info", {}).get("type", ""),
                                "gcj_lng": glng, "gcj_lat": glat, "source": "百度"
                            })
                        self.root.after(0, lambda p=bn+1, c=len(results): self.poi_log.insert(tk.END, f"[百度] 第{p}页: {c}条\n", "baidu"))
                        if len(results) < 10: break
                    except Exception as e:
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[百度] 解析异常: {e}\n", "error"))
                        break
                    bn += 1; time.sleep(0.3)
            # 合并去重
            merged = list(raw_results)
            dup_count = 0
            for br in baidu_results:
                is_dup = False
                for mr in merged:
                    if not br.get("gcj_lng") or not mr.get("gcj_lng"): continue
                    dx = abs(br["gcj_lng"] - mr["gcj_lng"]) * 111000 * math.cos(math.radians(br["gcj_lat"]))
                    dy = abs(br["gcj_lat"] - mr["gcj_lat"]) * 111000
                    dist = math.sqrt(dx*dx + dy*dy)
                    if dist < 50:
                        is_dup = True
                        if not mr.get("tel") and br.get("tel"):
                            mr["tel"] = br["tel"]
                            mr["source"] = mr.get("source", "") + "+百度补电话"
                        dup_count += 1
                        break
                if not is_dup:
                    merged.append(br)
            self.poi_results = merged
            self.root.after(0, lambda: self.poi_log.insert(tk.END, f"\n{'='*45}\n采集完成\n{'='*45}\n高德: {len(raw_results)}条 | 百度: {len(baidu_results)}条\n去重: {dup_count}条 | 合并: {len(merged)}条\n总调用: 高德{gaode_count}次 百度{baidu_count}次\n{'='*45}\n", "success"))
            self.root.after(0, self._refresh_poi_tree)
            self.root.after(0, lambda: (self.poi_excel_btn.config(state=tk.NORMAL), self.poi_kml_btn.config(state=tk.NORMAL)))
        except Exception as e:
            self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[系统异常] {str(e)}\n", "error"))
        finally:
            self.poi_running = False
            self.root.after(0, lambda: (self.poi_start_btn.config(state=tk.NORMAL), self.poi_stop_btn.config(state=tk.DISABLED)))  

    def _refresh_poi_tree(self):
        for item in self.poi_tree.get_children(): self.poi_tree.delete(item)
        for i, r in enumerate(self.poi_results, 1):
            gl, gt = r.get("gcj_lng", 0), r.get("gcj_lat", 0)
            wl, wt = gcj02_to_wgs84(gl, gt) if gl and gt else (0, 0)
            self.poi_tree.insert("", tk.END, values=(
                i, r.get("name",""), r.get("address",""), r.get("tel",""),
                round(gl, 6) if gl else "",
                round(gt, 6) if gt else "",
                round(wl, 6) if wl else "",
                round(wt, 6) if wt else "",
                r.get("type",""), r.get("source","高德")
            ))

    def _export_poi_excel(self):
        if not self.poi_results:
            messagebox.showwarning("提示", "暂无采集结果"); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                            initialfile=f"POI采集_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        if not path: return
        def task():
            try:
                ox, pf, ft, al, gcl = _get_openpyxl()
                wb = ox.Workbook(); ws = wb.active; ws.title = "Sheet1"
                headers = ["序号", "名称", "地址", "电话", "GCJ-02经度", "GCJ-02纬度", "WGS-84经度", "WGS-84纬度", "分类", "数据来源"]
                for col, h in enumerate(headers, 1):
                    c = ws.cell(row=1, column=col, value=h); c.font = ft(bold=True, color="FFFFFF")
                    c.fill = pf(start_color="1e293b", end_color="1e293b", fill_type="solid"); c.alignment = al(horizontal="center", vertical="center")
                for i, r in enumerate(self.poi_results, 2):
                    gl, gt = r.get("gcj_lng", 0), r.get("gcj_lat", 0)
                    wl, wt = gcj02_to_wgs84(gl, gt) if gl and gt else (0, 0)
                    def _s(v):
                        if v is None: return ""
                        if isinstance(v, list): return ", ".join(str(x) for x in v)
                        return str(v)
                    ws.cell(row=i, column=1, value=i-1)
                    ws.cell(row=i, column=2, value=_s(r.get("name","")))
                    ws.cell(row=i, column=3, value=_s(r.get("address","")))
                    ws.cell(row=i, column=4, value=_s(r.get("tel","")))
                    ws.cell(row=i, column=5, value=round(gl, 6) if gl else "")
                    ws.cell(row=i, column=6, value=round(gt, 6) if gt else "")
                    ws.cell(row=i, column=7, value=round(wl, 6) if wl else "")
                    ws.cell(row=i, column=8, value=round(wt, 6) if wt else "")
                    ws.cell(row=i, column=9, value=_s(r.get("type","")))
                    ws.cell(row=i, column=10, value=_s(r.get("source","高德")))
                for i, w in enumerate([8, 25, 35, 18, 15, 15, 15, 15, 25, 12], 1):
                    ws.column_dimensions[gcl(i)].width = w
                wb.save(path)
                self.root.after(0, lambda: messagebox.showinfo("导出成功", f"已导出 {len(self.poi_results)} 条到:\n{path}"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"导出失败: {e}"))
        threading.Thread(target=task, daemon=True).start()

    def _export_poi_kml(self):
        if not self.poi_results:
            messagebox.showwarning("提示", "暂无采集结果"); return
        path = filedialog.asksaveasfilename(defaultextension=".kml", filetypes=[("KML文件", "*.kml")],
                                            initialfile=f"POI采集_{time.strftime('%Y%m%d_%H%M%S')}.kml")
        if not path: return
        def task():
            try:
                kml_lines = [
                    '<?xml version="1.0" encoding="UTF-8"?>',
                    '<kml xmlns="http://www.opengis.net/kml/2.2">',
                    '  <Document>',
                    '    <name>POI采集结果</name>',
                    '    <Style id="redPin">',
                    '      <IconStyle>',
                    '        <Icon><href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href></Icon>',
                    '      </IconStyle>',
                    '    </Style>'
                ]
                for r in self.poi_results:
                    gl, gt = r.get("gcj_lng", 0), r.get("gcj_lat", 0)
                    if not gl or not gt: continue
                    wl, wt = gcj02_to_wgs84(gl, gt)
                    name = r.get("name", "")
                    desc = f"地址：{r.get('address','')}\n电话：{r.get('tel','')}\n分类：{r.get('type','')}\n来源：{r.get('source','高德')}"
                    kml_lines.append('    <Placemark>')
                    kml_lines.append(f'      <name>{_xml_esc(name)}</name>')
                    kml_lines.append(f'      <description><![CDATA[{desc}]]></description>')
                    kml_lines.append('      <styleUrl>#redPin</styleUrl>')
                    kml_lines.append(f'      <Point><coordinates>{round(wl,6)},{round(wt,6)},0</coordinates></Point>')
                    kml_lines.append('    </Placemark>')
                kml_lines += ['  </Document>', '</kml>']
                with open(path, "w", encoding="utf-8") as f:
                    f.write("\n".join(kml_lines))
                self.root.after(0, lambda: messagebox.showinfo("导出成功", f"已导出 {len(self.poi_results)} 个坐标点到 KML:\n{path}"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"导出失败: {e}"))
        threading.Thread(target=task, daemon=True).start()


    def _build_settings(self):
        f = self._page(); self.pages["settings"] = f
        tk.Label(f, text="API 密钥配置", font=("微软雅黑", 20, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w", pady=(0, 12))
        card = Card(f, title="密钥设置", accent=Config.C_ACCENT); card.pack(fill=tk.X, pady=0)
        form = tk.Frame(card, bg=Config.C_CARD); form.pack(fill=tk.X, padx=14, pady=4)
        for lbl, attr, val in [("* 高德 API Key：", "cfg_gaode", self.cfg.get("gaode_key","")), ("  百度 AK：", "cfg_baidu", self.cfg.get("baidu_ak",""))]:
            r = tk.Frame(form, bg=Config.C_CARD); r.pack(fill=tk.X, pady=8)
            tk.Label(r, text=lbl, font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, width=15, anchor="e").pack(side=tk.LEFT)
            e = tk.Entry(r, font=("微软雅黑", 11), width=55, show="*", relief=tk.SOLID, bd=1); e.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True, ipady=0)
            e.insert(0, val); setattr(self, attr, e)
            tk.Button(r, text="显示/隐藏", font=("微软雅黑", 10), bg=Config.C_BORDER, fg=Config.C_TEXT, bd=0, padx=10, pady=4,
                      command=lambda ent=e: ent.config(show="" if ent.cget("show")=="*" else "*")).pack(side=tk.LEFT)
        opt = tk.Frame(form, bg=Config.C_CARD); opt.pack(fill=tk.X, pady=8)
        self.cfg_eb = tk.BooleanVar(value=self.cfg.get("enable_baidu_fallback", True))
        tk.Checkbutton(opt, text="启用百度地图 Fallback（高德精度不足时自动切换）", variable=self.cfg_eb,
                       font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w")
        r3 = tk.Frame(form, bg=Config.C_CARD); r3.pack(fill=tk.X, pady=8)
        tk.Label(r3, text="百度触发等级：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, width=15, anchor="e").pack(side=tk.LEFT)
        self.cfg_trig = ttk.Combobox(r3, values=["门址","街道","乡镇","区县","城市"], font=("微软雅黑", 10), width=15, state="readonly")
        self.cfg_trig.set(self.cfg.get("baidu_trigger_level", "乡镇")); self.cfg_trig.pack(side=tk.LEFT, padx=10)
        tk.Label(r3, text="请求间隔（秒）：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT, padx=(30, 0))
        self.cfg_iv = tk.Spinbox(r3, from_=0.1, to=5.0, increment=0.1, width=8, font=("微软雅黑", 10), relief=tk.SOLID, bd=1)
        self.cfg_iv.delete(0, tk.END); self.cfg_iv.insert(0, str(self.cfg.get("base_interval", 0.2))); self.cfg_iv.pack(side=tk.LEFT, padx=10)
        bf = tk.Frame(card, bg=Config.C_CARD); bf.pack(fill=tk.X, padx=14, pady=(0, 14))
        SButton(bf, "  保存配置", self._save_cfg, "primary").pack(side=tk.LEFT)
        SButton(bf, "  测试高德", lambda: self._test_api("gaode"), "success").pack(side=tk.LEFT, padx=10)
        SButton(bf, "  测试百度", lambda: self._test_api("baidu"), "success").pack(side=tk.LEFT)
        tip = Card(f, title="使用说明", accent=Config.C_INFO); tip.pack(fill=tk.X, pady=6)
        for t in ["• 高德 Key 为必填项，请前往 https://console.amap.com 申请个人开发者 Key",
                  "• 百度 AK 为可选项，用于提升解析精度，前往 https://lbsyun.baidu.com 申请",
                  "• 配置将自动保存到当前目录的 addr_config.json 文件中，下次打开自动加载",
                  "• 高德个人开发者每天约有 5000 次免费地理编码额度，超出需付费",
                  "• 建议开启百度 Fallback，当高德返回精度低于「乡镇」时自动尝试百度"]:
            tk.Label(tip, text=t, font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT_M).pack(anchor="w", padx=15, pady=5)

    def _save_cfg(self):
        if self._check_expired(): return
        self.cfg["gaode_key"] = self.cfg_gaode.get().strip()
        self.cfg["baidu_ak"] = self.cfg_baidu.get().strip()
        self.cfg["enable_baidu_fallback"] = self.cfg_eb.get()
        self.cfg["baidu_trigger_level"] = self.cfg_trig.get()
        try: self.cfg["base_interval"] = float(self.cfg_iv.get())
        except: self.cfg["base_interval"] = 0.2
        Config.save(self.cfg); messagebox.showinfo("提示", "配置已保存！")

    def _test_api(self, api_type):
        if self._check_expired(): return
        key = self.cfg_gaode.get().strip() if api_type == "gaode" else self.cfg_baidu.get().strip()
        if not key: messagebox.showwarning("提示", "请先输入 API Key"); return
        self.status.set(f"正在测试 {api_type} ...")
        def task():
            if api_type == "gaode":
                url = f"{Config.GAODE_GEO}?address={url_enc('北京市天安门')}&output=json&key={key}"
                ok, resp = http_get(url, 2)
                if ok and json_val(resp, "status") == "1": self.root.after(0, lambda: messagebox.showinfo("测试成功", "高德 API 连接正常！"))
                else: self.root.after(0, lambda: messagebox.showerror("测试失败", f"错误：{gaode_err_msg(json_val(resp,'infocode'), json_val(resp,'info') if ok else resp)}"))
            else:
                url = f"{Config.BAIDU_GEO}?address={url_enc('北京市天安门')}&output=json&ak={key}"
                ok, resp = http_get(url, 2)
                if ok and json_val(resp, "status") == "0": self.root.after(0, lambda: messagebox.showinfo("测试成功", "百度 API 连接正常！"))
                else: self.root.after(0, lambda: messagebox.showerror("测试失败", f"错误：{baidu_err_msg(json_val(resp,'status'), resp if ok else resp)}"))
            self.root.after(0, lambda: self.status.set("就绪"))
        threading.Thread(target=task, daemon=True).start()

    # ---------- 自动更新 ----------
    def _auto_check(self):
        try:
            info = get_latest_version()
            if info and cmp_ver(Config.VERSION, info.get("latestVersion","")):
                self.root.after(0, lambda: self._confirm_upd(info))
        except: pass

    def _check_update(self):
        if self._check_expired(): return
        if self.is_updating: messagebox.showwarning("提示", "正在检查中..."); return
        self.is_updating = True; self.ubtn.config(text="检查中...", state=tk.DISABLED); self.status.set("正在检查更新...")
        def task():
            try:
                info = get_latest_version()
                if not info: self.status.set("连接更新服务器失败"); messagebox.showerror("失败", "无法连接更新服务器"); return
                if not cmp_ver(Config.VERSION, info.get("latestVersion","")): self.status.set(f"已是最新版本 v{Config.VERSION}"); messagebox.showinfo("提示", "已是最新版本"); return
                self.root.after(0, lambda: self._confirm_upd(info))
            except Exception as e: self.status.set("检查更新出错"); messagebox.showerror("错误", str(e))
            finally: self.is_updating = False; self.root.after(0, lambda: self.ubtn.config(text="🔍 检查更新", state=tk.NORMAL))
        threading.Thread(target=task, daemon=True).start()

    def _confirm_upd(self, info):
        lv = info.get("latestVersion", "")
        cl = info.get("changelog", "")
        fn = info.get("fileName", "")

        if messagebox.askyesno("发现新版本", f"发现新版本 v{lv}\n\n更新内容：\n{cl}\n\n是否立即下载更新？"):
            self.status.set(f"正在下载新版本 v{lv}...")
            download_url = f"{Config.UPDATE_URL}/download/{fn}"
            def _upd_progress(downloaded, total):
                if total > 0 and self.root.winfo_exists():
                    pct = min(100, int(downloaded * 100 / total))
                    self.root.after(0, lambda: self.upd_prog.config(value=pct))
            download_update(download_url, fn, lambda m: self.status.set(m), self.root, progress_cb=_upd_progress)
# ═══════════════════════════════════════════════════════════
# 10. 入口
# ═══════════════════════════════════════════════════════════
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()



