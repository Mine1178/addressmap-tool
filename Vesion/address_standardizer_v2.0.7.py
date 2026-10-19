#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
地址坐标地址转换工具 v2.0.6
功能：正向地理编码 / 逆向地理编码 / 坐标转换 / Excel 批量处理
基于 VB 核心逻辑重构，提升健壮性与准确性
"""

import sys
import os
import json
import math
import re
import time
import threading
import ssl
import urllib.request
import urllib.parse
from datetime import datetime
from tkinter import font as tkfont

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ===================== 常量配置 =====================
CONFIG_FILE = "addr_config.json"
ABNORMAL_COLOR = "FFFF00"

# Excel 列号（1-based）
COL_SERIAL = 1
COL_PROVINCE = 2
COL_CITY = 3
COL_DISTRICT = 4
COL_ADDR1 = 5
COL_ADDR2 = 6
COL_GCJ_LNG = 7
COL_GCJ_LAT = 8
COL_WGS_LNG = 9
COL_WGS_LAT = 10
COL_STD_ADDR = 11
COL_REGO_FULL = 12
COL_REGO_DIV = 13
COL_REGO_STREET = 14

# 坐标常量
PI = 3.14159265358979
X_PI = 52.3598775598299
EARTH_A = 6378245.0
EARTH_EE = 0.006693421622965943

# API 地址
GAODE_GEO_URL = "https://restapi.amap.com/v3/geocode/geo"
GAODE_REGO_URL = "https://restapi.amap.com/v3/geocode/regeo"
BAIDU_GEO_URL = "https://api.map.baidu.com/geocoding/v3/"
BAIDU_REGO_URL = "https://api.map.baidu.com/reverse_geocoding/v3/"

# ===================== 配色方案 =====================
COLOR_PRIMARY = "#1a1a2e"
COLOR_SECONDARY = "#16213e"
COLOR_ACCENT = "#e6b325"
COLOR_BG = "#f8f9fa"
COLOR_CARD = "#ffffff"
COLOR_TEXT = "#212529"
COLOR_TEXT_LIGHT = "#6c757d"
COLOR_BORDER = "#dee2e6"
COLOR_SUCCESS = "#28a745"
COLOR_DANGER = "#dc3545"
COLOR_WARNING = "#ffc107"
COLOR_INFO = "#17a2b8"

# ===================== 坐标转换（严格对照 VB） =====================
def transform_lat(lng, lat):
    ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * math.sqrt(abs(lng))
    ret += (20.0 * math.sin(6.0 * lng * PI) + 20.0 * math.sin(2.0 * lng * PI)) * 2.0 / 3.0
    ret += (20.0 * math.sin(lat * PI) + 40.0 * math.sin(lat / 3.0 * PI)) * 2.0 / 3.0
    ret += (160.0 * math.sin(lat / 12.0 * PI) + 320.0 * math.sin(lat * PI / 30.0)) * 2.0 / 3.0
    return ret

def transform_lng(lng, lat):
    ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * math.sqrt(abs(lng))
    ret += (20.0 * math.sin(6.0 * lng * PI) + 20.0 * math.sin(2.0 * lng * PI)) * 2.0 / 3.0
    ret += (20.0 * math.sin(lng * PI) + 40.0 * math.sin(lng / 3.0 * PI)) * 2.0 / 3.0
    ret += (150.0 * math.sin(lng / 12.0 * PI) + 300.0 * math.sin(lng * PI / 30.0)) * 2.0 / 3.0
    return ret

def wgs84_to_gcj02(wgs_lng, wgs_lat):
    """WGS-84 -> GCJ-02（对照 VB WGS84ToGCJ02）"""
    lng, lat = wgs_lng, wgs_lat
    gcj_lng, gcj_lat = wgs_lng, wgs_lat
    for _ in range(5):
        dlat = transform_lat(lng - 105, lat - 35)
        dlng = transform_lng(lng - 105, lat - 35)
        radlat = lat / 180.0 * PI
        magic = math.sin(radlat)
        magic = 1 - EARTH_EE * magic * magic
        sqrtmagic = math.sqrt(magic)
        dlat = (dlat * 180.0) / ((EARTH_A * (1 - EARTH_EE)) / (magic * sqrtmagic) * PI)
        dlng = (dlng * 180.0) / (EARTH_A / sqrtmagic * math.cos(radlat) * PI)
        gcj_lat = wgs_lat + dlat
        gcj_lng = wgs_lng + dlng
        lat = gcj_lat
        lng = gcj_lng
    return gcj_lng, gcj_lat

def gcj02_to_wgs84(gcj_lng, gcj_lat):
    """GCJ-02 -> WGS-84（对照 VB GCJ02ToWGS84）"""
    lng, lat = gcj_lng, gcj_lat
    wgs_lng, wgs_lat = gcj_lng, gcj_lat
    for _ in range(5):
        dlat = transform_lat(lng - 105, lat - 35)
        dlng = transform_lng(lng - 105, lat - 35)
        radlat = lat / 180.0 * PI
        magic = math.sin(radlat)
        magic = 1 - EARTH_EE * magic * magic
        sqrtmagic = math.sqrt(magic)
        dlat = (dlat * 180.0) / ((EARTH_A * (1 - EARTH_EE)) / (magic * sqrtmagic) * PI)
        dlng = (dlng * 180.0) / (EARTH_A / sqrtmagic * math.cos(radlat) * PI)
        wgs_lat = gcj_lat - dlat
        wgs_lng = gcj_lng - dlng
        lat = wgs_lat
        lng = wgs_lng
    return wgs_lng, wgs_lat

def bd09_to_gcj02(bd_lng, bd_lat):
    """百度坐标 BD-09 -> GCJ-02（对照 VB BD09ToGCJ02）"""
    x = bd_lng - 0.0065
    y = bd_lat - 0.006
    z = math.sqrt(x * x + y * y) - 0.00002 * math.sin(y * X_PI)
    theta = math.atan2(y, x) - 0.000003 * math.cos(x * X_PI)
    gcj_lng = z * math.cos(theta)
    gcj_lat = z * math.sin(theta)
    return gcj_lng, gcj_lat

# ===================== 工具函数（对照 VB 完善） =====================
def get_level_score(level):
    score_map = {
        "门址": 100, "门牌号": 100, "门址点": 100, "门牌": 100,
        "POI": 90, "兴趣点": 90, "楼宇": 90, "大厦": 90, "建筑物": 90, "购物": 90,
        "村庄": 80, "自然村": 80, "社区": 80, "村委会": 80, "住宅小区": 80,
        "道路": 70, "街道": 70, "路段": 70, "道路交叉口": 70,
        "乡镇": 60, "乡镇级": 60, "街道级": 60, "镇": 60,
        "区县": 50, "区县级": 50, "区县级别": 50,
        "城市": 30, "市级": 30, "地市级": 30, "市": 30,
        "省": 10, "省级": 10,
    }
    return score_map.get(level, 20)

def clean_address(addr):
    if not addr:
        return ""
    addr = str(addr)
    replacements = {
        "　": " ", "，": ",", "。": ".", "；": ";", "：": ":",
        "（": "(", "）": ")", "\r\n": "", "\n": "", "\t": "",
    }
    for k, v in replacements.items():
        addr = addr.replace(k, v)
    while "  " in addr:
        addr = addr.replace("  ", " ")
    return addr.strip()

def remove_parentheses(s):
    s = re.sub(r"[(（][^)）]*[)）]", "", s)
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip()

def extract_province_city(addr):
    m = re.search(r"(.+?(省|市))", addr)
    return m.group(1) if m else ""

def get_address_degrade_levels(addr):
    """地址降级策略（对照 VB GetAddressDegradeLevels）"""
    levels = [addr]
    addr2 = remove_parentheses(addr)
    if addr2 != addr and len(addr2) > 4:
        levels.append(addr2)
    addr3 = addr2
    for sep in [",", " ", "-"]:
        pos = addr3.rfind(sep)
        if pos > 3:
            addr3 = addr3[:pos]
            if addr3 not in levels:
                levels.append(addr3)
            break
    addr4 = extract_province_city(addr)
    if addr4 and addr4 not in levels:
        levels.append(addr4)
    return levels

def build_full_address(prov, city, dist, addr):
    result = ""
    if prov and prov not in addr:
        result += prov
    if city and city not in addr and city not in result:
        result += city
    if dist and dist not in addr and dist not in result:
        result += dist
    result += addr
    return result

def build_api_city_param(prov, city, dist):
    result = ""
    for part in [prov, city, dist]:
        if part:
            result += part
    return result

def extract_pure_city(city):
    return city if "市" in city else ""

def has_detail_content(addr):
    if not addr:
        return False
    keywords = ["路", "街", "号", "村", "小区", "大厦", "道", "巷", "弄", "镇", "乡", "社区"]
    return any(k in addr for k in keywords)

def standardize_addr_key(city, addr):
    key = f"{city}|{addr}"
    key = key.lower().replace(" ", "")
    key = key.replace("，", "").replace("。", "")
    return key

def check_address_drift(std_addr, prov, city, dist):
    if not std_addr:
        return True
    if dist and dist not in std_addr:
        return True
    if city and city not in std_addr:
        return True
    if prov and prov not in std_addr:
        return True
    return False

def extract_core_name(raw_addr, prov, city, dist):
    s = raw_addr
    for part in [prov, city, dist]:
        if part:
            s = s.replace(part, "")
    return s.strip()

def address_contains_core(std_addr, core_name):
    if not core_name:
        return True
    return core_name.lower() in std_addr.lower()

def deduplicate_address(addr):
    """去除地址中的重复片段，如'白水镇白水镇' → '白水镇'"""
    if not addr:
        return addr
    # 去除相邻重复词
    for _ in range(3):
        addr = re.sub(r'(\S+?)\1+', r'\1', addr)
    # 去除特定重复模式
    addr = re.sub(r'(\S{2,6})\s*\1', r'\1', addr)
    return addr.strip()

def clean_baidu_address(addr):
    addr = clean_address(addr)
    for ch in ["&", "#", "=", "?", "%"]:
        addr = addr.replace(ch, "")
    return addr.strip()

def url_encode_utf8(s):
    return urllib.parse.quote(str(s), safe="")

def http_get(url, max_retry=3, timeout=15, use_dynamic_interval=False):
    """HTTP GET（对照 VB HttpGet）"""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    last_err = ""
    for retry in range(1, max_retry + 1):
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
            )
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                data = resp.read().decode("utf-8")
                if data:
                    return True, data
        except Exception as e:
            last_err = str(e)
            if use_dynamic_interval:
                time.sleep(0.3 * retry)
            else:
                time.sleep(0.2)
    return False, last_err

def json_root_value(json_str, key):
    pattern = r'"' + re.escape(key) + r'"\s*:\s*("?)([^"\,}\s]+)\1'
    m = re.search(pattern, json_str)
    return m.group(2) if m else ""

def get_geocode_field(json_str, field_name):
    pattern = r'"geocodes"\s*:\s*\[\s*\{.*?"' + re.escape(field_name) + r'"\s*:\s*"([^"]*)"'
    m = re.search(pattern, json_str, re.DOTALL)
    return m.group(1) if m else ""

def get_regeo_field(json_str, field_name):
    pattern = r'"regeocode"\s*:\s*\{.*?"' + re.escape(field_name) + r'"\s*:\s*"([^"]*)"'
    m = re.search(pattern, json_str, re.DOTALL)
    return m.group(1) if m else ""

def get_regeo_component(json_str, field_name):
    pattern = r'"addressComponent"\s*:\s*\{.*?"' + re.escape(field_name) + r'"\s*:\s*"([^"]*)"'
    m = re.search(pattern, json_str, re.DOTALL)
    return m.group(1) if m else ""

def get_baidu_lng(json_str):
    m = re.search(r'"location"\s*:\s*\{[^{}]*"lng"\s*:\s*(-?\d+\.?\d*)', json_str)
    return m.group(1) if m else ""

def get_baidu_lat(json_str):
    m = re.search(r'"location"\s*:\s*\{[^{}]*"lat"\s*:\s*(-?\d+\.?\d*)', json_str)
    return m.group(1) if m else ""

def get_baidu_level(json_str):
    m = re.search(r'"result"\s*:\s*\{.*?"level"\s*:\s*"([^"]*)"', json_str, re.DOTALL)
    return m.group(1) if m else ""

def get_baidu_formatted_addr(json_str):
    """提取百度返回的实际地址，支持正向/逆地理编码"""
    try:
        data = json.loads(json_str)
        status = data.get("status")
        if (status == "0" or status == 0) and "result" in data and isinstance(data["result"], dict):
            result = data["result"]
            # 优先 formatted_address
            formatted = result.get("formatted_address")
            if formatted and str(formatted).strip():
                return str(formatted).strip()
            # 备选：sematic_description
            sematic = result.get("sematic_description")
            if sematic and str(sematic).strip():
                return str(sematic).strip()
            # 备选：addressComponent 组合
            ac = result.get("addressComponent", {})
            parts = []
            for key in ["province", "city", "district", "town", "street", "street_number"]:
                val = ac.get(key, "")
                if val and str(val).strip():
                    parts.append(str(val).strip())
            if parts:
                return "".join(parts)
    except Exception:
        pass
    return ""

# ===================== 配置管理 =====================
def load_config():
    default = {
        "gaode_key": "",
        "baidu_ak": "",
        "enable_baidu_fallback": True,
        "baidu_trigger_level": "区县",
        "daily_quota_gaode": 0,
        "daily_quota_baidu": 0,
        "max_retry": 3,
        "base_interval": 0.2,
        "skip_existing": True,
        "check_drift": True,
        "enable_cache": True,
        "baidu_debug": True,
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                default.update(cfg)
        except Exception:
            pass
    return default

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ===================== 核心地理编码类（对照 VB 重构） =====================
class GeoCoder:
    def __init__(self, config):
        self.cfg = config
        self.gaode_key = config.get("gaode_key", "")
        self.baidu_ak = config.get("baidu_ak", "")
        self.enable_baidu = config.get("enable_baidu_fallback", True)
        self.trigger_level = config.get("baidu_trigger_level", "区县")
        self.max_retry = config.get("max_retry", 3)
        self.base_interval = config.get("base_interval", 0.2)
        self.gaode_count = 0
        self.baidu_count = 0
        self.cache = {}
        self.regeo_cache = {}
        self.stopped = False
        self.key_checked = False
        self.gaode_key_valid = False
        self.baidu_key_valid = False

    def stop(self):
        self.stopped = True

    def init_cache(self):
        self.cache = {}
        self.regeo_cache = {}
        self.gaode_count = 0
        self.baidu_count = 0

    def check_keys(self):
        """对照 VB：检查密钥有效性"""
        if self.key_checked:
            return True, ""
        if len(self.gaode_key) != 32:
            return False, "高德密钥格式错误（应为32位）"
        self.gaode_key_valid = self._check_gaode_key()
        if not self.gaode_key_valid:
            return False, "高德密钥验证失败"
        if self.enable_baidu and self.baidu_ak:
            self.baidu_key_valid = self._check_baidu_key()
        self.key_checked = True
        return True, ""

    def _check_gaode_key(self):
        url = f"{GAODE_GEO_URL}?address={url_encode_utf8('北京市天安门')}&output=json&key={self.gaode_key}"
        ok, resp = http_get(url, 2)
        if ok:
            return json_root_value(resp, "status") == "1"
        return False

    def _check_baidu_key(self):
        url = f"{BAIDU_GEO_URL}?address={url_encode_utf8('北京市天安门')}&output=json&ak={self.baidu_ak}"
        ok, resp = http_get(url, 2)
        if ok:
            return json_root_value(resp, "status") == "0"
        return False

    def gaode_geocode(self, address, city=""):
        """高德地理编码（对照 VB 请求构建）"""
        enc_addr = url_encode_utf8(address)
        enc_city = url_encode_utf8(city)
        url = f"{GAODE_GEO_URL}?address={enc_addr}"
        if city:
            url += f"&city={enc_city}"
        url += f"&output=json&key={self.gaode_key}"
        ok, resp = http_get(url, self.max_retry)
        if not ok:
            return None, resp
        self.gaode_count += 1
        status = json_root_value(resp, "status")
        if status != "1":
            infocode = json_root_value(resp, "infocode")
            info = json_root_value(resp, "info")
            return None, f"高德错误[{infocode}]: {info}"
        location = get_geocode_field(resp, "location")
        std_addr = get_geocode_field(resp, "formatted_address")
        level = get_geocode_field(resp, "level")
        return {"location": location, "std_addr": std_addr, "level": level, "source": "高德"}, ""

    def baidu_geocode_final(self, address, city=""):
        """百度地理编码（对照 VB BaiduGeocodeFinal）—— 正向编码+逆地理编码获取实际地址"""
        clean = clean_baidu_address(address)
        if not clean:
            return None, "地址为空"
        enc_addr = url_encode_utf8(clean)
        url = f"{BAIDU_GEO_URL}?address={enc_addr}&output=json&ak={self.baidu_ak}"
        if city:
            url += f"&city={url_encode_utf8(city)}"
        ok, resp = http_get(url, 2)
        if not ok:
            return None, resp
        self.baidu_count += 1

        # 使用json解析
        try:
            data = json.loads(resp)
        except json.JSONDecodeError as e:
            return None, f"百度返回JSON解析失败: {e}"

        status = data.get("status")
        # 带城市失败则尝试不带城市（对照 VB 逻辑）
        if status not in (0, "0") and city and status in (1, "1"):
            url = f"{BAIDU_GEO_URL}?address={enc_addr}&output=json&ak={self.baidu_ak}"
            ok2, resp2 = http_get(url, 2)
            if ok2:
                self.baidu_count += 1
                try:
                    data2 = json.loads(resp2)
                    status2 = data2.get("status")
                    if status2 in (0, "0"):
                        data = data2
                        status = status2
                    else:
                        return None, f"百度错误码:{status2}"
                except json.JSONDecodeError:
                    return None, "百度二次返回JSON解析失败"
            else:
                return None, f"百度错误码:{status}"

        if status not in (0, "0"):
            return None, f"百度错误码:{status}"

        result = data.get("result", {})
        if not isinstance(result, dict):
            return None, "百度返回格式错误（无result）"

        # 提取坐标
        location = result.get("location", {})
        lng_val = location.get("lng")
        lat_val = location.get("lat")
        if lng_val is None or lat_val is None:
            return None, "提取坐标失败"
        try:
            bd_lng, bd_lat = float(lng_val), float(lat_val)
        except (ValueError, TypeError):
            return None, "坐标格式错误"
        if bd_lng == 0 or bd_lat == 0:
            return None, "坐标值为0"

        gcj_lng, gcj_lat = bd09_to_gcj02(bd_lng, bd_lat)
        level = result.get("level", "未知")

        # ===== 提取实际地址：优先正向编码返回的formatted_address =====
        std_addr = ""
        formatted = result.get("formatted_address")
        if formatted and str(formatted).strip():
            std_addr = str(formatted).strip()

        # ===== 关键修复：正向编码无地址时，调用百度逆地理编码获取实际地址 =====
        if not std_addr:
            try:
                regeo_url = f"{BAIDU_REGO_URL}?ak={self.baidu_ak}&output=json&coordtype=bd09ll&location={bd_lat},{bd_lng}"
                ok_re, resp_re = http_get(regeo_url, 2)
                if ok_re:
                    self.baidu_count += 1
                    data_re = json.loads(resp_re)
                    if data_re.get("status") in (0, "0"):
                        result_re = data_re.get("result", {})
                        formatted_re = result_re.get("formatted_address")
                        if formatted_re and str(formatted_re).strip():
                            std_addr = str(formatted_re).strip()
                        # 备选：addressComponent组合
                        if not std_addr:
                            ac = result_re.get("addressComponent", {})
                            parts = []
                            for key in ["province", "city", "district", "town", "street", "street_number"]:
                                val = ac.get(key, "")
                                if val and str(val).strip():
                                    parts.append(str(val).strip())
                            if parts:
                                std_addr = "".join(parts)
            except Exception:
                pass

        # 清洗去除重复内容
        std_addr = deduplicate_address(std_addr)
        # 最终回退：用户输入
        if not std_addr:
            std_addr = clean

        return {
            "location": f"{gcj_lng},{gcj_lat}",
            "std_addr": std_addr,
            "level": level,
            "source": "百度"
        }, ""

    def baidu_reverse_geocode(self, bd_lng, bd_lat):
        """百度逆地理编码（对照 VB BaiduReverseGeocode）"""
        url = f"{BAIDU_REGO_URL}?ak={self.baidu_ak}&output=json&coordtype=bd09ll&location={bd_lat},{bd_lng}"
        ok, resp = http_get(url, 2)
        if not ok:
            return None, "逆地理网络失败"
        self.baidu_count += 1
        status = json_root_value(resp, "status")
        if status != "0":
            return None, f"逆地理错误码{status}"
        formatted = get_baidu_formatted_addr(resp)
        return formatted, ""

    def resolve_address(self, addr_str, prov_str, city_str, dist_str):
        """
        地址解析核心函数（严格对照 VB ResolveAddress）
        返回: (result_dict, is_abnormal, err_msg, fatal_error, baidu_fallback_log)
        """
        baidu_fallback_log = ""  # 初始化百度Fallback日志
        if self.stopped:
            return None, False, "用户取消", False, baidu_fallback_log
        if not addr_str:
            return None, True, "详细地址为空", False, baidu_fallback_log

        api_city = build_api_city_param(prov_str, city_str, dist_str)
        gaode_city = extract_pure_city(city_str)
        baidu_city = extract_pure_city(city_str)
        query = build_full_address(prov_str, city_str, dist_str, addr_str)
        clean = clean_address(query)
        cache_key = standardize_addr_key(api_city, addr_str)

        # 缓存检查
        if self.cfg.get("enable_cache", True) and cache_key in self.cache:
            c = self.cache[cache_key]
            return c, c.get("is_abnormal", False), "", False, baidu_fallback_log

        has_detail = has_detail_content(addr_str)
        levels = get_address_degrade_levels(clean)

        gcj_lng, gcj_lat = 0.0, 0.0
        std_addr, level, source = "", "", "高德"
        is_drift, is_abnormal = False, False
        err_msg = ""
        location = ""

        # 高德逐级降级查询
        for lvl_addr in levels:
            if self.stopped:
                return None, False, "用户取消", False, baidu_fallback_log
            result, err = self.gaode_geocode(lvl_addr, gaode_city)
            if result:
                location = result["location"]
                std_addr = result["std_addr"]
                level = result["level"]
                source = result["source"]
                if location and "," in location:
                    if get_level_score(level) >= get_level_score("街道"):
                        break
            # 全局错误码检查（10001/10003）
            if err and ("10001" in err or "10003" in err):
                return None, True, f"高德全局错误：{err}", True, baidu_fallback_log
            time.sleep(self.base_interval)

        if not location:
            err_msg = "高德无匹配结果"
            is_abnormal = True
        else:
            parts = location.split(",")
            try:
                gcj_lng, gcj_lat = float(parts[0]), float(parts[1])
            except (ValueError, IndexError):
                gcj_lng, gcj_lat = 0.0, 0.0

        # 漂移检查
        if self.cfg.get("check_drift", True) and api_city and std_addr:
            is_drift = check_address_drift(std_addr, prov_str, city_str, dist_str)
            if is_drift:
                is_abnormal = True

        # 百度 Fallback（对照 VB 逻辑）
        baidu_fallback_log = ""
        if self.enable_baidu and location and self.baidu_ak and self.baidu_key_valid:
            gaode_score = get_level_score(level)
            trigger_score = get_level_score(self.trigger_level)
            need_baidu = is_drift or (gaode_score <= trigger_score and has_detail)
            if need_baidu:
                baidu_fallback_log = f"[百度Fallback] 高德精度不足({level})，尝试百度解析...\n"
                bd_result, bd_err = self.baidu_geocode_final(query, baidu_city)
                if bd_result:
                    bd_score = get_level_score(bd_result["level"])
                    bd_drift = check_address_drift(bd_result["std_addr"], prov_str, city_str, dist_str)
                    if bd_score > gaode_score and not bd_drift:
                        parts = bd_result["location"].split(",")
                        gcj_lng, gcj_lat = float(parts[0]), float(parts[1])
                        std_addr = bd_result["std_addr"]
                        level = bd_result["level"]
                        source = "百度"
                        is_drift = False
                        is_abnormal = False
                        baidu_fallback_log += f"[百度Fallback] 百度提升成功: {std_addr} [精度:{level}]\n"
                    else:
                        baidu_fallback_log += f"[百度Fallback] 百度未提升({bd_result.get('level','')})，保留高德结果\n"
                        # 百度未提升，保留高德结果但标记异常信息
                        if self.cfg.get("baidu_debug", True):
                            std_addr += f" [百度未提升:{bd_result.get('level','')}]"
                else:
                    baidu_fallback_log += f"[百度Fallback] 百度解析失败: {bd_err}\n"
                    if self.cfg.get("baidu_debug", True):
                        std_addr += f" [百度失败:{bd_err}]"
                    is_abnormal = True

        # 写入缓存（对照 VB：正常且精度大于区县）
        if gcj_lng and gcj_lat:
            if self.cfg.get("enable_cache", True) and not is_abnormal:
                if get_level_score(level) > get_level_score("区县"):
                    self.cache[cache_key] = {
                        "gcj_lng": gcj_lng, "gcj_lat": gcj_lat,
                        "std_addr": std_addr, "level": level,
                        "source": source, "is_drift": is_drift,
                        "is_abnormal": is_abnormal,
                    }
            return {
                "gcj_lng": gcj_lng, "gcj_lat": gcj_lat,
                "std_addr": std_addr, "level": level,
                "source": source, "is_drift": is_drift,
                "is_abnormal": is_abnormal,
            }, is_abnormal, err_msg or ("解析结果存在异常" if is_abnormal else ""), False, baidu_fallback_log
        else:
            return None, True, err_msg or "双平台均无法解析", False, baidu_fallback_log

    def gaode_regeocode(self, lng, lat):
        """高德逆地理编码（对照 VB GetAddressFromGaodeRegeo 核心逻辑）"""
        location = f"{lng},{lat}"
        cache_key = location
        if self.cfg.get("enable_cache", True) and cache_key in self.regeo_cache:
            return self.regeo_cache[cache_key], ""
        url = f"{GAODE_REGO_URL}?location={location}&output=json&extensions=all&key={self.gaode_key}"
        ok, resp = http_get(url, self.max_retry)
        if not ok:
            return None, resp
        self.gaode_count += 1
        status = json_root_value(resp, "status")
        if status != "1":
            infocode = json_root_value(resp, "infocode")
            info = json_root_value(resp, "info")
            return None, f"高德逆地理错误[{infocode}]: {info}"
        full = get_regeo_field(resp, "formatted_address")
        prov = get_regeo_component(resp, "province")
        city = get_regeo_component(resp, "city")
        dist = get_regeo_component(resp, "district")
        street = get_regeo_component(resp, "street")
        number = get_regeo_component(resp, "number")
        result = {
            "full": full,
            "division": f"{prov}{city}{dist}",
            "street": f"{street}{number}",
        }
        if self.cfg.get("enable_cache", True):
            self.regeo_cache[cache_key] = result
        return result, ""

# ===================== Excel 处理 =====================
class ExcelProcessor:
    def __init__(self, filepath):
        self.filepath = filepath
        self.wb = openpyxl.load_workbook(filepath)
        if "Sheet1" not in self.wb.sheetnames:
            raise ValueError("Excel 中未找到 Sheet1 工作表")
        self.ws = self.wb["Sheet1"]
        self.yellow_fill = PatternFill(start_color=ABNORMAL_COLOR, end_color=ABNORMAL_COLOR, fill_type="solid")
        self.none_fill = PatternFill(fill_type=None)

    def get_last_row(self, col=COL_ADDR1):
        for row in range(self.ws.max_row, 1, -1):
            val = self.ws.cell(row=row, column=col).value
            if val is not None and str(val).strip():
                return row
        return 1

    def is_result_normal(self, row):
        lng = self.ws.cell(row=row, column=COL_GCJ_LNG).value
        addr = str(self.ws.cell(row=row, column=COL_STD_ADDR).value or "")
        if not lng or not isinstance(lng, (int, float)) or lng == 0:
            return False
        bad_keywords = ["失败", "错误", "漂移", "无匹配", "空地址", "无效", "超出范围"]
        if any(k in addr for k in bad_keywords):
            return False
        return True

    def clear_forward_result(self, row):
        for c in range(COL_GCJ_LNG, COL_STD_ADDR + 1):
            self.ws.cell(row=row, column=c).value = None
            self.ws.cell(row=row, column=c).fill = self.none_fill

    def write_forward(self, row, result):
        if not result:
            self.ws.cell(row=row, column=COL_GCJ_LNG).value = "无匹配"
            self.ws.cell(row=row, column=COL_GCJ_LAT).value = "无匹配"
            self.ws.cell(row=row, column=COL_WGS_LNG).value = "无匹配"
            self.ws.cell(row=row, column=COL_WGS_LAT).value = "无匹配"
            self.ws.cell(row=row, column=COL_STD_ADDR).value = "解析失败"
            for c in range(COL_GCJ_LNG, COL_STD_ADDR + 1):
                self.ws.cell(row=row, column=c).fill = self.yellow_fill
            return
        gcj_lng = result["gcj_lng"]
        gcj_lat = result["gcj_lat"]
        self.ws.cell(row=row, column=COL_GCJ_LNG).value = gcj_lng
        self.ws.cell(row=row, column=COL_GCJ_LAT).value = gcj_lat
        if 73 <= gcj_lng <= 135 and 3 <= gcj_lat <= 53:
            wgs_lng, wgs_lat = gcj02_to_wgs84(gcj_lng, gcj_lat)
            self.ws.cell(row=row, column=COL_WGS_LNG).value = round(wgs_lng, 6)
            self.ws.cell(row=row, column=COL_WGS_LAT).value = round(wgs_lat, 6)
        else:
            self.ws.cell(row=row, column=COL_WGS_LNG).value = "超出范围"
            self.ws.cell(row=row, column=COL_WGS_LAT).value = "超出范围"
            result["is_abnormal"] = True
        final = f"{result['std_addr']} [精度:{result['level']}][来源:{result['source']}]"
        if result.get("is_drift"):
            final += " 疑似漂移"
        self.ws.cell(row=row, column=COL_STD_ADDR).value = final
        fill = self.yellow_fill if result.get("is_abnormal") else self.none_fill
        for c in range(COL_GCJ_LNG, COL_STD_ADDR + 1):
            self.ws.cell(row=row, column=c).fill = fill

    def clear_reverse_result(self, row):
        for c in range(COL_REGO_FULL, COL_REGO_STREET + 1):
            self.ws.cell(row=row, column=c).value = None

    def write_reverse(self, row, result):
        if not result:
            self.ws.cell(row=row, column=COL_REGO_FULL).value = "解析失败"
            return
        self.ws.cell(row=row, column=COL_REGO_FULL).value = result.get("full", "")
        self.ws.cell(row=row, column=COL_REGO_DIV).value = result.get("division", "")
        self.ws.cell(row=row, column=COL_REGO_STREET).value = result.get("street", "")

    def save(self):
        self.wb.save(self.filepath)

# ===================== 自定义样式组件 =====================
class StyledButton(tk.Button):
    def __init__(self, master, text, command, btn_type="primary", **kwargs):
        colors = {
            "primary": (COLOR_ACCENT, COLOR_PRIMARY),
            "secondary": (COLOR_INFO, "#ffffff"),
            "success": (COLOR_SUCCESS, "#ffffff"),
            "danger": (COLOR_DANGER, "#ffffff"),
            "warning": (COLOR_WARNING, COLOR_PRIMARY),
        }
        bg, fg = colors.get(btn_type, colors["primary"])
        super().__init__(master, text=text, command=command,
                        font=("微软雅黑", 11, "bold"),
                        bg=bg, fg=fg,
                        bd=0, padx=20, pady=8,
                        cursor="hand2",
                        relief=tk.FLAT,
                        **kwargs)

class Card(tk.Frame):
    def __init__(self, master, title=None, **kwargs):
        super().__init__(master, bg=COLOR_CARD, **kwargs)
        self.configure(highlightbackground=COLOR_BORDER, highlightthickness=1)
        if title:
            title_frame = tk.Frame(self, bg=COLOR_CARD)
            title_frame.pack(fill=tk.X, padx=15, pady=(15, 0))
            tk.Label(title_frame, text=title, font=("微软雅黑", 14, "bold"),
                    bg=COLOR_CARD, fg=COLOR_TEXT).pack(anchor="w")
            tk.Frame(self, bg=COLOR_BORDER, height=1).pack(fill=tk.X, padx=15, pady=10)

# ===================== GUI 界面 =====================
class AddressStandardizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("地址坐标地址转换工具 v2.0.6")
        self.root.geometry("1200x800")
        self.root.minsize(1100, 700)
        self.config = load_config()
        self.geo = GeoCoder(self.config)
        self.running = False
        self.stop_flag = False
        self.default_font = tkfont.Font(family="微软雅黑", size=10)
        self.root.option_add("*Font", self.default_font)
        self._build_ui()
        self.show_page("home")

    def _build_ui(self):
        main_frame = tk.Frame(self.root, bg=COLOR_BG)
        main_frame.pack(fill=tk.BOTH, expand=True)
        nav_width = 220
        self.nav_frame = tk.Frame(main_frame, bg=COLOR_PRIMARY, width=nav_width)
        self.nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_frame.pack_propagate(False)
        logo_frame = tk.Frame(self.nav_frame, bg=COLOR_PRIMARY, height=100)
        logo_frame.pack(fill=tk.X, pady=(30, 20))
        dot_frame = tk.Frame(logo_frame, bg=COLOR_PRIMARY)
        dot_frame.pack()
        cvs = tk.Canvas(dot_frame, width=50, height=50, bg=COLOR_PRIMARY, highlightthickness=0)
        cvs.pack()
        cvs.create_oval(5, 5, 45, 45, fill=COLOR_ACCENT, outline="")
        cvs.create_text(25, 25, text="从", font=("微软雅黑", 18, "bold"), fill=COLOR_PRIMARY)
        tk.Label(logo_frame, text="人人帮", font=("微软雅黑", 20, "bold"),
                bg=COLOR_PRIMARY, fg=COLOR_ACCENT).pack(pady=(10, 0))
        tk.Label(logo_frame, text="地址标准化工具", font=("微软雅黑", 10),
                bg=COLOR_PRIMARY, fg="#8892b0").pack()
        self.nav_buttons = {}
        nav_items = [
            ("home", "  首页", "🏠"),
            ("single", "  单条查询", "📍"),
            ("batch_fwd", "  批量正向编码", "📊"),
            ("batch_rev", "  批量逆向编码", "🔄"),
            ("settings", "  API 设置", "⚙️"),
        ]
        nav_btn_frame = tk.Frame(self.nav_frame, bg=COLOR_PRIMARY)
        nav_btn_frame.pack(fill=tk.X, padx=15, pady=20)
        for key, text, icon in nav_items:
            btn_frame = tk.Frame(nav_btn_frame, bg=COLOR_PRIMARY)
            btn_frame.pack(fill=tk.X, pady=3)
            btn = tk.Button(btn_frame, text=f"{icon}  {text.strip()}", 
                           font=("微软雅黑", 11),
                           bg=COLOR_PRIMARY, fg="#8892b0",
                           activebackground=COLOR_SECONDARY,
                           activeforeground=COLOR_ACCENT,
                           bd=0, padx=15, pady=12,
                           anchor="w", cursor="hand2",
                           relief=tk.FLAT,
                           command=lambda k=key: self.show_page(k))
            btn.pack(fill=tk.X)
            self.nav_buttons[key] = btn
        tk.Frame(self.nav_frame, bg=COLOR_PRIMARY).pack(expand=True)
        tk.Label(self.nav_frame, text="v2.0.6", font=("微软雅黑", 9),
                bg=COLOR_PRIMARY, fg="#5a6680").pack(pady=15)
        self.content_frame = tk.Frame(main_frame, bg=COLOR_BG)
        self.content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=25, pady=25)
        self.status_var = tk.StringVar(value="就绪")
        status_bar = tk.Frame(main_frame, bg=COLOR_PRIMARY, height=30)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        status_bar.pack_propagate(False)
        tk.Label(status_bar, textvariable=self.status_var, 
                font=("微软雅黑", 9), bg=COLOR_PRIMARY, fg="#8892b0").pack(side=tk.LEFT, padx=20)
        self.pages = {}
        self._build_home()
        self._build_single()
        self._build_batch_fwd()
        self._build_batch_rev()
        self._build_settings()

    def _nav_style(self, active_key):
        for k, btn in self.nav_buttons.items():
            if k == active_key:
                btn.config(bg=COLOR_SECONDARY, fg=COLOR_ACCENT, 
                          font=("微软雅黑", 11, "bold"))
            else:
                btn.config(bg=COLOR_PRIMARY, fg="#8892b0",
                          font=("微软雅黑", 11))

    def show_page(self, key):
        self._nav_style(key)
        for p in self.pages.values():
            p.pack_forget()
        self.pages[key].pack(fill=tk.BOTH, expand=True)
        self.current_page = key

    def _make_page_frame(self, parent):
        return tk.Frame(parent, bg=COLOR_BG)

    def _create_h_table(self, parent, rows, widths, align="center"):
        table_frame = tk.Frame(parent, bg=COLOR_CARD)
        table_frame.pack(fill=tk.X, padx=15, pady=5)
        for r_idx, row in enumerate(rows):
            if r_idx == 0:
                bg = COLOR_SECONDARY
                fg = "white"
                font = ("微软雅黑", 10, "bold")
            else:
                bg = "#f8f9fa" if r_idx % 2 == 0 else COLOR_CARD
                fg = COLOR_TEXT
                font = ("微软雅黑", 10)
            if r_idx > 0:
                sep_line = tk.Frame(table_frame, bg=COLOR_BORDER, height=1)
                sep_line.pack(fill=tk.X)
            row_container = tk.Frame(table_frame, bg=bg)
            row_container.pack(fill=tk.X)
            for c_idx, (cell, w) in enumerate(zip(row, widths)):
                anchor = "center" if align == "center" else "w"
                justify = tk.CENTER if align == "center" else tk.LEFT
                lbl = tk.Label(row_container, text=cell, font=font,
                              bg=bg, fg=fg, width=w, anchor=anchor, justify=justify)
                lbl.pack(side=tk.LEFT, padx=1, pady=5)

    def _build_home(self):
        outer_frame = self._make_page_frame(self.content_frame)
        self.pages["home"] = outer_frame

        canvas = tk.Canvas(outer_frame, bg=COLOR_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=COLOR_BG)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=1100)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        frame = scrollable_frame

        header = tk.Frame(frame, bg=COLOR_BG)
        header.pack(fill=tk.X, pady=(0, 20))
        tk.Label(header, text="欢迎使用地址转换工具", 
                font=("微软雅黑", 26, "bold"),
                bg=COLOR_BG, fg=COLOR_TEXT).pack(anchor="w")
        tk.Label(header, text="基于高德/百度地图 API 的地址智能解析与坐标转换工具", 
                font=("微软雅黑", 12),
                bg=COLOR_BG, fg=COLOR_TEXT_LIGHT).pack(anchor="w", pady=(5, 0))

        cards_frame = tk.Frame(frame, bg=COLOR_BG)
        cards_frame.pack(fill=tk.X, pady=10)

        features = [
            ("📍 单条查询", "快速查询单个地址的\nGCJ-02 / WGS-84 坐标", COLOR_INFO),
            ("📊 批量正向编码", "从 Excel 读取地址\n批量获取标准化结果", COLOR_SUCCESS),
            ("🔄 批量逆向编码", "从 Excel 读取坐标\n批量解析为标准地址", COLOR_WARNING),
            ("⚙️ 双平台配置", "支持高德 + 百度\n智能 Fallback 切换", COLOR_DANGER),
        ]

        for i, (title, desc, color) in enumerate(features):
            card = tk.Frame(cards_frame, bg=COLOR_CARD, highlightbackground=COLOR_BORDER,
                          highlightthickness=1)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0 if i==0 else 10, 0))
            tk.Frame(card, bg=color, height=4).pack(fill=tk.X)
            tk.Label(card, text=title, font=("微软雅黑", 13, "bold"),
                    bg=COLOR_CARD, fg=COLOR_TEXT).pack(anchor="w", padx=15, pady=(15, 5))
            tk.Label(card, text=desc, font=("微软雅黑", 10),
                    bg=COLOR_CARD, fg=COLOR_TEXT_LIGHT, justify=tk.LEFT).pack(anchor="w", padx=15, pady=(0, 15))

        step_card = Card(frame, title="使用步骤")
        step_card.pack(fill=tk.X, pady=20)

        steps = [
            "1. 点击左侧「API 设置」，配置您的高德/百度 API Key（支持自动保存）",
            "2. 准备好 Excel 文件，确保包含 Sheet1 工作表，按下方表格格式填写数据",
            "3. 选择「批量正向编码」或「批量逆向编码」，选择文件后开始处理",
            "4. 处理结果自动写入 Excel，异常行以黄色高亮标记，支持断点续传",
        ]
        for step in steps:
            tk.Label(step_card, text=step, font=("微软雅黑", 11),
                    bg=COLOR_CARD, fg=COLOR_TEXT).pack(anchor="w", padx=15, pady=5)

        # ===== 标题 + 下载按钮 同行布局（参考截图） =====
        format_header_frame = tk.Frame(frame, bg=COLOR_BG)
        format_header_frame.pack(fill=tk.X, pady=(20, 0))
        tk.Label(format_header_frame, text="Excel 数据格式说明",
                font=("微软雅黑", 14, "bold"),
                bg=COLOR_BG, fg=COLOR_TEXT).pack(side=tk.LEFT, anchor="w")
        StyledButton(format_header_frame, "  下载 Excel 模板", self.download_template, "primary").pack(side=tk.RIGHT, padx=(0, 0))
        tk.Label(format_header_frame, text=" 含示例数据→",
                font=("微软雅黑", 10),
                bg=COLOR_BG, fg=COLOR_TEXT_LIGHT).pack(side=tk.RIGHT, padx=10)

        format_card = Card(frame)
        format_card.pack(fill=tk.X, pady=(10, 20))

        fwd_frame = tk.Frame(format_card, bg=COLOR_CARD)
        fwd_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        tk.Label(fwd_frame, text="▎正向编码（地址 → 坐标）", font=("微软雅黑", 12, "bold"),
                bg=COLOR_CARD, fg=COLOR_SUCCESS).pack(anchor="w")

        # 调整列宽：压缩前几列，增大标准地址列，确保标题完整显示
        self._create_h_table(format_card, [
            ["序号", "省\n(选填)", "市\n(选填)", "县/区\n(选填)", "地址1\n(必填)", "地址2\n(选填)", 
             "GCJ-02\n经度", "GCJ-02\n纬度", "WGS-84\n经度", "WGS-84\n纬度", "标准地址"],
            ["输入", "输入", "输入", "输入", "输入", "输入", 
             "输出", "输出", "输出", "输出", "输出"],
        ], [6, 8, 8, 8, 16, 14, 11, 11, 11, 11, 28])

        rev_frame = tk.Frame(format_card, bg=COLOR_CARD)
        rev_frame.pack(fill=tk.X, padx=15, pady=(20, 5))
        tk.Label(rev_frame, text="▎逆向编码（坐标 → 地址）", font=("微软雅黑", 12, "bold"),
                bg=COLOR_CARD, fg=COLOR_WARNING).pack(anchor="w")

        self._create_h_table(format_card, [
            ["GCJ-02\n经度", "GCJ-02\n纬度", "WGS-84\n经度", "WGS-84\n纬度", 
             "逆地理\n完整地址", "省 - 市 - 区", "街道\n门牌号"],
            ["输入", "输入", "输入(备选)", "输入(备选)", 
             "输出", "输出", "输出"],
        ], [12, 12, 12, 12, 18, 18, 12])

    def download_template(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="地址标准化模板.xlsx"
        )
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            headers = ["序号", "省", "市", "区县", "地址1(必填)", "地址2(选填)",
                      "GCJ-02经度", "GCJ-02纬度", "WGS-84经度", "WGS-84纬度",
                      "标准地址", "完整地址", "省市区", "街道门牌号"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            examples = [
                [1, "四川省", "成都市", " ", "西南电子科技大学", "", "", "", "", "", "", "", "", ""],
                [2, "北京市", "北京市", "", "天安门", "人民英雄纪念碑", "", "", "", "", "", "", "", ""],
                [3, "云南省", "曲靖市", "麒麟区", "曲靖市农村信用合作社联合社", "云南省曲靖市麒麟区寥廓北路中段271-1号", "", "", "", "", "", "", "", ""],
            ]
            for row_idx, row_data in enumerate(examples, 2):
                for col_idx, value in enumerate(row_data, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            col_widths = [8, 12, 12, 12, 25, 20, 15, 15, 15, 15, 40, 30, 20, 20]
            for i, width in enumerate(col_widths, 1):
                ws.column_dimensions[get_column_letter(i)].width = width
            wb.save(path)
            messagebox.showinfo("成功", f"模板已保存到：\n{path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存模板失败：{e}")

    def _build_single(self):
        frame = self._make_page_frame(self.content_frame)
        self.pages["single"] = frame

        tk.Label(frame, text="单条地址查询", font=("微软雅黑", 22, "bold"),
                bg=COLOR_BG, fg=COLOR_TEXT).pack(anchor="w", pady=(0, 20))

        card = Card(frame, title="查询参数")
        card.pack(fill=tk.X, pady=10)

        form = tk.Frame(card, bg=COLOR_CARD)
        form.pack(fill=tk.X, padx=15, pady=10)

        fields = [
            ("* 地址（必填）", "single_addr", "曲靖市麒麟区农村信用合作联社三宝信用社"),
            ("  城市（选填）", "single_city", ""),
            ("  省份（选填）", "single_prov", ""),
            ("  区县（选填）", "single_dist", ""),
        ]

        for i, (label, attr, default) in enumerate(fields):
            tk.Label(form, text=label, font=("微软雅黑", 11),
                    bg=COLOR_CARD, fg=COLOR_TEXT).grid(row=i, column=0, sticky="w", pady=10, padx=5)
            entry = tk.Entry(form, font=("微软雅黑", 11), width=60, 
                           relief=tk.SOLID, bd=1, highlightcolor=COLOR_ACCENT,
                           highlightthickness=1)
            entry.grid(row=i, column=1, sticky="ew", padx=10)
            entry.insert(0, default)
            setattr(self, attr, entry)

        form.columnconfigure(1, weight=1)

        btn_frame = tk.Frame(card, bg=COLOR_CARD)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        StyledButton(btn_frame, "  查询", self.do_single_query, "primary").pack(side=tk.LEFT)
        tk.Button(btn_frame, text="清空", font=("微软雅黑", 11),
                 bg=COLOR_BORDER, fg=COLOR_TEXT, bd=0, padx=20, pady=8,
                 command=self.clear_single).pack(side=tk.LEFT, padx=10)

        result_card = Card(frame, title="查询结果")
        result_card.pack(fill=tk.BOTH, expand=True, pady=10)

        self.single_result = scrolledtext.ScrolledText(
            result_card, font=("Consolas", 11), wrap=tk.WORD, 
            height=15, bg="#fafafa", fg=COLOR_TEXT,
            relief=tk.SOLID, bd=1, padx=10, pady=10
        )
        self.single_result.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

    def clear_single(self):
        self.single_addr.delete(0, tk.END)
        self.single_city.delete(0, tk.END)
        self.single_prov.delete(0, tk.END)
        self.single_dist.delete(0, tk.END)
        self.single_result.delete("1.0", tk.END)

    def do_single_query(self):
        addr = self.single_addr.get().strip()
        if not addr:
            messagebox.showwarning("提示", "请输入地址")
            return
        if not self.config.get("gaode_key"):
            messagebox.showwarning("提示", "请先配置高德 API Key")
            self.show_page("settings")
            return

        self.status_var.set("正在查询...")
        self.single_result.delete("1.0", tk.END)
        self.single_result.insert(tk.END, "正在查询，请稍候...")

        def task():
            self.geo = GeoCoder(self.config)
            ok, msg = self.geo.check_keys()
            if not ok:
                self.root.after(0, lambda: messagebox.showerror("密钥错误", msg))
                self.root.after(0, lambda: self.status_var.set("就绪"))
                return
            result, is_abnormal, err_msg, fatal, baidu_fallback_log = self.geo.resolve_address(
                addr, self.single_prov.get().strip(), 
                self.single_city.get().strip(), 
                self.single_dist.get().strip()
            )
            self.root.after(0, lambda: self._show_single_result(result, is_abnormal, err_msg, fatal, baidu_fallback_log))

        threading.Thread(target=task, daemon=True).start()

    def _show_single_result(self, result, is_abnormal, err_msg, fatal, baidu_fallback_log=""):
        self.status_var.set("就绪")
        self.single_result.delete("1.0", tk.END)

        if fatal:
            self.single_result.insert(tk.END, f"查询失败（致命错误）：{err_msg}")
            return
        if not result:
            self.single_result.insert(tk.END, f"未找到结果：{err_msg}")
            return

        gcj_lng, gcj_lat = result["gcj_lng"], result["gcj_lat"]
        wgs_lng, wgs_lat = gcj02_to_wgs84(gcj_lng, gcj_lat)

        text = f"""【查询成功】

标准地址：{result['std_addr']}
精度等级：{result['level']}
数据来源：{result['source']}

【GCJ-02 坐标】
  经度：{gcj_lng}
  纬度：{gcj_lat}

【WGS-84 坐标】
  经度：{round(wgs_lng, 6)}
  纬度：{round(wgs_lat, 6)}
"""
        if baidu_fallback_log:
            text += f"\n【百度Fallback过程】\n{baidu_fallback_log}"
        if result.get("is_drift"):
            text += "\n【警告】疑似坐标漂移，请人工复核"
        if is_abnormal:
            text += "\n【注意】结果存在异常"

        self.single_result.insert(tk.END, text)

    def _build_batch_fwd(self):
        frame = self._make_page_frame(self.content_frame)
        self.pages["batch_fwd"] = frame

        tk.Label(frame, text="批量正向地理编码（地址 -> 坐标）", 
                font=("微软雅黑", 22, "bold"),
                bg=COLOR_BG, fg=COLOR_TEXT).pack(anchor="w", pady=(0, 20))

        param_card = Card(frame, title="处理参数")
        param_card.pack(fill=tk.X, pady=10)

        file_frame = tk.Frame(param_card, bg=COLOR_CARD)
        file_frame.pack(fill=tk.X, padx=15, pady=10)
        tk.Label(file_frame, text="Excel 文件：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT).pack(side=tk.LEFT)
        self.fwd_path_var = tk.StringVar()
        tk.Entry(file_frame, textvariable=self.fwd_path_var, 
                font=("微软雅黑", 11), width=60,
                relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(file_frame, text="浏览...", font=("微软雅黑", 10),
                 bg=COLOR_BORDER, fg=COLOR_TEXT, bd=0, padx=15, pady=5,
                 command=self.select_fwd_file).pack(side=tk.LEFT)

        param_row = tk.Frame(param_card, bg=COLOR_CARD)
        param_row.pack(fill=tk.X, padx=15, pady=(0, 15))
        tk.Label(param_row, text="起始行：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT).pack(side=tk.LEFT)
        self.fwd_start_row = tk.Spinbox(param_row, from_=2, to=100000, width=8, 
                                       font=("微软雅黑", 11), relief=tk.SOLID, bd=1)
        self.fwd_start_row.pack(side=tk.LEFT, padx=5)
        self.fwd_start_row.delete(0, tk.END)
        self.fwd_start_row.insert(0, "2")

        self.fwd_skip_var = tk.BooleanVar(value=True)
        tk.Checkbutton(param_row, text="跳过已有结果", variable=self.fwd_skip_var,
                      font=("微软雅黑", 11), bg=COLOR_CARD, fg=COLOR_TEXT,
                      selectcolor=COLOR_CARD, activebackground=COLOR_CARD).pack(side=tk.LEFT, padx=20)

        ctrl_frame = tk.Frame(param_card, bg=COLOR_CARD)
        ctrl_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        self.fwd_start_btn = StyledButton(ctrl_frame, "  开始处理", self.start_batch_fwd, "primary")
        self.fwd_start_btn.pack(side=tk.LEFT)
        self.fwd_stop_btn = StyledButton(ctrl_frame, "  停止", self.stop_batch, "danger")
        self.fwd_stop_btn.config(state=tk.DISABLED)
        self.fwd_stop_btn.pack(side=tk.LEFT, padx=10)

        progress_frame = tk.Frame(frame, bg=COLOR_BG)
        progress_frame.pack(fill=tk.X, pady=10)
        self.fwd_progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, 
                                           mode="determinate", length=400)
        self.fwd_progress.pack(side=tk.LEFT)
        self.fwd_progress_label = tk.Label(progress_frame, text="0 / 0", 
                                          font=("微软雅黑", 11),
                                          bg=COLOR_BG, fg=COLOR_TEXT)
        self.fwd_progress_label.pack(side=tk.LEFT, padx=15)

        log_card = Card(frame, title="处理日志")
        log_card.pack(fill=tk.BOTH, expand=True, pady=10)

        self.fwd_log = scrolledtext.ScrolledText(
            log_card, font=("Consolas", 10), wrap=tk.WORD, 
            height=12, bg="#fafafa", fg=COLOR_TEXT,
            relief=tk.SOLID, bd=1, padx=10, pady=10
        )
        self.fwd_log.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

    def select_fwd_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if path:
            self.fwd_path_var.set(path)

    def start_batch_fwd(self):
        path = self.fwd_path_var.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的 Excel 文件")
            return
        if not self.config.get("gaode_key"):
            messagebox.showwarning("提示", "请先配置高德 API Key")
            self.show_page("settings")
            return
        try:
            start_row = int(self.fwd_start_row.get())
        except ValueError:
            start_row = 2

        self.running = True
        self.stop_flag = False
        self.fwd_start_btn.config(state=tk.DISABLED)
        self.fwd_stop_btn.config(state=tk.NORMAL)
        self.fwd_log.delete("1.0", tk.END)

        def task():
            try:
                proc = ExcelProcessor(path)
                last_row = proc.get_last_row(COL_ADDR1)
                total = last_row - start_row + 1
                if total <= 0:
                    self.root.after(0, lambda: self.rev_log.insert(tk.END, "提示：没有找到需要处理的数据\n"))
                    self.root.after(0, lambda: self.rev_log.see(tk.END))
                    return

                self.geo = GeoCoder(self.config)
                ok, msg = self.geo.check_keys()
                if not ok:
                    self.root.after(0, lambda m=msg: self.fwd_log.insert(tk.END, f"[密钥错误] {m}\n"))
                    self.root.after(0, lambda: self.fwd_log.see(tk.END))
                    return

                self.root.after(0, lambda: self.fwd_progress.config(maximum=total, value=0))
                success = fail = skip = addr2_used = drift_cnt = baidu_upgrade = 0

                for i in range(start_row, last_row + 1):
                    if self.stop_flag:
                        self.root.after(0, lambda: self.fwd_log.insert(tk.END, "\n用户取消处理\n"))
                        break

                    progress_val = i - start_row + 1
                    self.root.after(0, lambda v=progress_val, t=total: self.fwd_progress_label.config(text=f"{v} / {t}"))
                    self.root.after(0, lambda v=progress_val: self.fwd_progress.config(value=v))

                    if self.fwd_skip_var.get() and proc.is_result_normal(i):
                        skip += 1
                        self.root.after(0, lambda r=i: self.fwd_log.insert(tk.END, f"第 {r} 行：跳过（已有结果）\n"))
                        self.root.after(0, lambda: self.fwd_log.see(tk.END))
                        continue

                    proc.clear_forward_result(i)
                    prov = str(proc.ws.cell(row=i, column=COL_PROVINCE).value or "").strip()
                    city = str(proc.ws.cell(row=i, column=COL_CITY).value or "").strip()
                    dist = str(proc.ws.cell(row=i, column=COL_DISTRICT).value or "").strip()
                    addr1 = str(proc.ws.cell(row=i, column=COL_ADDR1).value or "").strip()
                    addr2 = str(proc.ws.cell(row=i, column=COL_ADDR2).value or "").strip()

                    if not addr1:
                        proc.write_forward(i, None)
                        fail += 1
                        self.root.after(0, lambda r=i: self.fwd_log.insert(tk.END, f"第 {r} 行：地址1为空\n"))
                        self.root.after(0, lambda: self.fwd_log.see(tk.END))
                        continue

                    result, is_abnormal, err_msg, fatal, baidu_fallback_log = self.geo.resolve_address(addr1, prov, city, dist)
                    if fatal:
                        self.root.after(0, lambda e=err_msg: self.fwd_log.insert(tk.END, f"致命错误：{e}，已终止\n"))
                        break

                    # 内容匹配检查（对照 VB 主过程逻辑）
                    if result and not is_abnormal and result.get("source") == "高德" and addr2:
                        core = extract_core_name(addr1, prov, city, dist)
                        if core and not address_contains_core(result["std_addr"], core):
                            is_abnormal = True
                            result["is_abnormal"] = True
                            result["std_addr"] += " [内容不匹配]"

                    # 地址2 Fallback（对照 VB 主过程逻辑）
                    if (not result) or is_abnormal:
                        if addr2:
                            result2, ab2, err2, fatal2, _ = self.geo.resolve_address(addr2, prov, city, dist)
                            if fatal2:
                                self.root.after(0, lambda e=err2: self.fwd_log.insert(tk.END, f"致命错误：{e}，已终止\n"))
                                break
                            if result2 and not ab2:
                                result = result2
                                addr2_used += 1
                                if result.get("is_drift"):
                                    drift_cnt += 1
                                if result.get("source") == "百度":
                                    baidu_upgrade += 1
                            else:
                                result = None
                        else:
                            result = None

                    proc.write_forward(i, result)
                    if result:
                        success += 1
                        status = "成功"
                    else:
                        fail += 1
                        status = "失败"

                    # 输出百度Fallback日志（如果有）
                    if baidu_fallback_log:
                        self.root.after(0, lambda log=baidu_fallback_log: self.fwd_log.insert(tk.END, log))
                        self.root.after(0, lambda: self.fwd_log.see(tk.END))
                    self.root.after(0, lambda r=i, s=status: self.fwd_log.insert(tk.END, f"第 {r} 行：{s}\n"))
                    self.root.after(0, lambda: self.fwd_log.see(tk.END))
                    time.sleep(self.config.get("base_interval", 0.2))

                proc.save()
                msg = (f"\n========== 处理完成 ==========\n"
                       f"总行数：{total}\n"
                       f"成功：{success}\n"
                       f"使用地址2：{addr2_used}\n"
                       f"百度优化提升：{baidu_upgrade}\n"
                       f"疑似漂移：{drift_cnt}\n"
                       f"失败：{fail}\n"
                       f"跳过：{skip}\n"
                       f"高德调用：{self.geo.gaode_count}\n"
                       f"百度调用：{self.geo.baidu_count}\n"
                       f"==============================\n")
                self.root.after(0, lambda m=msg: self.fwd_log.insert(tk.END, m))
                self.root.after(0, lambda: self.fwd_log.see(tk.END))

            except Exception as e:
                err_msg = f"[处理异常] {str(e)}\n"
                self.root.after(0, lambda m=err_msg: self.fwd_log.insert(tk.END, m))
                self.root.after(0, lambda: self.fwd_log.see(tk.END))
            finally:
                self.running = False
                self.root.after(0, self._batch_finished)

        threading.Thread(target=task, daemon=True).start()

    def _batch_finished(self):
        self.fwd_start_btn.config(state=tk.NORMAL)
        self.fwd_stop_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")

    def stop_batch(self):
        self.stop_flag = True
        if self.geo:
            self.geo.stop()
        self.status_var.set("正在停止...")

    def _build_batch_rev(self):
        frame = self._make_page_frame(self.content_frame)
        self.pages["batch_rev"] = frame

        tk.Label(frame, text="批量逆向地理编码（坐标 -> 地址）", 
                font=("微软雅黑", 22, "bold"),
                bg=COLOR_BG, fg=COLOR_TEXT).pack(anchor="w", pady=(0, 20))

        param_card = Card(frame, title="处理参数")
        param_card.pack(fill=tk.X, pady=10)

        file_frame = tk.Frame(param_card, bg=COLOR_CARD)
        file_frame.pack(fill=tk.X, padx=15, pady=10)
        tk.Label(file_frame, text="Excel 文件：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT).pack(side=tk.LEFT)
        self.rev_path_var = tk.StringVar()
        tk.Entry(file_frame, textvariable=self.rev_path_var, 
                font=("微软雅黑", 11), width=60,
                relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(file_frame, text="浏览...", font=("微软雅黑", 10),
                 bg=COLOR_BORDER, fg=COLOR_TEXT, bd=0, padx=15, pady=5,
                 command=self.select_rev_file).pack(side=tk.LEFT)

        param_row = tk.Frame(param_card, bg=COLOR_CARD)
        param_row.pack(fill=tk.X, padx=15, pady=(0, 15))
        tk.Label(param_row, text="起始行：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT).pack(side=tk.LEFT)
        self.rev_start_row = tk.Spinbox(param_row, from_=2, to=100000, width=8,
                                       font=("微软雅黑", 11), relief=tk.SOLID, bd=1)
        self.rev_start_row.pack(side=tk.LEFT, padx=5)
        self.rev_start_row.delete(0, tk.END)
        self.rev_start_row.insert(0, "2")

        self.rev_skip_var = tk.BooleanVar(value=True)
        tk.Checkbutton(param_row, text="跳过已有结果", variable=self.rev_skip_var,
                      font=("微软雅黑", 11), bg=COLOR_CARD, fg=COLOR_TEXT,
                      selectcolor=COLOR_CARD, activebackground=COLOR_CARD).pack(side=tk.LEFT, padx=20)

        ctrl_frame = tk.Frame(param_card, bg=COLOR_CARD)
        ctrl_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        self.rev_start_btn = StyledButton(ctrl_frame, "  开始处理", self.start_batch_rev, "primary")
        self.rev_start_btn.pack(side=tk.LEFT)
        self.rev_stop_btn = StyledButton(ctrl_frame, "  停止", self.stop_batch, "danger")
        self.rev_stop_btn.config(state=tk.DISABLED)
        self.rev_stop_btn.pack(side=tk.LEFT, padx=10)

        progress_frame = tk.Frame(frame, bg=COLOR_BG)
        progress_frame.pack(fill=tk.X, pady=10)
        self.rev_progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL,
                                           mode="determinate", length=400)
        self.rev_progress.pack(side=tk.LEFT)
        self.rev_progress_label = tk.Label(progress_frame, text="0 / 0",
                                          font=("微软雅黑", 11),
                                          bg=COLOR_BG, fg=COLOR_TEXT)
        self.rev_progress_label.pack(side=tk.LEFT, padx=15)

        log_card = Card(frame, title="处理日志")
        log_card.pack(fill=tk.BOTH, expand=True, pady=10)

        self.rev_log = scrolledtext.ScrolledText(
            log_card, font=("Consolas", 10), wrap=tk.WORD,
            height=12, bg="#fafafa", fg=COLOR_TEXT,
            relief=tk.SOLID, bd=1, padx=10, pady=10
        )
        self.rev_log.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

    def select_rev_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if path:
            self.rev_path_var.set(path)

    def start_batch_rev(self):
        path = self.rev_path_var.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请选择有效的 Excel 文件")
            return
        if not self.config.get("gaode_key"):
            messagebox.showwarning("提示", "请先配置高德 API Key")
            self.show_page("settings")
            return
        try:
            start_row = int(self.rev_start_row.get())
        except ValueError:
            start_row = 2

        self.running = True
        self.stop_flag = False
        self.rev_start_btn.config(state=tk.DISABLED)
        self.rev_stop_btn.config(state=tk.NORMAL)
        self.rev_log.delete("1.0", tk.END)

        def task():
            try:
                proc = ExcelProcessor(path)
                last_row = max(proc.get_last_row(COL_GCJ_LNG), proc.get_last_row(COL_WGS_LNG))
                total = last_row - start_row + 1
                if total <= 0:
                    self.root.after(0, lambda: self.rev_log.insert(tk.END, "提示：没有找到需要处理的数据\n"))
                    self.root.after(0, lambda: self.rev_log.see(tk.END))
                    return

                self.geo = GeoCoder(self.config)
                ok, msg = self.geo.check_keys()
                if not ok:
                    self.root.after(0, lambda m=msg: self.rev_log.insert(tk.END, f"[密钥错误] {m}\n"))
                    self.root.after(0, lambda: self.rev_log.see(tk.END))
                    return

                self.root.after(0, lambda: self.rev_progress.config(maximum=total, value=0))
                success = fail = skip = 0
                global_error = False

                for i in range(start_row, last_row + 1):
                    if self.stop_flag:
                        self.root.after(0, lambda: self.rev_log.insert(tk.END, "\n用户取消处理\n"))
                        break

                    progress_val = i - start_row + 1
                    self.root.after(0, lambda v=progress_val, t=total: self.rev_progress_label.config(text=f"{v} / {t}"))
                    self.root.after(0, lambda v=progress_val: self.rev_progress.config(value=v))

                    if self.rev_skip_var.get():
                        existing = str(proc.ws.cell(row=i, column=COL_REGO_FULL).value or "").strip()
                        if existing and "失败" not in existing and "错误" not in existing:
                            skip += 1
                            self.root.after(0, lambda r=i: self.rev_log.insert(tk.END, f"第 {r} 行：跳过\n"))
                            self.root.after(0, lambda: self.rev_log.see(tk.END))
                            continue

                    proc.clear_reverse_result(i)

                    # 优先使用 GCJ-02，无效则用 WGS-84 并转换（对照 VB 逻辑）
                    gcj_lng = proc.ws.cell(row=i, column=COL_GCJ_LNG).value
                    gcj_lat = proc.ws.cell(row=i, column=COL_GCJ_LAT).value
                    try:
                        gcj_lng = float(gcj_lng) if gcj_lng else 0
                        gcj_lat = float(gcj_lat) if gcj_lat else 0
                    except (ValueError, TypeError):
                        gcj_lng, gcj_lat = 0, 0

                    if not (gcj_lng and gcj_lat and 73 <= gcj_lng <= 135 and 3 <= gcj_lat <= 53):
                        wgs_lng = proc.ws.cell(row=i, column=COL_WGS_LNG).value
                        wgs_lat = proc.ws.cell(row=i, column=COL_WGS_LAT).value
                        try:
                            wgs_lng = float(wgs_lng) if wgs_lng else 0
                            wgs_lat = float(wgs_lat) if wgs_lat else 0
                        except (ValueError, TypeError):
                            wgs_lng, wgs_lat = 0, 0
                        if wgs_lng and wgs_lat and 73 <= wgs_lng <= 135 and 3 <= wgs_lat <= 53:
                            gcj_lng, gcj_lat = wgs84_to_gcj02(wgs_lng, wgs_lat)
                        else:
                            proc.write_reverse(i, None)
                            fail += 1
                            self.root.after(0, lambda r=i: self.rev_log.insert(tk.END, f"第 {r} 行：无有效坐标\n"))
                            self.root.after(0, lambda: self.rev_log.see(tk.END))
                            continue

                    if not (gcj_lng and gcj_lat):
                        proc.ws.cell(row=i, column=COL_REGO_FULL).value = "无有效坐标"
                        fail += 1
                        self.root.after(0, lambda r=i: self.rev_log.insert(tk.END, f"第 {r} 行：无有效坐标\n"))
                        self.root.after(0, lambda: self.rev_log.see(tk.END))
                        continue

                    result, err = self.geo.gaode_regeocode(gcj_lng, gcj_lat)
                    if not result:
                        proc.ws.cell(row=i, column=COL_REGO_FULL).value = f"失败：{err}"
                        # 全局错误检查（10001/10003）
                        if "10001" in err or "10003" in err:
                            global_error = True
                            fail += 1
                            self.root.after(0, lambda r=i, e=err: self.rev_log.insert(tk.END, f"第 {r} 行：全局错误 {e}，已终止\n"))
                            self.root.after(0, lambda: self.rev_log.see(tk.END))
                            break
                        fail += 1
                        self.root.after(0, lambda r=i, e=err: self.rev_log.insert(tk.END, f"第 {r} 行：失败 - {e}\n"))
                    else:
                        proc.write_reverse(i, result)
                        success += 1
                        self.root.after(0, lambda r=i: self.rev_log.insert(tk.END, f"第 {r} 行：成功\n"))
                    self.root.after(0, lambda: self.rev_log.see(tk.END))
                    time.sleep(self.config.get("base_interval", 0.2))

                proc.save()
                if global_error:
                    self.root.after(0, lambda: self.rev_log.insert(tk.END, "[全局错误] 触发全局错误，已终止处理！\n"))
                    self.root.after(0, lambda: self.rev_log.see(tk.END))
                else:
                    msg = (f"\n========== 处理完成 ==========\n"
                           f"总行数：{total}\n"
                           f"成功：{success}\n"
                           f"失败：{fail}\n"
                           f"跳过：{skip}\n"
                           f"实际高德调用：{self.geo.gaode_count}\n"
                           f"运行耗时：请查看日志\n"
                           f"==============================\n")
                    self.root.after(0, lambda m=msg: self.rev_log.insert(tk.END, m))
                    self.root.after(0, lambda: self.rev_log.see(tk.END))

            except Exception as e:
                err_msg = f"[处理异常] {str(e)}\n"
                self.root.after(0, lambda m=err_msg: self.rev_log.insert(tk.END, m))
                self.root.after(0, lambda: self.rev_log.see(tk.END))
            finally:
                self.running = False
                self.root.after(0, self._rev_batch_finished)

        threading.Thread(target=task, daemon=True).start()

    def _rev_batch_finished(self):
        self.rev_start_btn.config(state=tk.NORMAL)
        self.rev_stop_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")

    def _build_settings(self):
        frame = self._make_page_frame(self.content_frame)
        self.pages["settings"] = frame

        tk.Label(frame, text="API 密钥配置", 
                font=("微软雅黑", 22, "bold"),
                bg=COLOR_BG, fg=COLOR_TEXT).pack(anchor="w", pady=(0, 20))

        card = Card(frame, title="密钥设置")
        card.pack(fill=tk.X, pady=10)

        form = tk.Frame(card, bg=COLOR_CARD)
        form.pack(fill=tk.X, padx=15, pady=15)

        row1 = tk.Frame(form, bg=COLOR_CARD)
        row1.pack(fill=tk.X, pady=8)
        tk.Label(row1, text="* 高德 API Key：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT, width=15, anchor="e").pack(side=tk.LEFT)
        self.cfg_gaode = tk.Entry(row1, font=("微软雅黑", 11), width=55, show="*",
                                 relief=tk.SOLID, bd=1)
        self.cfg_gaode.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.cfg_gaode.insert(0, self.config.get("gaode_key", ""))
        tk.Button(row1, text="显示/隐藏", font=("微软雅黑", 9),
                 bg=COLOR_BORDER, fg=COLOR_TEXT, bd=0, padx=10, pady=3,
                 command=lambda: self.cfg_gaode.config(show="" if self.cfg_gaode.cget("show") == "*" else "*")).pack(side=tk.LEFT)

        row2 = tk.Frame(form, bg=COLOR_CARD)
        row2.pack(fill=tk.X, pady=8)
        tk.Label(row2, text="  百度 AK：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT, width=15, anchor="e").pack(side=tk.LEFT)
        self.cfg_baidu = tk.Entry(row2, font=("微软雅黑", 11), width=55, show="*",
                                 relief=tk.SOLID, bd=1)
        self.cfg_baidu.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.cfg_baidu.insert(0, self.config.get("baidu_ak", ""))
        tk.Button(row2, text="显示/隐藏", font=("微软雅黑", 9),
                 bg=COLOR_BORDER, fg=COLOR_TEXT, bd=0, padx=10, pady=3,
                 command=lambda: self.cfg_baidu.config(show="" if self.cfg_baidu.cget("show") == "*" else "*")).pack(side=tk.LEFT)

        opt_frame = tk.Frame(form, bg=COLOR_CARD)
        opt_frame.pack(fill=tk.X, pady=10)

        self.cfg_enable_baidu = tk.BooleanVar(value=self.config.get("enable_baidu_fallback", True))
        tk.Checkbutton(opt_frame, text="启用百度地图 Fallback（高德精度不足时自动切换）",
                      variable=self.cfg_enable_baidu, font=("微软雅黑", 11),
                      bg=COLOR_CARD, fg=COLOR_TEXT, selectcolor=COLOR_CARD).pack(anchor="w")

        row3 = tk.Frame(form, bg=COLOR_CARD)
        row3.pack(fill=tk.X, pady=8)
        tk.Label(row3, text="百度触发等级：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT, width=15, anchor="e").pack(side=tk.LEFT)
        self.cfg_trigger = ttk.Combobox(row3, values=["门址", "街道", "乡镇", "区县", "城市"],
                                       font=("微软雅黑", 11), width=15, state="readonly")
        self.cfg_trigger.set(self.config.get("baidu_trigger_level", "区县"))
        self.cfg_trigger.pack(side=tk.LEFT, padx=10)

        tk.Label(row3, text="请求间隔（秒）：", font=("微软雅黑", 11),
                bg=COLOR_CARD, fg=COLOR_TEXT).pack(side=tk.LEFT, padx=(30, 0))
        self.cfg_interval = tk.Spinbox(row3, from_=0.1, to=5.0, increment=0.1, width=8,
                                      font=("微软雅黑", 11), relief=tk.SOLID, bd=1)
        self.cfg_interval.delete(0, tk.END)
        self.cfg_interval.insert(0, str(self.config.get("base_interval", 0.2)))
        self.cfg_interval.pack(side=tk.LEFT, padx=10)

        btn_frame = tk.Frame(card, bg=COLOR_CARD)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        StyledButton(btn_frame, "  保存配置", self.save_settings, "primary").pack(side=tk.LEFT)
        StyledButton(btn_frame, "  测试高德", lambda: self.test_api("gaode"), "success").pack(side=tk.LEFT, padx=10)
        StyledButton(btn_frame, "  测试百度", lambda: self.test_api("baidu"), "success").pack(side=tk.LEFT)

        tip_card = Card(frame, title="使用说明")
        tip_card.pack(fill=tk.X, pady=20)

        tips = [
            "• 高德 Key 为必填项，请前往 https://console.amap.com 申请个人开发者 Key",
            "• 百度 AK 为可选项，用于提升解析精度，前往 https://lbsyun.baidu.com 申请",
            "• 配置将自动保存到当前目录的 addr_config.json 文件中，下次打开自动加载",
            "• 高德个人开发者每天约有 5000 次免费地理编码额度，超出需付费",
            "• 建议开启百度 Fallback，当高德返回精度低于「区县」时自动尝试百度",
        ]
        for tip in tips:
            tk.Label(tip_card, text=tip, font=("微软雅黑", 11),
                    bg=COLOR_CARD, fg=COLOR_TEXT).pack(anchor="w", padx=15, pady=4)

    def save_settings(self):
        self.config["gaode_key"] = self.cfg_gaode.get().strip()
        self.config["baidu_ak"] = self.cfg_baidu.get().strip()
        self.config["enable_baidu_fallback"] = self.cfg_enable_baidu.get()
        self.config["baidu_trigger_level"] = self.cfg_trigger.get()
        try:
            self.config["base_interval"] = float(self.cfg_interval.get())
        except ValueError:
            self.config["base_interval"] = 0.2
        save_config(self.config)
        messagebox.showinfo("提示", "配置已保存！")

    def test_api(self, api_type):
        key = self.cfg_gaode.get().strip() if api_type == "gaode" else self.cfg_baidu.get().strip()
        if not key:
            messagebox.showwarning("提示", "请先输入 API Key")
            return
        self.status_var.set(f"正在测试 {api_type} ...")

        def task():
            if api_type == "gaode":
                url = f"{GAODE_GEO_URL}?address={url_encode_utf8('北京市天安门')}&output=json&key={key}"
                ok, resp = http_get(url, 2)
                if ok:
                    status = json_root_value(resp, "status")
                    if status == "1":
                        self.root.after(0, lambda: messagebox.showinfo("测试成功", "高德 API 连接正常！"))
                    else:
                        info = json_root_value(resp, "info")
                        self.root.after(0, lambda: messagebox.showerror("测试失败", f"错误：{info}"))
                else:
                    self.root.after(0, lambda: messagebox.showerror("测试失败", f"请求失败：{resp}"))
            else:
                url = f"{BAIDU_GEO_URL}?address={url_encode_utf8('北京市天安门')}&output=json&ak={key}"
                ok, resp = http_get(url, 2)
                if ok:
                    status = json_root_value(resp, "status")
                    if status == "0":
                        self.root.after(0, lambda: messagebox.showinfo("测试成功", "百度 API 连接正常！"))
                    else:
                        self.root.after(0, lambda: messagebox.showerror("测试失败", f"错误码：{status}"))
                else:
                    self.root.after(0, lambda: messagebox.showerror("测试失败", f"请求失败：{resp}"))
            self.root.after(0, lambda: self.status_var.set("就绪"))

        threading.Thread(target=task, daemon=True).start()

# ===================== 入口 =====================
def main():
    root = tk.Tk()
    app = AddressStandardizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
