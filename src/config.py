#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
地址坐标转换工具
━━━━━━━━━━━━━━━━━━━━━━━━━
模块化重构版 | 支持在线更新
"""

import sys, os, json, math, re, time, threading, random, ssl, urllib.request, urllib.parse, difflib, datetime
import io, platform, subprocess, tempfile, bisect
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
    VERSION       = "2.8.66"
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
        "加油站": ["加油", "petrol station", "油站", "中石化", "中石油"],
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
        "小区": ["花园", "家园", "公寓", "住宅区", "新村", "住宅小区"],
    }

    # AI 大模型平台预设（OpenAI 兼容协议）
    AI_PROVIDERS = {
        "doubao":   {"name": "豆包（火山方舟）",   "base": "https://ark.cn-beijing.volces.com/api/v3",                          "model": "doubao-seed-1-6-250615"},
        "deepseek": {"name": "DeepSeek",           "base": "https://api.deepseek.com/v1",                                     "model": "deepseek-chat"},
        "qwen":     {"name": "通义千问（百炼）",   "base": "https://dashscope.aliyuncs.com/compatible-mode/v1",               "model": "qwen-plus"},
        "kimi":     {"name": "Kimi（月之暗面）",   "base": "https://api.moonshot.cn/v1",                                      "model": "moonshot-v1-8k"},
        "glm":      {"name": "智谱 GLM",           "base": "https://open.bigmodel.cn/api/paas/v4",                            "model": "glm-4-flash"},
        "spark":    {"name": "讯飞星火",           "base": "https://spark-api-open.xf-yun.com/v1",                            "model": "generalv3.5"},
        "wenxin":   {"name": "百度千帆（文心）",   "base": "https://qianfan.baidubce.com/v2",                                 "model": "ernie-speed-8k"},
        "hunyuan":  {"name": "腾讯混元",           "base": "https://api.hunyuan.cloud.tencent.com/v1",                        "model": "hunyuan-lite"},
        "minimax":  {"name": "MiniMax",            "base": "https://api.minimax.chat/v1",                                     "model": "abab6.5s-chat"},
        "openai":   {"name": "OpenAI",             "base": "https://api.openai.com/v1",                                       "model": "gpt-4o-mini"},
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
        "401": "百度并发配额超限（请求过于频繁，请降低频率或稍后重试）",
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
            "auto_open_excel": False,
            "grid_workers": 2,
            # 本次升级：批量并发数 / 增量保存间隔 / 几何校验 / 区县码输出
            "fwd_workers": 3, "save_every": 200,
            "geo_drift_check": True, "write_adcode": True,
            "ai_enabled": False, "ai_provider": "doubao",
            "ai_base_url": "", "ai_api_key": "", "ai_model": "",
            "ai_expand_active": False, "ai_expand_pages": 2, "ai_expand_max_calls": 80
        }
        # 依次尝试：程序目录 → 用户目录（安装目录不可写时的回退位置）
        for path in cls.paths():
            if not os.path.exists(path):
                continue
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    default.update(data)
                break
            except Exception:
                continue
        return default

    @classmethod
    def user_config_path(cls):
        """用户目录下的配置路径（用于安装目录只读时的回退）"""
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, cls.APP_NAME, cls.CONFIG_FILE)

    @classmethod
    def paths(cls):
        """配置查找/保存顺序：程序目录优先，用户目录兜底"""
        return [os.path.abspath(cls.CONFIG_FILE), cls.user_config_path()]

    @classmethod
    def save(cls, cfg):
        """原子写入配置；程序目录不可写时自动回退到用户目录。

        修复：原实现直接 open(...,"w") 写入，一旦目标不可写（安装在 Program Files、
        目录只读、文件被占用）会抛 PermissionError 且调用方未捕获 → 用户点「保存配置」
        无任何反应；写到一半中断还会损坏已有配置。改为「写临时文件 + 原子替换」，
        并在程序目录不可写时回退到 %APPDATA%，返回实际保存路径。
        """
        data = json.dumps(cfg, ensure_ascii=False, indent=2)
        last_err = None
        for path in cls.paths():
            try:
                d = os.path.dirname(path)
                if d:
                    os.makedirs(d, exist_ok=True)
                tmp = path + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    f.write(data)
                os.replace(tmp, path)
                return path
            except Exception as e:
                last_err = e
                try:
                    if os.path.exists(path + ".tmp"):
                        os.remove(path + ".tmp")
                except Exception:
                    pass
        raise OSError(f"配置保存失败：{last_err}")


