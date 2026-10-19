# -*- coding: utf-8 -*-
"""核心业务层：本地数据(SQLite) / 地址清洗 / 地理编码 / POI采集 / AI / Excel / 更新"""
import sys, os, json, math, re, time, threading, random, ssl, urllib.request, urllib.parse, difflib, datetime
import io, platform, subprocess, tempfile, bisect, sqlite3
from src.config import *
from src.config import _get_openpyxl, _get_pil

# ═══════════════════════════════════════════════════════════
# 1.5 本地数据层（SQLite）：结果缓存 / 行政区划表 / 区县边界缓存
# ═══════════════════════════════════════════════════════════
import sqlite3

PROVINCES = [
    "北京市", "天津市", "河北省", "山西省", "内蒙古自治区", "辽宁省", "吉林省", "黑龙江省",
    "上海市", "江苏省", "浙江省", "安徽省", "福建省", "江西省", "山东省", "河南省",
    "湖北省", "湖南省", "广东省", "广西壮族自治区", "海南省", "重庆市", "四川省", "贵州省",
    "云南省", "西藏自治区", "陕西省", "甘肃省", "青海省", "宁夏回族自治区",
    "新疆维吾尔自治区", "台湾省", "香港特别行政区", "澳门特别行政区",
]

_DIV_SUFFIX = ("特别行政区", "维吾尔自治区", "壮族自治区", "回族自治区", "自治区",
               "省", "市辖区", "地区", "自治州", "盟", "省", "市", "自治县", "县", "区", "旗")


def _data_dir():
    """本地数据文件所在目录：优先程序目录（与 addr_config.json 同处），不可写则用用户目录"""
    for p in Config.paths():
        d = os.path.dirname(p) or "."
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".addr_write_probe")
            with open(probe, "w", encoding="utf-8") as f:
                f.write("x")
            os.remove(probe)
            return d
        except Exception:
            continue
    return tempfile.gettempdir()


def norm_division(name):
    """行政区划名归一化：去掉行政级别后缀，便于「沾益」=「沾益区」这类比对"""
    if not name:
        return ""
    s = str(name).strip().replace(" ", "")
    for suf in _DIV_SUFFIX:
        if len(s) > len(suf) and s.endswith(suf):
            s = s[:-len(suf)]
            break
    return s


class LocalDB:
    """SQLite 本地库（线程安全：每线程一个连接 + WAL）。

    - geo_cache：地理编码结果缓存，跨会话复用，直接省 API 配额
    - division ：离线行政区划表（省/市/区县），用于漂移校验与名称纠错
    - polygon  ：区县边界缓存，用于坐标级几何校验（避免重复调用区划接口）
    - stat/meta：命中率等统计
    """
    SCHEMA_VERSION = 1

    def __init__(self, path=None):
        self.path = path or os.path.join(_data_dir(), "addr_cache.db")
        self._local = threading.local()
        self._lock = threading.Lock()
        self.available = False
        self.init_error = ""
        try:
            self._ensure_schema()
            self.available = True
        except Exception as e:
            self.init_error = str(e)

    def conn(self):
        c = getattr(self._local, "conn", None)
        if c is None:
            c = sqlite3.connect(self.path, timeout=10)
            try:
                c.execute("PRAGMA journal_mode=WAL")
                c.execute("PRAGMA synchronous=NORMAL")
            except Exception:
                pass
            self._local.conn = c
        return c

    def _ensure_schema(self):
        with self._lock:
            c = self.conn()
            c.executescript("""
            CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
            CREATE TABLE IF NOT EXISTS geo_cache(
                k TEXT PRIMARY KEY, lng REAL, lat REAL, std_addr TEXT, level TEXT,
                source TEXT, is_drift INTEGER, is_abnormal INTEGER, adcode TEXT,
                hits INTEGER DEFAULT 0, ts INTEGER);
            CREATE TABLE IF NOT EXISTS polygon(adcode TEXT PRIMARY KEY, poly TEXT, ts INTEGER);
            CREATE TABLE IF NOT EXISTS division(
                adcode TEXT PRIMARY KEY, name TEXT, level TEXT, parent TEXT,
                citycode TEXT, norm TEXT);
            CREATE INDEX IF NOT EXISTS idx_div_norm ON division(norm);
            """)
            c.commit()

    # ---------- 通用 KV ----------
    def meta_get(self, k, default=None):
        try:
            r = self.conn().execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
            return r[0] if r else default
        except Exception:
            return default

    def meta_put(self, k, v):
        try:
            with self._lock:
                self.conn().execute("INSERT OR REPLACE INTO meta(k,v) VALUES(?,?)", (k, str(v)))
                self.conn().commit()
        except Exception:
            pass

    def stat_bump(self, k, n=1):
        try:
            with self._lock:
                cur = int(self.meta_get(k, 0) or 0) + n
                self.conn().execute("INSERT OR REPLACE INTO meta(k,v) VALUES(?,?)", (k, str(cur)))
                self.conn().commit()
        except Exception:
            pass

    # ---------- 地理编码结果缓存 ----------
    def cache_get(self, key):
        try:
            r = self.conn().execute(
                "SELECT lng,lat,std_addr,level,source,is_drift,is_abnormal,adcode FROM geo_cache WHERE k=?",
                (key,)).fetchone()
            if not r:
                self.stat_bump("cache_miss")
                return None
            self.stat_bump("cache_hit")
            try:
                self.conn().execute("UPDATE geo_cache SET hits=hits+1 WHERE k=?", (key,))
                self.conn().commit()
            except Exception:
                pass
            return {"gcj_lng": r[0], "gcj_lat": r[1], "std_addr": r[2] or "", "level": r[3] or "",
                    "source": r[4] or "", "is_drift": bool(r[5]), "is_abnormal": bool(r[6]),
                    "adcode": r[7] or "", "from_cache": True}
        except Exception:
            return None

    def cache_put(self, key, d):
        try:
            with self._lock:
                self.conn().execute(
                    "INSERT OR REPLACE INTO geo_cache(k,lng,lat,std_addr,level,source,is_drift,is_abnormal,adcode,hits,ts)"
                    " VALUES(?,?,?,?,?,?,?,?,?,COALESCE((SELECT hits FROM geo_cache WHERE k=?),0),?)",
                    (key, d.get("gcj_lng"), d.get("gcj_lat"), d.get("std_addr", ""), d.get("level", ""),
                     d.get("source", ""), 1 if d.get("is_drift") else 0, 1 if d.get("is_abnormal") else 0,
                     d.get("adcode", ""), key, int(time.time())))
                self.conn().commit()
        except Exception:
            pass

    def cache_count(self):
        try:
            return int(self.conn().execute("SELECT COUNT(*) FROM geo_cache").fetchone()[0])
        except Exception:
            return 0

    def cache_clear(self):
        try:
            with self._lock:
                self.conn().execute("DELETE FROM geo_cache")
                self.conn().execute("DELETE FROM meta WHERE k IN ('cache_hit','cache_miss')")
                self.conn().commit()
            return True
        except Exception:
            return False

    def cache_ratio(self):
        try:
            hit = int(self.meta_get("cache_hit", 0) or 0)
            miss = int(self.meta_get("cache_miss", 0) or 0)
            return hit, miss, (hit * 100.0 / (hit + miss)) if (hit + miss) else 0.0
        except Exception:
            return 0, 0, 0.0

    # ---------- 区县边界缓存 ----------
    def poly_get(self, adcode):
        """返回 (是否命中, 多边形列表)；命中但为空表示"该区县无边界数据"（负缓存）"""
        try:
            r = self.conn().execute("SELECT poly FROM polygon WHERE adcode=?", (adcode,)).fetchone()
            if not r:
                return False, None
            return True, (json.loads(r[0]) if r[0] else [])
        except Exception:
            return False, None

    def poly_put(self, adcode, polys):
        try:
            with self._lock:
                self.conn().execute("INSERT OR REPLACE INTO polygon(adcode,poly,ts) VALUES(?,?,?)",
                                    (adcode, json.dumps(polys or [], ensure_ascii=False), int(time.time())))
                self.conn().commit()
        except Exception:
            pass

    # ---------- 行政区划表 ----------
    def div_count(self):
        try:
            return int(self.conn().execute("SELECT COUNT(*) FROM division").fetchone()[0])
        except Exception:
            return 0

    def div_all(self):
        try:
            return self.conn().execute("SELECT adcode,name,level,parent,citycode,norm FROM division").fetchall()
        except Exception:
            return []

    def div_replace(self, records):
        """整表替换（同步区划时一次性写入）"""
        with self._lock:
            c = self.conn()
            c.execute("DELETE FROM division")
            c.executemany("INSERT OR REPLACE INTO division(adcode,name,level,parent,citycode,norm) VALUES(?,?,?,?,?,?)",
                          records)
            c.commit()


_DB = None
_DB_LOCK = threading.Lock()


def get_db():
    """进程内单例本地库；初始化失败时返回可用但降级的实例"""
    global _DB
    if _DB is None:
        with _DB_LOCK:
            if _DB is None:
                _DB = LocalDB()
    return _DB


class DivisionTable:
    """离线行政区划表：名称归一化 / 纠错 / adcode 查询。

    表内容通过高德行政区划接口一次性同步（每个省级 1 次调用，共 34 次）后写入本地库，
    之后完全离线可用。未同步时所有查询返回空结果，调用方自动降级为原有文本逻辑。
    """

    def __init__(self, db=None):
        self.db = db or get_db()
        self.by_norm = {}      # 归一化名 → [记录]
        self.by_adcode = {}
        self.loaded = False
        self.synced_at = 0
        self.sync_error = ""
        self.load()

    def load(self):
        try:
            rows = self.db.div_all()
        except Exception:
            rows = []
        self.by_norm, self.by_adcode = {}, {}
        for adcode, name, level, parent, citycode, norm in rows:
            rec = {"adcode": adcode or "", "name": name or "", "level": level or "",
                   "parent": parent or "", "citycode": citycode or "",
                   "norm": norm or norm_division(name)}
            self.by_adcode[rec["adcode"]] = rec
            self.by_norm.setdefault(rec["norm"], []).append(rec)
        self.loaded = bool(rows)
        try:
            self.synced_at = int(self.db.meta_get("division_ts", 0) or 0)
        except Exception:
            self.synced_at = 0
        return self.loaded

    # ---------- 同步 ----------
    def build(self, gkey, log_cb=None, stop_cb=None):
        """从高德同步行政区划（市 + 区县两级），写入本地库。返回 (ok, 消息)"""
        log = log_cb if callable(log_cb) else (lambda *a, **k: None)
        stop = stop_cb or (lambda: False)
        if not gkey:
            return False, "未配置高德 Key，无法同步行政区划"
        records, ok_prov, fail = [], 0, []
        for prov in PROVINCES:
            if stop():
                return False, "已取消同步"
            try:
                url = f"{Config.GAODE_DIST}?keywords={url_enc(prov)}&subdistrict=2&extensions=base&key={gkey}"
                ok, resp = http_get(url, 3, 15)
                if not ok:
                    fail.append(prov)
                    continue
                d = json.loads(resp)
                if d.get("status") != "1":
                    fail.append(prov)
                    continue
                ok_prov += 1
                for dist in (d.get("districts") or []):
                    self._walk(dist, "", records)
            except Exception:
                fail.append(prov)
        if not records:
            return False, "同步失败：未取到任何行政区划数据（请检查 Key 与网络）"
        self.db.div_replace(records)
        self.db.meta_put("division_ts", int(time.time()))
        self.load()
        msg = f"同步完成：{ok_prov}/{len(PROVINCES)} 个省级，共 {len(records)} 条区划"
        if fail:
            msg += f"；{len(fail)} 个省级失败：{'、'.join(fail[:5])}" + ("…" if len(fail) > 5 else "")
        log(f"[区划] {msg}\n", "info")
        return True, msg

    def _walk(self, node, parent_adcode, out):
        """递归收集行政区划节点（含市/区县）"""
        try:
            name = str(node.get("name") or "").strip()
            adcode = str(node.get("adcode") or "").strip()
            level = str(node.get("level") or "").strip()
            if name and adcode:
                out.append((adcode, name, level, parent_adcode,
                            str(node.get("citycode") or ""), norm_division(name)))
            for sub in (node.get("districts") or []):
                self._walk(sub, adcode or parent_adcode, out)
        except Exception:
            pass

    # ---------- 查询 ----------
    def find(self, name, level=None, parent_hint=""):
        """按名称查找区划记录（归一化匹配）；返回 {adcode,name,level,parent} 或 None"""
        if not name or not self.by_norm:
            return None
        cands = list(self.by_norm.get(norm_division(name), []))
        if not cands:
            return None
        if level:
            same = [c for c in cands if c["level"] == level]
            if same:
                cands = same
        if len(cands) > 1 and parent_hint:
            ph = norm_division(parent_hint)
            near = [c for c in cands
                    if ph and (norm_division(self.by_adcode.get(c["parent"], {}).get("name", "")) == ph
                               or ph in norm_division(c["parent"]))]
            if near:
                cands = near
        # 同级优先返回 adcode 最短（省 < 市 < 区县 的排序习惯）
        cands.sort(key=lambda c: ({"province": 0, "city": 1, "district": 2}.get(c["level"], 3), len(c["adcode"])))
        return cands[0]

    def known(self, name, level=None, parent_hint=""):
        return self.find(name, level, parent_hint) is not None

    def canonical(self, name, level=None, parent_hint=""):
        r = self.find(name, level, parent_hint)
        return r["name"] if r else ""

    def adcode_of(self, name, level=None, parent_hint=""):
        r = self.find(name, level, parent_hint)
        return r["adcode"] if r else ""

    def canonical_district(self, prov, city, dist):
        """把区县列/地址中的区县名纠错为规范全称（如「沾益」→「沾益区」）"""
        if not dist or not self.loaded:
            return ""
        for hint in (city, prov):
            r = self.find(dist, "district", hint)
            if r:
                return r["name"]
        return ""

    def districts_of(self, city):
        """返回某市下辖的区县名列表（用于漂移校验）"""
        if not city or not self.loaded:
            return []
        c = self.find(city, "city") or self.find(city)
        if not c:
            return []
        return [rec["name"] for rec in self.by_adcode.values()
                if rec["level"] == "district" and rec["parent"] == c["adcode"]]


_DIV = None
_DIV_LOCK = threading.Lock()


def get_div_table():
    """进程内单例行政区划表"""
    global _DIV
    if _DIV is None:
        with _DIV_LOCK:
            if _DIV is None:
                _DIV = DivisionTable()
    return _DIV


# ---------- 任务状态文件（断点续跑 / 重试失败行） ----------
def task_state_path(xlsx_path):
    return str(xlsx_path) + ".addr_task.json"


def load_task_state(xlsx_path):
    """读取未完成任务状态；文件不存在/不匹配/已损坏返回 None"""
    try:
        p = task_state_path(xlsx_path)
        if not os.path.exists(p):
            return None
        with open(p, "r", encoding="utf-8") as f:
            st = json.load(f)
        if not isinstance(st, dict) or st.get("v") != 1:
            return None
        if os.path.abspath(st.get("file", "")) != os.path.abspath(str(xlsx_path)):
            return None
        if st.get("completed"):
            return None
        return st
    except Exception:
        return None


def save_task_state(xlsx_path, state):
    """原子写入任务状态（与已有状态合并，保留 start_row/last_row 等字段）"""
    try:
        p = task_state_path(xlsx_path)
        old = {}
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    old = json.load(f) or {}
            except Exception:
                old = {}
        old.update(state or {})
        old.update({"v": 1, "file": os.path.abspath(str(xlsx_path)),
                    "updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(old, f, ensure_ascii=False, indent=2)
        os.replace(tmp, p)
        return True
    except Exception:
        return False


def clear_task_state(xlsx_path):
    try:
        p = task_state_path(xlsx_path)
        if os.path.exists(p):
            os.remove(p)
        return True
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════
# 2. 工具函数模块
# ═══════════════════════════════════════════════════════════
def url_enc(s): return urllib.parse.quote(str(s), safe="")

def http_get(url, max_retry=3, timeout=15):
    ctx = ssl.create_default_context()
    ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
    last_err = ""
    for retry in range(1, max_retry + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                data = resp.read().decode("utf-8", "replace")
                # 修复：HTTP 200 但响应体为空时不应按"网络失败"重试 3 次（浪费配额与时间），
                # 直接返回成功 + 空串，交由上层 JSON 解析逻辑判定。
                return True, data
        except Exception as e:
            last_err = str(e)
            if retry == max_retry:
                return False, last_err
            time.sleep(0.3 * retry)
    return False, last_err

def json_val(js, key):
    """提取 JSON 字段值：优先 json.loads，失败则正则兜底"""
    try:
        d = json.loads(js)
        if isinstance(d, dict):
            v = d.get(key)
            return str(v) if v is not None else ""
    except: pass
    # 修复：原正则用回引号 \1 匹配引号，遇到空字符串值（"info":""）或非引号值时会失配；
    # 改为显式分别匹配"带引号值"与"裸值"两种情形。
    m = re.search(r'"' + re.escape(key) + r'"\s*:\s*"([^"]*)"', js)
    if m:
        return m.group(1)
    m = re.search(r'"' + re.escape(key) + r'"\s*:\s*([^",}\s]+)', js)
    return m.group(1) if m else ""

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
    ③ 本次升级：借助离线行政区划表，输入区县确为「该市下辖某区县」而结果区县不是 → 漂移
    当输入未提供任何区县信息时，不做②③判定，避免误伤。"""
    if not std_addr: return True
    if any(p and p not in std_addr for p in [dist, city, prov]):
        return True
    # 归一化后比较（「沾益」与「沾益区」视为同名，避免因写法不同误判漂移）
    src = {norm_division(x) for x in extract_district("".join(filter(None, [prov, city, dist, raw_addr])))}
    dst = {norm_division(x) for x in extract_district(std_addr)}
    if src and dst and not (src & dst):
        return True
    try:
        tbl = get_div_table()
        if tbl.loaded and dist and dst:
            if tbl.known(dist, "district", city) and not any(tbl.known(x, "district", city) for x in dst):
                return True
    except Exception:
        pass
    return False

def extract_core(raw_addr, prov, city, dist):
    s = raw_addr
    for p in [prov, city, dist]:
        if p: s = s.replace(p, "")
    return s.strip()


# 前缀排除「省市区县」与数字/号：否则会把「云南省曲靖市麒麟区建宁街道」整段吞掉
_RE_TOWN = re.compile(r'([^省市区县\d号\s]{1,9}?(?:街道|镇|乡|开发区|工业园区|新区))')
_RE_ROAD = re.compile(r'([^省市区县\d号\s]{1,16}?(?:大道|大街|公路|路|街|巷|弄|道))')
_RE_NO = re.compile(r'(\d{1,5}(?:[-－—~至]\d{1,5})?\s*(?:号楼|号|栋|幢|座|单元|室|层|楼))')
_RE_COMM = re.compile(r'([^省市区县\d号\s]{2,20}?(?:小区|花园|公寓|大厦|公馆|广场|中心|学校|小学|中学|大学|医院|公司|信用社|合作社|事务所|市场|超市|酒店|宾馆|银行|苑|城|府|湾|里|村|寨|屯))')

# 机构类字眼：出现在道路候选里说明这段其实是单位名，不是路名
_ORG_CHARS = "社司心厂店院校部局所站场"


def _pick_road(text, limit):
    """从文本中挑出最可能的道路名。

    规则：候选需以「路/街/道/巷/弄…」结尾；优先取结束位置最贴近门牌号的候选，
    再剔除含机构字眼（社/公司/学校…）的候选，最后在这些候选中取最长的。
    这样「曲靖市农村信用合作社联合社寥廓北路中段271-1号」能正确取到「寥廓北路」，
    而不是把整段单位名吃掉。
    """
    cands = []
    for st in range(len(text)):
        m = re.match(r'([^省市区县\d号\s]{1,12}?(?:大道|大街|公路|路|街|巷|弄|道))', text[st:])
        if m:
            nm = m.group(1).strip()
            if nm:
                cands.append((st + m.end(), nm))
    if not cands:
        return ""
    capped = [c for c in cands if c[0] <= limit + 1] or cands
    clean = [c for c in capped if not any(ch in c[1] for ch in _ORG_CHARS)]
    pool = clean or capped
    best_end = max(c[0] for c in pool)
    same = [c for c in pool if c[0] == best_end]
    return max(same, key=lambda c: len(c[1]))[1]


def parse_addr_parts(addr, level="", div=None):
    """把标准地址拆解为结构化要素（省/市/区县/街道/道路/门牌号/小区）。

    有离线行政区划表时用表做规范化（省市区县取全称），没有时退化为纯正则提取。
    返回 dict；无法识别的字段为空串，`rest` 为剥离已识别片段后的残余文本。
    """
    out = {"province": "", "city": "", "district": "", "town": "", "road": "",
           "no": "", "community": "", "rest": ""}
    if not addr:
        return out
    s = clean_addr(str(addr))
    s = re.sub(r'\[[^\]]*\]', '', s)          # 去掉 [精度:xx][来源:xx] 等标注
    s = s.strip()
    if not s:
        return out
    tbl = div if div is not None else get_div_table()
    work = s

    # 1) 省 / 市 / 区县：离线表优先（取最长匹配），失败再走正则
    try:
        if tbl and tbl.loaded:
            for rec in sorted(tbl.by_adcode.values(), key=lambda r: -len(r["name"])):
                nm = rec["name"]
                if nm and nm in work and rec["level"] in ("province", "city", "district"):
                    key = {"province": "province", "city": "city", "district": "district"}[rec["level"]]
                    if not out[key]:
                        out[key] = nm
    except Exception:
        pass
    if not out["province"]:
        m = re.match(r'^(.{2,8}?(?:省|自治区|特别行政区))', s)
        if m:
            out["province"] = m.group(1)
        else:
            m2 = re.match(r'^(北京|天津|上海|重庆)市', s)
            if m2:
                out["province"] = m2.group(0)      # 直辖市保留「市」后缀，便于后续定位区县
    if not out["city"]:
        base = s[len(out["province"]):] if out["province"] and s.startswith(out["province"]) else s
        m = re.match(r'^(.{2,10}?(?:市|自治州|地区|盟))', base)
        if m and not any(k in m.group(1) for k in ("区", "县", "镇", "乡")):
            out["city"] = m.group(1)
    if not out["district"]:
        base = s
        for cut in (out["province"], out["city"]):
            if cut and cut in base:
                base = base[base.index(cut) + len(cut):]
        m = re.match(r'^(.{1,10}?(?:区|县|旗|市))', base)
        if m and not any(k in m.group(1) for k in ("镇", "乡", "街道")):
            out["district"] = m.group(1)

    # 2) 先剥离已识别的省/市/区县，避免「浦东新区」这类区县被当成街道「新区」
    for cut in (out["province"], out["city"], out["district"]):
        if cut and cut in work:
            work = work.replace(cut, " ", 1)
    # 3) 街道/乡镇
    m = _RE_TOWN.search(work)
    if m:
        out["town"] = m.group(1).strip()
        work = work.replace(out["town"], " ", 1)
    # 4) 主门牌号（用于给道路定位）
    no_list, no_pos = [], -1
    m = _RE_NO.search(work)
    if m:
        out["no"] = re.sub(r'\s+', '', m.group(1))
        no_pos = m.start()
        no_list.append(out["no"])
        work = work[:m.start()] + " " + work[m.end():]
    # 5) 道路：取「最靠近门牌号、剔除机构词后最长」的候选
    out["road"] = _pick_road(work, no_pos if no_pos >= 0 else len(work))
    if out["road"]:
        work = work.replace(out["road"], " ", 1)
    # 6) 小区/POI（此时道路已剥离，不会把路名并进小区名）
    m = _RE_COMM.search(work)
    if m:
        out["community"] = m.group(1).strip()
        work = work.replace(out["community"], " ", 1)
    # 7) 门牌号第二/三段（如「3栋」「2单元501室」）
    for _ in range(2):
        m = _RE_NO.search(work)
        if not m:
            break
        extra = re.sub(r'\s+', '', m.group(1))
        if extra and extra not in no_list:
            no_list.append(extra)
        work = work[:m.start()] + " " + work[m.end():]
    if no_list:
        out["no"] = " ".join(no_list)
    # 3) 残余
    rest = re.sub(r'\s+', '', work)
    for cut in (out["province"], out["city"], out["district"]):
        if cut:
            rest = rest.replace(cut, "")
    out["rest"] = rest.strip(" ,，。.、-")
    if not out["community"] and out["rest"] and level and level not in ("区县", "城市", "省", "乡镇"):
        out["community"] = out["rest"]
    return out


def parse_result_cell(v):
    """从标准地址结果单元格（含 [精度:xx][来源:xx][区县码:xx] 标注）解析出 (std_addr, level)。

    用于对"已有结果"的行补写结构化要素列：跳过已有结果 / 只重试异常行 /
    断点续跑时，前序行已带结果但未写结构化列，从此单元格解析地址与精度即可
    无接口调用地补写，避免「只有表头没有数据」。
    """
    if not v:
        return "", ""
    s = str(v)
    m = re.search(r'\[精度:([^\]]*)\]', s)
    level = m.group(1).strip() if m else ""
    std = re.split(r'\s*\[精度:', s)[0].strip()
    return std, level


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


def dedup_headers(names):
    """把原始表头列表去重、补齐（空→「列N」；重名→「名字(2)」）。

    read_headers / fast_preview 共用，保证两种读取路径产出的列名完全一致，
    避免预览下拉框选中列与批量处理列名错位。
    """
    out, seen = [], {}
    for idx, raw in enumerate(names, 1):
        name = str(raw).strip() if raw is not None and str(raw).strip() else f"列{idx}"
        if name in seen:
            seen[name] += 1
            name = f"{name}({seen[name]})"
        else:
            seen[name] = 1
        out.append(name)
    return out


def read_headers(ws):
    """读取表头行并保证列名唯一、非空。

    修复：原实现直接用首个数据行/表头文本，遇到表头为空、或表头与占位列重名
    （例如模板自带「序号」列，预览再插入一列「序号」就会撞名、字典键互相覆盖）时，
    预览与列选择会错位。这里统一为：空表头 → 「列N」；重名 → 「名字(2)」。
    """
    try:
        cells = list(ws[1])
    except Exception:
        cells = []
    return dedup_headers([getattr(c, "value", None) for c in cells])


def read_headers_rw(ws):
    """按 max_column 读取表头（读写模式工作表不保证 ws[1] 完整）；列名规则同 read_headers"""
    out, seen = [], {}
    try:
        ncol = max(1, int(ws.max_column or 1))
    except Exception:
        ncol = 1
    for idx in range(1, ncol + 1):
        raw = ws.cell(row=1, column=idx).value
        name = str(raw).strip() if raw is not None and str(raw).strip() else f"列{idx}"
        if name in seen:
            seen[name] += 1
            name = f"{name}({seen[name]})"
        else:
            seen[name] = 1
        out.append(name)
    return out


def _get_self_district_info(gkey, city):
    """获取 city 自身（区县级）的 {name, adcode, center, level}；失败返回 None"""
    try:
        url = f"{Config.GAODE_DIST}?keywords={url_enc(city)}&subdistrict=0&extensions=base&key={gkey}"
        ok, resp = http_get(url, 2)
        if not ok:
            return None
        d = json.loads(resp)
        if d.get("status") != "1":
            return None
        dists = d.get("districts") or []
        if not dists:
            return None
        s = dists[0]
        if not s.get("name"):
            return None
        return {"name": s.get("name", ""), "adcode": s.get("adcode", ""),
                "center": s.get("center", ""), "level": s.get("level", "")}
    except Exception:
        return None


def get_poi_district_info(gkey, city):
    """通过高德行政区划查询获取检索单位信息列表：
    元素 {name, adcode, center, level}；返回区县级单位（level==district）。
    若查询对象本身为区县级（其下级均为街道/镇），则返回其自身作为唯一检索单位，
    避免把街道误作检索单位导致数据互相覆盖（街道名作 city 参数不可靠）。
    查询失败或无下级时回退 [{name: city}]。"""
    try:
        url = f"{Config.GAODE_DIST}?keywords={url_enc(city)}&subdistrict=1&level=city&extensions=base&key={gkey}"
        ok, resp = http_get(url, 2)
        if not ok:
            return [{"name": city}]
        d = json.loads(resp)
        if d.get("status") != "1":
            return [{"name": city}]
        dists = d.get("districts") or []
        if not dists:
            return [{"name": city}]
        subs = dists[0].get("districts") or []
        units = [{"name": s.get("name", ""), "adcode": s.get("adcode", ""),
                  "center": s.get("center", ""), "level": s.get("level", "")} for s in subs]
        units = [u for u in units if u.get("name")]
        district_units = [u for u in units if u.get("level") == "district"]
        if district_units:
            return district_units
        # 下级均为街道/镇/商务区（无区县级）：city 本身为区县级，不再下钻，
        # 取 city 自身信息作为唯一检索单位（adcode/center 供深挖与边界裁剪）
        if units:
            own = _get_self_district_info(gkey, city)
            if own:
                return [own]
        return [{"name": city}]
    except Exception:
        return [{"name": city}]


def _static_expand_kws(main_kw):
    """静态关键词扩展兜底：精确命中优先；未命中时按"包含关系"取最贴近的一项
    （修复：原实现遍历 dict 取首个命中，dict 顺序变化或存在多个包含项时
    可能误取无关词库，例如主词"超市"同时命中"社区超市"与"超市"两条）。"""
    if not main_kw:
        return []
    hit = Config.POI_KEYWORD_EXPAND.get(main_kw)
    if hit:
        return list(hit)
    best = None
    for k, v in Config.POI_KEYWORD_EXPAND.items():
        if k in main_kw or main_kw in k:
            # 取得分最优（键与主词长度差最小）的一项
            score = abs(len(k) - len(main_kw))
            if best is None or score < best[0]:
                best = (score, k, v)
    return list(best[2]) if best else []


def _merge_expand_kws(ai_kws, main_kw):
    """AI 扩展词与静态词库取并集并去重；未配置 AI 或 AI 无结果时退回静态词库"""
    static_kws = _static_expand_kws(main_kw)
    if ai_kws:
        merged = list(dict.fromkeys([ek for ek in (list(ai_kws) + static_kws) if ek and ek != main_kw]))
    else:
        merged = list(dict.fromkeys([ek for ek in static_kws if ek and ek != main_kw]))
    return merged


def get_poi_district_units(gkey, city):
    """区县名称列表（供区县下钻与关键词拆分使用）"""
    return [u["name"] for u in get_poi_district_info(gkey, city)]


def _poi_near_any(items, lng, lat, threshold=50):
    """判断坐标是否与已有结果重合（距离 < threshold 米视为重复）；无坐标不做判定"""
    if not lng or not lat:
        return False
    for it in items:
        el, ea = it.get("gcj_lng"), it.get("gcj_lat")
        if not el or not ea:
            continue
        dx = abs(lng - el) * 111000 * math.cos(math.radians(lat))
        dy = abs(lat - ea) * 111000
        if math.sqrt(dx * dx + dy * dy) < threshold:
            return True
    return False


class PoiIndex:
    """基于固定网格哈希的坐标去重索引：将 O(n) 线性扫描降为 O(1) 邻格查询"""
    CELL = 0.001          # 格距（度），约 100-111 米
    THRESHOLD = 50.0      # 去重阈值（米）

    def __init__(self):
        self._cells = {}

    def _near(self, lng, lat, el, ea):
        dx = abs(lng - el) * 111000 * math.cos(math.radians(lat))
        dy = abs(lat - ea) * 111000
        return math.sqrt(dx * dx + dy * dy) < self.THRESHOLD

    def near_any(self, lng, lat):
        if not lng or not lat:
            return False
        cx = int(lng / self.CELL)
        cy = int(lat / self.CELL)
        for i in (cx - 1, cx, cx + 1):
            for j in (cy - 1, cy, cy + 1):
                for el, ea in self._cells.get((i, j), ()):
                    if self._near(lng, lat, el, ea):
                        return True
        return False

    def add(self, item):
        lng, lat = item.get("gcj_lng", 0), item.get("gcj_lat", 0)
        if lng and lat:
            self._cells.setdefault((int(lng / self.CELL), int(lat / self.CELL)), []).append((lng, lat))


def _http_retry(url, timeout=2, retries=2):
    """带指数退避的 HTTP GET；网络失败重试 retries 次（共 retries+1 次尝试）"""
    delay = 0.5
    last = (False, "")
    for attempt in range(retries + 1):
        ok, resp = http_get(url, timeout)
        if ok:
            return True, resp
        last = (False, resp)
        if attempt < retries:
            time.sleep(delay)
            delay *= 2
    return last


def _is_quota_error(text):
    """识别配额耗尽类错误（高德 QUERY_OVER_LIMIT 系列 / 百度 401 超限）"""
    if not text:
        return False
    t = str(text)
    return ("QUERY_OVER_LIMIT" in t or "QUOTA" in t.upper() or "DAILY_QUERY" in t
            or "超限" in t or t.strip() == "401" or "INSUFFICIENT" in t.upper())


def _is_qps_error(text):
    """识别 QPS/并发超限类错误（临时限流，退避重试可恢复，区别于日配额耗尽）"""
    if not text:
        return False
    t = str(text).upper()
    return "CUQPS" in t or "QPS_HAS_EXCEEDED" in t or "CONCURRENT" in t


def http_post(url, data, headers=None, timeout=30):
    """POST JSON 请求（OpenAI 兼容接口用）；返回 (ok, text)"""
    ctx = ssl.create_default_context()
    ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
    try:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        hd = {"User-Agent": "Mozilla/5.0", "Content-Type": "application/json"}
        if headers:
            hd.update(headers)
        req = urllib.request.Request(url, data=body, headers=hd, method="POST")
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            return True, resp.read().decode("utf-8")
    except Exception as e:
        return False, str(e)


class AiClient:
    """OpenAI 兼容大模型客户端（L0 智能增强，可选）。
    未启用 / 未配置 API Key / Base URL 时 available=False，
    所有增强方法返回 None，调用方自动降级为原有规则逻辑，
    行为与未接入 AI 时完全一致。"""
    PROVIDERS = Config.AI_PROVIDERS

    def __init__(self, cfg):
        self.enabled = bool(cfg.get("ai_enabled"))
        self.provider = str(cfg.get("ai_provider") or "doubao")
        self.base_url = str(cfg.get("ai_base_url") or "").strip().rstrip("/")
        self.api_key = str(cfg.get("ai_api_key") or "").strip()
        self.model = str(cfg.get("ai_model") or "").strip()
        if self.provider in self.PROVIDERS and not self.base_url:
            self.base_url = self.PROVIDERS[self.provider]["base"]
        if not self.model and self.provider in self.PROVIDERS:
            self.model = self.PROVIDERS[self.provider]["model"]

    @property
    def available(self):
        return bool(self.enabled and self.api_key and self.base_url and self.model)

    def _chat(self, system, user, timeout=30, temperature=0.1):
        """调用 chat/completions；成功返回内容文本，失败/未配置返回 None"""
        if not self.available:
            return None
        url = f"{self.base_url}/chat/completions"
        data = {"model": self.model, "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user}], "temperature": temperature}
        ok, resp = http_post(url, data, {"Authorization": "Bearer " + self.api_key}, timeout)
        if not ok:
            return None
        try:
            j = json.loads(resp)
            return str(j["choices"][0]["message"]["content"]).strip()
        except Exception:
            return None

    @staticmethod
    def _extract_json_array(text):
        """从 LLM 回复中提取 JSON 数组（容忍 markdown 围栏与前后杂质）"""
        if not text:
            return None
        t = text.strip()
        if t.startswith("```"):
            t = re.sub(r"^```[a-zA-Z]*\s*", "", t)
            t = re.sub(r"\s*```$", "", t)
        a, b = t.find("["), t.rfind("]")
        if a < 0 or b <= a:
            return None
        try:
            arr = json.loads(t[a:b + 1])
            return arr if isinstance(arr, list) else None
        except Exception:
            return None

    def expand_keywords(self, kw, max_kws=8):
        """AI 关键词词族扩展：返回扩展词列表（不含主词）；失败/未配置返回 None"""
        if not self.available or not kw:
            return None
        sys = "你是地理POI检索专家。根据主关键词，输出该关键词在高德地图POI检索中常见的同义词、别称、细分词，用于扩大检索覆盖。只输出JSON字符串数组，不要解释。"
        user = f"主关键词：{kw}\n请给出最多{max_kws}个扩展检索词（不包括主关键词本身）。示例输出：[\"词1\",\"词2\"]"
        arr = self._extract_json_array(self._chat(sys, user))
        if not arr:
            return None
        out = []
        for w in arr:
            w = str(w).strip()
            if w and w != kw and w not in out and len(w) <= 12:
                out.append(w)
        return out[:max_kws] or None

    def clean_address(self, text):
        """AI 单条地址清洗；失败/未配置返回 None（调用方用规则兜底）"""
        if not self.available or not text:
            return None
        sys = "你是中文地址标准化助手。修正地址中的错别字、缺字、多余字符；不要增删真实地名，不要编造门牌号。只输出清洗后的地址原文，不要解释。"
        user = f"原始地址：{text}\n清洗后地址："
        out = self._chat(sys, user)
        if not out:
            return None
        out = out.strip().strip("\"'，。")
        if len(out) < 2 or len(out) > 100:
            return None
        return out

    def batch_clean_addresses(self, texts, max_items=10):
        """AI 批量地址清洗（一次调用）：返回 {原文: 清洗后}；失败返回 None"""
        if not self.available or not texts:
            return None
        texts = [str(t).strip() for t in texts if str(t).strip()]
        if not texts:
            return None
        texts = texts[:max_items]
        sys = "你是中文地址标准化助手。对地址列表逐条清洗：修正错别字、缺字、冗余字符；不要编造门牌号或真实地名；无法处理的保留原文。只输出JSON对象，键为原文、值为清洗后地址，不要解释。"
        user = "地址列表：" + json.dumps(texts, ensure_ascii=False) + "\n请输出 {\"原文\":\"清洗后\"} 形式的JSON对象。"
        t = self._chat(sys, user)
        if not t:
            return None
        t = t.strip()
        if t.startswith("```"):
            t = re.sub(r"^```[a-zA-Z]*\s*", "", t)
            t = re.sub(r"\s*```$", "", t)
        a, b = t.find("{"), t.rfind("}")
        if a < 0 or b <= a:
            return None
        try:
            obj = json.loads(t[a:b + 1])
            return obj if isinstance(obj, dict) else None
        except Exception:
            return None

    def is_same_entity(self, name_a, name_b, addr_a="", addr_b=""):
        """AI 判定两个 POI 名称是否同一实体（别名）；无法判定返回 None"""
        if not self.available or not name_a or not name_b:
            return None
        if str(name_a) == str(name_b):
            return True
        sys = "你是POI数据清洗专家。判断两个POI名称是否指代同一实体（含别名、简称、后缀差异，如“阳光花园”与“阳光花园小区”）。只输出 true 或 false，不要解释。"
        user = f"名称A：{name_a}（地址：{addr_a or '未知'}）\n名称B：{name_b}（地址：{addr_b or '未知'}）\n是否同一实体："
        out = self._chat(sys, user)
        if not out:
            return None
        t = out.strip().lower()
        if t.startswith("true") or t == "是":
            return True
        if t.startswith("false") or t == "否":
            return False
        return None

    def analyze_collection(self, region, keyword, stats, total):
        """L1 采集完整性分析：单次调用输出 完整度结论/可能漏采区县/建议补充关键词。
        失败或未配置返回 None（调用方自动跳过，行为与原有逻辑一致）。"""
        if not self.available or not keyword:
            return None
        stats_sorted = sorted(stats, key=lambda x: -int(x.get("count", 0))) if stats else []
        zero_adms = [x.get("name", "") for x in stats if int(x.get("count", 0)) == 0][:8]
        sys = "你是地理POI数据质量分析专家。根据行政区划逐区检索的结果，评估采集完整度，指出异常区县（0条或达到上限225条），并给出可扩大覆盖的补充检索关键词。只输出JSON对象，不要解释。"
        user = (
            "区域：" + str(region) + "，关键词：" + str(keyword) + "，当前合计：" + str(total) + "条。\n"
            "各区县统计（名称/条数/是否达到上限）：" + json.dumps(stats_sorted, ensure_ascii=False) +
            "\n零结果区县：" + (json.dumps(zero_adms, ensure_ascii=False) if zero_adms else "无") + "\n"
            "请输出 {\"completeness\":\"完整度结论(30字内,说明是否可能未采全)\","
            "\"risk_adms\":[\"可能漏采的区县名(仅填统计中出现过的)\"],"
            "\"suggest_keywords\":[\"建议补充检索的关键词,最多2个,与主关键词不同\"]}"
        )
        t = self._chat(sys, user)
        if not t:
            return None
        t = t.strip()
        if t.startswith("```"):
            t = re.sub(r"^```[a-zA-Z]*\s*", "", t)
            t = re.sub(r"\s*```$", "", t)
        a, b = t.find("{"), t.rfind("}")
        if a < 0 or b <= a:
            return None
        try:
            obj = json.loads(t[a:b + 1])
        except Exception:
            return None
        if not isinstance(obj, dict):
            return None
        out = {
            "completeness": str(obj.get("completeness", "")).strip()[:60],
            "risk_adms": [str(x).strip() for x in (obj.get("risk_adms") or []) if str(x).strip()][:8],
            "suggest_keywords": [str(x).strip() for x in (obj.get("suggest_keywords") or []) if str(x).strip() and str(x).strip() != str(keyword)][:2],
        }
        return out


def ai_semantic_dedup(ai, items, max_pairs=200, dist_min=50.0, dist_max=800.0, sim_min=0.55, log_cb=None):
    """AI 语义别名去重标注：对坐标相近（50-800m）且名称相似（相似度>=0.55）的候选对
    调用 LLM 确认；确认的在后项加 ai_flag 标注（不删除，安全降级）。
    未配置 AI 或无可疑候选时直接返回 0，行为与原有逻辑完全一致。"""
    if not ai or not ai.available or not items or len(items) < 2:
        return 0
    log = ((lambda m, tag="info", **k: log_cb(m, tag=tag, **k)) if log_cb
           else (lambda *a, **k: None))
    cands = []
    n = len(items)
    for i in range(n):
        a = items[i]
        if not a.get("gcj_lng") or not a.get("gcj_lat"):
            continue
        for j in range(i + 1, n):
            b = items[j]
            if b.get("ai_flag"):
                continue
            if not b.get("gcj_lng") or not b.get("gcj_lat"):
                continue
            dx = abs(a["gcj_lng"] - b["gcj_lng"]) * 111000 * math.cos(math.radians(a["gcj_lat"]))
            dy = abs(a["gcj_lat"] - b["gcj_lat"]) * 111000
            dist = math.sqrt(dx * dx + dy * dy)
            if dist < dist_min or dist > dist_max:
                continue
            na, nb = str(a.get("name", "")), str(b.get("name", ""))
            if not na or not nb:
                continue
            if difflib.SequenceMatcher(None, na, nb).ratio() < sim_min:
                continue
            cands.append((i, j, a, b))
            if len(cands) >= max_pairs:
                break
        if len(cands) >= max_pairs:
            break
    # v2.8.49：候选对判定并发化（默认 4 路），显著降低 AI 调用串行耗时；结果按序标注保持一致
    marked = 0
    if cands:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        # 修复：原实现以 id(fmap[fut]) 作 verdicts 键，语义脆弱（依赖 CPython 对象 id 语义）。
        # 改为直接用 Future 对象本身作为键，语义清晰且不依赖实现细节。
        verdicts = {}
        with ThreadPoolExecutor(max_workers=min(4, len(cands))) as ex:
            fmap = {}
            for i, j, a, b in cands:
                fmap[ex.submit(ai.is_same_entity, a.get("name", ""), b.get("name", ""),
                               a.get("address", ""), b.get("address", ""))] = (i, j, a, b)
            for fut in as_completed(fmap):
                try:
                    r = fut.result()
                except Exception:
                    r = None
                verdicts[fut] = (fmap[fut], r)
        for (i, j, a, b), same in verdicts.values():
            if same is True:
                b["ai_flag"] = "疑似别名重复: " + str(a.get("name", ""))
                marked += 1
    if marked:
        log(f"[AI增强] 语义去重：识别 {marked} 条疑似别名重复（已标注，未删除）", "info")
    return marked


def ai_l1_postprocess(ai, region, keyword, stats, total, merged, poi_idx, gkey,
                      log_cb=None, stop_cb=None):
    """L1 采集智能决策（后台线程调用，不阻塞界面）：
    1. 采集完成后单次调用 AI 分析完整度 / 异常区县 / 建议补充关键词（仅日志展示）；
    2. 存在零结果区县且 AI 给出补充关键词时，执行严格受控的补扫：
       补充词 <=2 个、区县 <=3 个、每词每区县仅 1 页(25条)、高德总调用 <=6 次；
       补扫结果经 PoiIndex 空间去重后并入 merged。
    未配置 AI / 分析失败 / 无可补扫场景时直接返回 []，行为与原有逻辑完全一致。"""
    if not ai or not ai.available or not keyword:
        return []
    log = ((lambda m, tag="info", **k: log_cb(m, tag=tag, **k)) if log_cb
           else (lambda *a, **k: None))
    stop = stop_cb or (lambda: False)
    try:
        analysis = ai.analyze_collection(region, keyword, stats, total)
    except Exception:
        analysis = None
    if analysis:
        log(f"[AI采集分析] {analysis.get('completeness') or '（无结论）'}", "info")
        if analysis.get("risk_adms"):
            log("[AI采集分析] 可能未采全: " + "、".join(analysis["risk_adms"]), "info")
        if analysis.get("suggest_keywords"):
            log("[AI采集分析] 建议补充关键词: " + "、".join(analysis["suggest_keywords"]) + "（将自动补扫零结果区县）", "info")
    extra_kws = analysis.get("suggest_keywords") if analysis else None
    if not extra_kws:
        return []
    zero_adms = [x.get("name", "") for x in stats if int(x.get("count", 0)) == 0 and x.get("name")]
    if not zero_adms or len(zero_adms) > 5:
        log("[AI采集分析] 零结果区县数量不满足补扫条件，跳过补扫", "info")
        return []
    added = []
    used_calls = 0
    for adm in zero_adms[:3]:
        if stop():
            break
        for ek in extra_kws[:2]:
            if stop():
                break
            if used_calls >= 6:
                break
            try:
                url = f"{Config.GAODE_POI_TEXT}?key={gkey}&keywords={url_enc(ek)}&city={url_enc(adm)}&offset=25&page=1&extensions=all&output=json"
                ok, resp = _http_retry(url)
                if not ok:
                    log(f"[AI补扫] {adm}+{ek} 请求失败: {resp[:60]}", "error")
                    continue
                used_calls += 1
                d = json.loads(resp)
                if d.get("status") != "1":
                    continue
                pois = d.get("pois", [])
                if not pois:
                    continue
                for p in pois:
                    loc = p.get("location", "")
                    lng = lat = 0.0
                    if loc and "," in loc:
                        try:
                            lng, lat = map(float, loc.split(","))
                        except Exception:
                            pass
                    if not lng or not lat:
                        continue
                    if poi_idx.near_any(lng, lat):
                        continue
                    item = {"name": p.get("name", ""), "address": p.get("address", ""),
                            "tel": p.get("tel", ""), "type": p.get("type", ""),
                            "gcj_lng": lng, "gcj_lat": lat, "source": "高德(AI补扫)"}
                    poi_idx.add(item)
                    merged.append(item)
                    added.append(item)
                log(f"[AI补扫] {adm}+{ek}: 新增{len(added)}条(累计)（第{used_calls}次调用）", "info")
            except Exception as e:
                log(f"[AI补扫] {adm}+{ek} 异常: {e}", "error")
    if added:
        log(f"[AI补扫] 完成：补充检索 {used_calls} 次，新增 {len(added)} 条", "info")
    return added


_POLY_CACHE = {}


def _parse_polyline(pl):
    """解析高德行政区 polyline 为多边形列表。
    高德边界格式：多个子多边形用 '|' 分隔，多边形内点用 ';' 分隔，点内 'lng,lat'。
    返回 [[(lng, lat), ...], ...]（每个子多边形一个点列表）；无数据返回 []"""
    if not pl:
        return []
    polys = []
    for seg in pl.split("|"):
        pts = []
        for pair in seg.split(";"):
            if "," in pair:
                try:
                    lng, lat = map(float, pair.split(","))
                    pts.append((lng, lat))
                except Exception:
                    pass
        if len(pts) >= 3:
            polys.append(pts)
    return polys


def get_district_polygon(gkey, adcode):
    """获取区县边界多边形（extensions=all）。

    三级缓存：进程内 _POLY_CACHE → 本地 SQLite（跨会话，避免重复消耗区划接口）→ 网络。
    失败返回 None；"查到但无边界"会做负缓存，避免反复请求。
    """
    if not adcode:
        return None
    if adcode in _POLY_CACHE:
        return _POLY_CACHE[adcode]
    db = get_db()
    if db.available:
        found, polys = db.poly_get(adcode)
        if found:
            _POLY_CACHE[adcode] = polys or None
            return _POLY_CACHE[adcode]
    try:
        url = f"{Config.GAODE_DIST}?keywords={url_enc(adcode)}&subdistrict=0&extensions=all&key={gkey}"
        ok, resp = _http_retry(url, 2, 2)
        if not ok:
            return None
        d = json.loads(resp)
        if d.get("status") != "1":
            return None
        dists = d.get("districts") or []
        if not dists:
            return None
        poly = _parse_polyline(dists[0].get("polyline", ""))
        _POLY_CACHE[adcode] = poly or None
        if db.available:
            db.poly_put(adcode, poly or [])
        return _POLY_CACHE[adcode]
    except Exception:
        return None


def _point_in_polygon(lng, lat, poly):
    """射线法判断点是否在（任一）多边形内；空多边形视为全部通过"""
    if not poly:
        return True
    for ring in poly:
        inside = False
        n = len(ring)
        j = n - 1
        for i in range(n):
            xi, yi = ring[i]
            xj, yj = ring[j]
            if ((yi > lat) != (yj > lat)) and (lng < (xj - xi) * (lat - yi) / (yj - yi + 1e-12) + xi):
                inside = not inside
            j = i
        if inside:
            return True
    return False


def _calc_grid_radius(poly):
    """按区县边界外包矩形对角线估算网格半径（覆盖全区县），无边界返回 None"""
    if not poly:
        return None
    pts = [pt for ring in poly for pt in ring]
    if not pts:
        return None
    lngs = [p[0] for p in pts]
    lats = [p[1] for p in pts]
    mid = (max(lats) + min(lats)) / 2.0
    w = (max(lngs) - min(lngs)) * 111000 * math.cos(math.radians(mid))
    h = (max(lats) - min(lats)) * 111000
    r = math.sqrt((w / 2.0) ** 2 + (h / 2.0) ** 2)
    return max(5000.0, min(r, 50000.0))


class _QuotaStop(Exception):
    """内部信号：配额耗尽，整体停止网格扫描"""
    pass


def grid_collect_poi(gkey, clng, clat, kw, radius, stop_check=None, log_cb=None, poly=None, known=None, workers=None):
    """网格化周边扫描（含中心细分）：以区县中心为原点，8 方位 + 中心共 9 个采样点，
    每点做周边搜索（单点翻页上限 9 页/225 条）；若中心点仍被截断（高密度区域），
    自动将半径减半拆 4 子点继续补扫，进一步突破单点 225 条上限。
    支持 poly 边界过滤（剔除区县外的 POI）与网络失败自动重试。
    known：已采集 POI 列表（含 gcj_lng/gcj_lat）——某点连续 3 页无新增（数据全重复）
    即提前停止该点，大幅减少重复区域的无效 API 调用。
    v2.8.45：9 个网格点并发扫描、中心细分 4 子点并发（ThreadPoolExecutor），
    多路并行大幅缩短深挖耗时；got_idx 判重加锁、配额错误经 threading.Event 全局停止。
    返回 (poi列表, 调用次数)"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    stop = stop_check if callable(stop_check) else (lambda: False)
    log = log_cb if callable(log_cb) else (lambda *a, **k: None)
    # 修复：radius 可能是空串/非数字（界面自由输入或配置异常），原实现直接 float() 会抛异常
    try:
        d = float(radius)
    except (TypeError, ValueError):
        d = 0.0
    if not d or d <= 0:
        d = 12000.0
    d = min(d, 50000.0)
    # 扫描半径 = 网格间距（覆盖面积保证）：0.72 缩放虽省调用，但对 bbox 大/长宽比大的
    # 区县（如景洪市 d=50000 封顶）覆盖面积不足，实测深挖数据 -66%（景洪 502->216）。
    # 调用优化改由 known 增量停止（已实测无数据损失）承担。
    r_scan = max(2000.0, d)
    known_idx = PoiIndex()
    if known:
        for it in known:
            el = it.get("gcj_lng") or it.get("lng") or 0
            ea = it.get("gcj_lat") or it.get("lat") or 0
            if el and ea:
                known_idx.add({"gcj_lng": el, "gcj_lat": ea})
    got_idx = PoiIndex()
    got_lock = threading.Lock()
    quota_event = threading.Event()
    collected = []
    calls = [0]

    def scan_point(c_lng, c_lat, r):
        """单点周边 9 页扫描（并发安全）；返回 (本点收集, 本点调用数, truncated)"""
        truncated = False
        empty_new = 0
        page = 1
        local = []
        lcalls = 0
        qps_try = 0
        while not stop() and not quota_event.is_set() and page <= 9:
            url = f"{Config.GAODE_POI_AROUND}?key={gkey}&location={c_lng},{c_lat}&keywords={url_enc(kw)}&radius={int(r)}&offset=25&page={page}&extensions=all&output=json"
            ok, resp = _http_retry(url, 2, 1)
            if not ok:
                break
            lcalls += 1
            try:
                j = json.loads(resp)
                if j.get("status") != "1":
                    if _is_quota_error(j.get("info", "")):
                        quota_event.set()
                        break
                    if _is_qps_error(j.get("info", "")):
                        # 临时并发限流：退避后重试同一页（最多 3 次，退避 1/2/4s 增强韧性）
                        if qps_try >= 3:
                            break
                        qps_try += 1
                        time.sleep([0, 1.0, 2.0, 4.0][qps_try])
                        continue
                    break
                pois = j.get("pois") or []
                if not pois:
                    break
                new_found = 0
                for p in pois:
                    loc = p.get("location", "")
                    lng, lat = 0, 0
                    if loc and "," in loc:
                        try: lng, lat = map(float, loc.split(","))
                        except: pass
                    if poly is not None and (not lng or not lat or not _point_in_polygon(lng, lat, poly)):
                        continue
                    with got_lock:
                        if known_idx.near_any(lng, lat):
                            skip = "known"
                        elif got_idx.near_any(lng, lat):
                            skip = "got"
                        else:
                            got_idx.add({"gcj_lng": lng, "gcj_lat": lat})
                            skip = None
                    if skip == "known":
                        # 与外部已采集数据重复：该区域已覆盖，计入"无新增"累计
                        continue
                    if skip == "got":
                        # 与本轮兄弟点重复（并发抢先）：数据仍有价值，不计入"无新增"
                        new_found += 1
                        continue
                    local.append({
                        "name": p.get("name", ""), "address": p.get("address", ""),
                        "tel": p.get("tel", ""), "type": p.get("type", ""),
                        "gcj_lng": lng, "gcj_lat": lat, "source": "高德(深挖)"
                    })
                    new_found += 1
                if page >= 9:
                    truncated = True
                if new_found == 0:
                    empty_new += 1
                    if empty_new >= 3:
                        log("[深挖] 点(%s,%s)连续3页无新增，判定该区域已覆盖，提前结束" % (round(c_lng, 4), round(c_lat, 4)))
                        break
                else:
                    empty_new = 0
                if len(pois) < 25:
                    break
            except _QuotaStop:
                raise
            except Exception:
                break
            page += 1
            time.sleep(0.2 + random.random() * 0.3)   # 随机抖动错峰，降低并发扎堆触发 CUQPS
        return local, lcalls, truncated

    def run_points(pts, r):
        """并发执行一组扫描点；pts: {key: (lng, lat)}；返回 {key: (local, lcalls, truncated)}
        workers 由调用方经 grid_collect_poi 传入（默认 3，配置项 grid_workers 可调 1-5）；
        高德并发限流由页内 QPS 退避重试 + 随机抖动错峰兜底。"""
        results = {}
        if not pts or quota_event.is_set():
            return results
        w = workers if workers and workers > 0 else 2
        w = min(len(pts), w)
        with ThreadPoolExecutor(max_workers=w) as ex:
            fut_map = {ex.submit(scan_point, lng, lat, r): key for key, (lng, lat) in pts.items()}
            for fut in as_completed(fut_map):
                key = fut_map[fut]
                try:
                    results[key] = fut.result()
                except _QuotaStop:
                    quota_event.set()
                    results[key] = None
        return results

    try:
        dlat = d / 111000.0
        dlng = d / (111000.0 * math.cos(math.radians(clat))) if clat else 0.0
        # v2.8.48 调度修复：中心点先串行扫描（保证中心区数据归属中心点，不与方位点竞争），
        # 中心点截断则 4 子点并发补扫；最后 8 方位点并发扫描环形区。
        # 修复旧调度"9点并发 + 仅中心细分"的三方互斥丢失：方位点抢注中心区数据后，
        # 自身受 225 单点上限截断且不细分，中心/子点又因 got 判重跳过 → 三方都不收集而丢数据。
        # 同一高德数据下景洪市深挖实测 338→380 条，恢复 v2.8.44 召回级别，并发提速保留。
        center_local, center_calls, center_trunc = scan_point(clng, clat, r_scan)
        calls[0] += center_calls
        collected.extend(center_local)
        if center_trunc and r_scan / 2.0 >= 2000 and not quota_event.is_set():
            sub = r_scan / 2.0
            sdlat = sub / 111000.0
            sdlng = sub / (111000.0 * math.cos(math.radians(clat))) if clat else 0.0
            log("[深挖] 中心点已达单参数上限，细分 4 子点补扫（半径 %dm）..." % int(sub))
            sub_pts = {}
            for sx, sy in [(0.5, 0.5), (-0.5, 0.5), (0.5, -0.5), (-0.5, -0.5)]:
                sub_pts[(sx, sy)] = (clng + sx * sdlng * 2, clat + sy * sdlat * 2)
            res_sub = run_points(sub_pts, sub)
            for val2 in res_sub.values():
                if val2 is None:
                    continue
                calls[0] += val2[1]
                collected.extend(val2[0])
        offsets = [(1, 0), (-1, 0), (0, 1), (0, -1), (1, 1), (1, -1), (-1, 1), (-1, -1)]
        pts = {}
        for dx, dy in offsets:
            pts[(dx, dy)] = (clng + dx * dlng, clat + dy * dlat)
        res = run_points(pts, r_scan)
        for key, val in res.items():
            if val is None:
                continue
            local, lcalls, truncated = val
            calls[0] += lcalls
            collected.extend(local)
    except _QuotaStop:
        pass
    return collected, calls[0]

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


def _gen_kml(proc, srow, lrow, addr_idx, out_path, direction="fwd"):
    """从Excel生成KML文件，使用WGS-84坐标。

    direction="rev" 时地址名取「完整地址」列（第 12 列）：
    修复：原实现固定读「标准地址」列（第 11 列）作为名称与有效性判断，
    而逆向编码只写第 12~14 列，第 11 列为空 → 逆向导出 KML 恒为 0 个点。
    """
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
    name_col = Config.COLS["rego_full"] if direction == "rev" else Config.COLS["std_addr"]
    for r in range(srow, lrow + 1):
        std = str(proc.ws.cell(row=r, column=name_col).value or "").strip()
        if not std or any(k in std for k in ("失败", "错误", "无匹配", "无有效", "无效")):
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

    def _root_alive():
        """修复：winfo_exists() 必须只从 Tk 主线程调用（后台线程直接调用会触发
        'main thread is not in main loop' 或跨线程 Tcl 调用异常）。
        后台线程改为仅判断 root 是否存在，并在异常时保守视为存活，统一走 after() 调度。"""
        if not root:
            return False
        try:
            return True
        except Exception:
            return False

    def _do_download():
        """在后台线程执行下载"""
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
            req = urllib.request.Request(download_url, headers={"User-Agent": "Mozilla/5.0"})

            with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
                total = int(resp.headers.get('Content-Length', 0))
                downloaded = 0
                last_pct = -1
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
                        # 仅在整百进度变化时更新状态栏，避免刷屏与过度调度
                        if total > 0 and status_cb:
                            pct = min(100, int(downloaded * 100 / total))
                            if pct == last_pct:
                                continue
                            last_pct = pct
                            msg = f"正在下载... {pct}% ({downloaded//1024//1024}MB / {total//1024//1024}MB)"
                            if _root_alive():
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
            if _root_alive():
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
        if _root_alive():
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
class RateLimiter:
    """全局最小间隔限流器（多线程共享）。

    串行时代码用 time.sleep(interval) 控速，一旦并发就会同时打满 QPS 触发平台限流。
    这里用「下一次允许发请求的时刻」做统一节流，并对限流错误自适应减速、成功后逐步恢复，
    保证批量提速的同时不把配额/并发打到平台上限。
    """

    def __init__(self, min_interval=0.2):
        try:
            self.min_interval = max(0.0, float(min_interval))
        except Exception:
            self.min_interval = 0.2
        self.cur = self.min_interval
        self._lock = threading.Lock()
        self._next = 0.0
        self.slow_count = 0

    def wait(self):
        with self._lock:
            now = time.time()
            delay = self._next - now
            self._next = max(now, self._next) + self.cur
        if delay > 0:
            time.sleep(min(delay, 5.0))

    def slow_down(self, factor=1.6, cap=3.0):
        with self._lock:
            self.cur = min(cap, max(self.cur, self.min_interval) * factor)
            self.slow_count += 1

    def speed_up(self, factor=0.92):
        with self._lock:
            if self.cur > self.min_interval:
                self.cur = max(self.min_interval, self.cur * factor)


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
        self.gd_quota_dead = False   # v2.8.56：高德日配额耗尽后自动切换百度（优先高德）
        # 本次升级：本地落库缓存 / 离线区划表 / 并发保护
        self.db = get_db()
        self.div = get_div_table()
        self.cache_hits = 0
        self._lock = threading.Lock()
        self.limiter = RateLimiter(self.interval)   # 并发批量的统一节流器

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
        self.limiter.wait()
        ok, resp = http_get(url, self.retry)
        if not ok: return None, resp
        with self._lock:
            self.gcnt += 1
        if json_val(resp, "status") != "1":
            c = json_val(resp, "infocode")
            if _is_qps_error(json_val(resp, "info")) or _is_quota_error(c):
                self.limiter.slow_down()      # 触发平台限流 → 自动降速
            else:
                self.limiter.speed_up()
            return None, f"高德错误[{c}]: {gaode_err_msg(c, json_val(resp, 'info'))}"
        self.limiter.speed_up()
        return {"location": geo_field(resp, "location"), "std_addr": geo_field(resp, "formatted_address"),
                "level": geo_field(resp, "level"), "source": "高德"}, ""

    def gaode_regeo(self, lng, lat):
        loc = f"{lng},{lat}"
        if self.cfg.get("enable_cache") and loc in self.rcache: return self.rcache[loc], ""
        url = f"{Config.GAODE_REGO}?location={loc}&output=json&extensions=all&key={self.gkey}"
        self.limiter.wait()
        ok, resp = http_get(url, self.retry)
        if not ok: return None, resp
        with self._lock:
            self.gcnt += 1
        if json_val(resp, "status") != "1":
            c = json_val(resp, "infocode")
            if _is_qps_error(json_val(resp, "info")) or _is_quota_error(c):
                self.limiter.slow_down()
            else:
                self.limiter.speed_up()
            return None, f"高德逆地理错误[{c}]: {gaode_err_msg(c, json_val(resp, 'info'))}"
        self.limiter.speed_up()
        r = {"full": regeo_field(resp, "formatted_address"),
             "division": f"{regeo_comp(resp, 'province')}{regeo_comp(resp, 'city')}{regeo_comp(resp, 'district')}",
             "street": f"{regeo_comp(resp, 'street')}{regeo_comp(resp, 'number')}",
             "province": regeo_comp(resp, "province"), "city": regeo_comp(resp, "city"),
             "district": regeo_comp(resp, "district"), "township": regeo_comp(resp, "township")}
        if self.cfg.get("enable_cache"): self.rcache[loc] = r
        return r, ""

    def baidu_regeo(self, lng, lat):
        # v2.8.56：百度逆地理编码（GCJ-02 转 BD-09 后查询），返回与 gaode_regeo 相同结构
        loc = f"{lng},{lat}"
        if self.cfg.get("enable_cache") and loc in self.rcache: return self.rcache[loc], ""
        blng, blat = gcj02_to_bd09(lng, lat)
        url = f"{Config.BAIDU_REGO}?ak={self.bak}&output=json&coordtype=bd09ll&location={blat},{blng}"
        self.limiter.wait()
        ok, resp = http_get(url, self.retry)
        if not ok: return None, f"百度逆地理请求失败: {resp[:80]}"
        with self._lock:
            self.bcnt += 1
        try: d = json.loads(resp)
        except: return None, "百度逆地理JSON解析失败"
        if d.get("status") not in (0, "0"):
            return None, f"百度逆地理错误[{d.get('status')}]: {d.get('message', '未知错误')}"
        result = d.get("result", {}) or {}
        r = {"full": result.get("formatted_address", ""),
             "division": f"{result.get('province', '')}{result.get('city', '')}{result.get('district', '')}",
             "street": f"{result.get('street', '')}{result.get('street_number', '')}",
             "province": result.get("province", ""), "city": result.get("city", ""),
             "district": result.get("district", ""), "township": ""}
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
        self.limiter.wait()
        ok, resp = http_get(url, 2)
        if not ok: return None, f"百度请求失败: {resp[:80]}"
        with self._lock:
            self.bcnt += 1
        try: d = json.loads(resp)
        except: return None, f"百度JSON解析失败: {resp[:200]}"
        status = d.get("status")
        if status not in (0, "0"):
            if _is_qps_error(str(status)) or _is_quota_error(str(status)):
                self.limiter.slow_down()
            else:
                self.limiter.speed_up()
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
        self.limiter.speed_up()
        r = d.get("result", {})
        if not isinstance(r, dict): return None, "百度返回无result字段"
        loc = r.get("location", {})
        try: blng, blat = float(loc.get("lng", 0)), float(loc.get("lat", 0))
        except: return None, "百度坐标格式错误"
        if blng == 0 or blat == 0: return None, "百度返回坐标为0"
        glng, glat = bd09_to_gcj02(blng, blat)
        level = r.get("level", "未知")
        ck = (round(blng, 6), round(blat, 6))
        # v2.8.56：修复 std 未初始化（百度响应含 formatted_address 时原分支未赋值导致 UnboundLocalError）
        std = str(baidu_fmt(resp) or "")
        if not std and ck in self.ccache: std = self.ccache[ck]
        elif not std:
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
        # 本次升级：区县名纠错（「沾益」→「沾益区」），提升查询命中率
        dist_fix = dist_str
        try:
            if dist_str and self.div.loaded:
                canon = self.div.canonical_district(prov_str, city_str, dist_str)
                if canon and canon != dist_str:
                    dist_fix = canon
                    blog += f"[区划纠错] 区县「{dist_str}」→「{canon}」\n"
        except Exception:
            pass
        query = build_addr(prov_str, city_str, dist_fix, addr_str)
        cq = clean_addr(query)
        ckey = f"{api_city}|{addr_str}".lower().replace(" ", "").replace("，", "").replace("。", "")
        # 缓存：内存（本次会话）→ 本地 SQLite（跨会话，直接省配额）
        if self.cfg.get("enable_cache"):
            c = self.cache.get(ckey)
            if c is None and self.db.available:
                c = self.db.cache_get(ckey)
                if c: self.cache[ckey] = c
            if c:
                with self._lock:
                    self.cache_hits += 1
                if self.db.available:
                    self.db.stat_bump("cache_hit")   # 内存命中同样计入命中率统计
                # 返回副本：调用方（write_forward 等）会就地修改结果，避免污染缓存
                return dict(c), c.get("is_abnormal", False), "", False, "[缓存] 命中本地结果缓存，未消耗 API 配额\n"
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
            if self.gd_quota_dead: break   # v2.8.56：高德日配额已耗尽，跳过剩余层级直接走百度
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
                    if fc in err:
                        if fc == "10003" and self.use_baidu and self.bak and self.bok:
                            # v2.8.56：高德日配额耗尽，自动切换百度解析（优先高德，仅限额时启用）
                            self.gd_quota_dead = True
                            blog += "[百度Fallback] 高德日配额已耗尽，自动切换百度解析（优先高德）\n"
                            break
                        return None, True, f"高德全局错误：{err}", True, blog
                if not best_err: best_err = err
            # 速度控制交给 self.limiter（gaode_geo/baidu_geo 内部统一节流），
            # 串行模式下等价于原来的 time.sleep(interval)，并发模式下才能真的并行
        # v2.8.56：高德日配额耗尽 → 直接使用百度解析结果（优先高德，仅限额时切换）
        if self.gd_quota_dead and self.use_baidu and self.bak and self.bok:
            bd, be = self.baidu_geo(query, gcity)
            if bd:
                p = bd["location"].split(",")
                try: gcj_lng, gcj_lat = float(p[0]), float(p[1])
                except: gcj_lng = gcj_lat = 0.0
                std_addr, level, source = bd["std_addr"], bd["level"], "百度"
                is_drift = is_abnormal = False
                blog += f"[百度Fallback] 百度解析成功: {std_addr} [精度:{level}]\n"
            else:
                blog += f"[百度Fallback] 百度解析失败: {be}\n"
                err_msg, is_abnormal = be or "双平台均无法解析", True
        # 使用最优结果
        if best_res and not self.gd_quota_dead:
            location = best_res["location"]
            std_addr = best_res["std_addr"]
            level = best_res["level"]
            source = best_res["source"]
            try:
                p = location.split(",")
                gcj_lng, gcj_lat = float(p[0]), float(p[1])
            except: gcj_lng = gcj_lat = 0.0
        else:
            if not (self.gd_quota_dead and self.use_baidu and self.bak and self.bok):
                err_msg, is_abnormal = best_err or "高德无匹配结果", True
        if self.cfg.get("check_drift") and std_addr:
            is_drift = check_drift(std_addr, prov_str, city_str, dist_str, addr_str)
            if is_drift: is_abnormal = True
        # 百度Fallback：精度不足或漂移时自动切换（配额耗尽已切百度时不再重复）
        if self.use_baidu and location and self.bak and self.bok and not self.gd_quota_dead:
            gs = level_score(level); ts = level_score(self.trigger)
            should_fallback = is_drift or (gs <= ts)   # 去掉 has_d 限制
            if is_drift:
                blog += f"[百度Fallback] 地址疑似漂移，尝试百度...\n"
            elif gs <= ts:
                blog += f"[百度Fallback] 高德精度不足({level})，尝试百度...\n"
            else:
                blog += f"[百度Fallback] 高德精度满足要求({level})，未触发百度Fallback\n"
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
            # 本次升级：区县边界几何校验——坐标落在结果地址所属区县边界外即判漂移
            # （文本比对无法发现"地址写 A 区、坐标落在 B 区"这类错误）
            adcode = ""
            try:
                if self.cfg.get("geo_drift_check", True) and std_addr and level_score(level) >= level_score("道路"):
                    adcode, nm = self._district_adcode(std_addr)
                    if adcode:
                        poly = get_district_polygon(self.gkey, adcode)
                        if poly and not _point_in_polygon(gcj_lng, gcj_lat, poly):
                            is_drift = is_abnormal = True
                            blog += f"[几何校验] 坐标不在「{nm}」边界内（{gcj_lng:.6f},{gcj_lat:.6f}），判定漂移\n"
                        elif poly:
                            blog += f"[几何校验] 坐标位于「{nm}」边界内，通过\n"
                elif std_addr:
                    adcode, _nm = self._district_adcode(std_addr)   # 仅取 adcode 供结果/导出使用
            except Exception:
                pass
            rec = {"gcj_lng": gcj_lng, "gcj_lat": gcj_lat, "std_addr": std_addr, "level": level,
                   "source": source, "is_drift": is_drift, "is_abnormal": is_abnormal,
                   "adcode": adcode}
            if self.cfg.get("enable_cache") and not is_abnormal and level_score(level) > level_score("区县"):
                self.cache[ckey] = rec
                if self.db.available:
                    self.db.cache_put(ckey, rec)
            return dict(rec), is_abnormal, err_msg or ("解析异常" if is_abnormal else ""), False, blog
        return None, True, err_msg or "双平台均无法解析", False, blog

    def _district_adcode(self, std_addr):
        """从标准地址中提取区县名并查离线表得到 adcode；返回 (adcode, 名称)"""
        try:
            if not self.div.loaded:
                return "", ""
            names = sorted(extract_district(std_addr), key=len, reverse=True)
            for nm in names:
                ad = self.div.adcode_of(nm, "district")
                if ad:
                    return ad, nm
        except Exception:
            pass
        return "", ""


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
        # 本次升级：结构化要素输出列（按需在末尾追加，不影响 7~14 列既有布局）
        self.write_adcode = False
        self.struct_start = 0

    STRUCT_COLS = [("省(标准化)", "province"), ("市(标准化)", "city"), ("区县(标准化)", "district"),
                   ("街道/乡镇", "town"), ("道路", "road"), ("门牌号", "no"), ("小区/POI", "community")]

    def prepare_struct(self):
        """在表尾追加结构化输出列并写表头；返回起始列号（幂等）。

        使用「现有最后一列之后」作为起点，避免覆盖用户自己的数据；模板文件从第 15 列开始。
        若表中已存在本工具写入的表头块（二次运行），则复用原起始列，避免重复追加。
        """
        if self.struct_start:
            return self.struct_start
        label0 = self.STRUCT_COLS[0][0]
        try:
            for c in range(Config.COLS["rego_street"] + 1, int(self.ws.max_column or 0) + 1):
                if self.ws.cell(row=1, column=c).value == label0:
                    self.struct_start = c
                    return c
        except Exception:
            pass
        try:
            start = max(Config.COLS["rego_street"] + 1, int(self.ws.max_column or 0) + 1)
        except Exception:
            start = Config.COLS["rego_street"] + 1
        for i, (label, _k) in enumerate(self.STRUCT_COLS):
            self.ws.cell(row=1, column=start + i, value=label)
        self.struct_start = start
        return start

    def write_structured(self, row, parts):
        """写入结构化地址要素（未启用时忽略）"""
        if not self.struct_start:
            return
        for i, (_label, key) in enumerate(self.STRUCT_COLS):
            self.ws.cell(row=row, column=self.struct_start + i, value=(parts or {}).get(key, ""))

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
        return self._is_abnormal(row, "fwd")

    ABNORMAL_KWS = ("失败", "错误", "无匹配", "空地址", "无效", "无有效", "超出范围")

    def _is_abnormal(self, row, direction="fwd"):
        """异常判定：正向看「标准地址」列，逆向看「完整地址」列。
        修复：原实现固定读标准地址列（第 11 列）判断异常，而逆向编码的失败信息
        写在第 12 列（完整地址），导致逆向流程「导出异常行」永远提示「没有异常行」。"""
        col = Config.COLS["rego_full"] if direction == "rev" else Config.COLS["std_addr"]
        if direction != "rev" and self.is_normal(row):
            return False
        addr = str(self.ws.cell(row=row, column=col).value or "")
        return bool(addr) and any(k in addr for k in self.ABNORMAL_KWS)

    def last_data_row(self, direction="fwd", addr_col=None):
        """按方向取数据末行（供导出异常行 / 导出 KML 划定范围）。
        修复：原实现固定用第 5 列（模板地址1列）判断末行，而工具支持用户自选地址列，
        地址放在其它列时末行会被算成 1，导出范围直接为空。"""
        if direction == "rev":
            cols = [Config.COLS["gcj_lng"], Config.COLS["gcj_lat"], Config.COLS["wgs_lng"],
                    Config.COLS["wgs_lat"], Config.COLS["rego_full"]]
        else:
            cols = [addr_col or 0, Config.COLS["addr1"], Config.COLS["gcj_lng"], Config.COLS["std_addr"]]
        rows = [self.last_row(c) for c in cols if c and c > 0]
        return max(rows) if rows else 1

    def get_abnormal(self, start_row=2, direction="fwd", addr_col=None):
        abnormal, headers = [], read_headers_rw(self.ws)
        for r in range(start_row, self.last_data_row(direction, addr_col) + 1):
            if self._is_abnormal(r, direction):
                abnormal.append({headers[c - 1]: self.ws.cell(row=r, column=c).value
                                 for c in range(1, len(headers) + 1)})
                abnormal[-1]["_原始行号"] = r
        return abnormal, headers

    def export_abnormal(self, out_path, start_row=2, direction="fwd", addr_col=None):
        abnormal, headers = self.get_abnormal(start_row, direction, addr_col)
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
        if self.write_adcode and result.get("adcode"):
            final += f"[区县码:{result['adcode']}]"
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

    @staticmethod
    def _cell_str(ws, row, col):
        """安全读取单元格文本；col<=0 表示「该列未选择/不存在」，返回空串。

        修复：原实现按默认列号读取（如地址2恒取第 6 列），用户未选择可选的地址2列时
        会把第 6 列的无关数据当作地址2参与合并/查询；列号为 0 时 openpyxl 还会直接抛错。
        """
        if not col or col <= 0:
            return ""
        try:
            v = ws.cell(row=row, column=col).value
        except Exception:
            return ""
        return str(v).strip() if v is not None else ""

    def init(self, cfg):
        self.geo = GeoCoder(cfg)
        self.stop_flag = False
        ok, msg = self.geo.check()
        if not ok or msg:
            return False, msg or "密钥验证失败"
        return True, ""

    # ---------- 单行解析（并发安全：只做字符串/网络计算，不读写工作表） ----------
    def _resolve_row(self, a1, a2, city, dist):
        """解析单行（含地址2备用策略），返回 dict：res/ab/err/fatal/blog/msgs/a2u/drift/bdup"""
        out = {"res": None, "ab": True, "err": "", "fatal": False, "blog": "",
               "msgs": [], "a2u": 0, "drift": 0, "bdup": 0}
        res, ab, err, fatal, blog = self.geo.resolve(a1, "", city, dist)
        out.update({"res": res, "ab": ab, "err": err, "fatal": fatal, "blog": blog})
        if fatal:
            return out
        # 精度不足时做内容匹配校验（高精度结果不校验，避免误判）
        if res and not ab and res.get("source") == "高德" and a2:
            if level_score(res.get("level", "未知")) < level_score("街道"):
                std_lower = res["std_addr"].lower()
                a1_clean = _clean_single_addr(a1)
                a2_clean = _clean_single_addr(a2)
                match_a1 = a1_clean and (a1_clean.lower() in std_lower or std_lower in a1_clean.lower())
                match_a2 = a2_clean and (a2_clean.lower() in std_lower or std_lower in a2_clean.lower())
                if not match_a1 and not match_a2:
                    ab = True
                    res["is_abnormal"] = True
                    res["std_addr"] += " [内容不匹配]"
        # 地址2备用策略：① 地址1完全失败 → 无条件采用；② 精度更高 → 采用；
        #                 ③ 同精度 → 距离 ≤200 米才采用；④ 精度更低 → 保留地址1
        orig_res = res
        if (not res) or ab:
            if a2:
                r2, a22, e2, f2, blog2 = self.geo.resolve(a2, "", city, dist)
                if f2:
                    out["fatal"] = True
                    out["err"] = e2
                    return out
                if r2 and not a22:
                    r2_score = level_score(r2.get("level", "未知"))
                    res_score = level_score(res.get("level", "未知")) if res else 0
                    should_replace = False
                    if not res:
                        should_replace = True
                        out["msgs"].append(("[地址2备用] 地址1完全失败，采用地址2结果\n", "info"))
                    elif r2_score > res_score:
                        should_replace = True
                        out["msgs"].append((f"[地址2备用] 地址2精度更高({r2.get('level','')}/{r2_score} > {res.get('level','')}/{res_score})，直接采用地址2结果\n", "info"))
                    elif r2_score == res_score:
                        if res.get("gcj_lng") and res.get("gcj_lat") and r2.get("gcj_lng") and r2.get("gcj_lat"):
                            try:
                                dx = abs(r2["gcj_lng"] - res["gcj_lng"]) * 111000 * math.cos(math.radians(res["gcj_lat"]))
                                dy = abs(r2["gcj_lat"] - res["gcj_lat"]) * 111000
                                dd = math.sqrt(dx * dx + dy * dy)
                                if dd <= 200:
                                    should_replace = True
                                    out["msgs"].append((f"[地址2备用] 同精度({r2.get('level','')}/{r2_score})，距离{int(dd)}米≤200米，采用地址2结果\n", "info"))
                                else:
                                    out["msgs"].append((f"[地址2备用] 同精度({r2.get('level','')}/{r2_score})，距离{int(dd)}米>200米，保留地址1结果\n", "info"))
                            except Exception:
                                should_replace = True
                                out["msgs"].append(("[地址2备用] 同精度距离计算异常，回退采用地址2结果\n", "warning"))
                        else:
                            should_replace = True
                            out["msgs"].append((f"[地址2备用] 同精度({r2.get('level','')}/{r2_score})，无坐标对比距离，采用地址2结果\n", "info"))
                    else:
                        out["msgs"].append((f"[地址2备用] 地址2精度({r2.get('level','')}/{r2_score})低于地址1({res.get('level','')}/{res_score})，保留地址1结果\n", "info"))
                    if should_replace:
                        res = r2
                        out["a2u"] = 1
                        out["drift"] += 1 if r2.get("is_drift") else 0
                        out["bdup"] += 1 if r2.get("source") == "百度" else 0
                        if blog2:
                            out["blog"] = (out["blog"] + "\n" + blog2) if out["blog"] else blog2
                        out["msgs"].append((f"[地址2备用] 地址2查询成功: {r2.get('std_addr','')} [精度:{r2.get('level','')}]\n", "success"))
                else:
                    res = None
            else:
                res = None
            # 主地址仅「疑似漂移」且坐标有效时保留结果并标黄，避免坐标与地址一起丢失
            if res is None and orig_res and orig_res.get("is_drift"):
                res = orig_res
                out["drift"] += 1
                out["msgs"].append((f"[漂移保留] 结果疑似漂移（保留并标黄）：{orig_res.get('std_addr','')} [精度:{orig_res.get('level','')}]\n", "warning"))
        out["res"] = res
        out["ab"] = (not res) or bool(res.get("is_abnormal"))
        return out

    @staticmethod
    def _checkpoint(proc, save_cb, state_cb, direction, last_done, failed, done, total, stats):
        """批量过程中落盘：增量保存 Excel + 写任务状态（供断点续跑）"""
        if save_cb:
            try:
                save_cb()
            except Exception:
                pass
        if state_cb:
            try:
                state_cb({"direction": direction, "last_done": int(last_done), "failed": list(failed),
                          "done": int(done), "total": int(total), "stats": dict(stats)})
            except Exception:
                pass

    def forward(self, proc, srow, lrow, aidx, a2idx, cidx, didx, skip, only_ab,
                do_clean, do_merge, do_dedup, log_cb, prog_cb,
                workers=1, save_cb=None, state_cb=None, write_struct=False,
                only_rows=None, save_every=200):
        """批量正向编码。

        本次升级：① 小并发（workers>1 时分组并行解析，工作表读写仍在调用线程串行完成）；
        ② 每 save_every 行增量保存 + 写任务状态，支持断点续跑/只重试失败行；
        ③ 可选输出结构化地址要素列；④ only_rows 限定只处理指定行（重试失败行）。
        """
        rows = [r for r in range(max(2, int(srow)), int(lrow) + 1) if (not only_rows or r in only_rows)]
        total = len(rows)
        success = fail = skipc = a2u = drift = bdup = 0
        failed, result_cache, row_rep = [], {}, {}
        prog_done = 0
        stopped = False
        since_save = 0
        last_done_row = 0
        save_every = max(1, int(save_every or 200))

        def tick(last_row=0):
            """推进进度；累计到 save_every 行就增量保存 + 写断点状态"""
            nonlocal prog_done, since_save, last_done_row
            prog_done += 1
            since_save += 1
            if last_row:
                last_done_row = last_row
            try:
                prog_cb(prog_done, total)
            except Exception:
                pass
            if since_save >= save_every:
                since_save = 0
                self._checkpoint(proc, save_cb, state_cb, "fwd", last_row or prog_done, failed,
                                 prog_done, total,
                                 {"success": success, "fail": fail, "skip": skipc,
                                  "addr2_used": a2u, "drift_cnt": drift, "baidu_up": bdup})

        if write_struct:
            try:
                proc.prepare_struct()
                # 断点续跑/二次运行：补写起始行之前已有结果行的结构化列
                # （这些行此前跑过但未启用结构化输出，防止「前面行只有表头没有数据」）
                for _r in range(2, int(srow)):
                    _cv = proc.ws.cell(row=_r, column=Config.COLS["std_addr"]).value
                    if _cv and str(_cv).strip() and "失败" not in str(_cv) and "无匹配" not in str(_cv):
                        proc.write_structured(_r, parse_addr_parts(*parse_result_cell(_cv)))
            except Exception:
                pass

        # ---- 去重预处理（与处理逻辑完全一致的清洗链路；长度分桶剪枝保持等价） ----
        if do_dedup and rows:
            log_cb("[预处理] 正在分析地址去重...\n", "info")
            vals = {}
            for i in rows:
                a1 = self._cell_str(proc.ws, i, aidx)
                a2 = self._cell_str(proc.ws, i, a2idx)
                city = self._cell_str(proc.ws, i, cidx)
                dist = self._cell_str(proc.ws, i, didx)
                vals[i] = build_query_addr(a1, a2, city, dist, do_clean, do_merge)
            buckets, lens = {}, []
            for pos, i in enumerate(rows):
                L = len(vals[i])
                lens.append(L)
                buckets.setdefault(L, []).append(pos)
            bl = sorted(buckets)
            rep_pos = {}
            for p in range(len(rows)):
                if p in rep_pos:
                    continue
                rep_pos[p] = None
                li = lens[p]
                lo = max(0, (li * 3) // 4 - 1)
                hi = (li * 4) // 3 + 2
                for lj in bl[bisect.bisect_left(bl, lo):bisect.bisect_right(bl, hi)]:
                    for q in buckets[lj]:
                        if q <= p or q in rep_pos:
                            continue
                        if _sim_addr(vals[rows[p]], vals[rows[q]]) >= 0.85:
                            rep_pos[q] = rows[p]
            for q, rp in rep_pos.items():
                if rp:
                    row_rep[rows[q]] = rp
            log_cb(f"[预处理] 发现 {len(row_rep)} 个重复地址，将复用查询结果\n", "info")

        workers = max(1, int(workers or 1))
        cur = 0
        while cur < len(rows):
            if self.stop_flag:
                stopped = True
                log_cb("\n⚠ 用户取消\n", "warning")
                break
            chunk = rows[cur:cur + workers]
            cur += len(chunk)
            pending = []
            for i in chunk:
                try:
                    if skip and proc.is_normal(i):
                        skipc += 1; tick(i); log_cb(f"第 {i} 行：跳过\n", "skip")
                        # 修复：跳过已有结果的行也要补写结构化列（否则只有表头没有数据）
                        if write_struct:
                            _cv = proc.ws.cell(row=i, column=Config.COLS["std_addr"]).value
                            if _cv and str(_cv).strip():
                                proc.write_structured(i, parse_addr_parts(*parse_result_cell(_cv)))
                        continue
                    if only_ab and not proc.is_abnormal(i):
                        skipc += 1; tick(i)
                        # 修复：正常行（已有结果）跳过时同样补写结构化列
                        if write_struct:
                            _cv = proc.ws.cell(row=i, column=Config.COLS["std_addr"]).value
                            if _cv and str(_cv).strip():
                                proc.write_structured(i, parse_addr_parts(*parse_result_cell(_cv)))
                        continue
                    proc.clear_forward(i)
                    rep = row_rep.get(i, 0)
                    if rep and rep in result_cache:
                        res = result_cache[rep]
                        proc.write_forward(i, res)
                        if write_struct:
                            proc.write_structured(i, parse_addr_parts(res.get("std_addr", ""), res.get("level", "")) if res else {})
                        success += 1; tick(i)
                        log_cb(f"第 {i} 行：复用第 {rep} 行结果（去重）\n", "skip")
                        continue
                    city = self._cell_str(proc.ws, i, cidx)
                    a1 = self._cell_str(proc.ws, i, aidx)
                    a2 = self._cell_str(proc.ws, i, a2idx)
                    dist = self._cell_str(proc.ws, i, didx)
                    a1 = build_query_addr(a1, a2, city, dist, do_clean, do_merge)
                    if not a1:
                        proc.write_forward(i, None); fail += 1; failed.append(i); tick(i)
                        log_cb(f"第 {i} 行：地址为空\n", "error")
                        continue
                    pending.append((i, a1, a2, city, dist))
                except Exception as e:
                    fail += 1; failed.append(i); tick(i)
                    log_cb(f"第 {i} 行处理异常: {str(e)[:80]}\n", "error")

            # 并发解析本组（仅网络/字符串计算进入工作线程）
            # 组内去重：代表行所属代表行必然在本组且排在前面，故把"本组内可复用的行"放到
            # 第二波直接用缓存写入，既保持顺序一致，也不会为重复地址多花一次接口调用。
            chunk_set = set(chunk)
            first_jobs = [j for j in pending if not (row_rep.get(j[0]) and row_rep[j[0]] in chunk_set)]
            second_jobs = [j for j in pending if row_rep.get(j[0]) and row_rep[j[0]] in chunk_set]
            results = {}
            if first_jobs:
                if workers <= 1 or len(first_jobs) == 1:
                    for job in first_jobs:
                        try:
                            results[job[0]] = self._resolve_row(job[1], job[2], job[3], job[4])
                        except Exception as e:
                            results[job[0]] = {"res": None, "fatal": False, "blog": "",
                                               "msgs": [(f"处理异常: {str(e)[:80]}\n", "error")]}
                else:
                    from concurrent.futures import ThreadPoolExecutor
                    with ThreadPoolExecutor(max_workers=min(workers, len(first_jobs))) as ex:
                        import concurrent.futures as _cf
                        fut_map = {ex.submit(self._resolve_row, j[1], j[2], j[3], j[4]): j[0] for j in first_jobs}
                        for fut in _cf.as_completed(fut_map):
                            i = fut_map[fut]
                            try:
                                results[i] = fut.result()
                            except Exception as e:
                                results[i] = {"res": None, "fatal": False, "blog": "",
                                              "msgs": [(f"处理异常: {str(e)[:80]}\n", "error")]}

            for (i, _a1, _a2, _city, _dist) in first_jobs:
                r = results.get(i) or {}
                if r.get("fatal"):
                    fail += 1; failed.append(i); tick()
                    log_cb(f"第 {i} 行：致命错误：{r.get('err')}，已终止\n", "error")
                    stopped = True
                    break
                if r.get("blog"):
                    log_cb(r["blog"], "baidu")
                for m, t in r.get("msgs", []):
                    log_cb(m, t)
                res = r.get("res")
                proc.write_forward(i, res)
                if write_struct:
                    proc.write_structured(i, parse_addr_parts(res.get("std_addr", ""), res.get("level", "")) if res else {})
                if do_dedup and res:
                    result_cache[i] = res
                a2u += r.get("a2u", 0); drift += r.get("drift", 0); bdup += r.get("bdup", 0)
                if res:
                    success += 1; tag = "success"
                else:
                    fail += 1; failed.append(i); tag = "error"
                tick(i)
                log_cb(f"第 {i} 行：{'成功' if res else '失败'}\n", tag)

            # 组内重复行：直接复用代表行结果（不再消耗接口调用）
            for (i, _a1, _a2, _city, _dist) in second_jobs:
                rep = row_rep.get(i, 0)
                res = result_cache.get(rep)
                try:
                    proc.write_forward(i, res)
                    if write_struct:
                        proc.write_structured(i, parse_addr_parts(res.get("std_addr", ""), res.get("level", "")) if res else {})
                except Exception as e:
                    fail += 1; failed.append(i); tick(i)
                    log_cb(f"第 {i} 行处理异常: {str(e)[:80]}\n", "error")
                    continue
                if res:
                    success += 1
                else:
                    fail += 1; failed.append(i)
                tick(i)
                log_cb(f"第 {i} 行：复用第 {rep} 行结果（去重）\n", "skip")

            if stopped:
                self._checkpoint(proc, save_cb, state_cb, "fwd", last_done_row or srow,
                                 failed, prog_done, total,
                                 {"success": success, "fail": fail, "skip": skipc,
                                  "addr2_used": a2u, "drift_cnt": drift, "baidu_up": bdup})
                break

        return {"total": total, "success": success, "fail": fail, "skip": skipc,
                "addr2_used": a2u, "drift_cnt": drift, "baidu_up": bdup,
                "gaode_count": self.geo.gcnt, "baidu_count": self.geo.bcnt,
                "cache_hits": getattr(self.geo, "cache_hits", 0),
                "failed_rows": failed, "processed": prog_done,
                "completed": (not stopped) and (not self.stop_flag)}

    def reverse(self, proc, srow, lrow, glidx, gtidx, wlidx, wtidx, skip, log_cb, prog_cb,
                save_cb=None, state_cb=None, only_rows=None, save_every=200):
        """批量逆向编码（本次升级：增量保存 + 任务状态，支持断点续跑/重试失败行）"""
        rows = [r for r in range(max(2, int(srow)), int(lrow) + 1) if (not only_rows or r in only_rows)]
        total = len(rows); success = fail = skipc = 0; gerr = False
        failed, stopped = [], False
        prog_done = 0
        save_every = max(1, int(save_every or 200))
        since_save = 0

        def tick(n=1):
            nonlocal prog_done
            prog_done += n
            try:
                prog_cb(prog_done, total)
            except Exception:
                pass

        for i in rows:
            try:
                if self.stop_flag:
                    stopped = True; log_cb("\n⚠ 用户取消\n", "warning"); break
                tick()
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
                    else:
                        proc.write_reverse(i, None); fail += 1; failed.append(i)
                        log_cb(f"第 {i} 行：无有效坐标\n", "error"); continue
                if not (gl and gt):
                    proc.ws.cell(row=i, column=Config.COLS["rego_full"]).value = "无有效坐标"
                    fail += 1; failed.append(i); continue
                if self.geo.gd_quota_dead and self.geo.use_baidu and self.geo.bak:
                    res, err = self.geo.baidu_regeo(gl, gt)
                else:
                    res, err = self.geo.gaode_regeo(gl, gt)
                if not res:
                    quota_switchable = ("10003" in err) and self.geo.use_baidu and self.geo.bak
                    if quota_switchable:
                        # v2.8.56：高德日配额耗尽，自动切换百度逆地理（优先高德）
                        self.geo.gd_quota_dead = True
                        log_cb(f"第 {i} 行：高德日配额耗尽，自动切换百度逆地理...\n", "warning")
                        res, err = self.geo.baidu_regeo(gl, gt)
                        if res:
                            proc.write_reverse(i, res); success += 1; log_cb(f"第 {i} 行：成功（百度）\n", "success")
                            since_save += 1
                            if since_save >= save_every:
                                since_save = 0
                                self._checkpoint(proc, save_cb, state_cb, "rev", i, failed, prog_done, total,
                                                 {"success": success, "fail": fail, "skip": skipc})
                            continue
                    proc.ws.cell(row=i, column=Config.COLS["rego_full"]).value = f"失败：{err}"
                    # 修复：高德鉴权/配额类全局错误（且无可切换的百度通道）才终止整体流程；
                    # 若已切换百度仍失败，说明是高德额度问题已被绕过，不应误判为全局错误终止。
                    if (("10001" in err or "10003" in err) and not (self.geo.use_baidu and self.geo.bak)):
                        gerr = True; stopped = True; fail += 1; failed.append(i)
                        log_cb(f"第 {i} 行：全局错误 {err}\n", "error"); break
                    fail += 1; failed.append(i); log_cb(f"第 {i} 行：失败 - {err}\n", "error")
                else: proc.write_reverse(i, res); success += 1; log_cb(f"第 {i} 行：成功\n", "success")
                since_save += 1
                if since_save >= save_every:
                    since_save = 0
                    self._checkpoint(proc, save_cb, state_cb, "rev", i, failed, prog_done, total,
                                     {"success": success, "fail": fail, "skip": skipc})
            except Exception as e:
                log_cb(f"第 {i} 行处理异常: {str(e)[:80]}\n", "error")
                fail += 1; failed.append(i); continue
        if stopped or self.stop_flag:
            self._checkpoint(proc, save_cb, state_cb, "rev", srow, failed, prog_done, total,
                             {"success": success, "fail": fail, "skip": skipc})
        return {"total": total, "success": success, "fail": fail, "skip": skipc,
                "global_error": gerr, "gaode_count": self.geo.gcnt,
                "cache_hits": getattr(self.geo, "cache_hits", 0),
                "failed_rows": failed, "processed": prog_done,
                "completed": (not stopped) and (not self.stop_flag)}



__all__ = ['_data_dir', 'norm_division', 'LocalDB', 'get_db', 'DivisionTable', 'get_div_table', 'task_state_path', 'load_task_state', 'save_task_state', 'clear_task_state', 'url_enc', 'http_get', 'json_val', 'geo_field', 'regeo_field', 'regeo_comp', 'baidu_fmt', 'clean_addr', 'rm_paren', 'addr_levels', 'extract_city', 'build_addr', 'has_detail', 'extract_district', 'check_drift', 'extract_core', '_pick_road', 'parse_addr_parts', 'dedup_addr', 'build_query_addr', 'read_headers', 'read_headers_rw', 'dedup_headers', 'fast_preview', '_get_self_district_info', 'get_poi_district_info', '_static_expand_kws', '_merge_expand_kws', 'get_poi_district_units', '_poi_near_any', 'PoiIndex', '_http_retry', '_is_quota_error', '_is_qps_error', 'http_post', 'AiClient', 'ai_semantic_dedup', 'ai_l1_postprocess', '_parse_polyline', 'get_district_polygon', '_point_in_polygon', '_calc_grid_radius', '_QuotaStop', 'grid_collect_poi', '_clean_single_addr', '_sim_addr', '_xml_esc', '_gen_kml', 'level_score', 'gaode_err_msg', 'baidu_err_msg', '_delta', '_iter_transform', 'wgs84_to_gcj02', 'gcj02_to_wgs84', 'bd09_to_gcj02', 'gcj02_to_bd09', 'convert_coords', 'get_latest_version', 'cmp_ver', '_verify_exe', 'download_update', 'RateLimiter', 'GeoCoder', 'ExcelProcessor', 'BatchProcessor', 'PROVINCES']


def fast_preview(filepath, sheet_name="Sheet1", max_rows=200):
    """极速读取 Excel 前 max_rows 个有数据行（zipfile 直读 sheet XML，跳过 styles.xml）。

    背景：openpyxl load_workbook 无论 read_only 与否都会全量解析 styles.xml，
    表尾几十万行套用格式的文件 styles.xml 巨大，解析可达数秒~数十秒，
    是「点击导入 Excel 卡死」的根因。本函数只读 xl/worksheets/*.xml 数据流，
    完全不触碰样式，大文件也能秒开。

    返回：(headers_raw, rows, max_col) 或 None（非 xlsx / 无 Sheet1 / 解析失败，
    调用方回退 openpyxl）。headers_raw 为第 1 行原始值列表（未去重，用 dedup_headers
    处理）；rows 为 [ {列索引(1起): 字符串值}, ... ]，仅含有值单元格、仅含非全空行。
    兼容：共享字符串、内联字符串、公式缓存值（等价 data_only）、数字。
    """
    import zipfile
    import xml.etree.ElementTree as ET
    try:
        zf = zipfile.ZipFile(filepath)
    except Exception:
        return None
    M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    try:
        # workbook.xml → sheet 名 → rId
        root = ET.fromstring(zf.read("xl/workbook.xml"))
        rid = None
        for sh in root.iter("{%s}sheet" % M):
            if sh.get("name") == sheet_name:
                rid = sh.get("{%s}id" % R)
                break
        if rid is None:
            return None
        # rels → rId → sheet target
        rroot = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rroot:
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            return None
        # Target 三种形态：/xl/worksheets/sheet1.xml（绝对）、worksheets/sheet1.xml（相对xl/）、
        # xl/worksheets/sheet1.xml（相对包根）——统一为包内路径
        if target.startswith("/"):
            target = target[1:]
        elif not target.startswith("xl/"):
            target = "xl/" + target
        # sharedStrings（可能不存在）
        sst = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sroot = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sroot.iter("{%s}si" % M):
                sst.append("".join(t.text or "" for t in si.iter("{%s}t" % M)))
        # 流式解析 sheet：只取表头行 + 前 max_rows 个有数据行
        def colnum(ref):
            m = __import__('re').match(r'^([A-Z]+)(\d+)$', ref or '')
            if not m:
                return 0
            n = 0
            for ch in m.group(1):
                n = n * 26 + (ord(ch) - 64)
            return n
        headers = {}
        rows = []
        cur = {}
        max_col = 0
        for _ev, elem in ET.iterparse(zf.open(target), events=("end",)):
            tag = elem.tag.split("}")[-1]
            if tag == "row":
                rnum = elem.get("r")
                try:
                    rnum = int(rnum or 0)
                except Exception:
                    rnum = 0
                if rnum == 1:
                    headers = dict(cur)
                elif any(v for v in cur.values()):
                    rows.append(dict(cur))
                    if len(rows) >= max_rows:
                        elem.clear()
                        break
                # 只在此处 clear：释放整行内存（大文件关键）。
                # 不能在循环末尾对每个元素 clear——<is>/<t> 等子元素被提前清空后，
                # c 的 end 事件读取 inlineStr 内容会得到空值。
                elem.clear()
                cur = {}
            elif tag == "c":
                c = colnum(elem.get("r"))
                t = elem.get("t")
                val = None
                for ch in elem:
                    ct = ch.tag.split("}")[-1]
                    if ct == "v":
                        val = (ch.text or "").strip()
                        break
                    if ct == "is":
                        val = "".join(x.text or "" for x in ch.iter()
                                      if x.tag.split("}")[-1] == "t")
                if t == "s" and val is not None:
                    try:
                        val = sst[int(val)]
                    except Exception:
                        val = ""
                if val is not None and val != "":
                    cur[c] = str(val)
                    if c > max_col:
                        max_col = c
        zf.close()
        if not headers or not max_col:
            return None
        hlist = [""] * (max_col + 1)
        for c, v in headers.items():
            if 0 < c <= max_col:
                hlist[c] = v
        hnames = [hlist[i] if i and hlist[i] else f"列{i}" for i in range(1, max_col + 1)]
        return hnames, rows, max_col
    except Exception:
        try:
            zf.close()
        except Exception:
            pass
        return None
