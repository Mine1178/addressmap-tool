# -*- coding: utf-8 -*-
"""图形界面层：所有 tkinter 页面与主程序"""
import sys, os, json, math, re, time, threading, random, urllib.request, urllib.parse, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import font as tkfont
from src.config import *
from src.config import _get_pil, _get_openpyxl   # import * 不导入下划线函数，显式补导入（修复 NameError）
from src.core import *
from src.sponsor_widget import show_sponsor_button

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
        # 修复：原来直接用表头文本当列标识，行号列固定叫「序号」。
        # 当 Excel 首列本身就叫「序号」（自带模板即是）时会出现两列同名，
        # Tkinter 的列配置/取值会发生歧义。改为内部用 c0/c1/... 作列标识，文本仍显示表头。
        ids = ["c%d" % i for i in range(len(self.headers) + 1)]
        self.tree["columns"] = ids
        self.tree.column(ids[0], width=55, anchor="center"); self.tree.heading(ids[0], text="序号")
        for i, h in enumerate(self.headers):
            self.tree.column(ids[i + 1], width=120, anchor="w", minwidth=80, stretch=False)
            self.tree.heading(ids[i + 1], text=h)
        si = self.page * self.page_size; ei = min(si + self.page_size, len(self.data))
        for i in range(si, ei):
            vals = [i+1] + [str(self.data[i].get(h, ""))[:27]+"..." if len(str(self.data[i].get(h,"")))>30 else str(self.data[i].get(h,"")) for h in self.headers]
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
        self.cv_running = False; self.cv_stop_flag = False   # 坐标转换批量任务独立状态
        self._map_pending = None                             # 地图加载期间的"待补"请求
        self.poi_running = False
        self._last = {}; self.bp = BatchProcessor()
        self.df = None  # 延迟创建字体
        # 使用期限检查（2028-01-01）
        self.is_expired = datetime.datetime.now() > datetime.datetime(2028, 1, 1)
        self._expire_shown_at = 0   # 过期提示去抖时间戳
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
        """起始行：非法输入或 <2（表头行）一律回落到 2，避免把表头当数据覆盖"""
        try:
            v = int(float(getattr(self, f"{direction}_start").get()))
        except Exception:
            return 2
        return max(2, v)

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

    def _ui(self, fn):
        """从后台线程安全地把 UI 更新排到主线程执行。

        窗口已被关闭（用户中途退出）时 root.after 会抛 TclError，
        原实现会让后台线程在 except/finally 里二次抛异常、留下"卡死"状态，
        这里统一吞掉这类异常，保证后台任务能干净收尾。
        """
        try:
            self.root.after(0, fn)
        except Exception:
            pass

    def _log(self, direction, msg, tag="normal"):
        getattr(self, f"{direction}_log").insert(tk.END, msg, tag); getattr(self, f"{direction}_log").see(tk.END)

    # ---------- 首页 ----------
    def _build_home(self):
        f = self._page(); self.pages["home"] = f
        tf = tk.Frame(f, bg=Config.C_BG); tf.pack(fill=tk.X, pady=(0, 4))
        tk.Label(tf, text="欢迎使用地址转换工具", font=("微软雅黑", 24, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w")
        tk.Label(tf, text="基于高德/百度地图 API 的地址解析、坐标转换与 POI 数据采集，支持 AI 大模型智能增强",
                 font=("微软雅黑", 12), bg=Config.C_BG, fg=Config.C_TEXT_L).pack(anchor="w", pady=(4, 0))
        tk.Frame(tf, bg=Config.C_ACCENT, height=2).pack(fill=tk.X, pady=(8, 0))
        def _card(parent, t, d, c, first, target=None):
            cd = tk.Frame(parent, bg=Config.C_CARD, highlightbackground=Config.C_BORDER, highlightthickness=1)
            cd.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0 if first else 8, 0))
            tk.Frame(cd, bg=c, height=4).pack(fill=tk.X)
            tk.Label(cd, text=t, font=("微软雅黑", 12, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT).pack(anchor="w", padx=12, pady=(10, 4))
            tk.Label(cd, text=d, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_L, justify=tk.LEFT).pack(anchor="w", padx=12, pady=(0, 6))
            tk.Label(cd, text="进入 →", font=("微软雅黑", 9, "bold"), bg=Config.C_CARD, fg=c).pack(anchor="e", padx=12, pady=(0, 8))
            if target:
                def _enter(_e=None):
                    cd.config(bg="#eef2ff", highlightbackground=c, cursor="hand2")
                    for w in cd.winfo_children():
                        if w.cget("bg") == Config.C_CARD: w.config(bg="#eef2ff")
                def _leave(_e=None):
                    cd.config(bg=Config.C_CARD, highlightbackground=Config.C_BORDER, cursor="")
                    for w in cd.winfo_children():
                        if w.cget("bg") == "#eef2ff": w.config(bg=Config.C_CARD)
                cd.bind("<Enter>", _enter); cd.bind("<Leave>", _leave)
                for w in cd.winfo_children():
                    w.bind("<Enter>", _enter); w.bind("<Leave>", _leave)
                cd.bind("<Button-1>", lambda _e: self.show(target))
                for w in cd.winfo_children():
                    w.bind("<Button-1>", lambda _e: self.show(target))
        cf1 = tk.Frame(f, bg=Config.C_BG); cf1.pack(fill=tk.X, pady=(8, 6))
        _card(cf1, "📍 单条查询", "快速查询单个地址的 GCJ-02 / WGS-84 坐标\n支持地址查询与经纬度逆编码", Config.C_INFO, True, "single")
        _card(cf1, "📊 批量正向编码", "从 Excel 读取地址，批量获取\n标准化结果与坐标", Config.C_SUCCESS, False, "batch_fwd")
        _card(cf1, "🔄 批量逆向编码", "从 Excel 读取坐标，批量解析\n为标准地址信息", Config.C_WARNING, False, "batch_rev")
        _card(cf1, "🧭 坐标转换", "GCJ-02 / WGS-84 / BD-09\n三系坐标互转，支持单点与批量", "#7c3aed", False, "converter")
        cf2 = tk.Frame(f, bg=Config.C_BG); cf2.pack(fill=tk.X, pady=(0, 6))
        _card(cf2, "🗺️ POI 采集", "行政区划 / 周边检索\n自动下钻 + 网格深挖补全", "#0ea5e9", True, "poi")
        _card(cf2, "✨ AI 智能增强", "关键词扩展 / 地址清洗\n语义去重，提升采集精度", "#d946ef", False, "settings")
        _card(cf2, "⚙️ 双平台配置", "支持高德 + 百度 API\n智能 Fallback 自动切换", "#f59e0b", False, "settings")
        steps = Collapsible(f, title="▎使用步骤", expanded=True, tf=("微软雅黑", 13, "bold"), tfg=Config.C_ACCENT)
        steps.pack(fill=tk.X, pady=(6, 6))
        for i, s in enumerate(["点击左侧「API 设置」，配置高德/百度 API Key（自动保存到 addr_config.json）",
                               "可选：启用 AI 大模型智能增强，选择平台并填写 Key（未配置时自动使用规则模式）",
                               "准备 Excel 文件，确保包含 Sheet1，按下方表格格式填写数据",
                               "选择「批量正向编码」或「批量逆向编码」，选择文件后开始处理",
                               "POI 采集：行政区划或周边检索，自动下钻区县并网格深挖补全，突破单次检索 225 条上限",
                               "结果自动去重合并，异常/漂移黄色高亮；支持断点续传与失败重试",
                               "相同地址自动复用本地缓存（设置页可查看命中率/清空缓存），长时间跑批不再重复消耗配额",
                               "批量任务中断后重选同一文件，可选择「从断点续跑」或「只重试失败行」；正向编码可开并发加速",
                               "需要拆分地址要素时，勾选「输出结构化地址要素」即会在表尾追加 省/市/区县/街道/道路/门牌/小区 七列"]):
            sr = tk.Frame(steps.cf(), bg=Config.C_CARD); sr.pack(fill=tk.X, padx=6, pady=3)
            tk.Label(sr, text=str(i + 1), font=("微软雅黑", 10, "bold"), bg=Config.C_ACCENT, fg="#ffffff", width=2, anchor="center").pack(side=tk.LEFT, ipady=1)
            tk.Label(sr, text=s, font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT_M).pack(side=tk.LEFT, padx=(8, 0))
        fmt = Collapsible(f, title="▎Excel 数据格式说明", expanded=False, tf=("微软雅黑", 13, "bold"), tfg=Config.C_SUCCESS)
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
            # 修复：原示例行是「前 7 列 *2 + 4 列空」共 18 个值，而表头只有 14 列，
            # 导致下载的模板多出 4 个空列（列15~18），示例数据也被写错位。
            # 现在示例行严格按表头列数生成，只写入非空值。
            samples = [[1, "四川省", "成都市", "", "西南电子科技大学", ""],
                       [2, "北京市", "北京市", "", "天安门", "人民英雄纪念碑"],
                       [3, "云南省", "曲靖市", "麒麟区", "曲靖市农村信用合作社联合社", "寥廓北路中段271-1号"]]
            for ri, row in enumerate(samples, 2):
                for ci, v in enumerate(row[:len(hs)], 1):
                    if v not in (None, ""):
                        ws.cell(row=ri, column=ci, value=v)
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
        # 修复：Tk 输入在主线程读取后快照；异常提示用默认参数绑定，
        # 避免 `except ... as e` 结束后 e 被删除、延迟执行的 lambda 抛 NameError
        addr_in = self.saddr.get().strip()
        coord_in = self.scoord.get().strip()
        csys = self.csys.get()
        def task():
            try:
                ok, msg = self.geo.check()
                if not ok:
                    self._ui(lambda m=msg: self.serr.config(text=f"密钥错误：{m}"))
                    self._ui(lambda: self.status.set("就绪"))
                    self._ui(lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
                    return
                if mode == "address":
                    addr = addr_in
                    if not addr:
                        self._ui(lambda: self.serr.config(text="请输入地址"))
                        self._ui(lambda: self.status.set("就绪"))
                        self._ui(lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
                        return
                    res, ab, err, fatal, blog = self.geo.resolve(addr, "", "", "")
                    self._ui(lambda r=res, a=ab, e2=err, f=fatal, b=blog: self._show_addr(r, a, e2, f, b))
                else:
                    try:
                        lng, lat = self._parse_coord(coord_in)
                    except Exception as ex:
                        self._ui(lambda m=str(ex): self.serr.config(text=f"坐标格式错误：{m}"))
                        self._ui(lambda: self.status.set("就绪"))
                        self._ui(lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
                        return
                    gl, gt = {"WGS-84": wgs84_to_gcj02, "BD-09": bd09_to_gcj02, "GCJ-02": lambda x,y:(x,y)}[csys](lng, lat)
                    self._do_regeo(gl, gt)
            except Exception as ex:
                self._ui(lambda m=str(ex): self.serr.config(text=f"查询异常：{m}"))
                self._ui(lambda: self.status.set("就绪"))
                self._ui(lambda: self.squery_btn.config(state=tk.NORMAL) if hasattr(self, 'squery_btn') else None)
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
        try:
            parts = parse_addr_parts(res.get("std_addr", ""), res.get("level", ""))
            segs = [(t, parts[k]) for k, t in [("province", "省"), ("city", "市"), ("district", "区县"),
                                               ("town", "街道/乡镇"), ("road", "道路"), ("no", "门牌号"),
                                               ("community", "小区/POI")] if parts.get(k)]
            if segs:
                lines.append("【结构化】" + " | ".join(f"{t}: {v}" for t, v in segs))
            if res.get("adcode"):
                lines.append(f"【区县码】{res['adcode']}")
        except Exception:
            pass
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
        # 防并发锁：正在加载时记录"最新请求"，加载完成后补一次，
        # 避免缩放/平移/查询切换过程中的地图更新被直接丢弃（表现为"地图不跟随"）
        if self._map_loading:
            self._map_pending = (lng, lat)
            return
        self._map_loading = True
        self._map_pending = None
        # 修复：Tk 控件只能在主线程访问。尺寸在主线程读取后传入子线程，
        # 避免子线程 winfo_width() 触发 RuntimeError: main thread is not in main loop
        try:
            _rw = self.mlbl.winfo_width(); _rh = self.mlbl.winfo_height()
        except Exception:
            _rw = _rh = 0
        if _rw < 100 or _rh < 100:
            try:
                _rw = max(self.root.winfo_width() - 640, 600); _rh = max(self.root.winfo_height() - 120, 500)
            except Exception:
                _rw, _rh = 600, 500
        w, h = min(max(_rw, 600), 1024), min(max(_rh, 500), 1024)  # 高德静态地图最大尺寸限制
        def task():
            try:
                key = self.cfg.get("gaode_key", ""); zoom = self.zlvl; Image, ImageTk = _get_pil()
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
                pend = self._map_pending
                self._map_pending = None
                if pend:
                    self._ui(lambda p=pend: self._load_map(p[0], p[1]))
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
            # 本次升级：结构化地址要素输出（表尾追加列，不影响既有 7~14 列布局）
            self.fwd_struct = tk.BooleanVar(value=bool(self.cfg.get("write_struct", False)))
            tk.Checkbutton(pinner, text="输出结构化地址要素（表尾追加 省/市/区县/街道/道路/门牌/小区）",
                           variable=self.fwd_struct, wraplength=280, justify=tk.LEFT,
                           font=("微软雅黑", 9), bg="#f8fafc", fg=Config.C_TEXT, selectcolor="#f8fafc").pack(anchor="w", padx=8, pady=(0, 6))
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
        """后台线程加载 Excel 预览（界面永不冻结）。

        修复：原主线程同步 load_workbook——表尾几十万行套用格式的文件 styles.xml 巨大，
        openpyxl 会全量解析样式表，耗时数秒~数十秒，表现为「点击导入 Excel 卡死」。
        现在优先用 fast_preview（zipfile 直读数据流、跳过样式解析，大文件也秒开），
        非 xlsx 或解析失败自动回退 openpyxl；全程在后台线程执行，完成后回投主线程更新。
        """
        if getattr(self, "_preview_guard", False): return   # 上一次加载还在进行，防重复点击
        self._preview_guard = True
        self.status.set(f"正在加载 {os.path.basename(path)} …")
        def work():
            try:
                res = fast_preview(path, "Sheet1", 200)
                if res is not None:
                    hnames, rows, _mc = res
                    hs = dedup_headers(hnames)
                    data = [{hs[i - 1] if i - 1 < len(hs) else f"列{i}": v
                             for i, v in r.items()} for r in rows]
                    out = (hs, data)
                else:
                    ox = _get_openpyxl()[0]
                    wb = ox.load_workbook(path, read_only=True, data_only=True)
                    if "Sheet1" not in wb.sheetnames:
                        self.root.after(0, lambda: (messagebox.showwarning("提示", "Excel 中未找到 Sheet1"),
                                                    wb.close(), self._finish_preview()))
                        return
                    ws = wb["Sheet1"]
                    hs = read_headers(ws)
                    data = []
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if len(data) >= 200: break
                        if not any(v is not None for v in row): continue
                        data.append({hs[i] if i < len(hs) else f"列{i+1}": (str(v) if v is not None else "") for i, v in enumerate(row)})
                    wb.close()
                    out = (hs, data)
            except Exception as e:
                self.root.after(0, lambda e=e: (messagebox.showerror("错误", f"读取失败：{e}"),
                                                self._finish_preview()))
                return
            self.root.after(0, lambda out=out: self._apply_preview(out, path, direction))
        threading.Thread(target=work, daemon=True).start()

    def _finish_preview(self):
        self._preview_guard = False

    def _apply_preview(self, out, path, direction):
        """主线程回投：填充下拉框并刷新预览表格"""
        self._preview_guard = False
        hs, data = out
        if direction == "fwd":
            for attr in ["fwd_addr_col", "fwd_addr2_col", "fwd_city_col", "fwd_dist_col"]:
                cb = getattr(self, attr, None)
                if cb: cb["values"] = hs
        else:
            for attr in ["rev_gcj_lng_col", "rev_gcj_lat_col", "rev_wgs_lng_col", "rev_wgs_lat_col"]:
                cb = getattr(self, attr, None)
                if cb: cb["values"] = hs
        getattr(self, f"{direction}_preview").load(hs, data)
        self._auto_sel(hs, direction)
        getattr(self, f"{direction}_notebook").select(0)
        self.status.set(f"已加载：{os.path.basename(path)}")

    def _auto_sel(self, hs, direction):
        if direction == "fwd":
            am = a2m = cm = dm = None
            for h in hs:
                hl = h.lower()
                if am is None and any(k in hl for k in ["地址1","地址一","addr1","address1","详细地址"]): am = h
                if a2m is None and any(k in hl for k in ["地址2","地址二","addr2","address2","备用地址"]): a2m = h
                if cm is None and any(k in hl for k in ["市","城市","city"]): cm = h
                if dm is None and any(k in hl for k in ["区县","县区","区/县","county","district"]): dm = h
            if am: self.fwd_addr_col.set(am)
            if a2m: self.fwd_addr2_col.set(a2m)
            if cm: self.fwd_city_col.set(cm)
            if dm: self.fwd_dist_col.set(dm)
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
        """检查是否过期，过期则弹窗（去抖：10 秒内只提示一次，避免频繁操作时弹窗轰炸）"""
        if not self.is_expired:
            return False
        now = time.time()
        if now - getattr(self, "_expire_shown_at", 0) > 10:
            self._expire_shown_at = now
            self._show_expire_notice()
        return True

    def _refresh_preview(self, direction):
        """刷新数据预览：重新读取 Excel 文件并更新预览"""
        if self._check_expired(): return
        path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先导入 Excel 文件")
            return
        try:
            self._load_preview(path, direction)
        except Exception as e:
            messagebox.showerror("错误", f"刷新失败：{e}")

    def _export_abnormal(self, direction):
        if self._check_expired(): return
        path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请先选择有效的 Excel 文件"); return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                              initialfile=f"异常行_{'正向' if direction=='fwd' else '逆向'}编码.xlsx")
        if not out: return
        # 修复：Tk 变量必须在主线程读取，且异常行判定需要知道「地址列 / 处理方向」
        sr = self._get_start_row(direction)
        addr_choice = self.fwd_addr_col.get() if direction == "fwd" else ""
        def task():
            try:
                proc = ExcelProcessor(path)
                hs = read_headers_rw(proc.ws)
                a_idx = self._col_idx(hs, addr_choice, Config.COLS["addr1"]) if addr_choice else None
                ok, msg = proc.export_abnormal(out, sr, direction, a_idx)
                self._ui(lambda: messagebox.showinfo("导出结果", msg) if ok else messagebox.showwarning("导出失败", msg))
            except Exception as e: self._ui(lambda e=e: messagebox.showerror("错误", f"导出异常：{e}"))
            finally:
                try:
                    if 'proc' in locals() and proc is not None:
                        proc.close()
                except: pass
        threading.Thread(target=task, daemon=True).start()

    def _preview_clean(self):
        """预览地址清洗效果（严格遵循 do_clean / do_merge 开关；后台线程执行，界面不冻结）"""
        if self._check_expired(): return
        path = self.fwd_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先导入 Excel 文件"); return
        if not self.fwd_clean.get() and not self.fwd_merge.get():
            messagebox.showinfo("提示", "请先勾选「自动清洗」或「智能合并」后再预览"); return
        if getattr(self, "_preview_guard", False): return
        self._preview_guard = True
        self.status.set("正在计算清洗预览 …")
        # 主线程读取 Tk 变量与配置（线程安全）
        a1c_sel = self.fwd_addr_col.get(); a2c_sel = self.fwd_addr2_col.get()
        ctc_sel = self.fwd_city_col.get(); dtc_sel = self.fwd_dist_col.get()
        do_clean = self.fwd_clean.get(); do_merge = self.fwd_merge.get()
        try:
            cfg = dict(self.cfg)
        except Exception:
            cfg = self.cfg
        def work():
            try:
                data = None
                res = fast_preview(path, "Sheet1", 200)
                if res is not None:
                    hnames, rows, _mc = res
                    hs = dedup_headers(hnames)
                    data = [{hs[i - 1] if i - 1 < len(hs) else f"列{i}": v
                             for i, v in r.items()} for r in rows]
                else:
                    ox = _get_openpyxl()[0]
                    wb = ox.load_workbook(path, read_only=True, data_only=True)
                    if "Sheet1" not in wb.sheetnames:
                        self.root.after(0, lambda: (messagebox.showwarning("提示", "Excel 中未找到 Sheet1"),
                                                    wb.close(), self._finish_preview()))
                        return
                    ws = wb["Sheet1"]
                    hs = read_headers(ws)
                    data = []
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if len(data) >= 200: break
                        if not any(v is not None for v in row): continue
                        data.append({hs[i] if i < len(hs) else f"列{i+1}": (str(v) if v is not None else "") for i, v in enumerate(row)})
                    wb.close()
                a1c = self._pick_header(hs, a1c_sel, "地址1")
                a2c = self._pick_header(hs, a2c_sel, "地址2")
                ctc = self._pick_header(hs, ctc_sel, "市")
                dtc = self._pick_header(hs, dtc_sel, "区县", "区/县", "县", "区")
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
                # AI 智能清洗对照（仅预览，单次批量调用；批量处理仍走规则引擎保证性能与确定性）
                ai_lines = None
                try:
                    ai = AiClient(cfg)
                    if ai.available:
                        sample = [str(row.get(a1c, "") or "").strip() for row in data[:10]]
                        res_ai = ai.batch_clean_addresses(sample)
                        if res_ai:
                            lines = ["【AI 智能清洗对照（前10行，仅预览）】"]
                            for k, v in res_ai.items():
                                if str(v) != k:
                                    lines.append("'" + str(k) + "' → '" + str(v) + "'")
                            if len(lines) > 1:
                                ai_lines = lines
                except Exception:
                    ai_lines = None
                self.root.after(0, lambda msg=msg, ai_lines=ai_lines: self._apply_clean_preview(msg, ai_lines))
            except Exception as e:
                self.root.after(0, lambda e=e: (messagebox.showerror("错误", f"预览失败：{e}"),
                                                self._finish_preview()))
        threading.Thread(target=work, daemon=True).start()

    def _apply_clean_preview(self, msg, ai_lines):
        """主线程回投：显示清洗预览与 AI 对照"""
        self._preview_guard = False
        messagebox.showinfo("清洗预览", msg)
        if ai_lines:
            messagebox.showinfo("AI 清洗对照", "\n".join(ai_lines))
    def _export_kml(self, direction):
        """导出KML文件"""
        if self._check_expired(): return
        path = getattr(self, f"{direction}_path").get()
        if not path or not os.path.exists(path): messagebox.showwarning("提示", "请先选择有效的 Excel 文件"); return
        out = filedialog.asksaveasfilename(defaultextension=".kml", filetypes=[("KML文件", "*.kml")],
                                           initialfile=f"地址坐标结果_{time.strftime('%Y%m%d')}.kml")
        if not out: return
        # 修复：Tk 变量与起始行在主线程读取；逆向结果改用「完整地址」列且按方向取末行
        is_rev = direction == "rev"
        sr = self._get_start_row(direction)
        addr_choice = "" if is_rev else self.fwd_addr_col.get()
        def task():
            try:
                proc = ExcelProcessor(path)
                hs = read_headers_rw(proc.ws)
                if not is_rev:
                    addr_idx = self._col_idx(hs, addr_choice, Config.COLS["addr1"])
                else:
                    addr_idx = 0  # 逆向编码不需要地址列作为KML名称
                lr = proc.last_data_row(direction, addr_idx or None)
                count = _gen_kml(proc, sr, lr, addr_idx, out, direction)
                if count <= 0:
                    self._ui(lambda: messagebox.showwarning(
                        "无可导出数据",
                        "未找到可导出的坐标点。\n\n请确认：\n"
                        "① 已对该文件执行批量正向/逆向编码；\n"
                        "② 坐标列与结果列有值（非空、非“失败/无匹配”）。"))
                else:
                    self._ui(lambda c=count: messagebox.showinfo("导出成功", f"已导出 {c} 个坐标点到 KML\n文件：{out}"))
            except Exception as e: self._ui(lambda e=e: messagebox.showerror("错误", f"导出失败：{e}"))
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
        if self.running or self.cv_running:
            messagebox.showwarning("提示", "已有批量任务正在运行，请等待完成或先点击「停止」")
            return
        sr = self._get_start_row(direction)
        # ---- 本次升级：断点续跑 / 只重试失败行 ----
        only_rows = None
        resume_note = ""
        state = load_task_state(path)
        if state and state.get("direction") == direction:
            done0 = int(state.get("done", 0) or 0)
            total0 = int(state.get("total", 0) or 0)
            failed0 = [int(x) for x in (state.get("failed") or [])]
            ans = messagebox.askyesnocancel(
                "发现未完成任务",
                "该文件上次未跑完：\n\n"
                f"方向：{'正向编码' if is_fwd else '逆向编码'}\n"
                f"进度：{done0} / {total0} 行（最后处理到第 {state.get('last_done', '?')} 行）\n"
                f"失败待重试：{len(failed0)} 行\n\n"
                "「是」= 从断点继续跑剩余行\n"
                "「否」= 只重试上次失败的行\n"
                "「取消」= 忽略记录，完整重跑一遍")
            if ans is True:
                sr = max(sr, int(state.get("last_done", 0) or 0) + 1)
                resume_note = f"[续跑] 从第 {sr} 行继续（上次已完成 {done0}/{total0}）\n"
            elif ans is False and failed0:
                only_rows = set(failed0)
                resume_note = f"[重试] 仅重试上次失败的 {len(only_rows)} 行\n"
            else:
                clear_task_state(path)
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
        if resume_note:
            getattr(self, f"{direction}_log").insert(tk.END, resume_note)
        ao = getattr(self, f"{direction}_auto_open").get()
        self.cfg["auto_open_excel"] = ao
        if is_fwd and hasattr(self, "fwd_struct"):
            self.cfg["write_struct"] = bool(self.fwd_struct.get())
        try:
            Config.save(self.cfg)
        except Exception:
            pass   # 配置写不进去不应阻塞批量处理
        # 修复：Tk 变量（下拉框/勾选框）统一在主线程读取后快照给后台线程使用，
        # 避免后台线程直接读 Tk 变量（Tkinter 非线程安全，且窗口关闭时会抛异常）
        ui = {"skip": getattr(self, f"{direction}_skip").get(), "auto_open": ao,
              "addr": "", "addr2": "", "city": "", "dist": "",
              "clean": False, "merge": False, "dedup": False, "abnormal": False,
              "lng": "", "lat": "", "wlng": "", "wlat": "",
              "workers": max(1, min(5, int(self.cfg.get("fwd_workers", 3) or 3))),
              "save_every": max(1, int(self.cfg.get("save_every", 200) or 200)),
              "struct": bool(self.fwd_struct.get()) if (is_fwd and hasattr(self, "fwd_struct")) else False,
              "adcode": bool(self.cfg.get("write_adcode", True))}
        if is_fwd:
            ui.update({"addr": self.fwd_addr_col.get(), "addr2": self.fwd_addr2_col.get(),
                       "city": self.fwd_city_col.get(), "dist": self.fwd_dist_col.get(),
                       "clean": self.fwd_clean.get(), "merge": self.fwd_merge.get(),
                       "dedup": self.fwd_dedup.get(),
                       "abnormal": self.fwd_abnormal.get() if hasattr(self, "fwd_abnormal") else False})
        else:
            ui.update({"lng": self.rev_gcj_lng_col.get(), "lat": self.rev_gcj_lat_col.get(),
                       "wlng": self.rev_wgs_lng_col.get(), "wlat": self.rev_wgs_lat_col.get()})
        def task():
            # 先定义回调函数，确保后续日志可用
            def log_cb(msg, tag="normal"):
                self._ui(lambda m=msg, t=tag: self._log(direction, m, t))
            def prog_cb(val, total_val):
                self._ui(lambda v=val, t=total_val: (getattr(self, f"{direction}_prog_lbl").config(text=f"{v} / {t}"), getattr(self, f"{direction}_prog").config(value=v)))
            try:
                # 修复：ExcelProcessor 构造（openpyxl 加载，含 styles.xml 解析）在大样式表文件上
                # 可能耗时数秒；期间进度条先转圈并提示「正在读取 Excel…」，加载完成立即切换，
                # 避免用户误以为程序无响应
                self._ui(lambda: (getattr(self, f"{direction}_prog").config(mode="indeterminate"),
                                  getattr(self, f"{direction}_prog").start(12),
                                  getattr(self, f"{direction}_prog_lbl").config(text="正在读取 Excel …")))
                proc = ExcelProcessor(path)
                self._ui(lambda: (getattr(self, f"{direction}_prog").stop(),
                                  getattr(self, f"{direction}_prog").config(mode="determinate")))
                log_cb(f"[系统] Excel 读取成功，共 {proc.ws.max_column} 列（有效数据行数以进度条为准）\n", "info")
                proc.write_adcode = ui["adcode"]
                if is_fwd and ui["struct"]:
                    sc = proc.prepare_struct()
                    log_cb(f"[系统] 结构化地址要素输出已启用（第 {sc} 列起追加 7 列）\n", "info")
                hs = read_headers_rw(proc.ws)
                if is_fwd:
                    a_col = self._pick_header(hs, ui["addr"], "地址1")
                    if not a_col:
                        log_cb("[错误] 未能识别「地址1列」，请在下拉框中明确选择后重试，处理已终止\n", "error")
                        return
                    aidx = self._col_idx(hs, a_col, Config.COLS["addr1"])
                    # 修复：地址2 列本来就是选填（模板标注「地址2(选填)」、底层逻辑也支持为空），
                    # 原实现未选地址2列时直接终止整批处理，且会回退到固定第 6 列读到无关数据。
                    a2_col = self._pick_header(hs, ui["addr2"], "地址2")
                    a2idx = self._col_idx(hs, a2_col, 0) if a2_col else 0
                    if not a2idx:
                        log_cb("[提示] 未指定「地址2列」，本次仅使用地址1\n", "info")
                    cidx = self._col_idx(hs, self._pick_header(hs, ui["city"], "市", "城市", "city"), Config.COLS["city"])
                    didx = self._col_idx(hs, self._pick_header(hs, ui["dist"], "区县", "县区", "区/县", "county", "district", "县", "区"), Config.COLS["district"])
                    lr = proc.last_data_row("fwd", aidx)
                else:
                    lng_col = ui["lng"]
                    lat_col = ui["lat"]
                    if not lng_col or not lat_col:
                        log_cb("[错误] 请先选择 GCJ-02 经度列和纬度列，处理已终止\n", "error")
                        return
                    glidx = self._col_idx(hs, lng_col, Config.COLS["gcj_lng"])
                    gtidx = self._col_idx(hs, lat_col, Config.COLS["gcj_lat"])
                    # 校验列名是否真实存在于表头（防止回退到默认列导致读错数据）
                    if lng_col not in hs or lat_col not in hs:
                        log_cb(f"[错误] 所选列不存在于Excel表头，请重新选择（经度:{lng_col}, 纬度:{lat_col}）\n", "error")
                        return
                    wlidx = self._col_idx(hs, ui["wlng"], 0) if ui["wlng"] else 0
                    wtidx = self._col_idx(hs, ui["wlat"], 0) if ui["wlat"] else 0
                    lr = max(proc.last_row(glidx), proc.last_row(gtidx), proc.last_row(wlidx) if wlidx>0 else 1, proc.last_row(wtidx) if wtidx>0 else 1)
                    # 修复：原兜底回退 ws.max_row（表尾格式行把行号撑到几万导致卡死），
                    # 改为按五列扫描有效末行；确实无数据时保持 1，走下方「没有需要处理的数据」提示
                    if lr < 2: lr = proc.last_data_row("rev", 0)
                total = len(only_rows) if (is_fwd and only_rows is not None) else (lr - sr + 1)
                if total <= 0: log_cb("提示：没有需要处理的数据\n", "info"); return
                log_cb("[系统] 正在验证 API 密钥...\n", "info")
                ok, msg = self.bp.init(self.cfg)
                if not ok:
                    log_cb(f"[密钥错误] {msg}\n", "error")
                    return
                log_cb("[系统] API 密钥验证通过\n", "info")
                self._ui(lambda t=total: getattr(self, f"{direction}_prog").config(maximum=t, value=0))
                # 增量保存 + 任务状态（断点续跑）
                def _save_cb():
                    proc.save()
                def _state_cb(s):
                    s["start_row"] = sr
                    s["last_row"] = lr
                    save_task_state(path, s)
                    log_cb(f"[进度保存] 已处理 {s.get('done')}/{s.get('total')} 行（最后第 {s.get('last_done')} 行）\n", "info")
                if is_fwd:
                    do_clean = ui["clean"]
                    do_merge = ui["merge"]
                    do_dedup = ui["dedup"]
                    st = self.bp.forward(proc, sr, lr, aidx, a2idx, cidx, didx,
                                         ui["skip"],
                                         ui["abnormal"],
                                         do_clean, do_merge, do_dedup,
                                         log_cb, prog_cb,
                                         workers=ui["workers"], save_cb=_save_cb, state_cb=_state_cb,
                                         write_struct=ui["struct"], only_rows=only_rows,
                                         save_every=ui["save_every"])
                    self._ui(lambda s=st, w=ui["workers"]: self._log(direction, f"\n{'='*45}\n处理完成统计\n{'='*45}\n总行数：{s['total']}\n成功：{s['success']}\n使用地址2：{s['addr2_used']}\n百度优化：{s['baidu_up']}\n疑似漂移：{s['drift_cnt']}\n失败：{s['fail']}\n跳过：{s['skip']}\n缓存命中：{s.get('cache_hits', 0)}\n高德调用：{s['gaode_count']}\n百度调用：{s['baidu_count']}\n并发数：{w}\n{'='*45}\n", "info"))
                else:
                    st = self.bp.reverse(proc, sr, lr, glidx, gtidx, wlidx, wtidx, ui["skip"], log_cb, prog_cb,
                                         save_cb=_save_cb, state_cb=_state_cb, only_rows=only_rows,
                                         save_every=ui["save_every"])
                    if st.get("global_error"): self._ui(lambda: self._log(direction, "[全局错误] 已终止处理！\n", "error"))
                    else: self._ui(lambda s=st: self._log(direction, f"\n{'='*45}\n处理完成统计\n{'='*45}\n总行数：{s['total']}\n成功：{s['success']}\n失败：{s['fail']}\n跳过：{s['skip']}\n缓存命中：{s.get('cache_hits', 0)}\n高德调用：{s['gaode_count']}\n{'='*45}\n", "info"))
                # 修复：原文件被 Excel 占用/只读时 save() 会抛异常，整批结果直接丢失；
                # 改为失败时另存为副本，并在日志中明确提示实际保存位置。
                saved_path = path
                try:
                    proc.save()
                    if st.get("completed"):
                        clear_task_state(path)
                    else:
                        log_cb("[提示] 任务未完成，已记录断点：下次选择同一文件可续跑或只重试失败行\n", "warning")
                except Exception as se:
                    alt = os.path.join(os.path.dirname(path) or ".",
                                       f"{os.path.splitext(os.path.basename(path))[0]}_结果_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
                    try:
                        proc.fp = alt
                        proc.save()
                        saved_path = alt
                        log_cb(f"[警告] 原文件无法写入（{se}），结果已另存为：{alt}\n", "warning")
                    except Exception as se2:
                        log_cb(f"[错误] 结果保存失败：{se2}\n", "error")
                if ao and os.path.exists(saved_path): self._open(saved_path)
            except Exception as e:
                try:
                    self._ui(lambda: (getattr(self, f"{direction}_prog").stop(),
                                      getattr(self, f"{direction}_prog").config(mode="determinate")))
                except Exception: pass
                self._ui(lambda e=e: self._log(direction, f"[处理异常] {str(e)}\n", "error"))
            finally:
                self.running = False
                try:
                    if "proc" in locals() and proc:
                        proc.close()
                except: pass
                self._ui(lambda: (getattr(self, f"{direction}_start_btn").config(state=tk.NORMAL), getattr(self, f"{direction}_stop_btn").config(state=tk.DISABLED), self.status.set("就绪")))
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
            hs = read_headers(ws)
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

    def _stop_cv_batch(self):
        self.cv_stop_flag = True          # 独立标志：不再与批量编码共用 stop_flag
        self.status.set("正在停止...")

    def _start_cv_batch(self):
        if self._check_expired(): return
        if self.running or self.cv_running:
            messagebox.showwarning("提示", "已有批量任务正在运行，请等待完成或先点击「停止」")
            return
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
        self.cv_stop_flag = False; self.cv_running = True
        self.cv_sbtn.config(state=tk.DISABLED); self.cv_xbtn.config(state=tk.NORMAL); self.cv_log.delete("1.0", tk.END)
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
                # 修复：Excel 表尾常有大量无数据格式行，ws.max_row 会被撑到几万行；
                # 原逻辑按 max_row 遍历且每行都调度 UI 更新，导致导入/转换卡死。
                # 改为向上扫描经度/纬度列，取最后有数据的行作为处理末行。
                _mr = ws.max_row; _last = 2
                for _r in range(_mr, 1, -1):
                    if ws.cell(row=_r, column=li).value is not None or ws.cell(row=_r, column=ti).value is not None:
                        _last = _r; break
                lr = _last; self.root.after(0, lambda: self.cv_prog.config(maximum=lr-1, value=0))
                sc = fc = 0
                for r in range(2, lr + 1):
                    if self.cv_stop_flag: self._ui(lambda: self.cv_log.insert(tk.END, "\n用户取消\n")); break
                    pv = r - 1
                    self._ui(lambda v=pv, t=lr-1: (self.cv_plbl.config(text=f"{v} / {t}"), self.cv_prog.config(value=v)))
                    try:
                        lv = ws.cell(row=r, column=li).value; tv = ws.cell(row=r, column=ti).value
                        if lv is None or tv is None: continue
                        ol, ot = convert_coords(float(lv), float(tv), ss, ds)
                        ws.cell(row=r, column=oli, value=round(ol, 6)); ws.cell(row=r, column=oti, value=round(ot, 6)); sc += 1
                    except Exception as e: fc += 1; ws.cell(row=r, column=oli, value=f"错误:{str(e)[:20]}"); ws.cell(row=r, column=oti, value=f"错误:{str(e)[:20]}")
                # 修复：原文件被 Excel 占用/只读时 save() 失败会丢掉全部转换结果，改为另存为副本
                saved = path
                try:
                    wb.save(path)
                except Exception as se:
                    alt = os.path.join(os.path.dirname(path) or ".",
                                       f"{os.path.splitext(os.path.basename(path))[0]}_转换_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
                    try:
                        wb.save(alt); saved = alt
                        self._ui(lambda s=se, a=alt: self.cv_log.insert(tk.END, f"\n[警告] 原文件无法写入（{s}），已另存为：{a}\n"))
                    except Exception as se2:
                        self._ui(lambda s=se2: self.cv_log.insert(tk.END, f"\n[错误] 结果保存失败：{s}\n"))
                self._ui(lambda sp=saved: self.cv_log.insert(tk.END, f"\n{'='*45}\n转换完成统计\n{'='*45}\n总行数：{lr-1}\n成功：{sc}\n失败：{fc}\n输出列：{olc}, {olt}\n文件已保存：{sp}\n{'='*45}\n"))
            except Exception as e: self._ui(lambda e=e: self.cv_log.insert(tk.END, f"[处理异常] {str(e)}\n"))
            finally:
                self.cv_running = False
                self._ui(lambda: (self.cv_sbtn.config(state=tk.NORMAL), self.cv_xbtn.config(state=tk.DISABLED), self.status.set("就绪")))
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
        # 赞助作者按钮：标题行最右侧，扁平融入页面
        try:
            show_sponsor_button(title_row, bg=Config.C_BG, activebackground=Config.C_BG).pack(side=tk.RIGHT)
        except Exception:
            pass

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
        # v2.8.57：AI 智能延展开关（配置AI后主动展开细分词，默认开启；关闭时退回仅接近上限才扩展）
        self.poi_ai_expand = tk.BooleanVar(value=bool(self.cfg.get("ai_expand_active", False)))
        tk.Checkbutton(of, text="AI 智能延展（配置AI后自动展开细分词）", variable=self.poi_ai_expand,
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
        cols = ["序号", "名称", "地址", "电话", "GCJ-02经度", "GCJ-02纬度", "WGS-84经度", "WGS-84纬度", "分类", "数据来源", "AI判定"]
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
        self.poi_tree.column("AI判定", width=130, anchor="w"); self.poi_tree.heading("AI判定", text="AI判定")
        self.poi_tree.insert("", tk.END, values=("", "", "暂无数据", "点击「开始采集」后结果将显示在这里", "", "", "", "", "", "", ""))
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

    def _poi_search_page(self, mode, kwq, unit, clng, clat, radius, page, gkey, bak):
        # v2.8.56：POI 单页检索（优先高德，日配额耗尽自动切百度）
        # 返回 (items, provider, flag)；flag: ''正常 / 'gd_quota' / 'bd_quota' / 'err' / 'empty'
        provider = self._poi_provider
        if provider == "baidu":
            if not bak:
                return [], "baidu", "err"
            if mode == "district":
                url = f"{Config.BAIDU_POI}?query={url_enc(kwq)}&region={url_enc(unit)}&output=json&ak={bak}&page_size=10&page_num={max(0, page-1)}&scope=2"
            else:
                bd_lng, bd_lat = gcj02_to_bd09(clng, clat)
                url = f"{Config.BAIDU_POI}?query={url_enc(kwq)}&location={bd_lat},{bd_lng}&radius={radius}&output=json&ak={bak}&page_size=10&page_num={max(0, page-1)}&scope=2"
            ok, resp = _http_retry(url)
            if not ok:
                return [], "baidu", "empty"
            try: d = json.loads(resp)
            except: return [], "baidu", "err"
            if d.get("status") != 0:
                if _is_quota_error(str(d.get("status"))) or _is_quota_error(d.get("message", "")):
                    return [], "baidu", "bd_quota"
                return [], "baidu", "err"
            results = d.get("results", [])
            items = []
            for r in results:
                loc = r.get("location", {}) or {}
                try: blng, blat = float(loc.get("lng", 0)), float(loc.get("lat", 0))
                except: continue
                glng, glat = bd09_to_gcj02(blng, blat)
                items.append({"name": r.get("name", ""), "address": r.get("address", ""),
                              "tel": r.get("telephone", ""), "type": (r.get("detail_info", {}) or {}).get("type", ""),
                              "gcj_lng": glng, "gcj_lat": glat, "source": "百度"})
            return items, "baidu", ("" if items else "empty")
        # 默认高德
        if mode == "district":
            url = f"{Config.GAODE_POI_TEXT}?key={gkey}&keywords={url_enc(kwq)}&city={url_enc(unit)}&offset=25&page={page}&extensions=all&output=json"
        else:
            url = f"{Config.GAODE_POI_AROUND}?key={gkey}&location={clng},{clat}&keywords={url_enc(kwq)}&radius={radius}&offset=25&page={page}&extensions=all&output=json"
        ok, resp = _http_retry(url)
        if not ok:
            return [], "gaode", "empty"
        try: d = json.loads(resp)
        except: return [], "gaode", "err"
        if d.get("status") != "1":
            info = d.get("info", "")
            if _is_quota_error(info):
                return [], "gaode", "gd_quota"
            return [], "gaode", "err"
        pois = d.get("pois", [])
        items = []
        for p in pois:
            loc = p.get("location", "")
            lng, lat = 0, 0
            if loc and "," in loc:
                try: lng, lat = map(float, loc.split(","))
                except: pass
            items.append({"name": p.get("name", ""), "address": p.get("address", ""),
                          "tel": p.get("tel", ""), "type": p.get("type", ""),
                          "gcj_lng": lng, "gcj_lat": lat, "source": "高德"})
        return items, "gaode", ("" if items else "empty")

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
        # 修复：在主线程把界面参数一次性快照给后台线程，
        # 避免后台线程读 Tk 变量（非线程安全），也避免采集中途用户改选项导致行为不一致
        self._poi_opts = {
            "mode": self.poi_mode.get(),
            "city": self.poi_city.get().strip(),
            "kw_dist": self.poi_kw_dist.get().strip(),
            "center": self.poi_center.get().strip(),
            "kw_around": self.poi_kw_around.get().strip(),
            "radius": self.poi_radius.get(),
            "break_limit": self.poi_break_limit.get(),
            "ai_expand": self.poi_ai_expand.get(),
            "supplement_tel": self.poi_supplement_tel.get(),
        }
        threading.Thread(target=self._poi_task, daemon=True).start()

    def _stop_poi(self):
        self.poi_stop_flag = True; self.poi_log.insert(tk.END, "\n⚠ 用户请求停止...\n", "warning")

    def _poi_task(self):
        opt = getattr(self, "_poi_opts", {}) or {}
        brk = bool(opt.get("break_limit", False))
        ai_exp = bool(opt.get("ai_expand", False))
        sup_tel = bool(opt.get("supplement_tel", False))
        try:
            mode = opt.get("mode", "district")
            gkey = self.cfg.get("gaode_key", "")
            bak = self.cfg.get("baidu_ak", "")
            ai = AiClient(self.cfg)   # L0 智能增强（未配置时 available=False，全程降级为规则模式）
            geo = GeoCoder(self.cfg)
            self._poi_provider = "gaode"   # v2.8.56：当前检索平台（高德优先，日配额耗尽自动切百度）
            ok, msg = geo.check()
            if not ok:
                # v2.8.56：高德验证失败（如日配额耗尽）时，若百度可用则自动切百度继续采集
                if bak and geo.use_baidu and geo.bok:
                    self._poi_provider = "baidu"
                    self._ui(lambda m=msg: self.poi_log.insert(tk.END, f"[高德] 密钥验证失败({m})，已配置百度 AK，自动切换百度检索（优先高德）\n", "warning"))
                else:
                    self._ui(lambda m=msg: self.poi_log.insert(tk.END, f"[错误] 密钥验证失败: {m}\n", "error"))
                    return
            gaode_count = baidu_count = 0
            raw_results = []
            poi_idx = PoiIndex()
            clng = clat = 0.0
            kw = city = radius = ""
            if mode == "district":
                city = opt.get("city", "")
                kw = opt.get("kw_dist", "")
                if not city or not kw:
                    self.root.after(0, lambda: self.poi_log.insert(tk.END, "[错误] 城市和关键词不能为空\n", "error"))
                    return
                # 修复：区县下钻 —— 高德同请求参数翻页最多200条，且市级检索仅覆盖市辖区；
                # 通过行政区划查询获取下属区县，逐区县检索并合并，突破数量上限
                unit_stats = []   # L1: 逐区县统计（供采集后 AI 分析）
                units = get_poi_district_info(gkey, city)
                if len(units) > 1:
                    self.root.after(0, lambda u=units: self.poi_log.insert(tk.END, f"[高德] 已下钻 {len(u)} 个区县逐区检索: {'、'.join(x['name'] for x in u)}\n", "info"))
                for unit_info in units:
                    if self.poi_stop_flag: break
                    unit = unit_info["name"]
                    self.root.after(0, lambda un=unit: self.poi_log.insert(tk.END, f"[高德] 行政区划检索: {un} + {kw}\n", "info"))
                    unit_raw = 0
                    truncated = False
                    page = 1
                    while not self.poi_stop_flag and page <= 40:
                        items, used, flag = self._poi_search_page("district", kw, unit, 0, 0, 0, page, gkey, bak)
                        if flag == "gd_quota" and used == "gaode":
                            if bak:
                                self._poi_provider = "baidu"
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 日配额已耗尽，自动切换百度检索（优先高德）\n", "warning"))
                                continue
                            self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 配额耗尽（每日调用上限），且未配置百度 AK，采集提前结束\n", "error"))
                            self.poi_stop_flag = True
                            break
                        if flag == "bd_quota":
                            self.root.after(0, lambda: self.poi_log.insert(tk.END, "[百度] 配额耗尽（每日调用上限），本次采集提前结束\n", "error"))
                            self.poi_stop_flag = True
                            break
                        if flag == "err":
                            self.root.after(0, lambda u=used: self.poi_log.insert(tk.END, f"[{'高德' if u=='gaode' else '百度'}] 检索错误，跳过当前检索\n", "error"))
                            break
                        if flag == "empty":
                            self.root.after(0, lambda p=page, u=used: self.poi_log.insert(tk.END, f"[{'高德' if u=='gaode' else '百度'}] 第{p}页无数据，采集结束\n", "info"))
                            break
                        if used == "gaode": gaode_count += 1
                        else: baidu_count += 1
                        for p in items:
                            if poi_idx.near_any(p["gcj_lng"], p["gcj_lat"]):
                                continue
                            poi_idx.add(p)
                            raw_results.append(p)
                        unit_raw += len(items)
                        self.root.after(0, lambda p=page, c=len(items), u=used: self.poi_log.insert(tk.END, f"[{'高德' if u=='gaode' else '百度'}] 第{p}页: {c}条\n", "info"))
                        if len(items) < (25 if used == "gaode" else 10):
                            break
                        if page >= 9:
                            truncated = True
                        page += 1; time.sleep(0.3)
                    # 区县内数据可能被高德 225 条翻页上限截断：勾选突破限制时启用网格深挖
                    if truncated and brk:
                        if self._poi_provider == "baidu":
                            self.root.after(0, lambda un=unit: self.poi_log.insert(tk.END, f"[深挖] {un} 已切换百度检索，跳过网格深挖（百度无网格接口）\n", "info"))
                        else:
                            center = unit_info.get("center", "")
                            mclng = mclat = 0.0
                            if center and "," in center:
                                try: mclng, mclat = map(float, center.split(","))
                                except: pass
                            if mclng and mclat:
                                self.root.after(0, lambda un=unit, ur=unit_raw: self.poi_log.insert(tk.END, f"[深挖] {un} 已取{ur}条接近上限，启用网格扫描补全...\n", "info"))
                                poly = get_district_polygon(gkey, unit_info.get("adcode", ""))
                                grid_radius = _calc_grid_radius(poly) or 12000
                                grid_pois, grid_calls = grid_collect_poi(gkey, mclng, mclat, kw, grid_radius, lambda: self.poi_stop_flag,
                                                        log_cb=lambda m: self.root.after(0, lambda mm=m: self.poi_log.insert(tk.END, mm + "\n", "info")), poly=poly, known=raw_results, workers=self.cfg.get("grid_workers", 2))
                                gaode_count += grid_calls
                                added = 0
                                for gp in grid_pois:
                                    if not poi_idx.near_any(gp["gcj_lng"], gp["gcj_lat"]):
                                        poi_idx.add(gp)
                                        raw_results.append(gp)
                                        added += 1
                                self.root.after(0, lambda un=unit, gc=grid_calls, ad=added: self.poi_log.insert(tk.END, f"[深挖] {un} 网格扫描完成：{gc}次调用，新增{ad}条\n", "info"))
                        unit_stats.append({"name": unit, "count": unit_raw, "hit_limit": truncated})
            else:
                center = opt.get("center", "")
                kw = opt.get("kw_around", "")
                radius = opt.get("radius", "")
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
                raw_now = 0
                truncated = False
                page = 1
                while not self.poi_stop_flag and page <= 40:
                    items, used, flag = self._poi_search_page("around", kw, "", clng, clat, radius, page, gkey, bak)
                    if flag == "gd_quota" and used == "gaode":
                        if bak:
                            self._poi_provider = "baidu"
                            self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 日配额已耗尽，自动切换百度检索（优先高德）\n", "warning"))
                            continue
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 配额耗尽（每日调用上限），且未配置百度 AK，采集提前结束\n", "error"))
                        self.poi_stop_flag = True
                        break
                    if flag == "bd_quota":
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, "[百度] 配额耗尽（每日调用上限），本次采集提前结束\n", "error"))
                        self.poi_stop_flag = True
                        break
                    if flag == "err":
                        self.root.after(0, lambda u=used: self.poi_log.insert(tk.END, f"[{'高德' if u=='gaode' else '百度'}] 检索错误，跳过当前检索\n", "error"))
                        break
                    if flag == "empty":
                        self.root.after(0, lambda p=page, u=used: self.poi_log.insert(tk.END, f"[{'高德' if u=='gaode' else '百度'}] 第{p}页无数据，采集结束\n", "info"))
                        break
                    if used == "gaode": gaode_count += 1
                    else: baidu_count += 1
                    for p in items:
                        if poi_idx.near_any(p["gcj_lng"], p["gcj_lat"]):
                            continue
                        poi_idx.add(p)
                        raw_results.append(p)
                    raw_now += len(items)
                    self.root.after(0, lambda p=page, c=len(items), u=used: self.poi_log.insert(tk.END, f"[{'高德' if u=='gaode' else '百度'}] 第{p}页: {c}条\n", "info"))
                    if len(items) < (25 if used == "gaode" else 10):
                        break
                    if page >= 9:
                        truncated = True
                    page += 1; time.sleep(0.3)
                # 周边检索同样受高德 225 条翻页上限截断：勾选突破限制时以中心坐标网格深挖补全
                if truncated and brk and self._poi_provider == "gaode":
                    self.root.after(0, lambda ur=raw_now: self.poi_log.insert(tk.END, f"[深挖] 周边已取{ur}条接近上限，启用网格扫描补全...\n", "info"))
                    grid_pois, grid_calls = grid_collect_poi(gkey, clng, clat, kw, radius, lambda: self.poi_stop_flag,
                                                                   log_cb=lambda m: self.root.after(0, lambda mm=m: self.poi_log.insert(tk.END, mm + "\n", "info")), known=raw_results, workers=self.cfg.get("grid_workers", 2))
                    gaode_count += grid_calls
                    added = 0
                    for gp in grid_pois:
                        if not poi_idx.near_any(gp["gcj_lng"], gp["gcj_lat"]):
                            poi_idx.add(gp)
                            raw_results.append(gp)
                            added += 1
                    self.root.after(0, lambda gc=grid_calls, ad=added: self.poi_log.insert(tk.END, f"[深挖] 网格扫描完成：{gc}次调用，新增{ad}条\n", "info"))
            # ═══ 突破500条限制：关键词拆分检索 ═══
            expand_dup = 0   # 新增：初始化
            expand_block_done = False   # v2.8.57：标记"突破限制"拆分检索是否已执行（已执行则不再重复AI延展）
            if brk and len(raw_results) >= 480:
                expand_block_done = True
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[突破限制] 高德已返回{len(raw_results)}条（接近上限），启动关键词拆分检索...\n", "info"))
                main_kw = kw
                # 静态映射兜底（AI 未配置/失败时使用；与 AI 结果取并集补全细分词）
                static_kws = _static_expand_kws(main_kw)
                # AI 智能关键词扩展优先；v2.8.49：与静态映射取并集，避免遗漏静态词库细分词
                ai_kws = ai.expand_keywords(main_kw) if ai.available else None
                if ai_kws:
                    self.root.after(0, lambda ks=list(ai_kws): self.poi_log.insert(tk.END, "[AI增强] 智能关键词扩展(含静态兜底): " + "、".join(ks) + "\n", "info"))
                expand_kws = _merge_expand_kws(ai_kws, main_kw)
                for ek in expand_kws:
                    if self.poi_stop_flag: break
                    self.root.after(0, lambda k=ek: self.poi_log.insert(tk.END, f"[突破限制] 扩展关键词: {k}\n", "info"))
                    for unit_info in (units if mode == "district" else [{"name": ""}]):
                        if self.poi_stop_flag: break
                        unit = unit_info["name"]
                        eraw = 0
                        etruncated = False
                        ep = 1
                        while not self.poi_stop_flag and ep <= 40:  # 每个扩展词+区县最多40页（1000条）
                            emode = "district" if mode == "district" else "around"
                            eitems, eused, eflag = self._poi_search_page(emode, ek, unit, clng, clat, radius, ep, gkey, bak)
                            if eflag == "gd_quota" and eused == "gaode":
                                if bak:
                                    self._poi_provider = "baidu"
                                    self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 日配额已耗尽，自动切换百度检索（优先高德）\n", "warning"))
                                    continue
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 配额耗尽（每日调用上限），且未配置百度 AK，采集提前结束\n", "error"))
                                self.poi_stop_flag = True
                                break
                            if eflag == "bd_quota":
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, "[百度] 配额耗尽（每日调用上限），本次采集提前结束\n", "error"))
                                self.poi_stop_flag = True
                                break
                            if eflag in ("err", "empty"):
                                self.root.after(0, lambda u=eused, k=ek: self.poi_log.insert(tk.END, f"[突破限制] 扩展词[{k}] {['高德','百度'][u=='baidu']} 第{ep}页无数据，结束该词检索\n", "info"))
                                break
                            if eused == "gaode": gaode_count += 1
                            else: baidu_count += 1
                            for p in eitems:
                                if poi_idx.near_any(p["gcj_lng"], p["gcj_lat"]):
                                    expand_dup += 1
                                else:
                                    p["source"] = p.get("source", "") + "(扩展)"
                                    poi_idx.add(p)
                                    raw_results.append(p)
                            eraw += len(eitems)
                            if len(eitems) < (25 if eused == "gaode" else 10):
                                break
                            if ep >= 9:
                                etruncated = True
                            ep += 1; time.sleep(0.3)
                        # 扩展词同样受高德 225 条翻页上限截断：勾选突破限制时网格深挖补全
                        if etruncated and brk and self._poi_provider == "gaode":
                            if mode == "district":
                                ecen = unit_info.get("center", "")
                                elng = elat = 0.0
                                if ecen and "," in ecen:
                                    try: elng, elat = map(float, ecen.split(","))
                                    except: pass
                                eradi = 12000
                            else:
                                elng, elat, eradi = clng, clat, radius
                            if elng and elat:
                                self.root.after(0, lambda k=ek, u=unit, er=eraw: self.poi_log.insert(tk.END, f"[深挖] 扩展词[{k}]+{u} 已取{er}条接近上限，网格扫描补全...\n", "info"))
                                epoly = get_district_polygon(gkey, unit_info.get("adcode", "")) if mode == "district" else None
                                e_radius = _calc_grid_radius(epoly) if epoly else (eradi or 12000)
                                epois_g, ecalls = grid_collect_poi(gkey, elng, elat, ek, e_radius, lambda: self.poi_stop_flag,
                                                   log_cb=lambda m: self.root.after(0, lambda mm=m: self.poi_log.insert(tk.END, mm + "\n", "info")), poly=epoly, known=raw_results, workers=self.cfg.get("grid_workers", 2))
                                gaode_count += ecalls
                                eadd = 0
                                for gp in epois_g:
                                    if not poi_idx.near_any(gp["gcj_lng"], gp["gcj_lat"]):
                                        poi_idx.add(gp)
                                        raw_results.append(gp)
                                        eadd += 1
                                self.root.after(0, lambda k=ek, u=unit, ec=ecalls, ea=eadd: self.poi_log.insert(tk.END, f"[深挖] 扩展词[{k}]+{u} 网格扫描完成：{ec}次调用，新增{ea}条\n", "info"))
                        if self.poi_stop_flag: break
                    if self.poi_stop_flag: break
                self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[突破限制] 关键词拆分检索完成，当前共{len(raw_results)}条（去重{expand_dup}条）\n", "info"))
            # ── v2.8.57 AI 智能延展（主动）：配置 AI 后无论采集量多少都展开细分词，
            #    增量即停（每词每区县最多 N 页，无新增即止）+ 总调用预算封顶，防配额失控 ──
            if (not expand_block_done and ai_exp and ai.available
                    and not self.poi_stop_flag):
                main_kw = kw
                ai_kws = ai.expand_keywords(main_kw)
                expand_kws = _merge_expand_kws(ai_kws, main_kw)
                if expand_kws:
                    ai_pages = max(1, int(self.cfg.get("ai_expand_pages", 2)))
                    ai_budget = max(5, int(self.cfg.get("ai_expand_max_calls", 80)))
                    self.root.after(0, lambda ks=list(expand_kws): self.poi_log.insert(tk.END, "[AI延展] 智能关键词延展(与静态词库取并集): " + "、".join(ks) + "\n", "info"))
                    xcalls = 0
                    for ek in expand_kws:
                        if self.poi_stop_flag or xcalls >= ai_budget: break
                        self.root.after(0, lambda k=ek: self.poi_log.insert(tk.END, f"[AI延展] 扩展关键词: {k}\n", "info"))
                        x_added = 0
                        for unit_info in (units if mode == "district" else [{"name": ""}]):
                            if self.poi_stop_flag or xcalls >= ai_budget: break
                            unit = unit_info["name"]
                            ep = 1
                            while not self.poi_stop_flag and ep <= ai_pages and xcalls < ai_budget:
                                eitems, eused, eflag = self._poi_search_page("district" if mode == "district" else "around", ek, unit, clng, clat, radius, ep, gkey, bak)
                                xcalls += 1
                                if eflag == "gd_quota" and eused == "gaode":
                                    if bak:
                                        self._poi_provider = "baidu"
                                        self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 日配额已耗尽，自动切换百度检索（优先高德）\n", "warning"))
                                        continue
                                    self.root.after(0, lambda: self.poi_log.insert(tk.END, "[高德] 配额耗尽（每日调用上限），且未配置百度 AK，采集提前结束\n", "error"))
                                    self.poi_stop_flag = True
                                    break
                                if eflag == "bd_quota":
                                    self.root.after(0, lambda: self.poi_log.insert(tk.END, "[百度] 配额耗尽（每日调用上限），本次采集提前结束\n", "error"))
                                    self.poi_stop_flag = True
                                    break
                                if eflag in ("err", "empty"):
                                    self.root.after(0, lambda u=eused, k=ek: self.poi_log.insert(tk.END, f"[AI延展] 词[{k}] {['高德','百度'][u=='baidu']} 第{ep}页无数据，结束该词检索\n", "info"))
                                    break
                                if eused == "gaode": gaode_count += 1
                                else: baidu_count += 1
                                page_added = 0
                                for p in eitems:
                                    if poi_idx.near_any(p["gcj_lng"], p["gcj_lat"]):
                                        expand_dup += 1
                                    else:
                                        p["source"] = p.get("source", "") + "(AI延展)"
                                        poi_idx.add(p)
                                        raw_results.append(p)
                                        page_added += 1
                                x_added += page_added
                                self.root.after(0, lambda k=ek, u=unit, p=ep, c=len(eitems), a=page_added: self.poi_log.insert(tk.END, f"[AI延展] 词[{k}]+{u} 第{p}页: {c}条 新增{a}条\n", "info"))
                                if page_added == 0: break          # 增量即停：本页无新增即止
                                if len(eitems) < (25 if eused == "gaode" else 10): break
                                ep += 1; time.sleep(0.3)
                        self.root.after(0, lambda k=ek, a=x_added, c=xcalls: self.poi_log.insert(tk.END, f"[AI延展] 词[{k}] 完成：新增{a}条（累计{c}次调用）\n", "info"))
                    self.root.after(0, lambda c=len(raw_results), d=expand_dup, x=xcalls: self.poi_log.insert(tk.END, f"[AI延展] 智能延展完成：当前共{c}条（去重{d}条，延展调用{x}次）\n", "info"))

            # 百度补缺逻辑
            need_baidu = False
            if bak and self.cfg.get("enable_baidu_fallback", True):
                if len(raw_results) < 20:
                    need_baidu = True
                    self.root.after(0, lambda c=len(raw_results): self.poi_log.insert(tk.END, f"[百度Fallback] 高德结果仅{c}条，触发百度补缺\n", "baidu"))
                elif sup_tel:
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
                    ok, resp = _http_retry(url)
                    if not ok:
                        self.root.after(0, lambda: self.poi_log.insert(tk.END, f"[百度] 请求失败: {resp[:80]}\n", "error"))
                        break
                    baidu_count += 1
                    try:
                        d = json.loads(resp)
                        if d.get("status") != 0:
                            if _is_quota_error(str(d.get("status"))) or _is_quota_error(d.get("message", "")):
                                self.root.after(0, lambda: self.poi_log.insert(tk.END, "[百度] 配额耗尽（每日调用上限），停止百度补缺\n", "error"))
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
                        self._ui(lambda m=str(e): self.poi_log.insert(tk.END, f"[百度] 解析异常: {m}\n", "error"))
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
                            if "百度补电话" not in mr.get("source", ""):
                                mr["source"] = mr.get("source", "") + "+百度补电话"
                        dup_count += 1
                        break
                if not is_dup:
                    merged.append(br)
            # AI 语义去重标注（未配置 AI 时自动跳过，行为与原有逻辑完全一致）
            ai_semantic_dedup(ai, merged, log_cb=lambda m, **k: self.root.after(0, lambda mm=m: self.poi_log.insert(tk.END, mm + "\n", "info")))
            self.poi_results = merged
            self.root.after(0, lambda: self.poi_log.insert(tk.END, f"\n{'='*45}\n采集完成\n{'='*45}\n高德: {len(raw_results)}条 | 百度: {len(baidu_results)}条\n去重: {dup_count}条 | 合并: {len(merged)}条\n总调用: 高德{gaode_count}次 百度{baidu_count}次\n{'='*45}\n", "success"))
            self.root.after(0, self._refresh_poi_tree)
            # ── L1 采集智能分析：后台线程（单次 AI 分析 + 受控补扫；未配置 AI 时自动跳过，行为不变）──
            if ai.available and mode == "district" and not self.poi_stop_flag:
                l1_stats = list(unit_stats)
                l1_merged = merged
                l1_idx = poi_idx
                def l1_task():
                    if self.poi_stop_flag:
                        return
                    extra = ai_l1_postprocess(
                        ai, city, kw, l1_stats, len(l1_merged), l1_merged, l1_idx, gkey,
                        log_cb=lambda m, **k: self.root.after(0, lambda mm=m: self.poi_log.insert(tk.END, mm + "\n", "info")),
                        stop_cb=lambda: self.poi_stop_flag)
                    if extra:
                        self.root.after(0, lambda n=len(extra): self.poi_log.insert(tk.END, f"[AI补扫] 本次采集新增 {n} 条（共{len(l1_merged)}条）\n", "info"))
                        self.root.after(0, self._refresh_poi_tree)
                threading.Thread(target=l1_task, daemon=True).start()
            self._ui(lambda: (self.poi_excel_btn.config(state=tk.NORMAL), self.poi_kml_btn.config(state=tk.NORMAL)))
        except Exception as e:
            self._ui(lambda m=str(e): self.poi_log.insert(tk.END, f"[系统异常] {m}\n", "error"))
        finally:
            self.poi_running = False
            self._ui(lambda: (self.poi_start_btn.config(state=tk.NORMAL), self.poi_stop_btn.config(state=tk.DISABLED)))

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
                r.get("type",""), r.get("source","高德"), r.get("ai_flag","")
            ))

    def _export_poi_excel(self):
        if not self.poi_results:
            messagebox.showwarning("提示", "暂无采集结果"); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                            initialfile=f"POI采集_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        if not path: return
        results = list(self.poi_results)   # 主线程快照，避免后台线程读共享列表
        def task():
            try:
                ox, pf, ft, al, gcl = _get_openpyxl()
                wb = ox.Workbook(); ws = wb.active; ws.title = "Sheet1"
                headers = ["序号", "名称", "地址", "电话", "GCJ-02经度", "GCJ-02纬度", "WGS-84经度", "WGS-84纬度", "分类", "数据来源", "AI判定"]
                for col, h in enumerate(headers, 1):
                    c = ws.cell(row=1, column=col, value=h); c.font = ft(bold=True, color="FFFFFF")
                    c.fill = pf(start_color="1e293b", end_color="1e293b", fill_type="solid"); c.alignment = al(horizontal="center", vertical="center")
                for i, r in enumerate(results, 2):
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
                    ws.cell(row=i, column=11, value=_s(r.get("ai_flag","")))
                for i, w in enumerate([8, 25, 35, 18, 15, 15, 15, 15, 25, 12, 24], 1):
                    ws.column_dimensions[gcl(i)].width = w
                wb.save(path)
                self._ui(lambda n=len(results): messagebox.showinfo("导出成功", f"已导出 {n} 条到:\n{path}"))
            except Exception as e:
                self._ui(lambda m=str(e): messagebox.showerror("错误", f"导出失败: {m}"))
        threading.Thread(target=task, daemon=True).start()

    def _export_poi_kml(self):
        if not self.poi_results:
            messagebox.showwarning("提示", "暂无采集结果"); return
        path = filedialog.asksaveasfilename(defaultextension=".kml", filetypes=[("KML文件", "*.kml")],
                                            initialfile=f"POI采集_{time.strftime('%Y%m%d_%H%M%S')}.kml")
        if not path: return
        results = list(self.poi_results)
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
                n = 0
                for r in results:
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
                    n += 1
                kml_lines += ['  </Document>', '</kml>']
                with open(path, "w", encoding="utf-8") as f:
                    f.write("\n".join(kml_lines))
                self._ui(lambda n=n: messagebox.showinfo("导出成功", f"已导出 {n} 个坐标点到 KML:\n{path}"))
            except Exception as e:
                self._ui(lambda m=str(e): messagebox.showerror("错误", f"导出失败: {m}"))
        threading.Thread(target=task, daemon=True).start()


    def _build_settings(self):
        f = self._page(); self.pages["settings"] = f
        tk.Label(f, text="API 密钥配置", font=("微软雅黑", 20, "bold"), bg=Config.C_BG, fg=Config.C_TEXT).pack(anchor="w", pady=(0, 14))
        # 两栏等宽等高 grid：同一网格两行两列，卡片按行等高 -> 左右底部线天然齐平
        cols = tk.Frame(f, bg=Config.C_BG); cols.pack(fill=tk.X, pady=0)
        cols.columnconfigure(0, weight=1, uniform="col")
        cols.columnconfigure(1, weight=1, uniform="col")
        LBL = dict(font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, width=13, anchor="e")

        # ─────────── 左栏：密钥设置（高德 / 百度） ───────────
        card = Card(cols, title="密钥设置（高德 / 百度）", accent=Config.C_ACCENT)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        form = tk.Frame(card, bg=Config.C_CARD); form.pack(fill=tk.X, padx=16, pady=6)
        def _row(parent):
            r = tk.Frame(parent, bg=Config.C_CARD); r.pack(fill=tk.X, pady=7); return r
        # 高德 / 百度 Key（标签宽度统一右对齐，输入框自适应）
        for lbl, attr, val in [("* 高德 API Key：", "cfg_gaode", self.cfg.get("gaode_key","")), ("  百度 AK：", "cfg_baidu", self.cfg.get("baidu_ak",""))]:
            r = _row(form)
            tk.Label(r, text=lbl, **LBL).pack(side=tk.LEFT)
            e = tk.Entry(r, font=("微软雅黑", 11), show="*", relief=tk.SOLID, bd=1)
            e.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True, ipady=2)
            e.insert(0, val); setattr(self, attr, e)
            tk.Button(r, text="显示/隐藏", font=("微软雅黑", 10), bg=Config.C_BORDER, fg=Config.C_TEXT, bd=0, padx=10, pady=4,
                      command=lambda ent=e: ent.config(show="" if ent.cget("show")=="*" else "*")).pack(side=tk.LEFT)
        opt = _row(form)
        self.cfg_eb = tk.BooleanVar(value=self.cfg.get("enable_baidu_fallback", True))
        tk.Checkbutton(opt, text="启用百度地图 Fallback（高德精度不足时自动切换）", variable=self.cfg_eb,
                       font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", padx=(10, 0))
        # 触发等级 + 请求间隔：一行
        r3 = _row(form)
        tk.Label(r3, text="百度触发等级：", **LBL).pack(side=tk.LEFT)
        self.cfg_trig = ttk.Combobox(r3, values=["门址","街道","乡镇","区县","城市"], font=("微软雅黑", 10), width=11, state="readonly")
        self.cfg_trig.set(self.cfg.get("baidu_trigger_level", "乡镇")); self.cfg_trig.pack(side=tk.LEFT, padx=10)
        tk.Label(r3, text="请求间隔（秒）：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT, padx=(24, 0))
        self.cfg_iv = tk.Spinbox(r3, from_=0.1, to=5.0, increment=0.1, width=7, font=("微软雅黑", 10), relief=tk.SOLID, bd=1)
        self.cfg_iv.delete(0, tk.END); self.cfg_iv.insert(0, str(self.cfg.get("base_interval", 0.2))); self.cfg_iv.pack(side=tk.LEFT, padx=10)
        # 深挖并发数 / 批量并发数 / 批量保存间隔：一行三组等宽（避免每行单个控件稀疏）
        r5 = _row(form)
        r5.columnconfigure(0, weight=1, uniform="conc")
        r5.columnconfigure(1, weight=1, uniform="conc")
        r5.columnconfigure(2, weight=1, uniform="conc")
        for _col, (_lbl, _wd, _attr, _dft, _lo, _hi, _inc) in enumerate([
                ("深挖并发数：", 4, "cfg_gw", int(self.cfg.get("grid_workers", 2)), 1, 5, 1),
                ("批量并发数：", 4, "cfg_fw", int(self.cfg.get("fwd_workers", 3)), 1, 5, 1),
                ("批量保存间隔：", 5, "cfg_save", int(self.cfg.get("save_every", 200)), 20, 5000, 20)]):
            _g = tk.Frame(r5, bg=Config.C_CARD); _g.grid(row=0, column=_col, sticky="w")
            tk.Label(_g, text=_lbl, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT)
            _sb = tk.Spinbox(_g, from_=_lo, to=_hi, increment=_inc, width=_wd, font=("微软雅黑", 10), relief=tk.SOLID, bd=1)
            _sb.delete(0, tk.END); _sb.insert(0, str(_dft)); _sb.pack(side=tk.LEFT, padx=(8, 0))
            setattr(self, _attr, _sb)
        r6 = _row(form)
        self.cfg_geo = tk.BooleanVar(value=bool(self.cfg.get("geo_drift_check", True)))
        tk.Checkbutton(r6, text="启用区县边界几何校验（坐标落界外判漂移）", variable=self.cfg_geo,
                       font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(side=tk.LEFT)
        self.cfg_adcode = tk.BooleanVar(value=bool(self.cfg.get("write_adcode", True)))
        tk.Checkbutton(r6, text="标准地址写入区县码", variable=self.cfg_adcode,
                       font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(side=tk.LEFT, padx=(16, 0))
        bf = tk.Frame(card, bg=Config.C_CARD); bf.pack(fill=tk.X, padx=16, pady=(4, 14))
        SButton(bf, "  保存配置", self._save_cfg, "primary").pack(side=tk.LEFT)
        SButton(bf, "  测试高德", lambda: self._test_api("gaode"), "success").pack(side=tk.LEFT, padx=10)
        SButton(bf, "  测试百度", lambda: self._test_api("baidu"), "success").pack(side=tk.LEFT)
        tip = Card(cols, title="使用说明", accent=Config.C_ACCENT)
        tip.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=0)
        for t in ["• 高德 Key 必填（console.amap.com 申请）；百度 AK 可选（lbsyun.baidu.com）",
                  "• 配置自动保存到 addr_config.json，下次启动自动加载",
                  "• 并发建议 1-3，过大易触发地图 API 限流",
                  "• Fallback 开启时，高德精度低于「乡镇」自动切百度",
                  "• 漂移标黄：正向编码结果区县与输入不一致时该行标黄",
                  "• 单条/批量结果可导出 Excel，坐标支持 GCJ-02 / WGS-84 互转"]:
            tk.Label(tip, text=t, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_M).pack(anchor="w", padx=16, pady=5)
        # ─────────── 右栏：AI 大模型智能增强（可选） ───────────
        aic = Card(cols, title="AI 大模型智能增强（可选）", accent=Config.C_INFO)
        aic.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 8))
        af = tk.Frame(aic, bg=Config.C_CARD); af.pack(fill=tk.X, padx=16, pady=6)
        self.cfg_ai_on = tk.BooleanVar(value=self.cfg.get("ai_enabled", False))
        tk.Checkbutton(af, text="启用 AI 增强（未启用或未填写 Key 时，自动使用原有规则模式）", variable=self.cfg_ai_on,
                       font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT, selectcolor=Config.C_CARD).pack(anchor="w", pady=4)
        ra = tk.Frame(af, bg=Config.C_CARD); ra.pack(fill=tk.X, pady=5)
        tk.Label(ra, text="模型平台：", **LBL).pack(side=tk.LEFT)
        self.cfg_ai_prov = ttk.Combobox(ra, values=list(Config.AI_PROVIDERS.keys()) + ["custom"], font=("微软雅黑", 10), width=12, state="readonly")
        self.cfg_ai_prov.set(self.cfg.get("ai_provider", "doubao")); self.cfg_ai_prov.pack(side=tk.LEFT, padx=10)
        tk.Label(ra, text="模型：", font=("微软雅黑", 11), bg=Config.C_CARD, fg=Config.C_TEXT).pack(side=tk.LEFT, padx=(20, 0))
        self.cfg_ai_model = tk.Entry(ra, font=("微软雅黑", 10), relief=tk.SOLID, bd=1)
        self.cfg_ai_model.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True, ipady=2)
        self.cfg_ai_model.insert(0, self.cfg.get("ai_model", "") or Config.AI_PROVIDERS.get(self.cfg.get("ai_provider", "doubao"), {}).get("model", ""))
        def _on_ai_prov(*_a):
            pv = self.cfg_ai_prov.get()
            if pv in Config.AI_PROVIDERS:
                self.cfg_ai_base.delete(0, tk.END); self.cfg_ai_base.insert(0, Config.AI_PROVIDERS[pv]["base"])
                self.cfg_ai_model.delete(0, tk.END); self.cfg_ai_model.insert(0, Config.AI_PROVIDERS[pv]["model"])
        self.cfg_ai_prov.bind("<<ComboboxSelected>>", _on_ai_prov)
        rb = tk.Frame(af, bg=Config.C_CARD); rb.pack(fill=tk.X, pady=5)
        tk.Label(rb, text="Base URL：", **LBL).pack(side=tk.LEFT)
        self.cfg_ai_base = tk.Entry(rb, font=("微软雅黑", 10), relief=tk.SOLID, bd=1)
        self.cfg_ai_base.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True, ipady=2)
        self.cfg_ai_base.insert(0, self.cfg.get("ai_base_url", "") or Config.AI_PROVIDERS.get(self.cfg.get("ai_provider", "doubao"), {}).get("base", ""))
        rc = tk.Frame(af, bg=Config.C_CARD); rc.pack(fill=tk.X, pady=5)
        tk.Label(rc, text="API Key：", **LBL).pack(side=tk.LEFT)
        self.cfg_ai_key = tk.Entry(rc, font=("微软雅黑", 10), show="*", relief=tk.SOLID, bd=1)
        self.cfg_ai_key.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True, ipady=2)
        self.cfg_ai_key.insert(0, self.cfg.get("ai_api_key", ""))
        tk.Button(rc, text="显示/隐藏", font=("微软雅黑", 10), bg=Config.C_BORDER, fg=Config.C_TEXT, bd=0, padx=10, pady=4,
                  command=lambda ent=self.cfg_ai_key: ent.config(show="" if ent.cget("show") == "*" else "*")).pack(side=tk.LEFT)
        abt = tk.Frame(aic, bg=Config.C_CARD); abt.pack(fill=tk.X, padx=16, pady=(0, 10))
        SButton(abt, "  保存AI配置", self._save_ai_cfg, "primary").pack(side=tk.LEFT)
        SButton(abt, "  测试 AI", self._test_ai, "success").pack(side=tk.LEFT, padx=10)
        for t in ["• 兼容 OpenAI 协议：豆包/DeepSeek/通义/Kimi/智谱/OpenAI 等",
                  "• 预设平台可下拉选择；custom 可自填任意 OpenAI 兼容地址",
                  "• 测试通过后请先保存；按 token 计费，费用由所选平台收取"]:
            tk.Label(aic, text=t, font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_M).pack(anchor="w", padx=16, pady=4)
        mc = Card(cols, title="模型平台速查", accent=Config.C_INFO)
        mc.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=0)
        for _k, _info in Config.AI_PROVIDERS.items():
            _row2 = tk.Frame(mc, bg=Config.C_CARD); _row2.pack(fill=tk.X, padx=16, pady=0)
            tk.Label(_row2, text=_info["name"], font=("微软雅黑", 10, "bold"), bg=Config.C_CARD, fg=Config.C_TEXT, width=16, anchor="w").pack(side=tk.LEFT)
            tk.Label(_row2, text=_info["base"].replace("https://", ""), font=("微软雅黑", 10), bg=Config.C_CARD, fg=Config.C_TEXT_M).pack(side=tk.LEFT, padx=(8, 0))



    def _save_cfg(self):
        if self._check_expired(): return
        self.cfg["gaode_key"] = self.cfg_gaode.get().strip()
        self.cfg["baidu_ak"] = self.cfg_baidu.get().strip()
        self.cfg["enable_baidu_fallback"] = self.cfg_eb.get()
        self.cfg["baidu_trigger_level"] = self.cfg_trig.get()
        try: self.cfg["base_interval"] = float(self.cfg_iv.get())
        except: self.cfg["base_interval"] = 0.2
        try: self.cfg["grid_workers"] = max(1, min(5, int(self.cfg_gw.get())))
        except: self.cfg["grid_workers"] = 2
        try: self.cfg["fwd_workers"] = max(1, min(5, int(self.cfg_fw.get())))
        except: self.cfg["fwd_workers"] = 3
        try: self.cfg["save_every"] = max(1, min(100000, int(self.cfg_save.get())))
        except: self.cfg["save_every"] = 200
        self.cfg["geo_drift_check"] = bool(self.cfg_geo.get())
        self.cfg["write_adcode"] = bool(self.cfg_adcode.get())
        # 让已在运行的 GeoCoder 同步新的请求间隔（限流器按最新配置生效）
        try:
            g = getattr(self, "geo", None)
            if g is not None and getattr(g, "limiter", None) is not None:
                g.limiter.min_interval = max(0.0, float(self.cfg.get("base_interval", 0.2)))
                g.limiter.cur = max(g.limiter.min_interval, min(g.limiter.cur, 3.0))
        except Exception:
            pass
        self.cfg["ai_enabled"] = self.cfg_ai_on.get()
        self.cfg["ai_provider"] = self.cfg_ai_prov.get()
        self.cfg["ai_base_url"] = self.cfg_ai_base.get().strip()
        self.cfg["ai_api_key"] = self.cfg_ai_key.get().strip()
        self.cfg["ai_model"] = self.cfg_ai_model.get().strip()
        # 修复：原实现未捕获写入异常（安装目录只读时表现为"点了保存没反应"）
        try:
            saved = Config.save(self.cfg)
        except Exception as e:
            messagebox.showerror("保存失败", f"配置无法写入：\n{e}")
            return
        messagebox.showinfo("提示", f"配置已保存！\n{saved}")

    # ---------- 本地缓存 / 离线行政区划（本次升级） ----------
    def _refresh_cache_info(self):
        """刷新设置页的缓存与区划状态显示"""
        try:
            db = get_db(); div = get_div_table()
            n = db.cache_count()
            hit, miss, ratio = db.cache_ratio()
            divn = db.div_count()
            ts = getattr(div, "synced_at", 0)
            tstr = time.strftime("%Y-%m-%d %H:%M", time.localtime(ts)) if ts else "未同步"
            txt = (f"结果缓存：{n} 条；命中 {hit} 次 / 未命中 {miss} 次（命中率 {ratio:.1f}%）\n"
                   f"离线区划：{divn} 条（同步于 {tstr}）\n"
                   f"数据文件：{db.path}")
            if not getattr(db, "available", False):
                txt += f"\n注意：本地库不可用（{getattr(db, 'init_error', '')}），缓存与离线校验已自动降级"
            self.cache_info_lbl.config(text=txt)
        except Exception as e:
            try:
                self.cache_info_lbl.config(text=f"状态读取失败：{e}")
            except Exception:
                pass

    def _clear_cache(self):
        if self._check_expired(): return
        if not messagebox.askyesno("确认", "确定要清空已缓存的地理编码结果吗？\n\n"
                                   "清空后相同地址会重新请求地图接口（消耗配额）。"):
            return
        ok = get_db().cache_clear()
        self._refresh_cache_info()
        (messagebox.showinfo if ok else messagebox.showerror)("提示", "缓存已清空" if ok else "清空失败")

    def _sync_division(self):
        if self._check_expired(): return
        gkey = self.cfg_gaode.get().strip() if hasattr(self, "cfg_gaode") else self.cfg.get("gaode_key", "")
        if not gkey:
            messagebox.showwarning("提示", "请先填写并保存高德 API Key"); return
        if not messagebox.askyesno("同步行政区划",
                                   "将从高德行政区划接口拉取全国省/市/区县两级数据。\n\n"
                                   "约消耗 34 次接口调用，耗时 10~30 秒，同步后长期离线可用。\n\n是否继续？"):
            return
        self.status.set("正在同步行政区划...")
        self._refresh_cache_info()

        def task():
            try:
                ok, msg = get_div_table().build(gkey)
            except Exception as e:
                ok, msg = False, f"同步异常：{e}"
            self._ui(lambda: self.status.set("就绪"))
            self._ui(self._refresh_cache_info)
            self._ui(lambda: (messagebox.showinfo("同步完成", msg) if ok else messagebox.showerror("同步失败", msg)))
        threading.Thread(target=task, daemon=True).start()

    def _save_ai_cfg(self):
        # v2.8.56：仅保存 AI 大模型配置（不影响地图密钥等其他配置）
        if self._check_expired(): return
        self.cfg["ai_enabled"] = self.cfg_ai_on.get()
        self.cfg["ai_provider"] = self.cfg_ai_prov.get()
        self.cfg["ai_base_url"] = self.cfg_ai_base.get().strip()
        self.cfg["ai_api_key"] = self.cfg_ai_key.get().strip()
        self.cfg["ai_model"] = self.cfg_ai_model.get().strip()
        try:
            saved = Config.save(self.cfg)
        except Exception as e:
            messagebox.showerror("保存失败", f"配置无法写入：\n{e}")
            return
        messagebox.showinfo("提示", f"AI 配置已保存！\n{saved}")

    def _test_ai(self):
        if self._check_expired(): return
        key = self.cfg_ai_key.get().strip()
        if not key:
            messagebox.showwarning("提示", "请先填写 API Key"); return
        self.status.set("正在测试 AI ...")
        def task():
            ai = AiClient({"ai_enabled": True, "ai_provider": self.cfg_ai_prov.get(),
                           "ai_base_url": self.cfg_ai_base.get(), "ai_api_key": key,
                           "ai_model": self.cfg_ai_model.get()})
            if not ai.available:
                self.root.after(0, lambda: messagebox.showerror("测试失败", "Base URL 或模型名为空，请检查配置"))
                self.root.after(0, lambda: self.status.set("就绪"))
                return
            rep = ai._chat("你是测试助手。", "请只回复：OK")
            if rep:
                self.root.after(0, lambda r=rep: messagebox.showinfo("测试成功", "AI 连接正常！返回：" + r[:50]))
            else:
                self.root.after(0, lambda: messagebox.showerror("测试失败", "AI 未返回有效内容，请检查 Key / Base URL / 模型名"))
            self.root.after(0, lambda: self.status.set("就绪"))
        threading.Thread(target=task, daemon=True).start()

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
            # 修复：本函数运行于后台线程，所有 Tk 调用（status.set / messagebox）
            # 必须经 root.after 回到主线程，否则可能崩溃或界面错乱。
            try:
                info = get_latest_version()
                if not info:
                    self.root.after(0, lambda: (self.status.set("连接更新服务器失败"),
                                                messagebox.showerror("失败", "无法连接更新服务器")))
                    return
                if not cmp_ver(Config.VERSION, info.get("latestVersion","")):
                    self.root.after(0, lambda: (self.status.set(f"已是最新版本 v{Config.VERSION}"),
                                                messagebox.showinfo("提示", "已是最新版本")))
                    return
                self.root.after(0, lambda: self._confirm_upd(info))
            except Exception as e:
                self.root.after(0, lambda e=e: (self.status.set("检查更新出错"),
                                                messagebox.showerror("错误", str(e))))
            finally:
                self.is_updating = False
                self.root.after(0, lambda: self.ubtn.config(text="🔍 检查更新", state=tk.NORMAL))
        threading.Thread(target=task, daemon=True).start()

    def _confirm_upd(self, info):
        lv = info.get("latestVersion", "")
        cl = info.get("changelog", "")
        fn = info.get("fileName", "")

        if messagebox.askyesno("发现新版本", f"发现新版本 v{lv}\n\n更新内容：\n{cl}\n\n是否立即下载更新？"):
            self.status.set(f"正在下载新版本 v{lv}...")
            download_url = f"{Config.UPDATE_URL}/download/{fn}"
            last_pct = [-1]
            def _upd_progress(downloaded, total):
                # 修复：该回调运行在下载后台线程，不得直接调用 self.root.winfo_exists()
                #（跨线程调用 Tk 会抛异常）。改为仅做进度去抖 + after() 回主线程更新。
                if total <= 0:
                    return
                pct = min(100, int(downloaded * 100 / total))
                if pct == last_pct[0]:
                    return
                last_pct[0] = pct
                try:
                    self.root.after(0, lambda p=pct: self.upd_prog.config(value=p))
                except Exception:
                    pass
            download_update(download_url, fn,
                            lambda m: self.root.after(0, lambda mm=m: self.status.set(mm)),
                            self.root, progress_cb=_upd_progress)
# ═══════════════════════════════════════════════════════════
# 10. 入口
# ═══════════════════════════════════════════════════════════
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
#（注：内容由AI生成）

__all__ = ['SButton', 'Card', 'Collapsible', 'PreviewManager', 'App', 'main']
