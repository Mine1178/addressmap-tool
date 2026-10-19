"""
赞助弹窗模块（自包含，拷到任何 tkinter 项目直接用）
=====================================================

用法：
    from sponsor_widget import show_sponsor_button, show_sponsor_dialog

    # 方式一：一行搞定，自动在父容器里放按钮 + 绑弹窗
    btn = show_sponsor_button(parent_frame)   # 返回按钮本身，可 pack/grid

    # 方式二：自己控制按钮，点击时弹
    tk.Button(root, text="☕ 赞助作者",
              command=lambda: show_sponsor_dialog(root)).pack()

前置条件：
    1. 项目里有 assets/alipay.jpg 和 assets/wechatpay.jpg
    2. 安装了 Pillow（pip install pillow）
    3. PyInstaller 打包时加：--add-data "assets;assets"

文件位置约定：
    sponsor_widget.py 放项目里任意位置，自动按以下顺序找码图：
      - sys._MEIPASS/assets/alipay.jpg（PyInstaller 打包后）
      - sponsor_widget.py 所在目录/../assets/alipay.jpg
      - sponsor_widget.py 所在目录/assets/alipay.jpg
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
import tkinter as tk

# ============ 可调参数（复制后按需改） ============
DIALOG_W = 1060
DIALOG_H = 880
DIALOG_MINSIZE = (720, 560)
QR_SIZE = 480          # 码图显示尺寸（px，按原图比例缩放）
BG = "#f1f5f9"
CARD_BG = "#ffffff"
PRIMARY = "#1e293b"
ACCENT = "#d97706"
TEXT = "#1e293b"
TEXT_SUB = "#64748b"
BTN_BG = "#f1f5f9"
BTN_ACTIVE = "#e2e8f0"
FONT_FAMILY = "Microsoft YaHei"
# =================================================


def _find_asset(filename: str) -> Path | None:
    """按优先级找资源文件。"""
    # 1. PyInstaller 打包后
    if hasattr(sys, "_MEIPASS"):
        p = Path(sys._MEIPASS) / "assets" / filename  # type: ignore[attr-defined]
        if p.exists():
            return p
    # 2. 本文件上溯找 assets/
    here = Path(__file__).resolve().parent
    candidates = [
        here / "assets" / filename,
        here.parent / "assets" / filename,
        here.parent.parent / "assets" / filename,
        Path.cwd() / "assets" / filename,
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def show_sponsor_dialog(parent: tk.Misc) -> None:
    """弹赞助窗：居中、1060x880、双码图、一键放大。"""
    dlg = tk.Toplevel(parent)
    dlg.title("赞助作者")
    dlg.geometry(f"{DIALOG_W}x{DIALOG_H}")
    dlg.minsize(*DIALOG_MINSIZE)
    dlg.resizable(True, True)
    dlg.transient(parent)
    dlg.configure(bg=BG)
    dlg.grab_set()

    # 屏幕居中
    dlg.update_idletasks()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"+{(sw - DIALOG_W) // 2}+{(sh - DIALOG_H) // 2}")

    # 一键放大/还原
    state = {"zoomed": False, "geom": dlg.geometry()}

    def _toggle_zoom():
        if state["zoomed"]:
            dlg.geometry(state["geom"])
            zoom_btn.configure(text="⤢ 放大")
        else:
            state["geom"] = dlg.geometry()
            dlg.state("zoomed")
            zoom_btn.configure(text="🗗 还原")
        state["zoomed"] = not state["zoomed"]

    # 顶部文案
    tk.Label(dlg, text="☕", font=(FONT_FAMILY, 36),
             bg=BG, fg=ACCENT).pack(pady=(20, 0))
    tk.Label(dlg, text="如果你喜欢这个工具",
             font=(FONT_FAMILY, 16, "bold"), bg=BG, fg=TEXT).pack(pady=(4, 0))
    tk.Label(dlg, text="可以请作者喝杯咖啡，感谢赞助",
             font=(FONT_FAMILY, 11), bg=BG, fg=TEXT_SUB).pack(pady=(0, 20))

    # 码图区（expand 占满剩余空间）
    codes = tk.Frame(dlg, bg=BG)
    codes.pack(fill=tk.BOTH, expand=True, padx=30)
    codes.grid_columnconfigure(0, weight=1)
    codes.grid_columnconfigure(1, weight=1)
    codes.grid_rowconfigure(0, weight=1)

    def _load_qr(row: int, filename: str, label: str) -> None:
        cell = tk.Frame(codes, bg=BG)
        cell.grid(row=0, column=row, sticky="nsew", padx=30)
        path = _find_asset(filename)
        if path:
            try:
                from PIL import Image, ImageTk
                img = Image.open(path)
                img.thumbnail((QR_SIZE, QR_SIZE), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                lbl = tk.Label(cell, image=photo, bg=CARD_BG,
                               relief=tk.SOLID, bd=2)
                lbl.image = photo
                lbl.pack(expand=True)
            except Exception as e:
                tk.Label(cell, text=f"图片加载失败: {e}",
                         font=(FONT_FAMILY, 9), bg=BG, fg="#dc2626").pack()
        else:
            tk.Label(cell, text=f"[{filename} 待添加]",
                     font=(FONT_FAMILY, 9), width=16, height=12,
                     bg="#e2e8f0", fg=TEXT_SUB, relief=tk.SOLID, bd=1).pack()
        tk.Label(cell, text=label, font=(FONT_FAMILY, 11, "bold"),
                 bg=BG, fg=TEXT).pack(pady=(8, 0))

    _load_qr(0, "alipay.jpg", "📱 支付宝扫码")
    _load_qr(1, "wechatpay.jpg", "💬 微信扫码")

    # 底部按钮行
    btn_row = tk.Frame(dlg, bg=BG)
    btn_row.pack(pady=(0, 20))
    zoom_btn = tk.Button(btn_row, text="⤢ 放大", command=_toggle_zoom,
                         font=(FONT_FAMILY, 10), bg=BTN_BG, fg=TEXT,
                         relief=tk.SOLID, bd=1, padx=20, pady=6, cursor="hand2")
    zoom_btn.pack(side=tk.LEFT, padx=8)
    tk.Button(btn_row, text="关闭", command=dlg.destroy,
              font=(FONT_FAMILY, 11), bg=PRIMARY, fg="#ffffff",
              relief=tk.FLAT, padx=32, pady=6).pack(side=tk.LEFT, padx=8)


def show_sponsor_button(parent: tk.Misc, **btn_kwargs) -> tk.Button:
    """在 parent 上创建「☕ 赞助作者」按钮并自动绑弹窗，返回按钮对象。

    用法：
        btn = show_sponsor_button(about_page_frame)
        btn.pack(pady=10)   # 或 .grid / .place 自己布局
    """
    defaults = dict(
        text="☕ 赞助作者",
        font=(FONT_FAMILY, 10),
        bg=BG, fg=TEXT_SUB,
        relief=tk.FLAT, bd=0,
        activebackground=BG, activeforeground=ACCENT,
        padx=12, pady=4, cursor="hand2",
    )
    defaults.update(btn_kwargs)
    btn = tk.Button(parent, command=lambda: show_sponsor_dialog(parent), **defaults)
    # 悬停时文字变橙色，与页面其它链接式文字呼应，不突兀
    _bg = defaults.get("bg", BG)
    def _enter(e): btn.config(fg=ACCENT)
    def _leave(e): btn.config(fg=defaults.get("fg", TEXT_SUB))
    btn.bind("<Enter>", _enter)
    btn.bind("<Leave>", _leave)
    return btn


# ============ 自测：直接 python sponsor_widget.py 自动弹赞助窗 ============
if __name__ == "__main__":
    root = tk.Tk()
    root.title("sponsor_widget 预览")
    root.geometry("400x300")
    root.configure(bg=BG)
    tk.Label(root, text="点下面按钮预览赞助弹窗", font=(FONT_FAMILY, 11),
             bg=BG, fg=TEXT_SUB).pack(pady=(40, 10))
    show_sponsor_button(root).pack()
    # 启动后 0.3 秒自动弹一次，方便双击运行直接看效果
    root.after(300, lambda: show_sponsor_dialog(root))
    root.mainloop()
