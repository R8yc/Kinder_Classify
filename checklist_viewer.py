# -*- coding: utf-8 -*-
"""
Checklist Viewer — 清单树独立版（列宽70px & 真·单实例）
- 左列：文件类别
- 右列：状态 | 数量  (✅[N] / ⬜[N])，初始宽度 70px，min 50px，可拖动
- 窗口更紧凑：默认 300x540，可自行调整
- F5 刷新、双击类别打开目录
- 单实例：Windows 命名互斥量 + 本地端口双保险；多次/快速双击只保留一个窗口

【ver12.2本次改动】
1) 手动“本月没有的文档”标记 ❎ ：右侧清单“状态 | 数量”列单击切换 ⬜ ↔ ❎（当自动为 ✅ 时不可改）
2) 与 Checklist Viewer 通过 na_overrides.json 同步
3) 新增支持占位符：{MM-1}、{YYYY-1}、{YYYYMM-1}（rename / path_template / dest_subdir / default_path_template）
"""

import json
import os
import socket
import socketserver
import sys
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

# =============== 配置定位（与 Kinder Classify 相同顺序） ===============
SCRIPT_DIR = Path(__file__).resolve().parent
CANDIDATES = [
    SCRIPT_DIR,
    Path(r"E:/ShangJin_Kindergarten/KinderClassify"),
    Path(r"E:\ShangJin_Kindergarten\KinderClassify"),
]
CONFIG = None
for base in CANDIDATES:
    c = base / "config.json"
    if c.exists():
        CONFIG = c
        break
if CONFIG is None:
    CONFIG = SCRIPT_DIR / "config.json"  # 兜底

# —— 本月不适用（❎）相关：与 config.json 同目录的持久化文件 —— #
NA_FILE = CONFIG.with_name("na_overrides.json")

def _load_na() -> dict:
    try:
        with open(NA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _save_na(data: dict):
    try:
        tmp = NA_FILE.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, NA_FILE)
    except Exception:
        pass

# =============== 与 Kinder Classify 一致的工具函数 ===============
def load_config() -> dict:
    with open(CONFIG, "r", encoding="utf-8") as f:
        return json.load(f)

def now_ym():
    n = datetime.now()
    return n.year, n.month

def fmt_ym(y: int, m: int):
    return f"{y:04d}", f"{m:02d}", f"{y:04d}{m:02d}"

# ====== 新增：支持 {MM-1}/{YYYY-1}/{YYYYMM-1} 的轻量替换 ======
def _prev_ym(y:int, m:int):
    if m == 1:
        return y-1, 12
    return y, m-1

def _expand_month_minus_one(s: str, y: int, m: int):
    """Replace {YYYY-1}, {MM-1}, {YYYYMM-1} in the template string.
    Keep other {placeholders} untouched for .format(...) later.
    """
    if not isinstance(s, str):
        return s
    y0, m0 = y, m
    yp, mp = _prev_ym(y0, m0)
    YYYY  = f"{y0:04d}";  MM  = f"{m0:02d}";  YYYYMM  = f"{y0:04d}{m0:02d}"
    YYYYp = f"{yp:04d}"; MMp = f"{mp:02d}"; YYYYMMp = f"{yp:04d}{mp:02d}"
    return (s.replace("{YYYY-1}", YYYYp)
             .replace("{MM-1}", MMp)
             .replace("{YYYYMM-1}", YYYYMMp))

def target_dir(cfg: dict, it: dict, y: int, m: int) -> Path:
    YYYY, MM, YYYYMM = fmt_ym(y, m)
    # 1) 选择路径模板：item 覆盖全局
    if it.get("path_template"):
        tpl = it["path_template"]
    elif cfg.get("default_path_template"):
        tpl = cfg["default_path_template"]
    else:
        return Path(cfg["out_root"]) / f"{YYYYMM}_Unclassified"

    # 2) 先把 dest_subdir 自己也做一次占位符替换
    raw_sub = it.get("dest_subdir", "")
    if raw_sub:
        raw_sub = _expand_month_minus_one(raw_sub, y, m)
        sub = raw_sub.format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM)
    else:
        sub = ""

    # 3) 套模板本身（支持模板中直接写 {dest_subdir}）
    tpl = _expand_month_minus_one(tpl, y, m)
    base = Path(tpl.format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM, dest_subdir=sub))

    # 4) 如果模板里没写 {dest_subdir}，但配置给了 sub，则在末尾追加
    if sub and "{dest_subdir}" not in tpl:
        base = base / sub

    return base

def expected_prefix(it: dict, y: int, m: int) -> str:
    YYYY, MM, _ = fmt_ym(y, m)
    tpl_raw = _expand_month_minus_one(it["rename"], y, m)
    tpl = tpl_raw.replace("{YYYY}", YYYY).replace("{MM}", MM)
    cut = len(tpl)
    for token in ("{orig}", "{DD}", "{ext}", "{YYYYMM}"):
        i = tpl.find(token)
        if i != -1 and i < cut:
            cut = i
    return tpl[:cut]

def compute_status_and_count(cfg: dict, y: int, m: int):
    status, counts = {}, {}
    for it in cfg["items"]:
        d = target_dir(cfg, it, y, m)
        prefix = expected_prefix(it, y, m)
        exts = [e.lower() for e in it.get("exts", [])] if it.get("exts") else None
        cnt = 0
        if d.exists():
            for p in d.iterdir():
                if not p.is_file():      continue
                if not p.name.startswith(prefix): continue
                if exts and p.suffix.lower() not in exts: continue
                cnt += 1
        rule = it.get("present_rule", {"mode": "any"})
        ok = (cnt >= int(rule.get("n", 1))) if rule.get("mode") == "count_at_least" else (cnt > 0)
        status[it["key"]] = ok
        counts[it["key"]] = cnt
    return status, counts

def group_of(k: str) -> str:
    l, r = k.find("【"), k.find("】")
    return k[l+1:r].strip() if (l != -1 and r != -1 and r > l) else "其他"

def grouped_items(items: list):
    groups, order = {}, []
    for it in items:
        g = group_of(it["key"])
        if g not in groups:
            groups[g] = []
            order.append(g)
        groups[g].append(it)
    return order, groups

# =============== 单实例：互斥量 + 端口双保险 ===============
IPC_HOST, IPC_PORT = "127.0.0.1", 53452  # 与 Kinder Classify 不同
MUTEX_NAME = r"Global\Checklist_Viewer_Singleton_v1"

def already_running_raise_then_exit() -> bool:
    """
    返回 True 表示检测到已有实例，且已请求其置顶；当前进程应立即退出。
    先用 Windows 互斥量判定；若已存在，再通过端口唤起已有窗口。
    """
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
        if not handle:
            raise OSError("CreateMutexW failed")
        err = kernel32.GetLastError()
        if err == 183:  # ERROR_ALREADY_EXISTS
            try:
                with socket.create_connection((IPC_HOST, IPC_PORT), timeout=0.8) as s:
                    s.sendall(b"RAISE")
            except OSError:
                pass
            return True
        global _MUTEX_HANDLE
        _MUTEX_HANDLE = handle
        return False
    except Exception:
        try:
            srv = socket.socket()
            srv.bind((IPC_HOST, IPC_PORT))
            srv.listen(1)
            srv.setblocking(False)
            global _SIMPLE_SRV
            _SIMPLE_SRV = srv
            return False
        except OSError:
            try:
                with socket.create_connection((IPC_HOST, IPC_PORT), timeout=0.8) as s:
                    s.sendall(b"RAISE")
            except OSError:
                pass
            return True

# =============== UI ===============
class Viewer(tk.Tk):
    def __init__(self, cfg: dict):
        super().__init__()
        self.cfg = cfg
        self.title("Checklist Viewer")
        self.geometry("300x540")
        self.minsize(300, 340)

        if "_SIMPLE_SRV" in globals():
            import threading
            t = threading.Thread(target=self._serve_raise_once, daemon=True)
            t.start()

        y0, m0 = now_ym()
        self.year_var = tk.StringVar(value=str(y0))
        self.month_var = tk.StringVar(value=f"{m0:02d}")

        self.bind("<F5>", lambda e: self.refresh())

        top = ttk.Frame(self, padding=(10, 8, 10, 0))
        top.pack(fill=tk.X)
        ttk.Label(top, text="操作年月：").pack(side=tk.LEFT)
        ycb = ttk.Combobox(top, textvariable=self.year_var, width=6, state="readonly",
                           values=[str(y) for y in range(2000, y0+6)])
        ycb.pack(side=tk.LEFT, padx=(0,6))
        mcb = ttk.Combobox(top, textvariable=self.month_var, width=4, state="readonly",
                           values=[f"{i:02d}" for i in range(1, 13)])
        mcb.pack(side=tk.LEFT)
        ycb.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        mcb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        self.root_hint = tk.StringVar()
        ttk.Label(top, textvariable=self.root_hint, foreground="#0066cc").pack(side=tk.RIGHT)

        mid = ttk.Frame(self, padding=(10, 6, 10, 6))
        mid.pack(fill=tk.BOTH, expand=True)

        wrap = ttk.Frame(mid); wrap.pack(fill=tk.BOTH, expand=True)
        yscroll = ttk.Scrollbar(wrap, orient="vertical"); yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            wrap, columns=("status",), show="tree headings",
            yscrollcommand=yscroll.set
        )
        yscroll.config(command=self.tree.yview)

        self.tree.heading("#0", text="文件类别")
        self.tree.heading("status", text="状态[数量]")
        self.tree.column("#0",     width=180, minwidth=160, stretch=True,  anchor="w")
        self.tree.column("status", width=50,  minwidth=10,  stretch=True,  anchor="center")

        self.tree.tag_configure("group", foreground="#666")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.open_dir_of_selected)

        # —— 新增：单击“状态[数量]”列切换 ⬜ ↔ ❎（✅ 不可手改） —— #
        self.tree.bind("<Button-1>", self._on_status_click, add="+")

        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(fill=tk.X)
        ttk.Button(bottom, text="刷新 (F5)", command=self.refresh).pack(side=tk.LEFT)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom, textvariable=self.status_var, anchor="w", relief="groove").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0)
        )

        self.refresh()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _serve_raise_once(self):
        srv = _SIMPLE_SRV
        while True:
            try:
                conn, _ = srv.accept()
            except BlockingIOError:
                self.after(200, lambda: None)
                continue
            with conn:
                _ = conn.recv(1024)
                self.after(0, self.raise_to_front)

    def raise_to_front(self):
        try:
            self.deiconify(); self.lift()
            self.attributes("-topmost", True)
            self.after(250, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

    # ---------- 基本动作 ----------
    def set_status(self, s: str): self.status_var.set(s)

    def current_ym(self):
        try: return int(self.year_var.get()), int(self.month_var.get())
        except Exception: return now_ym()

    def refresh(self):
        y, m = self.current_ym()
        YYYY, MM, YYYYMM = fmt_ym(y, m)

        if self.cfg.get("default_path_template"):
            tpl = _expand_month_minus_one(self.cfg["default_path_template"], y, m)
            self.root_hint.set(tpl.format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM))
        else:
            self.root_hint.set(f"{self.cfg.get('out_root', '')}/{YYYYMM}_Unclassified")

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        status, counts = compute_status_and_count(self.cfg, y, m)

        # —— 本月不适用（❎）读取 —— #
        na_for_month = _load_na().get(f"{y:04d}-{m:02d}", {})

        order, groups = grouped_items(self.cfg["items"])
        ok_ct, tot_ct = 0, 0
        for g in order:
            parent = self.tree.insert("", tk.END, text=f"—— {g} ——", tags=("group",), open=True)
            for it in groups[g]:
                key = it["key"]

                # —— 状态优先级：❎ > ✅ > ⬜ —— #
                if na_for_month.get(key, False):
                    mark = "❎"
                elif status.get(key, False):
                    mark = "✅"
                else:
                    mark = "⬜"

                cnt = counts.get(key, 0)
                self.tree.insert(parent, tk.END, text=key, values=(f"{mark}[{cnt}]",))
                tot_ct += 1
                if status.get(key, False): ok_ct += 1

        self.set_status(f"刷新成功：{ok_ct}/{tot_ct} 类满足")

    def open_dir_of_selected(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        iid = sel[0]
        if self.tree.get_children(iid):  # 组标题
            return
        key = self.tree.item(iid, "text")
        y, m = self.current_ym()
        it = next((x for x in self.cfg["items"] if x["key"] == key), None)
        if not it: return
        d = target_dir(self.cfg, it, y, m)
        try:
            d.mkdir(parents=True, exist_ok=True)
            os.startfile(str(d))
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    # —— 新增：手动 ❎ 切换（不影响原有双击/按钮逻辑） —— #
    def _on_status_click(self, event):
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if self.tree.get_children(iid):  # 组标题
            return

        key = self.tree.item(iid, "text")
        y, m = self.current_ym()
        ym = f"{y:04d}-{m:02d}"

        # 自动为 ✅ 时不可手改
        status, _ = compute_status_and_count(self.cfg, y, m)
        if status.get(key, False):
            return

        data = _load_na()
        d = data.get(ym, {})
        if d.get(key, False):
            d.pop(key, None)
            if not d:
                data.pop(ym, None)
        else:
            d[key] = True
            data[ym] = d
        _save_na(data)
        self.refresh()

    def on_close(self):
        try:
            if "_SIMPLE_SRV" in globals():
                _SIMPLE_SRV.close()
        except Exception:
            pass
        self.destroy()

# =============== 入口 ===============
def main():
    if already_running_raise_then_exit():
        return
    try:
        cfg = load_config()
    except Exception as e:
        messagebox.showerror("Checklist Viewer 启动失败", f"无法读取配置：\n{CONFIG}\n{e}")
        return
    app = Viewer(cfg)
    app.mainloop()

if __name__ == "__main__":
    main()
