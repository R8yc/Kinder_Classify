# -*- coding: utf-8 -*-
"""
Checklist Viewer — 清单树独立版（列宽70px & 真·单实例）
- 左列：文件类别
- 右列：状态 | 数量  (✅[N] / ⬜[N])，初始宽度 70px，min 50px，可拖动
- 窗口更紧凑：默认 300x540，可自行调整
- F5 刷新、双击类别打开目录
- 单实例：Windows 命名互斥量 + 本地端口双保险；多次/快速双击只保留一个窗口
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

# =============== 与 Kinder Classify 一致的工具函数 ===============
def load_config() -> dict:
    with open(CONFIG, "r", encoding="utf-8") as f:
        return json.load(f)

def now_ym():
    n = datetime.now()
    return n.year, n.month

def fmt_ym(y: int, m: int):
    return f"{y:04d}", f"{m:02d}", f"{y:04d}{m:02d}"

def target_dir(cfg: dict, it: dict, y: int, m: int) -> Path:
    YYYY, MM, YYYYMM = fmt_ym(y, m)
    if it.get("path_template"):
        return Path(it["path_template"].format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM))
    if cfg.get("default_path_template"):
        return Path(cfg["default_path_template"].format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM))
    return Path(cfg["out_root"]) / f"{YYYYMM}_Unclassified"

def expected_prefix(it: dict, y: int, m: int) -> str:
    YYYY, MM, _ = fmt_ym(y, m)
    tpl = it["rename"].replace("{YYYY}", YYYY).replace("{MM}", MM)
    cut = len(tpl)
    for token in ("{orig}", "{DD}", "{ext}"):
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
        # 创建命名互斥量（若已存在，GetLastError 将返回 ERROR_ALREADY_EXISTS=183）
        handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
        if not handle:
            # 创建失败，退而求其次用端口判定
            raise OSError("CreateMutexW failed")
        err = kernel32.GetLastError()
        if err == 183:  # ERROR_ALREADY_EXISTS
            try:
                with socket.create_connection((IPC_HOST, IPC_PORT), timeout=0.8) as s:
                    s.sendall(b"RAISE")
            except OSError:
                pass
            return True
        # 持有句柄到进程结束
        global _MUTEX_HANDLE
        _MUTEX_HANDLE = handle
        return False
    except Exception:
        # 非 Windows 或 ctypes 失败，退化为端口方案
        try:
            srv = socket.socket()
            srv.bind((IPC_HOST, IPC_PORT))
            srv.listen(1)
            srv.setblocking(False)
            # 建立一个轻量线程用于收 RAISE，见 Viewer.__init__
            global _SIMPLE_SRV
            _SIMPLE_SRV = srv
            return False
        except OSError:
            # 端口被占用 → 已有实例
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
        self.geometry("300x540")         # 更紧凑
        self.minsize(300, 340)

        # 若仅使用端口方案，启动一个微型线程处理 RAISE
        if "_SIMPLE_SRV" in globals():
            import threading
            t = threading.Thread(target=self._serve_raise_once, daemon=True)
            t.start()

        y0, m0 = now_ym()
        self.year_var = tk.StringVar(value=str(y0))
        self.month_var = tk.StringVar(value=f"{m0:02d}")

        # 快捷键
        self.bind("<F5>", lambda e: self.refresh())

        # 顶部：年月 + 路径提示
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

        # 中部：树（滚动条；列宽可拖动）
        mid = ttk.Frame(self, padding=(10, 6, 10, 6))
        mid.pack(fill=tk.BOTH, expand=True)

        wrap = ttk.Frame(mid); wrap.pack(fill=tk.BOTH, expand=True)
        yscroll = ttk.Scrollbar(wrap, orient="vertical"); yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            wrap, columns=("status",), show="tree headings",
            yscrollcommand=yscroll.set
        )
        yscroll.config(command=self.tree.yview)

        # 左“文件类别”，右“状态[数量]”；右列初始 50px，最小 10px；两列 stretch=True 以避免“回弹”
        self.tree.heading("#0", text="文件类别")
        self.tree.heading("status", text="状态[数量]")
        self.tree.column("#0",     width=180, minwidth=160, stretch=True,  anchor="w")
        self.tree.column("status", width=50,  minwidth=10,  stretch=True,  anchor="center")

        self.tree.tag_configure("group", foreground="#666")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.open_dir_of_selected)

        # 底部：刷新 + 状态栏
        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(fill=tk.X)
        ttk.Button(bottom, text="刷新 (F5)", command=self.refresh).pack(side=tk.LEFT)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom, textvariable=self.status_var, anchor="w", relief="groove").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0)
        )

        self.refresh()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # 端口方案的 RAISE 处理（仅当互斥量不可用时）
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
            self.root_hint.set(self.cfg["default_path_template"].format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM))
        else:
            self.root_hint.set(f"{self.cfg.get('out_root', '')}/{YYYYMM}_Unclassified")

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        status, counts = compute_status_and_count(self.cfg, y, m)
        order, groups = grouped_items(self.cfg["items"])
        ok_ct, tot_ct = 0, 0
        for g in order:
            parent = self.tree.insert("", tk.END, text=f"—— {g} ——", tags=("group",), open=True)
            for it in groups[g]:
                key = it["key"]
                mark = "✅" if status.get(key, False) else "⬜"
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

    def on_close(self):
        # 互斥量句柄让系统回收即可；端口方案需主动关闭
        try:
            if "_SIMPLE_SRV" in globals():
                _SIMPLE_SRV.close()
        except Exception:
            pass
        self.destroy()

# =============== 入口 ===============
def main():
    # 单实例：若已有 → 置顶后退出
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
