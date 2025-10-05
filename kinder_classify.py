#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kinder Classify
- 单实例 + IPC：二次启动只把文件追加到现有窗口
- 拖拽（tkinterdnd2 可用则用；不可用则正常运行）
- 撤销/重做：一次一步；撤销回到列表，重做从列表移除
- 清单树：左“文件类别”，右“状态 | 数量”（勾选依据磁盘现状）
- 状态栏提示（成功/刷新/撤销/重做等），仅错误弹窗

【ver10本次改动】
1) cmd_redo() 修复：重做后从待处理列表移除“移动前的旧路径”，不再误留。
2) 总是默认选中第一个待处理文件：任何刷新文件列表的时刻都会自动选中第一个。
"""

import os, sys, json, shutil, logging, socket, threading, socketserver
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ---------- 拖拽支持 ----------
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    class TkinterDnD:
        class Tk(tk.Tk): ...
    DND_FILES = "DND_FALLBACK"
    HAS_DND = False

# ---------- 配置与日志 ----------
SCRIPT_DIR = Path(__file__).resolve().parent
CANDIDATES = [
    SCRIPT_DIR,
    Path(r"E:/ShangJin_Kindergarten/KinderClassify"),
    Path(r"E:\ShangJin_Kindergarten\KinderClassify"),
]
CONFIG = None; LOG = None
for base in CANDIDATES:
    cfg = base / "config.json"
    if cfg.exists():
        CONFIG = cfg; LOG = base / "kinder_classify.log"; break
if CONFIG is None:
    CONFIG = SCRIPT_DIR / "config.json"
    LOG = SCRIPT_DIR / "kinder_classify.log"

logging.basicConfig(filename=str(LOG), level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")

# 单实例 / IPC
IPC_HOST, IPC_PORT = "127.0.0.1", 53451


# ---------- 通用工具 ----------
def load_config():
    with open(CONFIG, "r", encoding="utf-8") as f:
        return json.load(f)

def now_ym():
    n = datetime.now(); return n.year, n.month

def fmt_ym(y, m):
    return f"{y:04d}", f"{m:02d}", f"{y:04d}{m:02d}"

def safe_name(s: str) -> str:
    for c in '<>:"/\\|?*': s = s.replace(c, "_")
    return s.rstrip(" .")

def move_with_conflict(src: Path, dst: Path) -> Path:
    dst.parent.mkdir(parents=True, exist_ok=True)
    t, i = dst, 1
    while t.exists():
        t = dst.with_name(f"{dst.stem}-{i}{dst.suffix}"); i += 1
    shutil.move(str(src), str(t))
    return t

def target_dir(cfg: dict, it: dict, y: int, m: int) -> Path:
    YYYY, MM, YYYYMM = fmt_ym(y, m)

    # 1) 选择路径模板：item 覆盖全局
    if it.get("path_template"):
        tpl = it["path_template"]
    elif cfg.get("default_path_template"):
        tpl = cfg["default_path_template"]
    else:
        # 老规则兜底（保持你原来的行为）
        return Path(cfg["out_root"]) / f"{YYYYMM}_Unclassified"

    # 2) 先把 dest_subdir 自己也做一次占位符替换
    raw_sub = it.get("dest_subdir", "")
    sub = raw_sub.format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM) if raw_sub else ""

    # 3) 套模板本身（支持模板中直接写 {dest_subdir}）
    base = Path(tpl.format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM, dest_subdir=sub))

    # 4) 如果模板里没写 {dest_subdir}，但配置给了 sub，则在末尾追加
    if sub and "{dest_subdir}" not in tpl:
        base = base / sub

    return base


def expected_prefix(it: dict, y: int, m: int) -> str:
    YYYY, MM, _ = fmt_ym(y, m)
    tpl = it["rename"].replace("{YYYY}", YYYY).replace("{MM}", MM)
    cut = len(tpl)
    for token in ("{orig}", "{DD}", "{ext}"):
        i = tpl.find(token)
        if i != -1 and i < cut: cut = i
    return tpl[:cut]

def compute_status(cfg: dict, y: int, m: int) -> dict:
    status = {}
    for it in cfg["items"]:
        d = target_dir(cfg, it, y, m)
        prefix = expected_prefix(it, y, m)
        exts = [e.lower() for e in it.get("exts", [])] if it.get("exts") else None
        cnt = 0
        if d.exists():
            for p in d.iterdir():
                if not p.is_file(): continue
                if not p.name.startswith(prefix): continue
                if exts and p.suffix.lower() not in exts: continue
                cnt += 1
        rule = it.get("present_rule", {"mode": "any"})
        ok = (cnt >= int(rule.get("n", 1))) if rule.get("mode") == "count_at_least" else (cnt > 0)
        status[it["key"]] = ok
    return status

def count_for_item(cfg: dict, it: dict, y: int, m: int) -> int:
    d = target_dir(cfg, it, y, m)
    prefix = expected_prefix(it, y, m)
    exts = [e.lower() for e in it.get("exts", [])] if it.get("exts") else None
    cnt = 0
    if d.exists():
        for p in d.iterdir():
            if not p.is_file(): continue
            if not p.name.startswith(prefix): continue
            if exts and p.suffix.lower() not in exts: continue
            cnt += 1
    return cnt

def group_of(k: str) -> str:
    l, r = k.find("【"), k.find("】")
    return k[l+1:r].strip() if (l != -1 and r != -1 and r > l) else "其他"

def grouped_items(items: list[dict]):
    groups, order = {}, []
    for it in items:
        g = group_of(it["key"])
        if g not in groups:
            groups[g] = []; order.append(g)
        groups[g].append(it)
    return order, groups


# ---------- 主应用 ----------
class App(TkinterDnD.Tk):
    def __init__(self, cfg: dict, files_cli: list[str]):
        super().__init__()
        self.cfg = cfg
        self.files = [Path(f) for f in files_cli if Path(f).exists()]
        self.title("Kinder Classify")
        self.geometry("1120x700")

        # 撤销/重做栈：entry = {"orig": str, "dst": str, "current": str}
        self.undo_stack: list[dict] = []
        self.redo_stack: list[dict] = []

        y0, m0 = now_ym()
        self.year_var = tk.StringVar(value=str(y0))
        self.month_var = tk.StringVar(value=f"{m0:02d}")

        # 快捷键（窗口与主要控件）
        for w in (self,):
            w.bind("<Control-z>", lambda e: self.cmd_undo())
            w.bind("<Control-y>", lambda e: self.cmd_redo())
            w.bind("<F5>",       lambda e: self.refresh_status())

        # ===== 布局 =====
        frm = ttk.Frame(self, padding=10); frm.pack(fill=tk.BOTH, expand=True)

        # ---------------- 左栏：固定头部 + 固定标题 + “类别按钮列表”可滚动 ----------------
        left_col = ttk.Frame(frm)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # (1) 固定头部：操作年月（不随滚动）
        head = ttk.Frame(left_col)
        head.pack(fill=tk.X)

        ttk.Label(head, text="操作年月").pack(anchor="w")
        ybox = ttk.Combobox(head, textvariable=self.year_var, width=8, state="readonly",
                            values=[str(y) for y in range(2000, y0+6)])
        ybox.pack(anchor="w")
        mbox = ttk.Combobox(head, textvariable=self.month_var, width=8, state="readonly",
                            values=[f"{i:02d}" for i in range(1, 13)])
        mbox.pack(anchor="w", pady=(0, 8))
        ybox.bind("<<ComboboxSelected>>", lambda e: self.refresh_status())
        mbox.bind("<<ComboboxSelected>>", lambda e: self.refresh_status())

        # (2) 固定小标题：不滚动
        ttk.Label(left_col, text="类别（点击分类）").pack(anchor="w", pady=(8, 4))

        # (3) 可滚动区：仅“类别下面”的所有按钮在这里
        scroll_wrap = ttk.Frame(left_col)
        scroll_wrap.pack(fill=tk.BOTH, expand=True)

        # —— 关键改动 A：固定宽度 200，并确保只竖向滚动 —— #
        left_canvas = tk.Canvas(scroll_wrap, borderwidth=0, highlightthickness=0, width=200)
        left_vsb = ttk.Scrollbar(scroll_wrap, orient="vertical", command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_vsb.set)

        # 画布宽度固定，填充竖向即可；不随父级横向拉伸
        left_canvas.pack(side=tk.LEFT, fill=tk.Y, expand=False)
        left_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 真正承载分类按钮的内部 Frame（固定宽度 200）
        left = ttk.Frame(left_canvas, width=200)
        win_id = left_canvas.create_window((0, 0), window=left, anchor="nw", width=200)

        # 更新滚动区域
        def _update_scrollregion(event=None):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        left.bind("<Configure>", _update_scrollregion)

        # 固定内部 Frame 宽度为 200（不随画布伸缩）
        def _keep_fixed_width(event=None):
            left_canvas.itemconfigure(win_id, width=200)
        left_canvas.bind("<Configure>", _keep_fixed_width)

        # —— 改进：把滚轮事件直接绑定到可滚动区域及其子控件，避免 bind_all 抖动 —— #
        SCROLL_UNITS = 3  # 滚动步长：2 更细腻，4/5 更快

        def _on_mousewheel(event):
            """
            Windows：event.delta 为 ±120 的倍数；向下滚 delta<0
            我们遵循 Windows 习惯：向下滚 = 内容向下（滚动条向下）
            """
            if hasattr(event, "delta") and event.delta != 0:
                move = int(-event.delta / 120) * SCROLL_UNITS
                if move != 0:
                    left_canvas.yview_scroll(move, "units")
            elif hasattr(event, "num") and event.num in (4, 5):
                # Linux: 4=上，5=下
                move = (-1 if event.num == 4 else 1) * SCROLL_UNITS
                left_canvas.yview_scroll(move, "units")

        def _bind_scrollwheel(target):
            target.bind("<MouseWheel>", _on_mousewheel, add="+")
            target.bind("<Button-4>",  _on_mousewheel, add="+")  # Linux
            target.bind("<Button-5>",  _on_mousewheel, add="+")  # Linux
            for child in target.winfo_children():
                _bind_scrollwheel(child)
        # —— 改进结束 —— #

        # 样式与按钮生成（保持不变）
        s = ttk.Style(self); s.configure("Left.TButton", anchor="w", padding=(6, 4))
        order, groups = grouped_items(cfg["items"])
        for g in order:
            ttk.Label(left, text=f"—— {g} ——", foreground="#666").pack(anchor="w", pady=(8, 2))
            ttk.Separator(left, orient="horizontal").pack(fill="x", pady=(0, 4))
            for it in groups[g]:
                ttk.Button(left, text=it["key"], style="Left.TButton",
                           command=lambda _it=it: self.assign(_it)).pack(fill=tk.X, pady=1)

        # 绑定滚轮到可滚动区及其子控件，使鼠标停在列表上即可滚动
        _bind_scrollwheel(left)

        # ---------------- 中栏：待分类文件 + 状态栏 + 按钮（保持原样） ----------------
        mid = ttk.Frame(frm); mid.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        ttk.Label(mid, text="待分类文件（可拖拽或“添加文件…”）").pack(anchor="w")

        self.file_list = tk.Listbox(mid, selectmode=tk.EXTENDED, height=24)
        self.file_list.pack(fill=tk.BOTH, expand=True)

        if HAS_DND:
            self.file_list.drop_target_register(DND_FILES)
            self.file_list.dnd_bind('<<Drop>>', self._on_drop)

        btn_mid = ttk.Frame(mid); btn_mid.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        ttk.Button(btn_mid, text="添加文件…", command=self.add_files_dialog).pack(side=tk.LEFT)
        ttk.Button(btn_mid, text="移除所选", command=self.remove_selected).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_mid, text="清空列表", command=self.clear_files).pack(side=tk.LEFT)

        status_frame = ttk.Frame(mid); status_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 6))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var, anchor="w", relief="groove").pack(fill=tk.X)

        self.refresh_files()

        # ---------------- 右栏：清单树 + 按钮（保持原样） ----------------
        right = ttk.Frame(frm); right.pack(side=tk.RIGHT, fill=tk.Y)

        self.root_hint = tk.StringVar()
        ttk.Label(right, textvariable=self.root_hint, foreground="#0066cc").pack(anchor="w")

        self.tree = ttk.Treeview(right, columns=("status",), show="tree headings", height=24)
        self.tree.heading("#0", text="文件类别")
        self.tree.heading("status", text="状态 | 数量")
        self.tree.column("#0", width=220, minwidth=180, stretch=False, anchor="w")
        self.tree.column("status", width=70, minwidth=50, stretch=False, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True, pady=(6, 6))
        self.tree.tag_configure("group", foreground="#666")
        self.tree.bind("<Double-1>", self.open_dir_of_selected)

        btn_right = ttk.Frame(right); btn_right.pack(fill=tk.X)
        ttk.Button(btn_right, text="刷新 (F5)", command=self.refresh_status).pack(fill=tk.X, pady=4)
        ttk.Button(btn_right, text="撤销 (Ctrl+Z)", command=self.cmd_undo).pack(fill=tk.X, pady=4)
        ttk.Button(btn_right, text="重做 (Ctrl+Y)", command=self.cmd_redo).pack(fill=tk.X, pady=4)

        self.refresh_status()

    # ---------- 辅助 ----------
    def set_status(self, text: str): 
        self.status_var.set(text)

    def current_ym(self):
        try: return int(self.year_var.get()), int(self.month_var.get())
        except Exception: return now_ym()

    def _on_drop(self, event):
        try: paths = self.tk.splitlist(event.data)
        except Exception: paths = [event.data]
        self._add_files(paths)

    def add_files_dialog(self):
        paths = filedialog.askopenfilenames(title="选择要分类的文件")
        self._add_files(paths)

    # —— 统一的“默认选中第一个”工具 ——
    def _select_first_if_any(self):
        try:
            self.file_list.selection_clear(0, tk.END)
            n = self.file_list.size()
            if n > 0:
                self.file_list.selection_set(0)
                self.file_list.activate(0)
                self.file_list.see(0)
        except Exception:
            pass

    def _add_files(self, paths):
        added, exist = 0, {str(p) for p in self.files}
        for p in paths:
            pth = Path(os.path.expandvars(p))
            if pth.exists() and pth.is_file() and str(pth) not in exist:
                self.files.append(pth); exist.add(str(pth)); added += 1
        if added:
            self.refresh_files()
            self.set_status(f"已添加 {added} 个文件")

    def remove_selected(self):
        for i in reversed(self.file_list.curselection()):
            del self.files[i]
        self.refresh_files()

    def clear_files(self):
        self.files.clear()
        self.refresh_files()

    def refresh_files(self):
        self.file_list.delete(0, tk.END)
        for p in self.files:
            self.file_list.insert(tk.END, str(p))
        # 总是默认选中第一个
        self._select_first_if_any()

    def refresh_status(self):
        y, m = self.current_ym()
        YYYY, MM, YYYYMM = fmt_ym(y, m)
        if self.cfg.get("default_path_template"):
            self.root_hint.set(self.cfg["default_path_template"].format(YYYY=YYYY, MM=MM, YYYYMM=YYYYMM))
        else:
            self.root_hint.set(f"{self.cfg['out_root']}/{YYYYMM}_Unclassified")

        for i in self.tree.get_children():
            self.tree.delete(i)

        status = compute_status(self.cfg, y, m)
        order, groups = grouped_items(self.cfg["items"])
        for g in order:
            parent = self.tree.insert("", tk.END, text=f"—— {g} ——", tags=("group",), open=True)
            for it in groups[g]:
                cnt = count_for_item(self.cfg, it, y, m)
                mark = "✅" if status.get(it["key"], False) else "⬜"
                self.tree.insert(parent, tk.END, text=it["key"], values=(f"{mark}[{cnt}]",))
        self.set_status("刷新成功")

    # ---------- 撤销 / 重做（一次一步） ----------
    def cmd_undo(self):
        if not self.undo_stack:
            self.set_status("没有可撤销的操作"); return
        entry = self.undo_stack.pop()

        cur = Path(entry["current"])  # 当前真实位置（通常为目标位置）
        if not cur.exists():
            self.redo_stack.append(entry)
            self.set_status("撤销跳过：文件不存在"); return

        orig = Path(entry["orig"])
        tgt, i = orig, 1
        while tgt.exists():
            tgt = orig.with_name(f"{orig.stem}-undo{i}{orig.suffix}"); i += 1

        cur = cur.rename(tgt)
        entry["current"] = cur
        # 撤销：回到待办列表
        if all(Path(x) != cur for x in self.files):
            self.files.append(cur)
        self.redo_stack.append(entry)
        self.refresh_files()
        self.refresh_status()
        self.set_status("撤销成功：1 个")

    def cmd_redo(self):
        if not self.redo_stack:
            self.set_status("没有可重做的操作"); return
        entry = self.redo_stack.pop()

        cur = Path(entry["current"])  # 撤销后当前应位于“orig 或其 -undoK”路径
        if not cur.exists():
            self.undo_stack.append(entry)
            self.set_status("重做跳过：文件不存在"); return

        dst = Path(entry["dst"])
        old = cur                               # ← 记录移动前路径（关键）
        final = move_with_conflict(cur, dst)    # 执行重做移动
        entry["current"] = final
        # 修复点：从待办列表移除“旧路径 old”，而不是 final
        try:
            self.files = [p for p in self.files if Path(p) != old]
        except Exception:
            pass

        self.undo_stack.append(entry)
        self.refresh_files()
        self.refresh_status()
        self.set_status("重做成功：1 个")

    # ---------- 分类 ----------
    def assign(self, it: dict):
        sel = list(self.file_list.curselection())
        if not sel:
            sel = list(range(len(self.files)))
        if not sel:
            self.set_status("没有可分类的文件"); return

        y, m = self.current_ym()
        cnt_ok = cnt_skip_ext = cnt_skip_missing = 0
        for idx in sorted(sel, reverse=True):
            src = self.files[idx]
            if not Path(src).exists():
                cnt_skip_missing += 1; del self.files[idx]; continue
            exts = [e.lower() for e in it.get("exts", [])] if it.get("exts") else None
            if exts and Path(src).suffix.lower() not in exts:
                cnt_skip_ext += 1; continue

            YYYY, MM, YYYYMM = fmt_ym(y, m)
            newname = it["rename"].format(
                key=it["key"], YYYY=YYYY, MM=MM, YYYYMM=YYYYMM,
                DD=f"{datetime.now().day:02d}",
                orig="_"+safe_name(Path(src).stem), ext=Path(src).suffix
            )
            dst = target_dir(self.cfg, it, y, m) / newname
            final = move_with_conflict(Path(src), dst)

            entry = {"orig": str(src), "dst": str(dst), "current": str(final)}
            self.undo_stack.append(entry); cnt_ok += 1
            del self.files[idx]
            logging.info(f"MOVED: {src} -> {final}")

        if cnt_ok: self.redo_stack.clear()
        self.refresh_files()
        self.refresh_status()
        msg = f"{it['key']}：分类成功 {cnt_ok} 个"
        if cnt_skip_ext: msg += f"；扩展名不匹配跳过 {cnt_skip_ext} 个"
        if cnt_skip_missing: msg += f"；源文件不存在移除 {cnt_skip_missing} 个"
        self.set_status(msg)

    def open_dir_of_selected(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        iid = sel[0]
        if self.tree.get_children(iid): return   # 组标题行
        key = self.tree.item(iid, "text")        # 左列“文件类别”
        y, m = self.current_ym()
        it = next((x for x in self.cfg["items"] if x["key"] == key), None)
        if not it: return
        d = target_dir(self.cfg, it, y, m)
        try: d.mkdir(parents=True, exist_ok=True); os.startfile(str(d))
        except Exception as e: messagebox.showerror("打开失败", str(e))


# ---------- IPC：确保单窗口 & “发送到”只追加 ----------
class _Handler(socketserver.BaseRequestHandler):
    def handle(self):
        import json as _json
        try:
            data = self.request.recv(1024 * 1024)
            payload = _json.loads(data.decode("utf-8", "ignore"))
            if payload.get("cmd") == "add":
                files = payload.get("files", [])
                app: App = self.server.app_ref
                app.after(0, app._add_files, files)
        except Exception as e:
            logging.exception(f"IPC error: {e}")

class _Server(socketserver.TCPServer):
    allow_reuse_address = True
    def __init__(self, addr, handler, app_ref):
        super().__init__(addr, handler, bind_and_activate=True)
        self.app_ref = app_ref

def start_ipc(app: App):
    try:
        srv = _Server((IPC_HOST, IPC_PORT), _Handler, app)
        t = threading.Thread(target=srv.serve_forever, daemon=True); t.start(); return srv
    except OSError:
        return None

def send_to_existing(files: list[str]) -> bool:
    import json as _json
    try:
        with socket.create_connection((IPC_HOST, IPC_PORT), timeout=0.8) as s:
            s.sendall(_json.dumps({"cmd": "add", "files": files}).encode("utf-8"))
        return True
    except OSError:
        return False


# ---------- 入口 ----------
def main():
    cfg = load_config()
    cli_files = [str(Path(p)) for p in sys.argv[1:] if Path(p).exists()]

    # 若已有实例，直接把文件发过去并退出；否则启动新实例
    if cli_files and send_to_existing(cli_files):
        return

    app = App(cfg, cli_files)
    srv = start_ipc(app)

    def on_close():
        try:
            if srv: srv.shutdown(); srv.server_close()
        except Exception:
            pass
        app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_close)
    app.mainloop()

if __name__ == "__main__":
    main()
