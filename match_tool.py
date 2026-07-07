#!/usr/bin/env python3
"""
Main HUB -> DHL Match Tool
GUI: tkinter + openpyxl
Build: pyinstaller --onefile --noconsole --name "DHL_Match_Tool" match_tool.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, json, os, re, shutil
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

# ============================================================ HELPERS ============================================================

def norm(text):
    if text is None: return ""
    return " ".join(str(text).strip().split())

def fuzzy_find(headers, hint):
    hlow = hint.lower()
    for i, h in enumerate(headers):
        n = norm(h).lower()
        if not n or n == "none": continue
        if n == hlow: return i
        if hlow in n or n in hlow: return i
    return -1

def read_sheet_rows(wb, sheet_name):
    ws = wb[sheet_name]
    data = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
    if not data: return [], []
    headers = [str(c) if c is not None else "" for c in data[0]]
    rows = []
    for ws_ri in range(1, len(data)):
        r = data[ws_ri]
        if r is None: continue
        row = [c if c is not None else "" for c in r]
        if all(str(v).strip() == "" for v in row): continue
        rows.append((ws_ri, row))
    return headers, rows

def _find_candidate_sheets(wb):
    candidates = []
    for name in wb.sheetnames:
        try:
            headers = read_sheet_headers(wb, name)
            if not headers: continue
            has_lr = any("loading reference" in norm(h).lower() or "提货码" in norm(h).lower() for h in headers)
            if has_lr and len(headers) >= 20:
                candidates.append(name)
        except Exception: pass
    return candidates

def is_strikethrough(cell):
    """判断 openpyxl Cell 是否有删除线"""
    return cell.font and cell.font.strike

def is_valid_reference(val):
    v = str(val).strip()
    if not v: return False
    placeholders = {"-", "/", "none", "cancel", "need cancel", "empty", "n/a", "na", "tbd", "x", ""}
    if v.lower() in placeholders: return False
    if "TEM" not in v.upper(): return False
    return True

def read_sheet_headers(wb, sheet_name):
    """只读表头，不加载数据行"""
    ws = wb[sheet_name]
    data = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    if not data: return []
    return [str(c) if c is not None else "" for c in data[0]]

# ============================================================ DATA LOADING ============================================================

def read_source_data(wb, selected_sheets):
    """Returns (headers, combined_rows, image_map).
    combined_rows: [(sheet_name, ws_row_idx, row_data), ...]
    ws_row_idx is 0-based worksheet data row index, matching image anchor ._from.row"""
    combined_rows, image_map = [], {}
    first_headers = None
    for sn in selected_sheets:
        ws = wb[sn]
        raw = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
        if not raw: continue
        headers = [str(c) if c is not None else "" for c in raw[0]]
        if first_headers is None: first_headers = headers
        for ws_ri in range(1, len(raw)):
            r = raw[ws_ri]
            if r is None: continue
            row = [c if c is not None else "" for c in r]
            if all(str(v).strip() == "" for v in row): continue
            combined_rows.append((sn, ws_ri, row))
        if hasattr(ws, "_images") and ws._images:
            for img in ws._images:
                try:
                    fc = img.anchor._from
                    image_map.setdefault((sn, fc.row), []).append((img, fc.col))
                except Exception: pass
    return first_headers, combined_rows, image_map

def check_duplicates(selected_sheets, wb):
    if len(selected_sheets) <= 1: return {}
    all_refs = defaultdict(list)
    for sn in selected_sheets:
        ws = wb[sn]
        headers = read_sheet_headers(wb, sn)
        if not headers: continue
        li = fuzzy_find(headers, "Loading Reference")
        if li == -1: li = fuzzy_find(headers, "提货码")
        if li == -1: continue
        for ri, row in enumerate(ws.iter_rows(min_row=2)):
            cell = row[li] if li < len(row) else None
            if cell is None or cell.value is None: continue
            if is_strikethrough(cell): continue
            ref = str(cell.value).strip()
            if is_valid_reference(ref):
                all_refs[ref].append((sn, ri + 2))
    result = {}
    for ref, locs in all_refs.items():
        sheets_ = set(s[0] for s in locs)
        if len(sheets_) >= 2: result[ref] = locs
    return result

# ============================================================ IMAGE COPY ============================================================

def copy_images_for_row(source_wb, sheet_name, src_row_0based, target_ws, tgt_row_1based,
                        src_col_idx, tgt_col_idx, image_map):
    key = (sheet_name, src_row_0based)
    if key not in image_map: return 0
    copied = 0
    for img, img_col in image_map[key]:
        if img_col != src_col_idx: continue
        try:
            img_data = img.ref
            if hasattr(img_data, 'seek'):
                img_data.seek(0)
            new_img = XLImage(img_data)
            new_img.width = float(img.width) if img.width else 100
            new_img.height = float(img.height) if img.height else 100
            marker = AnchorMarker(col=tgt_col_idx, colOff=0, row=tgt_row_1based - 1, rowOff=0)
            ext = img.anchor.ext if hasattr(img.anchor, 'ext') else None
            new_img.anchor = OneCellAnchor(_from=marker, ext=ext)
            target_ws.add_image(new_img)
            copied += 1
        except Exception: pass
    return copied

# ============================================================ DEFAULT MAPPINGS ============================================================

DEFAULT_MAPPINGS = [
    ("Destination", "Destination"),
    ("AWB", "Bill of Lading No. (AWB, B/L)"),
    ("Plate/车牌", "License Plate No."),
    ("Actual Date & Time \nfor arrival", "ATA (Actual Arrival)"),
    ("Acual Date & Time \nfor arrival", "ATA (Actual Arrival)"),
    ("Planned Date & Time \nfor loading", "ETA (Planned Pickup)"),
    ("Planned Time \nfor loading", "ETA (Planned Pickup)"),
    ("Cartons", "Number of Boxes"),
    ("Note", "Note"),
    ("Weight", "Weight"),
    ("Price/车费", "Price"),
    ("Trucker/车队", "Trucker"),
]
def extract_destination(ref, dp_map=None):
    """从 Reference 解析目的地城市。兼容旧调用方式"""
    if dp_map is None:
        dp_map = DP_TO_CITY
    m = re.search(r'DP([A-Z]{3,5})', str(ref))
    if m:
        code = m.group(1)
        return dp_map.get(code, f"??{code}")
    return ""

# ============================================================ DP 映射表 (REF DP代码 -> 目的地城市) ============================================================
# Loading Reference 格式: TEM-SKY-DP{CODE}-{DATE}-{N}
# 提取 DP 后的 3-4 个字母, 映射到目的地城市名
# 映射文件存储在 app data 目录, 用户可自行编辑保存

_DEFAULT_DP_TO_CITY = {
    "BOCH": "Bochum",
    "BRUC": "Bruchsal",
    "EUTI": "Eutingen",
    "HAGE": "Hagen",
    "KONG": "Köngen",
    "KITZ": "Kitzingen",
    "LAHR": "Lahr",
    "NEWI": "Neuwied",
    "SAUL": "Saulheim",
    "SPEY": "Speyer",
    "OBER": "Obertshausen",
    "DORS": "Dorsten",
    "STAU": "Staufenberg",
    "HANN": "Hannover",
    "GREV": "Greven",
}

def _get_dp_map_path():
    """获取 DP 映射文件的存储路径 (跨平台)"""
    if os.name == "nt":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    elif os.sys.platform == "darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config"))
    folder = os.path.join(base, "DHL_Match_Tool")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "dp_map.json")

def load_dp_map():
    """加载 DP 映射, 优先从文件读取, 不存在则创建默认文件"""
    dp_path = _get_dp_map_path()
    if os.path.exists(dp_path):
        try:
            with open(dp_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # 文件不存在或损坏, 写入默认值
    save_dp_map(_DEFAULT_DP_TO_CITY)
    return dict(_DEFAULT_DP_TO_CITY)

def save_dp_map(dp_map):
    """保存 DP 映射到文件"""
    dp_path = _get_dp_map_path()
    with open(dp_path, "w", encoding="utf-8") as f:
        json.dump(dp_map, f, ensure_ascii=False, indent=2)

# 全局 DP 映射 (启动时加载)
DP_TO_CITY = load_dp_map()

def extract_dp_code(ref):
    """从 Loading Reference 提取 DP 代码.
    例: TEM-SKY-DPBOCH-030726-1 -> BOCH"""
    import re
    m = re.search(r'DP([A-Z]{3,5})', str(ref))
    return m.group(1) if m else None

def dp_to_city(dp_code):
    """DP 代码 -> 城市名"""
    return DP_TO_CITY.get(dp_code)


# ============================================================ GUI ============================================================

class MatchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Main HUB → DHL 用车记录 字段匹配工具")
        self.root.geometry("1000x800")
        self.root.minsize(850, 650)
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.src_wb = None; self.src_path = None
        self.src_headers = None; self.src_rows = None; self.src_images = None
        self.selected_sheets = ["LEJ", "LGG"]
        self.tgt_wb = None; self.tgt_path = None
        self.tgt_headers = None; self.tgt_rows = None; self.tgt_sheet = None
        self.mapping_rows = []
        self.dp_map = load_dp_map()          # 可编辑的 DP→城市映射（从文件加载）
        self._has_cross_sheet_dups = False   # 是否有跨Sheet重复

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # --- File selection ---
        ff = ttk.LabelFrame(main, text="📂 文件选择", padding=10)
        ff.pack(fill=tk.X, pady=(0, 6))

        sf = ttk.Frame(ff); sf.pack(fill=tk.X, pady=2)
        ttk.Label(sf, text="源文件 (Main HUB):", width=18).pack(side=tk.LEFT)
        self.src_path_var = tk.StringVar()
        ttk.Entry(sf, textvariable=self.src_path_var, width=50).pack(side=tk.LEFT, padx=4)
        ttk.Button(sf, text="选择...", command=self._select_src, width=8).pack(side=tk.LEFT)
        self.src_st = tk.StringVar(value="未加载")
        ttk.Label(sf, textvariable=self.src_st, foreground="gray", width=22).pack(side=tk.LEFT, padx=8)

        tf = ttk.Frame(ff); tf.pack(fill=tk.X, pady=2)
        ttk.Label(tf, text="目标文件 (DHL 记录):", width=18).pack(side=tk.LEFT)
        self.tgt_path_var = tk.StringVar()
        ttk.Entry(tf, textvariable=self.tgt_path_var, width=50).pack(side=tk.LEFT, padx=4)
        ttk.Button(tf, text="选择...", command=self._select_tgt, width=8).pack(side=tk.LEFT)
        self.tgt_st = tk.StringVar(value="未加载")
        ttk.Label(tf, textvariable=self.tgt_st, foreground="gray", width=22).pack(side=tk.LEFT, padx=8)

        # --- Sheet selection ---
        self.shf = ttk.LabelFrame(main, text="📑 源数据 Sheet（可多选，数据自动合并）", padding=10)
        self.shf.pack(fill=tk.X, pady=(0, 6))
        self.shcf = ttk.Frame(self.shf); self.shcf.pack(fill=tk.X)

        # --- Match config ---
        cf = ttk.LabelFrame(main, text="🔗 匹配配置", padding=10)
        cf.pack(fill=tk.X, pady=(0, 6))

        kf = ttk.Frame(cf); kf.pack(fill=tk.X, pady=2)
        ttk.Label(kf, text="源匹配键:").pack(side=tk.LEFT)
        self.src_key_var = tk.StringVar()
        self.src_key_cb = ttk.Combobox(kf, textvariable=self.src_key_var, width=28, state="readonly")
        self.src_key_cb.pack(side=tk.LEFT, padx=4)
        ttk.Label(kf, text=" ⬌ ").pack(side=tk.LEFT, padx=8)
        ttk.Label(kf, text="目标匹配键:").pack(side=tk.LEFT)
        self.tgt_key_var = tk.StringVar()
        self.tgt_key_cb = ttk.Combobox(kf, textvariable=self.tgt_key_var, width=28, state="readonly")
        self.tgt_key_cb.pack(side=tk.LEFT, padx=4)

        # Mapping header
        mh = ttk.Frame(cf); mh.pack(fill=tk.X, pady=(10, 2))
        ttk.Label(mh, text="#", width=3).pack(side=tk.LEFT)
        ttk.Label(mh, text="源字段 (Main HUB)", width=38).pack(side=tk.LEFT, padx=4)
        ttk.Label(mh, text="→", width=3).pack(side=tk.LEFT)
        ttk.Label(mh, text="目标字段 (DHL)", width=38).pack(side=tk.LEFT, padx=4)

        # Scrollable mapping area
        map_container = ttk.Frame(cf)
        map_container.pack(fill=tk.BOTH, expand=True, pady=2)
        self.map_canvas = tk.Canvas(map_container, height=140, highlightthickness=0, bg="#2a2a3e")
        self.map_scrollbar = ttk.Scrollbar(map_container, orient=tk.VERTICAL, command=self.map_canvas.yview)
        self.map_canvas.configure(yscrollcommand=self.map_scrollbar.set)
        self.map_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.map_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.map_inner = ttk.Frame(self.map_canvas)
        self.map_canvas_window = self.map_canvas.create_window((0, 0), window=self.map_inner, anchor="nw")
        self.map_inner.bind("<Configure>", self._on_map_inner_configure)
        self.map_canvas.bind("<Configure>", self._on_map_canvas_configure)

        # Mapping buttons + status
        bf = ttk.Frame(cf); bf.pack(fill=tk.X, pady=4)
        ttk.Button(bf, text="+ 添加映射", command=self._add_mapping_row).pack(side=tk.LEFT)
        ttk.Button(bf, text="🔄 重新自动检测", command=self._auto_map_fields).pack(side=tk.LEFT, padx=8)
        ttk.Button(bf, text="📝 编辑 DP 映射", command=self._edit_dp_map).pack(side=tk.LEFT, padx=4)
        self.detect_st = tk.StringVar()
        ttk.Label(bf, textvariable=self.detect_st, foreground="#81c784").pack(side=tk.LEFT, padx=12)
        self.dup_st = tk.StringVar()
        ttk.Label(bf, textvariable=self.dup_st, foreground="#ff6d00",
                  font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, padx=12)

        # --- Action buttons ---
        af = ttk.Frame(main); af.pack(fill=tk.X, pady=(0, 6))
        self.start_btn = ttk.Button(af, text="🚀 开始匹配", command=self._start_matching, state="disabled")
        self.start_btn.pack()
        self.prog_var = tk.StringVar()
        ttk.Label(af, textvariable=self.prog_var, foreground="#90caf9").pack(pady=2)
        self.result_bar = ttk.Frame(af)
        self.result_bar.pack(pady=4)
        self.result_bar.pack_forget()

        # --- Log ---
        lf = ttk.LabelFrame(main, text="📋 运行日志", padding=6)
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(lf, height=10, font=("Consolas", 10), wrap=tk.WORD,
                                bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
                                relief=tk.FLAT, borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        sb = ttk.Scrollbar(lf, command=self.log_text.yview)
        sb.pack(fill=tk.Y, side=tk.RIGHT)
        self.log_text.configure(yscrollcommand=sb.set)
        for tag, color in [("info", "#90caf9"), ("good", "#81c784"), ("warn", "#ff6d00"), ("err", "#ef5350")]:
            self.log_text.tag_configure(tag, foreground=color)

        self._update_sheet_checkboxes([])

    def _on_map_inner_configure(self, event):
        self.map_canvas.configure(scrollregion=self.map_canvas.bbox("all"))
    def _on_map_canvas_configure(self, event):
        self.map_canvas.itemconfig(self.map_canvas_window, width=event.width)

    def log(self, msg, tag="info"):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # ==================== FILE LOADING ====================

    def _select_src(self):
        path = filedialog.askopenfilename(title="选择 Main HUB XLSX 文件", filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if not path: return
        self.src_path_var.set(path); self.src_path = path
        try:
            self.src_wb = openpyxl.load_workbook(path)
            self.log(f"已加载源文件: {os.path.basename(path)}", "good")
            candidates = _find_candidate_sheets(self.src_wb)
            self.log(f"可用 Sheet: {', '.join(candidates)}", "info")
            self.selected_sheets = [s for s in ["LEJ", "LGG"] if s in candidates]
            if not self.selected_sheets and candidates: self.selected_sheets = [candidates[0]]
            self._update_sheet_checkboxes(candidates)
            self._load_source()
        except Exception as e:
            self.log(f"加载失败: {e}", "err")
            messagebox.showerror("Error", str(e))

    def _select_tgt(self):
        path = filedialog.askopenfilename(title="选择 DHL 用车记录 XLSX 文件", filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if not path: return
        self.tgt_path_var.set(path); self.tgt_path = path
        try:
            self.tgt_wb = openpyxl.load_workbook(path)
            self.tgt_sheet = self.tgt_wb.sheetnames[0]
            self.tgt_headers, self.tgt_rows = read_sheet_rows(self.tgt_wb, self.tgt_sheet)
            self.log(f"已加载目标文件: {os.path.basename(path)} ({len(self.tgt_rows)} 行)", "good")
            self._refresh_ui()
        except Exception as e:
            self.log(f"加载失败: {e}", "err")
            messagebox.showerror("Error", str(e))

    def _update_sheet_checkboxes(self, sheets):
        for w in self.shcf.winfo_children(): w.destroy()
        self.sheet_vars = {}
        for s in sheets:
            v = tk.BooleanVar(value=s in self.selected_sheets)
            self.sheet_vars[s] = v
            ttk.Checkbutton(self.shcf, text=s, variable=v, command=self._on_sheets_changed).pack(side=tk.LEFT, padx=4, pady=4)
        if not sheets:
            ttk.Label(self.shcf, text="请先加载源文件", foreground="gray").pack(side=tk.LEFT)

    def _on_sheets_changed(self):
        self.selected_sheets = [s for s, v in self.sheet_vars.items() if v.get()]
        if self.selected_sheets: self._load_source()
        else: self.src_st.set("未选择 Sheet"); self._refresh_ui()

    def _load_source(self):
        if not self.src_wb or not self.selected_sheets: return
        try:
            self.src_headers, self.src_rows, self.src_images = read_source_data(self.src_wb, self.selected_sheets)
            total = len(self.src_rows)
            self.src_st.set(f"{', '.join(self.selected_sheets)} - {total} 行")

            self._has_cross_sheet_dups = False
            self.dup_st.set("")
            dups = check_duplicates(self.selected_sheets, self.src_wb)
            if dups:
                self._has_cross_sheet_dups = True
                self.dup_st.set(f"⚠ 跨 Sheet 重复: {len(dups)} 个 Reference")
                self.log(f"⚠ 检测到 {len(dups)} 个跨 Sheet 重复的 Loading Reference:", "warn")
                for ref, locs in list(dups.items())[:10]:
                    self.log(f"  · {ref} → {' / '.join(f'{s}第{r}行' for s,r in locs)}", "warn")
                if len(dups) > 10: self.log(f"  ... 还有 {len(dups)-10} 个", "warn")
                self.log("  请先在源文件中解决重复问题，再重新匹配。", "warn")
            else:
                self.log("✓ 无跨 Sheet 重复", "good")

            self.log(f"源数据: {total} rows, {len(self.src_images)} 图片位置", "info")
            self._refresh_ui()
        except Exception as e:
            self.log(f"读取源数据失败: {e}", "err")

    # ==================== UI REFRESH ====================

    def _refresh_ui(self):
        ok = bool(self.src_headers and self.tgt_headers)
        self.start_btn.configure(state="normal" if ok else "disabled")
        if not ok: return
        src_opts = [norm(h) for h in self.src_headers if h and h != "None"]
        tgt_opts = [norm(h) for h in self.tgt_headers if h and h != "None"]
        self.src_key_cb["values"] = src_opts; self.tgt_key_cb["values"] = tgt_opts
        ds = next((h for h in src_opts if "loading reference" in h.lower() or "提货码" in h.lower()), src_opts[0] if src_opts else "")
        dt = next((h for h in tgt_opts if "dhl reference" in h.lower()), tgt_opts[0] if tgt_opts else "")
        self.src_key_var.set(ds); self.tgt_key_var.set(dt)
        self._auto_map_fields()

    def _auto_map_fields(self):
        for w in self.map_inner.winfo_children(): w.destroy()
        self.mapping_rows.clear()
        sk = self.src_key_var.get(); tk = self.tgt_key_var.get()
        found = 0
        for sh, th in DEFAULT_MAPPINGS:
            si = fuzzy_find(self.src_headers, sh); ti = fuzzy_find(self.tgt_headers, th)
            if si != -1 and ti != -1:
                sn = norm(self.src_headers[si]); tn = norm(self.tgt_headers[ti])
                if sn == sk or tn == tk: continue
                self._add_mapping_row(sn, tn); found += 1
        self.detect_st.set(f"Auto-detected {found} mapping 对映射")

    def _add_mapping_row(self, sv="", tv=""):
        rf = ttk.Frame(self.map_inner); rf.pack(fill=tk.X, pady=1)
        idx = len(self.mapping_rows) + 1
        ttk.Label(rf, text=str(idx), width=3).pack(side=tk.LEFT)
        src_var = tk.StringVar(value=sv)
        src_cb = ttk.Combobox(rf, textvariable=src_var, width=36, state="readonly")
        src_cb["values"] = [norm(h) for h in self.src_headers if h and h != "None"] if self.src_headers else []
        src_cb.pack(side=tk.LEFT, padx=4)
        ttk.Label(rf, text="→", width=3).pack(side=tk.LEFT)
        tgt_var = tk.StringVar(value=tv)
        tgt_cb = ttk.Combobox(rf, textvariable=tgt_var, width=36, state="readonly")
        tgt_cb["values"] = [norm(h) for h in self.tgt_headers if h and h != "None"] if self.tgt_headers else []
        tgt_cb.pack(side=tk.LEFT, padx=4)
        def rm():
            if len(self.mapping_rows) <= 1: return
            rf.destroy(); self.mapping_rows.remove(rd); self._renumber()
        rd = {"frame": rf, "src_var": src_var, "tgt_var": tgt_var}
        ttk.Button(rf, text="X", width=3, command=rm).pack(side=tk.LEFT)
        self.mapping_rows.append(rd)
        self.map_canvas.after(50, lambda: self.map_canvas.yview_moveto(1.0))

    def _renumber(self):
        for i, mr in enumerate(self.mapping_rows):
            ch = mr["frame"].winfo_children()
            if ch: ch[0].configure(text=str(i + 1))

    # ==================== MATCHING ====================

    def _edit_dp_map(self):
        """打开 DP 映射编辑对话框"""
        dp_win = tk.Toplevel(self.root)
        dp_win.title("编辑 DP 映射表")
        dp_win.geometry("500x500")
        dp_win.transient(self.root)
        dp_win.grab_set()

        # 读当前映射
        current = load_dp_map()

        # 主框架
        mf = ttk.Frame(dp_win, padding=10)
        mf.pack(fill=tk.BOTH, expand=True)

        # 列标题
        hf = ttk.Frame(mf)
        hf.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(hf, text="DP 代码", width=14, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(hf, text="目的地城市", width=20, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT)

        # 可滚动列表
        lf = ttk.Frame(mf)
        lf.pack(fill=tk.BOTH, expand=True, pady=4)
        tree = ttk.Treeview(lf, columns=("code", "city"), show="headings", height=14)
        tree.heading("code", text="DP 代码")
        tree.heading("city", text="目的地城市")
        tree.column("code", width=120)
        tree.column("city", width=200)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tsb = ttk.Scrollbar(lf, orient=tk.VERTICAL, command=tree.yview)
        tsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=tsb.set)

        # 填充数据
        for code, city in sorted(current.items()):
            tree.insert("", tk.END, values=(code, city))

        # 编辑区域
        ef = ttk.Frame(mf)
        ef.pack(fill=tk.X, pady=8)
        ttk.Label(ef, text="DP 代码:").pack(side=tk.LEFT, padx=(0, 4))
        code_var = tk.StringVar()
        code_entry = ttk.Entry(ef, textvariable=code_var, width=12)
        code_entry.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(ef, text="城市:").pack(side=tk.LEFT, padx=(0, 4))
        city_var = tk.StringVar()
        city_entry = ttk.Entry(ef, textvariable=city_var, width=18)
        city_entry.pack(side=tk.LEFT)

        # 按钮
        bf = ttk.Frame(mf)
        bf.pack(fill=tk.X, pady=4)

        def refresh_tree():
            for item in tree.get_children():
                tree.delete(item)
            for code, city in sorted(current.items()):
                tree.insert("", tk.END, values=(code, city))
            status_var.set(f"共 {len(current)} 条")

        def on_tree_select(event):
            sel = tree.selection()
            if sel:
                vals = tree.item(sel[0], "values")
                code_var.set(vals[0])
                city_var.set(vals[1])

        tree.bind("<<TreeviewSelect>>", on_tree_select)

        def add_or_update():
            code = code_var.get().strip().upper()
            city = city_var.get().strip()
            if not code or not city:
                messagebox.showwarning("提示", "DP 代码和城市名不能为空", parent=dp_win)
                return
            current[code] = city
            save_dp_map(current)
            refresh_tree()
            code_var.set("")
            city_var.set("")

        def delete_selected():
            sel = tree.selection()
            if not sel: return
            vals = tree.item(sel[0], "values")
            code = vals[0]
            if messagebox.askyesno("确认删除", f"删除 {code} → {vals[1]}?", parent=dp_win):
                if code in current:
                    del current[code]
                    save_dp_map(current)
                    refresh_tree()
                    code_var.set("")
                    city_var.set("")

        ttk.Button(bf, text="➕ 添加/更新", command=add_or_update).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="🗑 删除选中", command=delete_selected).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="↩ 恢复默认", command=lambda: [current.clear(), current.update(_DEFAULT_DP_TO_CITY), save_dp_map(current), refresh_tree()]).pack(side=tk.LEFT, padx=4)

        status_var = tk.StringVar(value=f"共 {len(current)} 条  |  路径: {_get_dp_map_path()}")
        ttk.Label(mf, textvariable=status_var, foreground="gray", font=("TkDefaultFont", 9)).pack(pady=4)

        ttk.Button(mf, text="关闭", command=dp_win.destroy).pack(pady=4)

    def _start_matching(self):
        # 每次匹配前重新加载 DP 映射 (用户可能已编辑文件)
        global DP_TO_CITY
        DP_TO_CITY = load_dp_map()
        if not self.mapping_rows:
            messagebox.showwarning("Warning", "Please configure at least one field mapping pair"); return
        if self._has_cross_sheet_dups:
            messagebox.showwarning("无法匹配", f"{self.dup_st.get()}\n\n请先解决跨 Sheet 重复问题，再开始匹配。"); return
        self.start_btn.configure(state="disabled", text="匹配中...")
        self.prog_var.set("正在处理...")
        self.result_bar.pack_forget()
        threading.Thread(target=self._do_match, daemon=True).start()

    def _do_match(self):
        try:
            # 复用 _select_src 阶段已打开的源文件，不再重复加载
            _src_wb = self.src_wb
            _headers, _rows, _images = read_source_data(_src_wb, self.selected_sheets)

            # 输出文件：先复制目标文件，保留其中已有的图片（邮件截图等）
            out = os.path.splitext(self.tgt_path)[0] + "_matched.xlsx"
            shutil.copy2(self.tgt_path, out)
            os.chmod(out, 0o666)

            skn = self.src_key_var.get(); tkn = self.tgt_key_var.get()
            ski = next((i for i, h in enumerate(_headers) if norm(h) == skn), -1)
            tki = next((i for i, h in enumerate(self.tgt_headers) if norm(h) == tkn), -1)
            if ski == -1 or tki == -1: self._uilog("❌ 匹配键未找到", "err"); self._uidone(); return

            self._uilog(f"源键: {skn} (列{ski})  |  目标键: {tkn} (列{tki})", "info")

            cmaps = []
            for mr in self.mapping_rows:
                si = next((i for i, h in enumerate(_headers) if norm(h) == mr["src_var"].get()), -1)
                ti = next((i for i, h in enumerate(self.tgt_headers) if norm(h) == mr["tgt_var"].get()), -1)
                if si != -1 and ti != -1: cmaps.append((si, ti, mr["src_var"].get(), mr["tgt_var"].get()))
            self._uilog(f"字段映射: {len(cmaps)} 对", "info")
            for si, ti, sn, tn in cmaps: self._uilog(f"  {sn} → {tn}", "info")

            # 源索引 — 流式遍历，跳过删除线行；遇到真实重复直接终止
            idx = {}
            for sn_ in self.selected_sheets:
                ws = _src_wb[sn_]
                for ri_, row in enumerate(ws.iter_rows(min_row=2)):
                    ref_cell = row[ski] if ski < len(row) else None
                    if ref_cell is None or ref_cell.value is None: continue
                    if is_strikethrough(ref_cell): continue
                    k = str(ref_cell.value).strip()
                    if k and is_valid_reference(k):
                        if k in idx:
                            prev_sn, prev_ri = idx[k][1], idx[k][2]
                            self._uilog(
                                f"❌ 重复 Reference: {k}\n"
                                f"    位置1: {prev_sn} 第{prev_ri + 2}行\n"
                                f"    位置2: {sn_} 第{ri_ + 2}行\n"
                                f"  请解决重复后再匹配。", "err")
                            self._uidone()
                            self.root.after(0, lambda: messagebox.showerror(
                                "匹配失败",
                                f"源数据中存在重复的 Loading Reference:\n\n{k}\n\n"
                                f"位置1: {prev_sn} 第{prev_ri + 2}行\n"
                                f"位置2: {sn_} 第{ri_ + 2}行\n\n"
                                f"请先在源文件中解决重复问题，再重新匹配。"))
                            return
                        row_vals = [c.value if c.value is not None else "" for c in row]
                        idx[k] = (row_vals, sn_, ri_)
            self._uilog(f"源索引: {len(idx)} 条唯一 Key", "info")

            # 匹配 — 打开复制后的目标文件进行修改
            matched = unmatched = skipped = 0
            self.tgt_wb = openpyxl.load_workbook(out)
            tws = self.tgt_wb[self.tgt_sheet]
            _, tgt_rows_new = read_sheet_rows(self.tgt_wb, self.tgt_sheet)

            note_si = next((i for i, h in enumerate(_headers) if norm(h).lower() == "note"), -1)
            note_ti = next((i for i, h in enumerate(self.tgt_headers) if norm(h).lower() == "note"), -1)
            note_mapped = any(si == note_si and ti == note_ti for si, ti, _, _ in cmaps)

            # Destination 字段不从源文件取值，改为从 Reference 解析
            dest_ti = next((ti for si, ti, sn, tn in cmaps if "destination" in tn.lower()), -1)

            for ti_, (tgt_ri, row) in enumerate(tgt_rows_new):
                k = str(row[tki]).strip() if tki < len(row) else ""
                if not k: skipped += 1; continue
                if k not in idx: unmatched += 1; continue
                srow, shn, sri = idx[k]
                matched += 1
                for sc, tc, sn, tn in cmaps:
                    if tc == dest_ti:
                        v = extract_destination(k, self.dp_map)
                    else:
                        v = srow[sc] if sc < len(srow) else ""
                    if v is not None and str(v).strip():
                        row[tc] = v
                        tws.cell(row=tgt_ri + 1, column=tc + 1).value = v

                if note_mapped and note_si != -1 and note_ti != -1:
                    copy_images_for_row(_src_wb, shn, sri, tws, tgt_ri + 1,
                                       note_si, note_ti, _images)

            self._uilog("", ""); self._uilog("=" * 40, "info")
            self._uilog(f"  ✅ 匹配成功: {matched} 行", "good")
            if unmatched: self._uilog(f"  ⚠ 未匹配: {unmatched} 行", "warn")
            if skipped: self._uilog(f"  ⊝ 跳过空键: {skipped} 行", "info")
            self._uilog("=" * 40, "info")

            self.tgt_wb.save(out)
            self._uilog(f"📥 输出: {os.path.basename(out)}", "good")
            self._output_path = out
            self._uidone()
            # 显示结果栏
            self.root.after(0, lambda: self._show_result_bar(out))
            self.root.after(0, lambda: messagebox.showinfo(
                "完成",
                f"✅ 成功: {matched} 行  |  ⚠ 未匹配: {unmatched} 行  |  ⊝ 跳过: {skipped} 行\n"
                f"\n输出: {os.path.basename(out)}"))

        except Exception as e:
            import traceback
            self._uilog(f"Error: {e}", "err")
            self._uilog(traceback.format_exc()[-600:], "err")
            self._uidone()

    # ==================== RESULT BAR ====================

    def _show_result_bar(self, output_path):
        for w in self.result_bar.winfo_children(): w.destroy()
        ttk.Label(self.result_bar, text="📥 结果就绪：",
                  font=("TkDefaultFont", 10, "bold"), foreground="#81c784").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(self.result_bar, text="📂 打开文件位置",
                   command=lambda: self._open_folder(output_path)).pack(side=tk.LEFT, padx=4)
        ttk.Button(self.result_bar, text="📄 打开匹配文件",
                   command=lambda: self._open_file(output_path)).pack(side=tk.LEFT, padx=4)
        self.result_bar.pack(pady=4)

    @staticmethod
    def _open_file(path):
        import subprocess, sys
        if sys.platform == "darwin": subprocess.Popen(["open", path])
        elif sys.platform == "win32": os.startfile(path)
        else: subprocess.Popen(["xdg-open", path])

    @staticmethod
    def _open_folder(path):
        import subprocess, sys
        folder = os.path.dirname(os.path.abspath(path))
        if sys.platform == "darwin": subprocess.Popen(["open", folder])
        elif sys.platform == "win32": os.startfile(folder)
        else: subprocess.Popen(["xdg-open", folder])

    def _uilog(self, m, t="info"): self.root.after(0, lambda: self.log(m, t))
    def _uidone(self):
        self.root.after(0, lambda: self.start_btn.configure(state="normal", text="🚀 开始匹配"))
        self.root.after(0, lambda: self.prog_var.set(""))

# ============================================================ MAIN ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    MatchApp(root)
    root.mainloop()
