#!/usr/bin/env python3
"""
Main HUB -> DHL Match Tool
GUI: tkinter + openpyxl
Build: pyinstaller --onefile --noconsole --name "DHL_Match_Tool" match_tool.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, io, tempfile
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
            headers, _ = read_sheet_rows(wb, name)
            if not headers: continue
            has_lr = any("loading reference" in norm(h).lower() or "提货码" in norm(h).lower() for h in headers)
            if has_lr and len(headers) >= 20:
                candidates.append(name)
        except Exception: pass
    return candidates

def is_valid_reference(val):
    v = str(val).strip()
    if not v: return False
    placeholders = {"-", "/", "none", "cancel", "need cancel", "empty", "n/a", "na", "tbd", "x", ""}
    if v.lower() in placeholders: return False
    if "TEM" not in v.upper(): return False
    return True

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
        headers, rows = read_sheet_rows(wb, sn)
        if not headers: continue
        li = fuzzy_find(headers, "Loading Reference")
        if li == -1: li = fuzzy_find(headers, "提货码")
        if li == -1: continue
        for ws_ri, row in rows:
            ref = str(row[li]).strip() if li < len(row) else ""
            if is_valid_reference(ref):
                all_refs[ref].append((sn, ws_ri + 1))
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
            raw = img._data()
            new_img = XLImage(io.BytesIO(raw))
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
    ("Actual Date & Time \nfor arrival", "ATA (Actual Arrival"),
    ("Acual Date & Time \nfor arrival", "ATA (Actual Arrival"),
    ("Planned Date & Time \nfor loading", "ETA (Planned Pickup"),
    ("Planned Time \nfor loading", "ETA (Planned Pickup"),
    ("Cartons", "Number of Boxes"),
    ("Note", "Note"),
    ("Weight", "Weight"),
    ("Price/车费", "Price"),
    ("Trucker/车队", "Trucker"),
]

# ============================================================ GUI ============================================================

class MatchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Main HUB -> DHL Match Tool")
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
        self._has_cross_sheet_dups = False

        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # --- File selection ---
        ff = ttk.LabelFrame(main, text="File Selection", padding=10)
        ff.pack(fill=tk.X, pady=(0, 6))

        sf = ttk.Frame(ff); sf.pack(fill=tk.X, pady=2)
        ttk.Label(sf, text="Source (Main HUB):", width=18).pack(side=tk.LEFT)
        self.src_path_var = tk.StringVar()
        ttk.Entry(sf, textvariable=self.src_path_var, width=50).pack(side=tk.LEFT, padx=4)
        ttk.Button(sf, text="Browse...", command=self._select_src, width=8).pack(side=tk.LEFT)
        self.src_st = tk.StringVar(value="Not loaded")
        ttk.Label(sf, textvariable=self.src_st, foreground="gray", width=22).pack(side=tk.LEFT, padx=8)

        tf = ttk.Frame(ff); tf.pack(fill=tk.X, pady=2)
        ttk.Label(tf, text="Target (DHL Record):", width=18).pack(side=tk.LEFT)
        self.tgt_path_var = tk.StringVar()
        ttk.Entry(tf, textvariable=self.tgt_path_var, width=50).pack(side=tk.LEFT, padx=4)
        ttk.Button(tf, text="Browse...", command=self._select_tgt, width=8).pack(side=tk.LEFT)
        self.tgt_st = tk.StringVar(value="Not loaded")
        ttk.Label(tf, textvariable=self.tgt_st, foreground="gray", width=22).pack(side=tk.LEFT, padx=8)

        # --- Sheet selection ---
        self.shf = ttk.LabelFrame(main, text="Source Sheets (check to include)", padding=10)
        self.shf.pack(fill=tk.X, pady=(0, 6))
        self.shcf = ttk.Frame(self.shf); self.shcf.pack(fill=tk.X)

        # --- Match config ---
        cf = ttk.LabelFrame(main, text="Match Configuration", padding=10)
        cf.pack(fill=tk.X, pady=(0, 6))

        kf = ttk.Frame(cf); kf.pack(fill=tk.X, pady=2)
        ttk.Label(kf, text="Source Key:").pack(side=tk.LEFT)
        self.src_key_var = tk.StringVar()
        self.src_key_cb = ttk.Combobox(kf, textvariable=self.src_key_var, width=28, state="readonly")
        self.src_key_cb.pack(side=tk.LEFT, padx=4)
        ttk.Label(kf, text=" <=> ").pack(side=tk.LEFT, padx=8)
        ttk.Label(kf, text="Target Key:").pack(side=tk.LEFT)
        self.tgt_key_var = tk.StringVar()
        self.tgt_key_cb = ttk.Combobox(kf, textvariable=self.tgt_key_var, width=28, state="readonly")
        self.tgt_key_cb.pack(side=tk.LEFT, padx=4)

        # Mapping header
        mh = ttk.Frame(cf); mh.pack(fill=tk.X, pady=(10, 2))
        ttk.Label(mh, text="#", width=3).pack(side=tk.LEFT)
        ttk.Label(mh, text="Source Field", width=38).pack(side=tk.LEFT, padx=4)
        ttk.Label(mh, text="->", width=3).pack(side=tk.LEFT)
        ttk.Label(mh, text="Target Field", width=38).pack(side=tk.LEFT, padx=4)

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
        ttk.Button(bf, text="+ Add Mapping", command=self._add_mapping_row).pack(side=tk.LEFT)
        ttk.Button(bf, text="Auto-Detect", command=self._auto_map_fields).pack(side=tk.LEFT, padx=8)
        self.detect_st = tk.StringVar()
        ttk.Label(bf, textvariable=self.detect_st, foreground="#81c784").pack(side=tk.LEFT, padx=12)
        self.dup_st = tk.StringVar()
        ttk.Label(bf, textvariable=self.dup_st, foreground="#ff6d00",
                  font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, padx=12)

        # --- Action buttons ---
        af = ttk.Frame(main); af.pack(fill=tk.X, pady=(0, 6))
        self.start_btn = ttk.Button(af, text="Start Matching", command=self._start_matching, state="disabled")
        self.start_btn.pack()
        self.prog_var = tk.StringVar()
        ttk.Label(af, textvariable=self.prog_var, foreground="#90caf9").pack(pady=2)
        self.result_bar = ttk.Frame(af)
        self.result_bar.pack(pady=4)
        self.result_bar.pack_forget()

        # --- Log ---
        lf = ttk.LabelFrame(main, text="Log", padding=6)
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
        path = filedialog.askopenfilename(title="Select Main HUB XLSX", filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if not path: return
        self.src_path_var.set(path); self.src_path = path
        try:
            self.src_wb = openpyxl.load_workbook(path)
            self.log(f"Loaded source: {os.path.basename(path)}", "good")
            candidates = _find_candidate_sheets(self.src_wb)
            self.log(f"Available sheets: {', '.join(candidates)}", "info")
            self.selected_sheets = [s for s in ["LEJ", "LGG"] if s in candidates]
            if not self.selected_sheets and candidates: self.selected_sheets = [candidates[0]]
            self._update_sheet_checkboxes(candidates)
            self._load_source()
        except Exception as e:
            self.log(f"Load failed: {e}", "err")
            messagebox.showerror("Error", str(e))

    def _select_tgt(self):
        path = filedialog.askopenfilename(title="Select DHL Record XLSX", filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if not path: return
        self.tgt_path_var.set(path); self.tgt_path = path
        try:
            self.tgt_wb = openpyxl.load_workbook(path)
            self.tgt_sheet = self.tgt_wb.sheetnames[0]
            self.tgt_headers, self.tgt_rows = read_sheet_rows(self.tgt_wb, self.tgt_sheet)
            self.log(f"Loaded target: {os.path.basename(path)} ({len(self.tgt_rows)} rows)", "good")
            self._refresh_ui()
        except Exception as e:
            self.log(f"Load failed: {e}", "err")
            messagebox.showerror("Error", str(e))

    def _update_sheet_checkboxes(self, sheets):
        for w in self.shcf.winfo_children(): w.destroy()
        self.sheet_vars = {}
        for s in sheets:
            v = tk.BooleanVar(value=s in self.selected_sheets)
            self.sheet_vars[s] = v
            ttk.Checkbutton(self.shcf, text=s, variable=v, command=self._on_sheets_changed).pack(side=tk.LEFT, padx=4, pady=4)
        if not sheets:
            ttk.Label(self.shcf, text="Load source file first", foreground="gray").pack(side=tk.LEFT)

    def _on_sheets_changed(self):
        self.selected_sheets = [s for s, v in self.sheet_vars.items() if v.get()]
        if self.selected_sheets: self._load_source()
        else: self.src_st.set("No sheet selected"); self._refresh_ui()

    def _load_source(self):
        if not self.src_wb or not self.selected_sheets: return
        try:
            self.src_headers, self.src_rows, self.src_images = read_source_data(self.src_wb, self.selected_sheets)
            total = len(self.src_rows)
            self.src_st.set(f"{', '.join(self.selected_sheets)} - {total} rows")

            self._has_cross_sheet_dups = False
            self.dup_st.set("")
            dups = check_duplicates(self.selected_sheets, self.src_wb)
            if dups:
                self._has_cross_sheet_dups = True
                self.dup_st.set(f"WARNING: {len(dups)} cross-sheet duplicate Refs")
                self.log(f"WARNING: {len(dups)} cross-sheet duplicate Loading References:", "warn")
                for ref, locs in list(dups.items())[:10]:
                    self.log(f"  {ref} -> {' / '.join(f'{s} row {r}' for s,r in locs)}", "warn")
                if len(dups) > 10: self.log(f"  ... and {len(dups)-10} more", "warn")
            else:
                self.log("No cross-sheet duplicates", "good")

            self.log(f"Source: {total} rows, {len(self.src_images)} image locations", "info")
            self._refresh_ui()
        except Exception as e:
            self.log(f"Read source failed: {e}", "err")

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
        self.detect_st.set(f"Auto-detected {found} mapping pairs")

    def _add_mapping_row(self, sv="", tv=""):
        rf = ttk.Frame(self.map_inner); rf.pack(fill=tk.X, pady=1)
        idx = len(self.mapping_rows) + 1
        ttk.Label(rf, text=str(idx), width=3).pack(side=tk.LEFT)
        src_var = tk.StringVar(value=sv)
        src_cb = ttk.Combobox(rf, textvariable=src_var, width=36, state="readonly")
        src_cb["values"] = [norm(h) for h in self.src_headers if h and h != "None"] if self.src_headers else []
        src_cb.pack(side=tk.LEFT, padx=4)
        ttk.Label(rf, text="->", width=3).pack(side=tk.LEFT)
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

    def _start_matching(self):
        if not self.mapping_rows:
            messagebox.showwarning("Warning", "Please configure at least one field mapping pair"); return
        if self._has_cross_sheet_dups:
            messagebox.showwarning("Cannot Match", f"{self.dup_st.get()}\n\nPlease resolve cross-sheet duplicates first."); return
        self.start_btn.configure(state="disabled", text="Matching...")
        self.prog_var.set("Processing...")
        self.result_bar.pack_forget()
        threading.Thread(target=self._do_match, daemon=True).start()

    def _do_match(self):
        try:
            _src_wb = openpyxl.load_workbook(self.src_path)
            _headers, _rows, _images = read_source_data(_src_wb, self.selected_sheets)

            skn = self.src_key_var.get(); tkn = self.tgt_key_var.get()
            ski = next((i for i, h in enumerate(_headers) if norm(h) == skn), -1)
            tki = next((i for i, h in enumerate(self.tgt_headers) if norm(h) == tkn), -1)
            if ski == -1 or tki == -1: self._uilog("Key column not found", "err"); self._uidone(); return

            self._uilog(f"Source key: {skn} (col {ski})  |  Target key: {tkn} (col {tki})", "info")

            cmaps = []
            for mr in self.mapping_rows:
                si = next((i for i, h in enumerate(_headers) if norm(h) == mr["src_var"].get()), -1)
                ti = next((i for i, h in enumerate(self.tgt_headers) if norm(h) == mr["tgt_var"].get()), -1)
                if si != -1 and ti != -1: cmaps.append((si, ti, mr["src_var"].get(), mr["tgt_var"].get()))
            self._uilog(f"Field mappings: {len(cmaps)} pairs", "info")
            for si, ti, sn, tn in cmaps: self._uilog(f"  {sn} -> {tn}", "info")

            idx = {}
            for sn_, ri_, row in _rows:
                k = str(row[ski]).strip() if ski < len(row) else ""
                if k and is_valid_reference(k): idx[k] = (row, sn_, ri_)
            self._uilog(f"Source index: {len(idx)} unique keys", "info")

            matched = unmatched = skipped = img_copied = 0
            tws = self.tgt_wb[self.tgt_sheet]

            note_si = next((i for i, h in enumerate(_headers) if norm(h).lower() == "note"), -1)
            note_ti = next((i for i, h in enumerate(self.tgt_headers) if norm(h).lower() == "note"), -1)
            note_mapped = any(si == note_si and ti == note_ti for si, ti, _, _ in cmaps)

            for ti_, (tgt_ri, row) in enumerate(self.tgt_rows):
                k = str(row[tki]).strip() if tki < len(row) else ""
                if not k: skipped += 1; continue
                if k not in idx: unmatched += 1; continue
                srow, shn, sri = idx[k]
                matched += 1
                for sc, tc, _, _ in cmaps:
                    v = srow[sc] if sc < len(srow) else ""
                    if v is not None and str(v).strip():
                        row[tc] = v
                        tws.cell(row=tgt_ri + 1, column=tc + 1).value = v

                if note_mapped and note_si != -1 and note_ti != -1:
                    img_copied += copy_images_for_row(_src_wb, shn, sri, tws, tgt_ri + 1,
                                                      note_si, note_ti, _images)

            self._uilog("", ""); self._uilog("=" * 40, "info")
            self._uilog(f"  Matched: {matched} rows", "good")
            if unmatched: self._uilog(f"  Unmatched: {unmatched} rows", "warn")
            if skipped: self._uilog(f"  Skipped empty key: {skipped} rows", "info")
            if img_copied: self._uilog(f"  Images copied: {img_copied}", "good")
            self._uilog("=" * 40, "info")

            out = self.tgt_path.replace(".xlsx", "_matched.xlsx")
            # Save to temp, reload to fix Windows BytesIO image loss
            tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
            tmp.close()
            self.tgt_wb.save(tmp.name)
            self.tgt_wb = openpyxl.load_workbook(tmp.name)
            tws = self.tgt_wb[self.tgt_sheet]
            for ti_, (tgt_ri, row) in enumerate(self.tgt_rows):
                for sc, tc, _, _ in cmaps:
                    v = row[tc] if tc < len(row) else ""
                    if v is not None and str(v).strip():
                        tws.cell(row=tgt_ri + 1, column=tc + 1).value = v
            self.tgt_wb.save(out)
            os.unlink(tmp.name)

            self._uilog(f"Output: {os.path.basename(out)}", "good")
            self._output_path = out
            self._uidone()
            self.root.after(0, lambda: self._show_result_bar(out))
            self.root.after(0, lambda: messagebox.showinfo(
                "Complete",
                f"Matched: {matched}  |  Unmatched: {unmatched}  |  Skipped: {skipped}\n"
                f"{'Images: ' + str(img_copied) + '  |  ' if img_copied else ''}"
                f"\nOutput: {os.path.basename(out)}"))

        except Exception as e:
            import traceback
            self._uilog(f"Error: {e}", "err")
            self._uilog(traceback.format_exc()[-600:], "err")
            self._uidone()

    # ==================== RESULT BAR ====================

    def _show_result_bar(self, output_path):
        for w in self.result_bar.winfo_children(): w.destroy()
        ttk.Label(self.result_bar, text="Result ready:",
                  font=("TkDefaultFont", 10, "bold"), foreground="#81c784").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(self.result_bar, text="Open Folder",
                   command=lambda: self._open_folder(output_path)).pack(side=tk.LEFT, padx=4)
        ttk.Button(self.result_bar, text="Open File",
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
        self.root.after(0, lambda: self.start_btn.configure(state="normal", text="Start Matching"))
        self.root.after(0, lambda: self.prog_var.set(""))

# ============================================================ MAIN ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    MatchApp(root)
    root.mainloop()
