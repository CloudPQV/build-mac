import json
import os
import re
import threading
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl


APP_NAME = "TraCuuExcel"
EXCEL_EXTS = {".xlsx", ".xlsm"}
MAX_RECENT_FOLDERS = 6


def get_config_path() -> Path:
    appdata = os.getenv("APPDATA")
    base_dir = Path(appdata) if appdata else Path.home()
    cfg_dir = base_dir / APP_NAME
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir / "config.json"


def load_folder_history() -> List[str]:
    cfg_file = get_config_path()
    folders: List[str] = []
    if cfg_file.exists():
        try:
            data = json.loads(cfg_file.read_text(encoding="utf-8"))
            raw_list = data.get("recent_folders", [])
            if isinstance(raw_list, list):
                folders = [p for p in raw_list if p and Path(p).exists()]
            old_single = data.get("root_folder", "")
            if old_single and Path(old_single).exists() and old_single not in folders:
                folders.insert(0, old_single)
        except Exception:
            pass
    default_path = Path(r"D:\MomTL")
    if default_path.exists() and str(default_path) not in folders:
        folders.append(str(default_path))
    return folders[:MAX_RECENT_FOLDERS]


def save_folder_history(folders: List[str]):
    cfg_file = get_config_path()
    try:
        clean_list = folders[:MAX_RECENT_FOLDERS]
        first_item = next(iter(clean_list), "")
        payload = {
            "root_folder": first_item,
            "recent_folders": clean_list,
        }
        cfg_file.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass
def strip_accents(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))


def norm_phrase(s: str) -> str:
    s = strip_accents(s or "").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return f" {s} " if s else " "


def clean_val(v) -> str:
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", s)


def iter_excel_files(root: Path):
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in EXCEL_EXTS and not p.name.startswith("~$"):
            yield p


@dataclass
class Hit:
    file_path: Path
    sheet_name: str
    row_idx: int
    matched_col_idx: int
    headers: List[str]
    values: List[str]


def get_merged_cell_val(ws, r: int, c: int) -> str:
    val = ws.cell(row=r, column=c).value
    if val is not None:
        return clean_val(val)
    for m_range in ws.merged_cells.ranges:
        if m_range.min_row <= r <= m_range.max_row and m_range.min_col <= c <= m_range.max_col:
            top_left = ws.cell(row=m_range.min_row, column=m_range.min_col).value
            return clean_val(top_left)
    return ""


def get_merged_span(ws, r: int, c: int) -> int:
    for m_range in ws.merged_cells.ranges:
        if m_range.min_row <= r <= m_range.max_row and m_range.min_col <= c <= m_range.max_col:
            return m_range.max_col - m_range.min_col + 1
    return 1


def row_has_merged_cells(ws, r: int, max_col: int) -> bool:
    for m_range in ws.merged_cells.ranges:
        if m_range.min_row <= r <= m_range.max_row and m_range.min_col <= max_col:
            if m_range.max_col > m_range.min_col or m_range.max_row > m_range.min_row:
                return True
    return False
def build_column_labels(ws, header_row: int, max_col: int) -> List[str]:
    labels: List[str] = []
    if header_row >= 1:
        has_parent_row = header_row > 1 and row_has_merged_cells(ws, header_row - 1, max_col)
        for idx in range(1, max_col + 1):
            curr_val = get_merged_cell_val(ws, header_row, idx)
            parent_val = ""
            if has_parent_row:
                span = get_merged_span(ws, header_row - 1, idx)
                if 2 <= span <= max(5, max_col // 2):
                    cand = get_merged_cell_val(ws, header_row - 1, idx)
                    if not norm_phrase(cand).strip().startswith("nam hoc"):
                        parent_val = cand

            if parent_val and curr_val and norm_phrase(parent_val) != norm_phrase(curr_val):
                t = f"{parent_val} - {curr_val}"
            else:
                t = curr_val or parent_val

            labels.append(t if t else f"Cột {openpyxl.utils.get_column_letter(idx)}")
    else:
        for idx in range(1, max_col + 1):
            labels.append(f"Cột {openpyxl.utils.get_column_letter(idx)}")
    return labels


EXACT_HEADER_TOKENS = {
    "stt", "tt", "so thu tu",
    "ma dinh danh", "ma dinh danh bo gd dt", "ma dinh danh bo gd", "ma hoc sinh", "ma hs",
    "so dinh danh ca nhan", "so dinh danh",
    "ho ten", "ho va ten", "ten", "ho ten hoc sinh", "ho ten cha", "ho ten me", "ho ten nguoi giam ho",
    "ngay sinh", "nam sinh", "gioi tinh", "dan toc", "ton giao", "quoc tich",
    "trang thai", "sdt", "so dien thoai", "sdt lien he", "dien thoai", "dien thoai lien he",
    "so buoi hoc tren tuan", "so buoi hoc", "lop", "khoi", "ten lop", "lop hoc",
    "to thon ho khau thuong tru", "ho khau thuong tru", "to thon", "thon xom",
    "xa phuong", "phuong xa", "quan huyen", "huyen quan", "tinh thanh pho", "tinh tp",
    "cho o hien nay", "noi o hien nay", "dia chi", "que quan", "thuong tru", "tam tru",
    "noi sinh", "benh vien", "nghe nghiep", "trinh do", "cccd", "cmnd", "so cccd", "so cmnd",
    "don vi", "chuc vu", "ghi chu"
}

DATA_VALUE_MARKERS = {"nam", "nu", "kinh", "dang hoc", "chuyen di", "thoi hoc", "bo hoc"}


def is_definitely_data_row(vals: List[str]) -> bool:
    if not vals:
        return True
    first = next(iter(vals), "").strip()
    if first.isdigit():
        return True

    for v in vals:
        if re.search(r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b", v):
            return True
        digits_only = re.sub(r"\s+", "", v)
        if digits_only.isdigit() and len(digits_only) >= 8:
            return True
        if norm_phrase(v).strip() in DATA_VALUE_MARKERS:
            return True
    return False


def score_header_row(ws, r: int, max_col: int) -> int:
    vals = []
    bold_cnt = 0
    for c in range(1, max_col + 1):
        cell = ws.cell(row=r, column=c)
        v = get_merged_cell_val(ws, r, c)
        if v:
            vals.append(v)
            if cell.font and cell.font.bold:
                bold_cnt += 1

    if len(vals) < 3 or is_definitely_data_row(vals):
        return 0

    hint_hits = 0
    for v in vals:
        nv = norm_phrase(v).strip()
        if nv in EXACT_HEADER_TOKENS:
            hint_hits += 1
        elif any(
            kw in nv
            for kw in ("ho ten", "ngay sinh", "ma dinh danh", "so dinh danh", "sdt", "dien thoai", "cho o", "xa phuong", "tinh thanh", "dan toc", "noi sinh")
        ):
            hint_hits += 1

    if hint_hits < 2:
        return 0

    return hint_hits * 20 + bold_cnt * 2 + len(vals)


def find_header_row_above(ws, target_row: int, max_col: int, lookback: int = 2000) -> int:
    start = max(1, target_row - lookback)
    best_row = 0
    best_score = 0
    for r in range(target_row - 1, start - 1, -1):
        sc = score_header_row(ws, r, max_col)
        if sc > best_score:
            best_score = sc
            best_row = r
        if best_score >= 80:
            break
    return best_row
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tra cứu Excel")
        self.geometry("1220x720")

        self.recent_folders: List[str] = load_folder_history()
        first_folder = next(iter(self.recent_folders), "")
        self.root_folder: Optional[Path] = Path(first_folder) if first_folder else None
        self.hits: List[Hit] = []
        self.current_hit: Optional[Hit] = None
        self.stop_flag = False
        self.search_thread: Optional[threading.Thread] = None

        self._build_ui()
        self._refresh_folder_combo()

    def _build_ui(self):
        folder_bar = ttk.Frame(self, padding=(10, 8, 10, 4))
        folder_bar.pack(fill=tk.X)

        ttk.Label(folder_bar, text="Thư mục tra cứu:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(
            folder_bar,
            textvariable=self.folder_var,
            state="readonly",
            width=72,
        )
        self.folder_combo.pack(side=tk.LEFT, padx=(6, 6))
        self.folder_combo.bind("<<ComboboxSelected>>", self._on_combo_folder_selected)

        ttk.Button(folder_bar, text="Chọn thư mục...", command=self.choose_folder).pack(side=tk.LEFT)

        top = ttk.Frame(self, padding=(10, 4, 10, 8))
        top.pack(fill=tk.X)

        ttk.Label(top, text="Từ khóa:").pack(side=tk.LEFT)
        self.q_var = tk.StringVar()
        self.q_entry = ttk.Entry(top, textvariable=self.q_var, width=42)
        self.q_entry.pack(side=tk.LEFT, padx=(6, 8))
        self.q_entry.bind("<Return>", lambda e: self.start_search())

        self.btn_search = ttk.Button(top, text="Tìm", command=self.start_search)
        self.btn_search.pack(side=tk.LEFT)

        self.btn_stop = ttk.Button(top, text="Dừng", command=self.stop_search, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=(6, 12))

        self.btn_copy_selected = ttk.Button(
            top, text="Copy dòng đã chọn", command=self.copy_selected_hit, state=tk.DISABLED
        )
        self.btn_copy_selected.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_copy_all = ttk.Button(
            top, text="Copy tất cả kết quả", command=self.copy_all_hits, state=tk.DISABLED
        )
        self.btn_copy_all.pack(side=tk.LEFT, padx=(0, 12))

        self.status_var = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(top, textvariable=self.status_var).pack(side=tk.RIGHT)

        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        left = ttk.Frame(paned, padding=6)
        paned.add(left, weight=2)

        ttk.Label(left, text="Kết quả tìm thấy:").pack(anchor="w")
        self.listbox = tk.Listbox(left, activestyle="none", selectmode=tk.EXTENDED)
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self.listbox.bind("<<ListboxSelect>>", self.on_select_hit)

        right = ttk.Frame(paned, padding=6)
        paned.add(right, weight=4)

        self.meta_var = tk.StringVar(value="Chọn một kết quả bên trái để xem chi tiết.")
        ttk.Label(right, textvariable=self.meta_var, wraplength=760).pack(anchor="w", pady=(0, 6))

        table_wrap = ttk.Frame(right)
        table_wrap.pack(fill=tk.BOTH, expand=True)

        self.table_canvas = tk.Canvas(table_wrap, highlightthickness=0)
        self.v_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=self.table_canvas.yview)
        self.h_scroll = ttk.Scrollbar(table_wrap, orient=tk.HORIZONTAL, command=self.table_canvas.xview)

        self.table_canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.table_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.table_inner = ttk.Frame(self.table_canvas)
        self.table_window = self.table_canvas.create_window((0, 0), window=self.table_inner, anchor="nw")

        self.table_inner.bind("<Configure>", self._on_inner_configure)
        self.table_canvas.bind("<Configure>", self._on_canvas_configure)

        self.bind_all("<Control-c>", self._on_ctrl_c)
        self.bind_all("<Control-C>", self._on_ctrl_c)
        self.q_entry.focus_set()
    def _set_active_folder(self, folder: Path):
        folder_str = str(folder)
        self.root_folder = folder
        if folder_str in self.recent_folders:
            self.recent_folders.remove(folder_str)
        self.recent_folders.insert(0, folder_str)
        self.recent_folders = self.recent_folders[:MAX_RECENT_FOLDERS]
        save_folder_history(self.recent_folders)
        self._refresh_folder_combo()

    def _refresh_folder_combo(self):
        self.folder_combo["values"] = self.recent_folders
        if self.root_folder and self.root_folder.exists():
            self.folder_var.set(str(self.root_folder))
        else:
            self.folder_var.set("-- Chưa chọn thư mục (nhấn 'Chọn thư mục...') --")

    def _on_combo_folder_selected(self, _event=None):
        chosen = self.folder_var.get()
        if chosen and Path(chosen).exists():
            self._set_active_folder(Path(chosen))
            self.status_var.set("Đã chuyển thư mục tra cứu.")

    def choose_folder(self):
        initial = str(self.root_folder) if (self.root_folder and self.root_folder.exists()) else None
        selected = filedialog.askdirectory(
            title="Chọn thư mục chứa các tệp Excel cần tra cứu",
            initialdir=initial,
        )
        if selected:
            self._set_active_folder(Path(selected))
            self.status_var.set("Đã cập nhật thư mục tra cứu.")

    def _on_inner_configure(self, _event=None):
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

    def _on_canvas_configure(self, _event=None):
        self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

    def clear_table(self):
        for w in self.table_inner.winfo_children():
            w.destroy()

    def make_selectable_cell(
        self,
        parent,
        text: str,
        width_chars: int,
        height_lines: int,
        bold: bool = False,
        bg: str = "#ffffff",
        highlight: bool = False,
    ) -> tk.Text:
        if highlight:
            bg = "#fff2cc"
        font = ("Segoe UI", 10, "bold") if bold else ("Segoe UI", 10)
        t = tk.Text(
            parent,
            width=width_chars,
            height=height_lines,
            wrap="word",
            font=font,
            bg=bg,
            relief="solid",
            borderwidth=1,
            padx=6,
            pady=4,
            exportselection=True,
            takefocus=1,
        )
        t.insert("1.0", text)
        t.configure(state="disabled")
        t.bind("<Button-1>", lambda e, w=t: w.focus_set())
        return t

    def render_horizontal_row_table(self, hit: Hit):
        self.clear_table()
        n = max(len(hit.headers), len(hit.values))
        if n == 0:
            return

        headers = hit.headers + [""] * (n - len(hit.headers))
        values = hit.values + [""] * (n - len(hit.values))

        col_widths: List[int] = []
        val_lines: List[int] = []
        hdr_lines: List[int] = []

        for i in range(n):
            h = headers[i] or f"Cột {openpyxl.utils.get_column_letter(i+1)}"
            v = values[i] or ""
            max_len = max(len(h), len(v))
            w = max(12, min(32, max_len + 2))
            col_widths.append(w)

            hl = max(1, min(3, (len(h) + w - 1) // w))
            vl = max(1, min(6, (len(v) + w - 1) // w))
            hdr_lines.append(hl)
            val_lines.append(vl)

        max_hdr_h = max(hdr_lines) if hdr_lines else 1
        max_val_h = max(val_lines) if val_lines else 1

        for c in range(n):
            h = headers[c] or f"Cột {openpyxl.utils.get_column_letter(c+1)}"
            is_match = (c + 1 == hit.matched_col_idx)
            cell_h = self.make_selectable_cell(
                self.table_inner,
                text=h,
                width_chars=col_widths[c],
                height_lines=max_hdr_h,
                bold=True,
                bg="#f0f0f0",
                highlight=is_match,
            )
            cell_h.grid(row=0, column=c, sticky="nsew")

        for c in range(n):
            v = values[c] or ""
            is_match = (c + 1 == hit.matched_col_idx)
            cell_v = self.make_selectable_cell(
                self.table_inner,
                text=v,
                width_chars=col_widths[c],
                height_lines=max_val_h,
                bold=False,
                bg="#ffffff",
                highlight=is_match,
            )
            cell_v.grid(row=1, column=c, sticky="nsew")

        self.table_canvas.xview_moveto(0)
        self.table_canvas.yview_moveto(0)

    def _on_ctrl_c(self, _event=None):
        w = self.focus_get()
        if isinstance(w, tk.Text):
            try:
                sel = w.get("sel.first", "sel.last")
                if sel:
                    self.clipboard_clear()
                    self.clipboard_append(sel)
                    return "break"
            except tk.TclError:
                pass
        if w == self.listbox and self.current_hit is not None:
            self.copy_selected_hit()
            return "break"
        return None

    def format_hit_tsv(self, hit: Hit) -> str:
        n = max(len(hit.headers), len(hit.values))
        headers = hit.headers + [""] * (n - len(hit.headers))
        values = hit.values + [""] * (n - len(hit.values))
        meta = f"File: {hit.file_path} | Sheet: {hit.sheet_name} | Dòng: {hit.row_idx}"
        return "\n".join([meta, "\t".join(headers), "\t".join(values)])

    def copy_selected_hit(self):
        if self.current_hit is None:
            return
        txt = self.format_hit_tsv(self.current_hit)
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.status_var.set("Đã copy dòng (kèm tiêu đề) vào clipboard.")

    def copy_all_hits(self):
        if not self.hits:
            return
        blocks = [self.format_hit_tsv(h) for h in self.hits]
        txt = ("\n" + ("-" * 60) + "\n").join(blocks)
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.status_var.set(f"Đã copy {len(self.hits)} kết quả vào clipboard.")
    def stop_search(self):
        self.stop_flag = True
        self.status_var.set("Đang dừng...")

    def start_search(self):
        q_raw = self.q_var.get().strip()
        if not q_raw:
            messagebox.showinfo("Thiếu từ khóa", "Nhập từ khóa cần tìm.")
            return

        if not self.root_folder or not self.root_folder.exists():
            messagebox.showwarning(
                "Chưa chọn thư mục",
                "Vui lòng chọn thư mục chứa các tệp Excel trước khi tìm kiếm.",
            )
            self.choose_folder()
            if not self.root_folder or not self.root_folder.exists():
                return

        if self.search_thread and self.search_thread.is_alive():
            return

        self.stop_flag = False
        self.hits.clear()
        self.current_hit = None
        self.listbox.delete(0, tk.END)
        self.clear_table()
        self.meta_var.set("Đang tìm...")

        self.btn_search.configure(state=tk.DISABLED)
        self.btn_stop.configure(state=tk.NORMAL)
        self.btn_copy_selected.configure(state=tk.DISABLED)
        self.btn_copy_all.configure(state=tk.DISABLED)

        q_phrase = norm_phrase(q_raw)
        target_folder = self.root_folder

        self.search_thread = threading.Thread(
            target=self._worker_search, args=(target_folder, q_phrase), daemon=True
        )
        self.search_thread.start()

    def _worker_search(self, target_folder: Path, q_phrase: str):
        files = list(iter_excel_files(target_folder))
        total_files = len(files)

        for idx, fp in enumerate(files, start=1):
            if self.stop_flag:
                break

            self.after(
                0,
                lambda i=idx, n=total_files, name=fp.name: self.status_var.set(
                    f"Đang quét ({i}/{n}): {name} | Trùng: {len(self.hits)}"
                ),
            )

            try:
                wb = openpyxl.load_workbook(fp, read_only=False, data_only=True)
            except Exception:
                continue

            try:
                for ws in wb.worksheets:
                    if self.stop_flag:
                        break

                    max_r = ws.max_row or 0
                    max_c = ws.max_column or 0
                    if max_r == 0 or max_c == 0:
                        continue

                    header_cache: Dict[int, Tuple[int, List[str]]] = {}

                    for r in range(1, max_r + 1):
                        if self.stop_flag:
                            break

                        matched_col = 0
                        row_vals: List[str] = []
                        has_any = False

                        for c in range(1, max_c + 1):
                            v = clean_val(ws.cell(row=r, column=c).value)
                            row_vals.append(v)
                            if v:
                                has_any = True
                                if matched_col == 0 and (q_phrase in norm_phrase(v)):
                                    matched_col = c

                        if not has_any or matched_col == 0:
                            continue

                        while row_vals and row_vals[-1] == "":
                            row_vals.pop()
                        eff_cols = max(len(row_vals), matched_col)

                        block_key = r // 10
                        if block_key not in header_cache:
                            hr = find_header_row_above(ws, r, eff_cols, lookback=2000)
                            labels = build_column_labels(ws, hr, eff_cols)
                            header_cache[block_key] = (hr, labels)
                        else:
                            hr, labels = header_cache[block_key]
                            if len(labels) < eff_cols:
                                labels = build_column_labels(ws, hr, eff_cols)
                                header_cache[block_key] = (hr, labels)

                        hit = Hit(
                            file_path=fp,
                            sheet_name=ws.title,
                            row_idx=r,
                            matched_col_idx=matched_col,
                            headers=labels[:eff_cols],
                            values=row_vals[:eff_cols],
                        )
                        self.after(0, lambda h=hit: self._add_hit(h))
            finally:
                try:
                    wb.close()
                except Exception:
                    pass

        self.after(0, self._finish_search)

    def _add_hit(self, hit: Hit):
        self.hits.append(hit)
        rel = hit.file_path.name
        try:
            if self.root_folder:
                rel = str(hit.file_path.relative_to(self.root_folder))
        except Exception:
            pass
        col_letter = openpyxl.utils.get_column_letter(hit.matched_col_idx)
        self.listbox.insert(tk.END, f"[{len(self.hits)}] {rel} | {hit.sheet_name} | Dòng {hit.row_idx} ({col_letter})")
        if len(self.hits) == 1:
            self.btn_copy_all.configure(state=tk.NORMAL)

    def _finish_search(self):
        self.btn_search.configure(state=tk.NORMAL)
        self.btn_stop.configure(state=tk.DISABLED)
        if self.stop_flag:
            self.status_var.set(f"Đã dừng. Tìm thấy {len(self.hits)} dòng.")
        else:
            self.status_var.set(f"Hoàn tất. Tìm thấy {len(self.hits)} dòng.")
        if not self.hits:
            self.meta_var.set("Không tìm thấy kết quả phù hợp.")

    def on_select_hit(self, _event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = next(iter(sel), -1)
        if idx < 0 or idx >= len(self.hits):
            return

        hit = self.hits[idx]
        self.current_hit = hit
        self.btn_copy_selected.configure(state=tk.NORMAL)

        col_letter = openpyxl.utils.get_column_letter(hit.matched_col_idx)
        self.meta_var.set(
            f"File: {hit.file_path}\nSheet: {hit.sheet_name} | Dòng: {hit.row_idx} | Ô khớp: {col_letter}{hit.row_idx}"
        )
        self.render_horizontal_row_table(hit)


if __name__ == "__main__":
    app = App()
    app.mainloop()
