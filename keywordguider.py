# ============================================================
# Keyword Guide UI POC
# ------------------------------------------------------------
# Version: 0.5.2 (2025-12-31)
#
# Release Notes (v0.5.2)
# - (I/O) 저장 안정성(WinError5 방지): atomic write(tmp->replace) + retry(backoff)
# - (UI) KeywordDialog 마우스휠 바인딩 부작용 패치: bind_all 제거, 로컬 바인딩만 사용
# - (UI) KeywordDialog Description 스크롤바 추가
# - (UX) 로그성 메시지(Log panel) 추가 + 저장/Export/Import 등 이벤트 기록
#
# Base: v0.5.0 (2025-12-30)
# ============================================================

import json
import re
import os
import time
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

# ------------------------------------------------------------
# Paths / Constants
# ------------------------------------------------------------
APP_VERSION = "0.5.2"

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "keywords_db.json"
UI_STATE_PATH = BASE_DIR / "ui_state.json"
ISSUES_PATH = BASE_DIR / "issues_config.json"

PLACEHOLDER_RE = re.compile(r"\{([A-Za-z0-9_]+)\}")
DEFAULT_GEOMETRY = "1280x780"

# Keyword list view: Summary | Group | Info | Copy | CopyNP | Preview
KEYWORD_COLS = ("summary", "group", "info", "copy", "copynp", "preview")

DEFAULT_KEYWORD_COL_WIDTHS = {
    "summary": 240,
    "group": 140,
    "info": 70,
    "copy": 70,
    "copynp": 90,
    "preview": 560,
}

PARAM_COLS = ("pname", "pval")
DEFAULT_PARAM_COL_WIDTHS = {"pname": 140, "pval": 280}

COPY_FEEDBACK_MS = 900
DEFAULT_DELIMITER = ";"

# KeywordDialog UI
PREVIEW_MAX_LINES = 4
PARTS_AREA_HEIGHT_PX = 160

# Description Rich tags
DESC_COLOR_KEYS = ("black", "red", "blue")

# Export schema
EXPORT_SCHEMA_VERSION = "1.0"


# ------------------------------------------------------------
# Utility: Safe JSON I/O (WinError 5 방지)
# ------------------------------------------------------------
def load_json(path: Path):
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _safe_atomic_write_text(path: Path, text: str, encoding="utf-8", retries: int = 7):
    """
    Windows에서 간헐적으로 발생하는 PermissionError/WinError 5 대응:
    - tmp 파일에 먼저 write + flush + fsync(가능하면)
    - os.replace로 원본 교체 (atomic)
    - 여러 번 retry + backoff
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    delays = [0.0, 0.03, 0.06, 0.12, 0.25, 0.5, 0.9]  # seconds
    last_err = None

    for attempt in range(max(1, retries)):
        tmp_path = None
        try:
            suffix = f".tmp.{os.getpid()}.{int(time.time()*1000)}"
            tmp_path = path.with_name(path.name + suffix)

            with open(tmp_path, "w", encoding=encoding, newline="\n") as f:
                f.write(text)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except Exception:
                    pass

            os.replace(tmp_path, path)
            return

        except Exception as e:
            last_err = e
            try:
                if tmp_path and tmp_path.exists():
                    tmp_path.unlink(missing_ok=True)
            except Exception:
                pass

            delay = delays[attempt] if attempt < len(delays) else delays[-1]
            if delay > 0:
                time.sleep(delay)

    raise last_err


def save_json(path: Path, data: dict):
    txt = json.dumps(data, indent=2, ensure_ascii=False)
    _safe_atomic_write_text(path, txt, encoding="utf-8")


def detect_placeholders(text: str):
    seen, ordered = set(), []
    for m in PLACEHOLDER_RE.finditer(text or ""):
        name = m.group(1)
        if name not in seen:
            seen.add(name)
            ordered.append(name)
    return ordered


def render_keyword(template: str, params: dict):
    """placeholder를 params 값으로 치환"""
    out = template or ""
    params = params or {}
    for k, v in params.items():
        out = out.replace("{" + str(k) + "}", str(v))
    return out


def render_keyword_without_params(template: str):
    """placeholder는 제거(빈값) 처리: {CH} -> "" """
    if not template:
        return ""
    return PLACEHOLDER_RE.sub("", template)


def default_issues():
    return ["데이터 이슈", "망등록 이슈"]


def _clean_str_list_keep_order(items):
    out = []
    for it in items or []:
        s = str(it).strip()
        if s:
            out.append(s)
    return out


def ensure_issue_config_vendor_scoped(cfg: dict, vendors: list[str]):
    """
    issues_config.json (vendor-scoped) schema:
      {
        "vendors": {
          "MTK":  {"issues": [...], "delimiter": ";" },
          "SLSI": {"issues": [...], "delimiter": ";" }
        }
      }

    Backward compatible:
      - old schema: {"issues":[...]}  -> apply to all vendors
    """
    if not isinstance(cfg, dict):
        cfg = {}

    if "issues" in cfg and "vendors" not in cfg:
        shared = cfg.get("issues")
        if not isinstance(shared, list) or not shared:
            shared = default_issues()
        cfg = {"vendors": {v: {"issues": list(shared), "delimiter": DEFAULT_DELIMITER} for v in vendors}}

    cfg.setdefault("vendors", {})
    if not isinstance(cfg["vendors"], dict):
        cfg["vendors"] = {}

    for v in vendors:
        vobj = cfg["vendors"].get(v)
        if not isinstance(vobj, dict):
            vobj = {}
            cfg["vendors"][v] = vobj

        issues = vobj.get("issues")
        if not isinstance(issues, list) or not issues:
            vobj["issues"] = default_issues()
        else:
            seen = set()
            cleaned = []
            for it in issues:
                s = str(it).strip()
                if not s or s in seen:
                    continue
                seen.add(s)
                cleaned.append(s)
            vobj["issues"] = cleaned if cleaned else default_issues()

        delim = vobj.get("delimiter", DEFAULT_DELIMITER)
        delim = str(delim) if delim is not None else DEFAULT_DELIMITER
        vobj["delimiter"] = delim

    return cfg


def _desc_plain_from_rich(runs):
    if not isinstance(runs, list):
        return ""
    out = []
    for r in runs:
        if isinstance(r, dict):
            t = r.get("text", "")
            if t:
                out.append(str(t))
    return "".join(out)


def normalize_keywords(lst):
    """
    Normalize keyword list items to dict:
      legacy: {"summary":..., "desc":..., "text":...}
      new:    {"summary":..., "group":..., "desc":..., "desc_rich":[...], "parts":[...]}
    """
    out = []
    for item in lst or []:
        if isinstance(item, str):
            t = item.strip()
            if t:
                out.append({"text": t, "summary": "", "group": "", "desc": ""})
            continue

        if not isinstance(item, dict):
            continue

        summary = str(item.get("summary", "")).strip()
        group = str(item.get("group", "")).strip()

        desc_rich = item.get("desc_rich", None)
        desc = str(item.get("desc", item.get("description", ""))).strip()
        if isinstance(desc_rich, list) and desc_rich and not desc:
            desc = _desc_plain_from_rich(desc_rich).strip()

        if "parts" in item and isinstance(item.get("parts"), list):
            parts = _clean_str_list_keep_order(item.get("parts"))
            if parts:
                kw = {"parts": parts, "summary": summary, "group": group, "desc": desc}
                if isinstance(desc_rich, list) and desc_rich:
                    kw["desc_rich"] = desc_rich
                out.append(kw)
                continue

        text = str(item.get("text", "")).strip()
        if text:
            kw = {"text": text, "summary": summary, "group": group, "desc": desc}
            if isinstance(desc_rich, list) and desc_rich:
                kw["desc_rich"] = desc_rich
            out.append(kw)

    return out


def keyword_joined_template(kw: dict, delimiter: str) -> str:
    delimiter = DEFAULT_DELIMITER if delimiter is None else str(delimiter)
    if isinstance(kw, dict) and isinstance(kw.get("parts"), list):
        parts = _clean_str_list_keep_order(kw.get("parts"))
        return delimiter.join(parts)
    return str((kw or {}).get("text", "")).strip()


def keyword_parts_from_kw(kw: dict, delimiter: str) -> list[str]:
    delimiter = DEFAULT_DELIMITER if delimiter is None else str(delimiter)

    if isinstance(kw, dict) and isinstance(kw.get("parts"), list) and kw.get("parts"):
        return _clean_str_list_keep_order(kw.get("parts"))

    text = str((kw or {}).get("text", "")).strip()
    if not text:
        return []
    if delimiter and delimiter in text:
        return [p.strip() for p in text.split(delimiter) if p.strip()]
    return [text]


def export_package(db: dict, issue_cfg: dict, ui_state: dict | None = None) -> dict:
    return {
        "schema": "KeywordGuideExport",
        "schema_version": EXPORT_SCHEMA_VERSION,
        "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "db": db if isinstance(db, dict) else {},
        "issue_cfg": issue_cfg if isinstance(issue_cfg, dict) else {},
        "ui_state": ui_state if isinstance(ui_state, dict) else {},
    }


def validate_import_package(pkg: dict) -> tuple[bool, str]:
    if not isinstance(pkg, dict):
        return False, "Import file is not a JSON object."
    if pkg.get("schema") != "KeywordGuideExport":
        return False, "Invalid schema (not KeywordGuideExport)."
    if "db" not in pkg or "issue_cfg" not in pkg:
        return False, "Missing required keys: db / issue_cfg."
    if not isinstance(pkg.get("db"), dict):
        return False, "db must be a dict."
    if not isinstance(pkg.get("issue_cfg"), dict):
        return False, "issue_cfg must be a dict."
    return True, "OK"


# ------------------------------------------------------------
# Dialogs
# ------------------------------------------------------------
class KeywordDialog(tk.Toplevel):
    def __init__(self, parent, title, init=None, delimiter=";"):
        super().__init__(parent)
        self.title(title)
        self.resizable(True, True)
        self.result = None

        self.delimiter = str(delimiter) if delimiter is not None else DEFAULT_DELIMITER
        init = init or {"summary": "", "group": "", "desc": ""}

        self.var_summary = tk.StringVar(value=str(init.get("summary", "")))
        self.var_group = tk.StringVar(value=str(init.get("group", "")))

        self._last_split_offer_text = None
        self._split_offer_inflight = False

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Summary (ListView에 표시될 짧은 요약)").grid(row=0, column=0, sticky="w")
        ent_sum = ttk.Entry(frm, textvariable=self.var_summary)
        ent_sum.grid(row=1, column=0, sticky="ew", pady=(2, 8))

        ttk.Label(frm, text="Group (ListView Summary 우측에 표시)").grid(row=2, column=0, sticky="w")
        ent_grp = ttk.Entry(frm, textvariable=self.var_group)
        ent_grp.grid(row=3, column=0, sticky="ew", pady=(2, 10))

        ttk.Label(frm, text=f"Import Joined String (delimiter: '{self.delimiter}')").grid(row=4, column=0, sticky="w")

        import_row = ttk.Frame(frm)
        import_row.grid(row=5, column=0, sticky="ew", pady=(2, 10))
        import_row.columnconfigure(0, weight=1)

        self.var_import_joined = tk.StringVar(value="")
        self.ent_import = ttk.Entry(import_row, textvariable=self.var_import_joined)
        self.ent_import.grid(row=0, column=0, sticky="ew")
        ttk.Button(import_row, text="Split", command=lambda: self._import_joined_to_parts(ask_confirm=True)).grid(
            row=0, column=1, padx=(6, 0)
        )
        self.ent_import.bind("<Return>", lambda _e: self._import_joined_to_parts(ask_confirm=True))

        ttk.Label(frm, text=f"Keyword Parts (구분자: '{self.delimiter}')").grid(row=6, column=0, sticky="w")

        parts_outer = ttk.Frame(frm)
        parts_outer.grid(row=7, column=0, sticky="ew", pady=(2, 6))
        parts_outer.columnconfigure(0, weight=1)
        parts_outer.rowconfigure(0, weight=1)

        self.parts_canvas = tk.Canvas(
            parts_outer,
            height=PARTS_AREA_HEIGHT_PX,
            highlightthickness=1,
            highlightbackground="#c0c0c0",
        )
        self.parts_canvas.grid(row=0, column=0, sticky="ew")

        self.parts_scroll = ttk.Scrollbar(parts_outer, orient="vertical", command=self.parts_canvas.yview)
        self.parts_scroll.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        self.parts_canvas.configure(yscrollcommand=self.parts_scroll.set)

        self.parts_container = ttk.Frame(self.parts_canvas)
        self._parts_window_id = self.parts_canvas.create_window((0, 0), window=self.parts_container, anchor="nw")

        self.parts_container.bind("<Configure>", self._on_parts_container_configure)
        self.parts_canvas.bind("<Configure>", self._on_parts_canvas_configure)

        # (PATCH) 마우스휠: bind_all/unbind_all 제거 -> 로컬 바인딩만
        self._bind_mousewheel_local(self.parts_canvas, target_canvas=self.parts_canvas)
        self._bind_mousewheel_local(self.parts_container, target_canvas=self.parts_canvas)

        add_btn_row = ttk.Frame(frm)
        add_btn_row.grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Button(add_btn_row, text="+ Add Part", command=self._add_part_row).pack(side=tk.LEFT)

        ttk.Label(frm, text="Joined Keyword Preview (auto wrap, scrollable)").grid(row=9, column=0, sticky="w", pady=(6, 0))

        preview_box = ttk.Frame(frm)
        preview_box.grid(row=10, column=0, sticky="ew", pady=(2, 10))
        preview_box.columnconfigure(0, weight=1)

        self.preview_text = tk.Text(
            preview_box,
            height=PREVIEW_MAX_LINES,
            wrap="word",
            relief="solid",
            borderwidth=1,
        )
        self.preview_text.grid(row=0, column=0, sticky="ew")
        self.preview_text.configure(state="disabled")

        self.preview_scroll = ttk.Scrollbar(preview_box, orient="vertical", command=self.preview_text.yview)
        self.preview_scroll.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        self.preview_text.configure(yscrollcommand=self.preview_scroll.set)

        ttk.Label(frm, text="Description (Bold + Color, Info 팝업으로 표시)").grid(row=11, column=0, sticky="w")
        desc_toolbar = ttk.Frame(frm)
        desc_toolbar.grid(row=12, column=0, sticky="w", pady=(4, 2))

        ttk.Button(desc_toolbar, text="B", width=3, command=self._toggle_bold).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(desc_toolbar, text="Black", command=lambda: self._apply_color("black")).pack(side=tk.LEFT, padx=2)
        ttk.Button(desc_toolbar, text="Red", command=lambda: self._apply_color("red")).pack(side=tk.LEFT, padx=2)
        ttk.Button(desc_toolbar, text="Blue", command=lambda: self._apply_color("blue")).pack(side=tk.LEFT, padx=2)

        # (PATCH) Description: 스크롤바 추가
        desc_box = ttk.Frame(frm)
        desc_box.grid(row=13, column=0, sticky="nsew", pady=(2, 10))
        desc_box.columnconfigure(0, weight=1)
        desc_box.rowconfigure(0, weight=1)

        self.txt_desc = tk.Text(desc_box, height=10, wrap="word")
        self.txt_desc.grid(row=0, column=0, sticky="nsew")

        desc_sb = ttk.Scrollbar(desc_box, orient="vertical", command=self.txt_desc.yview)
        desc_sb.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        self.txt_desc.configure(yscrollcommand=desc_sb.set)

        self.txt_desc.tag_configure("b", font=("TkDefaultFont", 9, "bold"))
        self.txt_desc.tag_configure("c_black", foreground="black")
        self.txt_desc.tag_configure("c_red", foreground="red")
        self.txt_desc.tag_configure("c_blue", foreground="blue")

        btns = ttk.Frame(frm)
        btns.grid(row=14, column=0, sticky="e")
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="OK", command=self._ok).pack(side=tk.LEFT)

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(13, weight=1)

        initial_parts = keyword_parts_from_kw(init, self.delimiter)
        if not initial_parts:
            initial_parts = [""]

        for p in initial_parts:
            self._add_part_row(initial_text=p)

        try:
            self.var_import_joined.set(self.delimiter.join([x for x in initial_parts if str(x).strip()]))
        except Exception:
            pass

        self._set_preview_text(self.delimiter.join(self._get_parts()))

        desc_rich = init.get("desc_rich", None)
        if isinstance(desc_rich, list) and desc_rich:
            self._apply_desc_rich(desc_rich)
        else:
            self.txt_desc.insert("1.0", str(init.get("desc", "")))

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        ent_sum.focus_set()
        self.wait_window(self)

    def _on_parts_container_configure(self, _event=None):
        try:
            self.parts_canvas.configure(scrollregion=self.parts_canvas.bbox("all"))
        except Exception:
            pass

    def _on_parts_canvas_configure(self, event=None):
        try:
            w = event.width if event else self.parts_canvas.winfo_width()
            self.parts_canvas.itemconfigure(self._parts_window_id, width=w)
        except Exception:
            pass

    # (PATCH) 로컬 마우스휠 바인딩
    def _bind_mousewheel_local(self, widget, target_canvas: tk.Canvas):
        def on_mousewheel(event):
            try:
                delta = int(-1 * (event.delta / 120))
                target_canvas.yview_scroll(delta, "units")
                return "break"
            except Exception:
                return None

        def on_mousewheel_linux(event):
            try:
                if event.num == 4:
                    target_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    target_canvas.yview_scroll(1, "units")
                return "break"
            except Exception:
                return None

        widget.bind("<MouseWheel>", on_mousewheel, add="+")
        widget.bind("<Button-4>", on_mousewheel_linux, add="+")
        widget.bind("<Button-5>", on_mousewheel_linux, add="+")

    def _add_part_row(self, initial_text=""):
        row = ttk.Frame(self.parts_container)
        row.pack(fill=tk.X, pady=2)

        ent = ttk.Entry(row)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if initial_text:
            ent.insert(0, str(initial_text))

        btn = ttk.Button(row, text="-", width=3, command=lambda r=row: self._remove_part_row(r))
        btn.pack(side=tk.LEFT, padx=(6, 0))

        ent.bind("<KeyRelease>", lambda e, w=ent: self._on_part_entry_keyrelease(e, w))
        ent.bind("<Control-v>", lambda e, w=ent: self._after_paste_offer_split(w))
        ent.bind("<Control-V>", lambda e, w=ent: self._after_paste_offer_split(w))

        self._on_parts_container_configure()
        self._update_preview()

    def _remove_part_row(self, row_frame):
        children = [w for w in self.parts_container.winfo_children() if isinstance(w, ttk.Frame)]
        if len(children) <= 1:
            try:
                ent = row_frame.winfo_children()[0]
                if isinstance(ent, ttk.Entry):
                    ent.delete(0, tk.END)
            except Exception:
                pass
            self._update_preview()
            return

        try:
            row_frame.destroy()
        except Exception:
            pass

        self._on_parts_container_configure()
        self._update_preview()

    def _get_parts(self) -> list[str]:
        parts = []
        for row in self.parts_container.winfo_children():
            if not isinstance(row, ttk.Frame):
                continue
            kids = row.winfo_children()
            if not kids:
                continue
            ent = kids[0]
            if isinstance(ent, ttk.Entry):
                s = ent.get().strip()
                if s:
                    parts.append(s)
        return parts

    def _clear_all_part_rows(self):
        for child in list(self.parts_container.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass
        self._on_parts_container_configure()

    def _set_parts_rows(self, parts: list[str]):
        self._clear_all_part_rows()
        if not parts:
            parts = [""]
        for p in parts:
            self._add_part_row(initial_text=p)

        self._on_parts_container_configure()
        self._update_preview()
        try:
            self.parts_canvas.yview_moveto(0.0)
        except Exception:
            pass

    def _set_preview_text(self, text: str):
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", text or "")
        self.preview_text.configure(state="disabled")
        try:
            self.preview_text.yview_moveto(0.0)
        except Exception:
            pass

    def _update_preview(self):
        joined = self.delimiter.join(self._get_parts())
        self._set_preview_text(joined)

    def _split_by_delimiter(self, text: str) -> list[str]:
        delim = self.delimiter if self.delimiter is not None else DEFAULT_DELIMITER
        delim = str(delim)
        if delim == "":
            return []
        parts = [p.strip() for p in (text or "").split(delim)]
        return [p for p in parts if p]

    def _confirm_apply_split(self, parts: list[str], source_label: str) -> bool:
        if not parts:
            return False
        sample = self.delimiter.join(parts[:6])
        msg = (
            f"{source_label} 입력에서 구분자('{self.delimiter}') 기준으로 {len(parts)}개 Part가 감지되었습니다.\n\n"
            f"Parts로 분해하여 적용할까요?\n\n"
            f"예시(앞부분): {sample}"
        )
        return messagebox.askyesno("Split into Parts?", msg)

    def _import_joined_to_parts(self, ask_confirm: bool = True):
        raw = (self.var_import_joined.get() or "").strip()
        if not raw:
            return

        parts = self._split_by_delimiter(raw)
        if not parts:
            messagebox.showinfo("Info", "분해할 수 있는 Part가 없습니다. (Delimiter가 비어있거나 split 결과가 비어있음)")
            return

        if ask_confirm and not self._confirm_apply_split(parts, source_label="Import Joined String"):
            return

        self._set_parts_rows(parts)
        self._last_split_offer_text = raw

    def _after_paste_offer_split(self, entry_widget):
        self.after(1, lambda: self._offer_split_from_part_entry(entry_widget))

    def _on_part_entry_keyrelease(self, _event, entry_widget):
        self._update_preview()
        self._offer_split_from_part_entry(entry_widget)

    def _offer_split_from_part_entry(self, entry_widget):
        if self._split_offer_inflight:
            return

        try:
            raw = (entry_widget.get() or "").strip()
        except Exception:
            return

        if not raw:
            return

        delim = str(self.delimiter if self.delimiter is not None else DEFAULT_DELIMITER)
        if delim == "" or delim not in raw:
            return

        parts = self._split_by_delimiter(raw)
        if len(parts) <= 1:
            return

        if raw == self._last_split_offer_text:
            return

        current_parts = self._get_parts()
        likely_joined = (len(current_parts) <= 1)
        if not likely_joined and len(raw) < 30:
            return

        self._split_offer_inflight = True
        try:
            ok = self._confirm_apply_split(parts, source_label="Part 입력")
            if ok:
                self._set_parts_rows(parts)
                try:
                    self.var_import_joined.set(self.delimiter.join(parts))
                except Exception:
                    pass
                self._last_split_offer_text = raw
            else:
                self._last_split_offer_text = raw
        finally:
            self._split_offer_inflight = False

    def _get_sel_range(self):
        try:
            start = self.txt_desc.index("sel.first")
            end = self.txt_desc.index("sel.last")
            return start, end
        except Exception:
            return None

    def _toggle_tag(self, tag: str, start: str, end: str):
        has_any = False
        idx = start
        while self.txt_desc.compare(idx, "<", end):
            tags = self.txt_desc.tag_names(idx)
            if tag in tags:
                has_any = True
                break
            idx = self.txt_desc.index(f"{idx}+1c")

        if has_any:
            self.txt_desc.tag_remove(tag, start, end)
        else:
            self.txt_desc.tag_add(tag, start, end)

    def _toggle_bold(self):
        rng = self._get_sel_range()
        if rng:
            self._toggle_tag("b", rng[0], rng[1])
            return

        start = self.txt_desc.index("insert linestart")
        end = self.txt_desc.index("insert lineend")
        self._toggle_tag("b", start, end)

    def _apply_color(self, color_name: str):
        if color_name not in DESC_COLOR_KEYS:
            return

        rng = self._get_sel_range()
        if rng:
            start, end = rng
        else:
            start = self.txt_desc.index("insert linestart")
            end = self.txt_desc.index("insert lineend")

        for t in ("c_black", "c_red", "c_blue"):
            self.txt_desc.tag_remove(t, start, end)

        self.txt_desc.tag_add(f"c_{color_name}", start, end)

    def _serialize_desc_rich(self):
        text = self.txt_desc.get("1.0", "end-1c")
        if not text:
            return []

        runs = []
        cur = None

        for i, ch in enumerate(text):
            idx = f"1.0+{i}c"
            tags = set(self.txt_desc.tag_names(idx))

            b = ("b" in tags)
            c = "black"
            if "c_red" in tags:
                c = "red"
            elif "c_blue" in tags:
                c = "blue"
            elif "c_black" in tags:
                c = "black"

            if cur and cur["b"] == b and cur["c"] == c:
                cur["text"] += ch
            else:
                cur = {"text": ch, "b": b, "c": c}
                runs.append(cur)

        return [r for r in runs if r.get("text")]

    def _apply_desc_rich(self, runs: list[dict]):
        self.txt_desc.delete("1.0", "end")
        for r in runs:
            if not isinstance(r, dict):
                continue
            t = str(r.get("text", ""))
            if not t:
                continue

            start = self.txt_desc.index("end-1c")
            self.txt_desc.insert("end", t)
            end = self.txt_desc.index("end-1c")

            if r.get("b"):
                self.txt_desc.tag_add("b", start, end)

            c = r.get("c", "black")
            if c in DESC_COLOR_KEYS:
                self.txt_desc.tag_add(f"c_{c}", start, end)

    def _ok(self):
        parts = self._get_parts()
        if not parts:
            messagebox.showwarning("Warning", "At least one non-empty part is required.")
            return

        desc_plain = self.txt_desc.get("1.0", "end-1c").strip()
        desc_rich = self._serialize_desc_rich()

        self.result = {
            "parts": parts,
            "summary": self.var_summary.get().strip(),
            "group": self.var_group.get().strip(),
            "desc": desc_plain,
            "desc_rich": desc_rich,
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ------------------------------------------------------------
# Info Popup
# ------------------------------------------------------------
class InfoPopup(tk.Toplevel):
    def __init__(self, parent, title: str, summary: str, group: str, keyword: str, desc: str, desc_rich=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("840x520")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Summary").grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text=summary or "(empty)").grid(row=1, column=0, sticky="w", pady=(0, 8))

        ttk.Label(frm, text="Group").grid(row=2, column=0, sticky="w")
        ttk.Label(frm, text=group or "(empty)").grid(row=3, column=0, sticky="w", pady=(0, 10))

        ttk.Label(frm, text="Keyword").grid(row=4, column=0, sticky="w")
        txt_kw = tk.Text(frm, wrap="word", height=4)
        txt_kw.grid(row=5, column=0, sticky="nsew", pady=(0, 10))
        txt_kw.insert("1.0", keyword or "")
        txt_kw.configure(state="disabled")

        ttk.Label(frm, text="Description").grid(row=6, column=0, sticky="w")

        # InfoPopup도 길어질 수 있어 스크롤 가능하도록
        desc_box = ttk.Frame(frm)
        desc_box.grid(row=7, column=0, sticky="nsew")
        desc_box.columnconfigure(0, weight=1)
        desc_box.rowconfigure(0, weight=1)

        txt = tk.Text(desc_box, wrap="word")
        txt.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(desc_box, orient="vertical", command=txt.yview)
        sb.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        txt.configure(yscrollcommand=sb.set)

        txt.tag_configure("b", font=("TkDefaultFont", 9, "bold"))
        txt.tag_configure("c_black", foreground="black")
        txt.tag_configure("c_red", foreground="red")
        txt.tag_configure("c_blue", foreground="blue")

        if isinstance(desc_rich, list) and desc_rich:
            for r in desc_rich:
                if not isinstance(r, dict):
                    continue
                t = str(r.get("text", ""))
                if not t:
                    continue
                start = txt.index("end-1c")
                txt.insert("end", t)
                end = txt.index("end-1c")
                if r.get("b"):
                    txt.tag_add("b", start, end)
                c = r.get("c", "black")
                if c in DESC_COLOR_KEYS:
                    txt.tag_add(f"c_{c}", start, end)
        else:
            txt.insert("1.0", desc or "")

        txt.configure(state="disabled")

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Close", command=self.destroy).pack()

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(7, weight=1)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)


# ------------------------------------------------------------
# Main App
# ------------------------------------------------------------
class KeywordGuideApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Chipset Log Keyword Guide  v{APP_VERSION}")

        self.db = load_json(DB_PATH) or self._default_db()
        self.ui_state = load_json(UI_STATE_PATH)

        vendors = list(self.db.keys()) if isinstance(self.db, dict) else []
        if not vendors:
            self.db = self._default_db()
            vendors = list(self.db.keys())

        raw_cfg = load_json(ISSUES_PATH)
        self.issue_cfg = ensure_issue_config_vendor_scoped(raw_cfg, vendors)
        self._sync_vendor_scoped_config_with_db()

        # current selection state
        self.vendor_var = tk.StringVar()
        self.issue_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.delim_var = tk.StringVar(value=DEFAULT_DELIMITER)

        # param in-place editor state
        self._param_editor = None
        self._param_editing = None

        # copy feedback
        self._copy_feedback_after_id = None
        self._copy_feedback_row = None

        self.geometry(self.ui_state.get("geometry", DEFAULT_GEOMETRY) if isinstance(self.ui_state, dict) else DEFAULT_GEOMETRY)
        self._build_ui()
        self._make_checkbox_images()

        self._apply_saved_widths()

        # init selection from UI state if possible
        self._init_nav_default_selection()
        self._ensure_current_obj_migrated()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.log_info("App started.")

    # --------------------------------------------------------
    # Logging (NEW)
    # --------------------------------------------------------
    def _ts(self):
        return datetime.now().strftime("%H:%M:%S")

    def _append_log(self, level: str, msg: str):
        line = f"[{self._ts()}] {level}: {msg}"
        self.status_var.set(line)
        try:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", line + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        except Exception:
            pass

        try:
            print(line)
        except Exception:
            pass

    def log_info(self, msg: str):
        self._append_log("INFO", msg)

    def log_warn(self, msg: str):
        self._append_log("WARN", msg)

    def log_error(self, msg: str):
        self._append_log("ERROR", msg)

    # --------------------------------------------------------
    # Defaults
    # --------------------------------------------------------
    def _default_db(self):
        vendors = ["MTK", "SLSI"]
        issues = default_issues()
        out = {}
        for v in vendors:
            out[v] = {}
            for i in issues:
                out[v][i] = {"_COMMON": {"_keywords": [], "_params": {}}}
        return out

    def _default_issue_obj(self):
        return {"_COMMON": {"_keywords": [], "_params": {}}}

    # --------------------------------------------------------
    # Config helpers
    # --------------------------------------------------------
    def _get_vendor_cfg(self, vendor: str) -> dict:
        return self.issue_cfg.get("vendors", {}).get(vendor, {}) if vendor else {}

    def _get_vendor_issues(self, vendor: str) -> list[str]:
        vobj = self._get_vendor_cfg(vendor)
        issues = vobj.get("issues", [])
        return list(issues) if isinstance(issues, list) else []

    def _set_vendor_issues(self, vendor: str, issues: list[str]):
        self.issue_cfg.setdefault("vendors", {})
        self.issue_cfg["vendors"].setdefault(vendor, {})
        self.issue_cfg["vendors"][vendor]["issues"] = issues

    def _get_vendor_delimiter(self, vendor: str) -> str:
        vobj = self._get_vendor_cfg(vendor)
        delim = vobj.get("delimiter", DEFAULT_DELIMITER)
        if delim is None:
            return DEFAULT_DELIMITER
        return str(delim)

    def _set_vendor_delimiter(self, vendor: str, delim: str):
        self.issue_cfg.setdefault("vendors", {})
        self.issue_cfg["vendors"].setdefault(vendor, {})
        self.issue_cfg["vendors"][vendor]["delimiter"] = "" if delim is None else str(delim)

    def _persist_issues(self, msg: str):
        try:
            self.issue_cfg = ensure_issue_config_vendor_scoped(self.issue_cfg, list(self.db.keys()))
            save_json(ISSUES_PATH, self.issue_cfg)
            self.log_info(f"{msg} (saved: {ISSUES_PATH.name})")
        except Exception as e:
            self.log_error(f"{msg} FAILED (save {ISSUES_PATH.name}) -> {e}")

    def _persist_db(self, msg: str):
        try:
            save_json(DB_PATH, self.db)
            self.log_info(f"{msg} (saved: {DB_PATH.name})")
        except Exception as e:
            self.log_error(f"{msg} FAILED (save {DB_PATH.name}) -> {e}")

    def _sync_vendor_scoped_config_with_db(self):
        if not isinstance(self.db, dict):
            self.db = self._default_db()

        vendors = list(self.db.keys())
        self.issue_cfg = ensure_issue_config_vendor_scoped(self.issue_cfg, vendors)

        changed_cfg = False
        changed_db = False

        for v in vendors:
            self.db.setdefault(v, {})
            if not isinstance(self.db[v], dict):
                self.db[v] = {}

            cfg_list = self._get_vendor_issues(v)
            if not cfg_list:
                cfg_list = default_issues()
                self._set_vendor_issues(v, cfg_list)
                changed_cfg = True

            existing_db_issues = [str(x) for x in self.db[v].keys()]
            seen = set(cfg_list)
            appended = False
            for it in existing_db_issues:
                s = str(it).strip()
                if not s or s in seen:
                    continue
                cfg_list.append(s)
                seen.add(s)
                appended = True
            if appended:
                self._set_vendor_issues(v, cfg_list)
                changed_cfg = True

            for issue_name in cfg_list:
                if issue_name not in self.db[v]:
                    self.db[v][issue_name] = self._default_issue_obj()
                    changed_db = True

            delim = self._get_vendor_delimiter(v)
            if delim is None:
                self._set_vendor_delimiter(v, DEFAULT_DELIMITER)
                changed_cfg = True

        if changed_cfg:
            try:
                save_json(ISSUES_PATH, self.issue_cfg)
                self.log_info(f"issues_config normalized (saved: {ISSUES_PATH.name})")
            except Exception as e:
                self.log_warn(f"issues_config normalize save failed -> {e}")
        if changed_db:
            try:
                save_json(DB_PATH, self.db)
                self.log_info(f"DB normalized (saved: {DB_PATH.name})")
            except Exception as e:
                self.log_warn(f"DB normalize save failed -> {e}")

    # --------------------------------------------------------
    # UI Build
    # --------------------------------------------------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style(self)
        self._font_default = ("TkDefaultFont", 9, "normal")
        self._font_bold = ("TkDefaultFont", 9, "bold")
        style.configure("Treeview", font=self._font_default)

        # LEFT: Navigation tree + controls
        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky="ns")

        ttk.Label(left, text="Navigation (Vendor / Issue / Detail)").pack(anchor="w")

        nav_box = ttk.Frame(left)
        nav_box.pack(fill=tk.BOTH, expand=False)

        self.nav_tree = ttk.Treeview(nav_box, show="tree", height=18, selectmode="browse")
        self.nav_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        nav_sb = ttk.Scrollbar(nav_box, orient="vertical", command=self.nav_tree.yview)
        nav_sb.pack(side=tk.LEFT, fill=tk.Y)
        self.nav_tree.configure(yscrollcommand=nav_sb.set)

        self.nav_tree.bind("<<TreeviewSelect>>", self.on_nav_select)

        # Vendor delimiter
        ttk.Label(left, text="Keyword Delimiter (Selected Vendor)").pack(anchor="w", pady=(10, 0))
        delim_row = ttk.Frame(left)
        delim_row.pack(fill=tk.X)
        self.delim_entry = ttk.Entry(delim_row, textvariable=self.delim_var, width=8)
        self.delim_entry.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(delim_row, text="Apply", command=self.apply_vendor_delimiter).pack(side=tk.LEFT)

        # Issue / Category CRUD (acts on current selection)
        crud = ttk.LabelFrame(left, text="CRUD (Selected Node)")
        crud.pack(fill=tk.X, pady=(10, 0))

        btn_row1 = ttk.Frame(crud)
        btn_row1.pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Button(btn_row1, text="+ Issue", command=self.add_issue).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row1, text="- Issue", command=self.delete_issue).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row1, text="R Issue", command=self.rename_issue).pack(side=tk.LEFT, padx=3)

        btn_row2 = ttk.Frame(crud)
        btn_row2.pack(fill=tk.X, padx=6, pady=(2, 6))
        ttk.Button(btn_row2, text="+ Category", command=self.add_category).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row2, text="- Category", command=self.delete_category).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_row2, text="R Category", command=self.rename_category).pack(side=tk.LEFT, padx=3)

        # Export / Import
        io_box = ttk.LabelFrame(left, text="Data I/O")
        io_box.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(io_box, text="Export...", command=self.export_data).pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Button(io_box, text="Import (Replace)...", command=self.import_data_replace).pack(fill=tk.X, padx=6, pady=(2, 6))

        # RIGHT: Keyword + Params
        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        ttk.Label(right, text="Keywords: [ ] / Summary / Group / Info / Copy / CopyNP / Preview(Rendered)").grid(row=0, column=0, sticky="w")

        self.tree = ttk.Treeview(
            right,
            columns=KEYWORD_COLS,
            show="tree headings",
            height=13,
            selectmode="extended",
        )

        self.tree.heading("#0", text="")
        self.tree.column("#0", width=34, anchor="center", stretch=False)

        headings = {
            "summary": "Summary",
            "group": "Group",
            "info": "Info",
            "copy": "Copy",
            "copynp": "CopyNP",
            "preview": "Preview",
        }
        for c in KEYWORD_COLS:
            self.tree.heading(c, text=headings.get(c, c))

        self.tree.column("summary", width=DEFAULT_KEYWORD_COL_WIDTHS["summary"], anchor="w", stretch=False)
        self.tree.column("group", width=DEFAULT_KEYWORD_COL_WIDTHS["group"], anchor="w", stretch=False)
        self.tree.column("info", width=DEFAULT_KEYWORD_COL_WIDTHS["info"], anchor="center", stretch=False)
        self.tree.column("copy", width=DEFAULT_KEYWORD_COL_WIDTHS["copy"], anchor="center", stretch=False)
        self.tree.column("copynp", width=DEFAULT_KEYWORD_COL_WIDTHS["copynp"], anchor="center", stretch=False)
        self.tree.column("preview", width=DEFAULT_KEYWORD_COL_WIDTHS["preview"], anchor="w", stretch=True)

        self.tree.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        sb = ttk.Scrollbar(right, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=1, column=1, sticky="ns", pady=(4, 0))

        try:
            self.tree.tag_configure("copied", font=self._font_bold)
        except Exception:
            pass

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_keyword_select)
        self.tree.bind("<Control-c>", self.copy_selected_keywords)
        self.tree.bind("<Control-C>", self.copy_selected_keywords)

        btns = ttk.Frame(right)
        btns.grid(row=2, column=0, sticky="e", pady=8)

        ttk.Button(btns, text="Add", command=self.add_keyword).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Edit", command=self.edit_keyword).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Remove", command=self.remove_keyword).pack(side=tk.LEFT, padx=4)

        ttk.Button(btns, text="Up", command=self.move_keyword_up).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Down", command=self.move_keyword_down).pack(side=tk.LEFT, padx=4)

        ttk.Separator(btns, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(btns, text="Select All", command=self.select_all_keywords).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Clear All", command=self.clear_all_keywords_selection).pack(side=tk.LEFT, padx=4)

        ttk.Separator(btns, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(btns, text="Copy Selected", command=self.copy_selected_keywords).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Copy Selected NP", command=self.copy_selected_keywords_no_params).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Reset UI Layout", command=self.reset_ui_layout).pack(side=tk.LEFT, padx=12)

        self.inline_box = ttk.LabelFrame(right, text="Inline Parameters (selected keyword placeholders)")
        self.inline_box.grid(row=3, column=0, sticky="ew", pady=(6, 0))

        param_box = ttk.LabelFrame(right, text="Parameters (Category-level)")
        param_box.grid(row=4, column=0, sticky="nsew", pady=(10, 0))

        self.param_tree = ttk.Treeview(param_box, columns=PARAM_COLS, show="headings", height=7)
        self.param_tree.heading("pname", text="Name")
        self.param_tree.heading("pval", text="Value")
        self.param_tree.column("pname", width=DEFAULT_PARAM_COL_WIDTHS["pname"], anchor="w", stretch=False)
        self.param_tree.column("pval", width=DEFAULT_PARAM_COL_WIDTHS["pval"], anchor="w", stretch=True)
        self.param_tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        psb = ttk.Scrollbar(param_box, orient="vertical", command=self.param_tree.yview)
        self.param_tree.configure(yscrollcommand=psb.set)
        psb.grid(row=0, column=1, sticky="ns", pady=8)

        pbtns = ttk.Frame(param_box)
        pbtns.grid(row=1, column=0, sticky="e", padx=8, pady=(0, 8))
        ttk.Button(pbtns, text="Add", command=self.add_param).pack(side=tk.LEFT, padx=4)
        ttk.Button(pbtns, text="Remove", command=self.remove_param).pack(side=tk.LEFT, padx=4)

        self.param_tree.bind("<Double-1>", self.on_param_cell_double_click)
        self.param_tree.bind("<Button-1>", self._param_editor_cancel_on_click_elsewhere)

        # (NEW) Log panel
        log_box = ttk.LabelFrame(root, text="Log")
        log_box.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        log_box.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_box, height=6, wrap="none")
        self.log_text.grid(row=0, column=0, sticky="ew", padx=(8, 0), pady=8)
        self.log_text.configure(state="disabled")

        log_sb = ttk.Scrollbar(log_box, orient="vertical", command=self.log_text.yview)
        log_sb.grid(row=0, column=1, sticky="ns", padx=(8, 8), pady=8)
        self.log_text.configure(yscrollcommand=log_sb.set)

        # status bar
        status = ttk.Frame(root)
        status.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(status, textvariable=self.status_var).pack(side=tk.LEFT)

        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        param_box.columnconfigure(0, weight=1)
        param_box.rowconfigure(0, weight=1)

    # --------------------------------------------------------
    # Navigation tree helpers
    # --------------------------------------------------------
    def _nav_iid_vendor(self, v: str) -> str:
        return f"v|{v}"

    def _nav_iid_issue(self, v: str, i: str) -> str:
        return f"i|{v}|{i}"

    def _nav_iid_detail(self, v: str, i: str, d: str) -> str:
        return f"d|{v}|{i}|{d}"

    def _parse_nav_iid(self, iid: str):
        try:
            parts = iid.split("|")
            kind = parts[0]
            if kind == "v" and len(parts) >= 2:
                return kind, parts[1], "", ""
            if kind == "i" and len(parts) >= 3:
                return kind, parts[1], parts[2], ""
            if kind == "d" and len(parts) >= 4:
                return kind, parts[1], parts[2], parts[3]
        except Exception:
            pass
        return "", "", "", ""

    def build_nav_tree(self, select_default=True, restore_path=None):
        self.nav_tree.delete(*self.nav_tree.get_children(""))
        vendors = list(self.db.keys()) if isinstance(self.db, dict) else []
        if not vendors:
            self.db = self._default_db()
            vendors = list(self.db.keys())

        self.issue_cfg = ensure_issue_config_vendor_scoped(self.issue_cfg, vendors)

        for v in vendors:
            vid = self._nav_iid_vendor(v)
            self.nav_tree.insert("", "end", iid=vid, text=v, open=True)

            issues = self._get_vendor_issues(v)
            if not issues and isinstance(self.db.get(v), dict):
                issues = list(self.db[v].keys())
            for issue in issues:
                iid = self._nav_iid_issue(v, issue)
                self.nav_tree.insert(vid, "end", iid=iid, text=issue, open=True)

                self.db.setdefault(v, {})
                self.db[v].setdefault(issue, self._default_issue_obj())
                details = list(self.db[v][issue].keys())
                if "_COMMON" in details:
                    details = ["_COMMON"] + [x for x in details if x != "_COMMON"]

                for d in details:
                    did = self._nav_iid_detail(v, issue, d)
                    self.nav_tree.insert(iid, "end", iid=did, text=d, open=False)

        if restore_path:
            v, i, d = restore_path
            target = None
            if v and i and d:
                t = self._nav_iid_detail(v, i, d)
                if self.nav_tree.exists(t):
                    target = t
            if not target and v and i:
                t = self._nav_iid_issue(v, i)
                if self.nav_tree.exists(t):
                    target = t
            if not target and v:
                t = self._nav_iid_vendor(v)
                if self.nav_tree.exists(t):
                    target = t

            if target:
                try:
                    self.nav_tree.selection_set(target)
                    self.nav_tree.see(target)
                    self._apply_nav_selection(target)
                    return
                except Exception:
                    pass

        if select_default and vendors:
            v = vendors[0]
            issues = self._get_vendor_issues(v)
            if not issues:
                issues = list(self.db.get(v, {}).keys())
            if issues:
                i = issues[0]
                d = "_COMMON" if "_COMMON" in self.db[v][i] else (list(self.db[v][i].keys())[0] if self.db[v][i] else "_COMMON")
                target = self._nav_iid_detail(v, i, d)
                if not self.nav_tree.exists(target):
                    target = self._nav_iid_issue(v, i)
                try:
                    self.nav_tree.selection_set(target)
                    self.nav_tree.see(target)
                    self._apply_nav_selection(target)
                except Exception:
                    pass

    def _init_nav_default_selection(self):
        restore = None
        if isinstance(self.ui_state, dict):
            p = self.ui_state.get("nav_path", None)
            if isinstance(p, dict):
                restore = (p.get("vendor", ""), p.get("issue", ""), p.get("detail", ""))
        self.build_nav_tree(select_default=True, restore_path=restore)

    def on_nav_select(self, _event=None):
        sel = self.nav_tree.selection()
        if not sel:
            return
        self._apply_nav_selection(sel[0])

    def _apply_nav_selection(self, nav_iid: str):
        kind, v, i, d = self._parse_nav_iid(nav_iid)
        if not kind:
            return

        if kind == "v":
            self.vendor_var.set(v)
            self.delim_var.set(self._get_vendor_delimiter(v))

            issues = self._get_vendor_issues(v)
            if not issues:
                issues = list(self.db.get(v, {}).keys())

            cur_i = self.issue_var.get()
            if cur_i in issues:
                self.issue_var.set(cur_i)
            else:
                self.issue_var.set(issues[0] if issues else "")

            issue = self.issue_var.get()
            details = list(self.db.get(v, {}).get(issue, {}).keys())
            if "_COMMON" in details:
                self.detail_var.set("_COMMON")
            else:
                self.detail_var.set(details[0] if details else "_COMMON")

        elif kind == "i":
            self.vendor_var.set(v)
            self.issue_var.set(i)
            self.delim_var.set(self._get_vendor_delimiter(v))

            details = list(self.db.get(v, {}).get(i, {}).keys())
            if "_COMMON" in details:
                self.detail_var.set("_COMMON")
            else:
                self.detail_var.set(details[0] if details else "_COMMON")

        elif kind == "d":
            self.vendor_var.set(v)
            self.issue_var.set(i)
            self.detail_var.set(d)
            self.delim_var.set(self._get_vendor_delimiter(v))

        self._ensure_path_exists(v, self.issue_var.get(), self.detail_var.get())
        self.refresh_all()

    def _ensure_path_exists(self, v, i, d):
        if not v or not i or not d:
            return
        self.db.setdefault(v, {})
        self.db[v].setdefault(i, self._default_issue_obj())
        self.db[v][i].setdefault(d, {"_keywords": [], "_params": {}})
        self._sync_vendor_scoped_config_with_db()

    # --------------------------------------------------------
    # Checkbox Images + Sync
    # --------------------------------------------------------
    def _make_checkbox_images(self):
        def make_img(checked: bool):
            img = tk.PhotoImage(width=12, height=12)
            img.put("white", to=(0, 0, 12, 12))
            for x in range(12):
                img.put("black", (x, 0))
                img.put("black", (x, 11))
            for y in range(12):
                img.put("black", (0, y))
                img.put("black", (11, y))
            if checked:
                pts = [(3, 6), (4, 7), (5, 8), (6, 7), (7, 6), (8, 5)]
                for (x, y) in pts:
                    img.put("black", (x, y))
                    if x + 1 < 12:
                        img.put("black", (x + 1, y))
            return img

        self._img_cb_off = make_img(False)
        self._img_cb_on = make_img(True)

    def _set_checkbox_for_iid(self, iid: str, checked: bool):
        try:
            self.tree.item(iid, image=(self._img_cb_on if checked else self._img_cb_off))
        except Exception:
            pass

    def _sync_checkboxes_with_selection(self):
        sel = set(self.tree.selection())
        for iid in self.tree.get_children(""):
            self._set_checkbox_for_iid(iid, iid in sel)

    def _toggle_checkbox_row(self, iid: str):
        sel = set(self.tree.selection())
        if iid in sel:
            self.tree.selection_remove(iid)
        else:
            self.tree.selection_add(iid)
        self._sync_checkboxes_with_selection()

    # --------------------------------------------------------
    # Selection restore helpers
    # --------------------------------------------------------
    def _safe_tree_restore_selection(self, iids: list[str] | None, focus_iid: str | None = None):
        if iids is None:
            iids = []
        existing = set(self.tree.get_children(""))
        valid = [iid for iid in iids if iid in existing]

        try:
            self.tree.selection_remove(self.tree.selection())
        except Exception:
            pass

        if valid:
            try:
                self.tree.selection_set(valid[0])
                for iid in valid[1:]:
                    self.tree.selection_add(iid)
            except Exception:
                pass

            tgt = focus_iid if (focus_iid in existing) else valid[0]
            try:
                self.tree.see(tgt)
            except Exception:
                pass

        self._sync_checkboxes_with_selection()

    def _safe_param_restore_selection(self, iid: str | None):
        if not iid:
            return
        try:
            if iid in self.param_tree.get_children(""):
                self.param_tree.selection_set(iid)
                self.param_tree.see(iid)
        except Exception:
            pass

    def _tree_selected_iids_sorted(self) -> list[str]:
        sel = list(self.tree.selection())
        if not sel:
            return []
        try:
            return sorted(sel, key=lambda x: int(x))
        except Exception:
            return sel

    # --------------------------------------------------------
    # Current object
    # --------------------------------------------------------
    def _current_obj(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if not v or not i or not d:
            return {"_keywords": [], "_params": {}}

        self.db.setdefault(v, {})
        self.db[v].setdefault(i, self._default_issue_obj())
        self.db[v][i].setdefault(d, {"_keywords": [], "_params": {}})

        obj = self.db[v][i][d]

        if isinstance(obj, list):
            obj = {"_keywords": normalize_keywords(obj), "_params": {}}
            self.db[v][i][d] = obj

        obj.setdefault("_keywords", [])
        obj.setdefault("_params", {})
        obj["_keywords"] = normalize_keywords(obj["_keywords"])
        if not isinstance(obj["_params"], dict):
            obj["_params"] = {}
        return obj

    def _ensure_current_obj_migrated(self):
        _ = self._current_obj()
        self._persist_db("DB migrated/normalized")

    # --------------------------------------------------------
    # Refresh
    # --------------------------------------------------------
    def refresh_all(self, *_):
        self.refresh_keywords()
        self.refresh_params()
        self.clear_inline()

        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        self.log_info(f"Selected: {v} > {i} > {d}")

    def refresh_keywords(self):
        self._clear_copy_feedback(force=True)
        self.tree.delete(*self.tree.get_children())

        obj = self._current_obj()
        params = obj["_params"]
        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        for idx, kw in enumerate(obj["_keywords"]):
            raw_joined = keyword_joined_template(kw, delim)
            summary = kw.get("summary", "")
            group = kw.get("group", "")
            preview = render_keyword(raw_joined, params)

            iid = str(idx)
            self.tree.insert("", "end", iid=iid, text="", values=(summary, group, "Info", "Copy", "CopyNP", preview))
            self._set_checkbox_for_iid(iid, checked=False)

        self._sync_checkboxes_with_selection()

    def refresh_keyword_previews_only(self):
        obj = self._current_obj()
        params = obj["_params"]
        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        for iid in self.tree.get_children(""):
            try:
                idx = int(iid)
            except Exception:
                continue
            if idx < 0 or idx >= len(obj["_keywords"]):
                continue

            kw = obj["_keywords"][idx]
            raw_joined = keyword_joined_template(kw, delim)
            preview = render_keyword(raw_joined, params)
            try:
                self.tree.set(iid, "preview", preview)
            except Exception:
                pass

    def refresh_params(self):
        self.param_tree.delete(*self.param_tree.get_children())
        params = self._current_obj()["_params"]
        for k in sorted(params.keys()):
            self.param_tree.insert("", "end", iid=k, values=(k, str(params.get(k, ""))))

    # --------------------------------------------------------
    # Inline params
    # --------------------------------------------------------
    def clear_inline(self):
        for w in self.inline_box.winfo_children():
            w.destroy()
        ttk.Label(self.inline_box, text="Select a keyword containing placeholders like {CH}, {ABC}.").pack(anchor="w", padx=8, pady=8)

    def on_keyword_select(self, *_):
        self._sync_checkboxes_with_selection()

        for w in self.inline_box.winfo_children():
            w.destroy()

        sel = self.tree.selection()
        if not sel:
            self.clear_inline()
            return

        try:
            idx = int(sel[0])
        except Exception:
            self.clear_inline()
            return

        obj = self._current_obj()
        if idx < 0 or idx >= len(obj["_keywords"]):
            self.clear_inline()
            return

        kw = obj["_keywords"][idx]
        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)
        raw_joined = keyword_joined_template(kw, delim)

        placeholders = detect_placeholders(raw_joined)
        if not placeholders:
            self.clear_inline()
            return

        for p in placeholders:
            obj["_params"].setdefault(p, "")

        header = ttk.Label(self.inline_box, text="Detected placeholders (Apply updates category-level params):")
        header.pack(anchor="w", padx=8, pady=(8, 4))

        for p in placeholders:
            row = ttk.Frame(self.inline_box)
            row.pack(anchor="w", padx=8, pady=3, fill=tk.X)
            ttk.Label(row, text=p, width=14).pack(side=tk.LEFT)

            ent = ttk.Entry(row, width=28)
            ent.insert(0, str(obj["_params"].get(p, "")))
            ent.pack(side=tk.LEFT, padx=(6, 6))

            ttk.Button(row, text="Apply", command=lambda k=p, e=ent: self.apply_inline_param(k, e.get())).pack(side=tk.LEFT)

        self.refresh_params()
        self.refresh_keyword_previews_only()

    def apply_inline_param(self, key, value):
        obj = self._current_obj()
        obj["_params"][key] = value
        self._persist_db(f"Param updated: {key}={value}")
        self.refresh_params()
        self.refresh_keyword_previews_only()

    # --------------------------------------------------------
    # Vendor delimiter
    # --------------------------------------------------------
    def apply_vendor_delimiter(self):
        v = self.vendor_var.get()
        if not v:
            return
        delim = self.delim_var.get()
        if delim is None:
            delim = DEFAULT_DELIMITER
        if str(delim) == "":
            if not messagebox.askyesno("Confirm", "Delimiter is empty. Continue?"):
                return

        saved_sel = self._tree_selected_iids_sorted()
        focus = saved_sel[0] if saved_sel else None

        self._set_vendor_delimiter(v, str(delim))
        self._persist_issues(f"Delimiter updated for {v}: '{delim}'")
        self.refresh_keywords()
        self._safe_tree_restore_selection(saved_sel, focus_iid=focus)
        self.on_keyword_select()

        self._sync_vendor_scoped_config_with_db()

    # --------------------------------------------------------
    # Select All / Clear All
    # --------------------------------------------------------
    def select_all_keywords(self):
        kids = self.tree.get_children("")
        if not kids:
            return
        self.tree.selection_set(kids)
        self._sync_checkboxes_with_selection()
        self.log_info(f"Selected all ({len(kids)})")

    def clear_all_keywords_selection(self):
        self.tree.selection_remove(self.tree.selection())
        self._sync_checkboxes_with_selection()
        self.log_info("Cleared all selections")

    # --------------------------------------------------------
    # Bulk copy selected keywords
    # --------------------------------------------------------
    def _collect_selected_joined_templates(self) -> list[str]:
        sel = list(self.tree.selection())
        if not sel:
            return []

        try:
            sel_sorted = sorted(sel, key=lambda x: int(x))
        except Exception:
            sel_sorted = sel

        obj = self._current_obj()
        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        joined_list = []
        for iid in sel_sorted:
            try:
                idx = int(iid)
            except Exception:
                continue
            if idx < 0 or idx >= len(obj["_keywords"]):
                continue
            kw = obj["_keywords"][idx]
            raw_joined = keyword_joined_template(kw, delim).strip()
            if raw_joined:
                joined_list.append(raw_joined)

        return joined_list

    def copy_selected_keywords(self, _event=None):
        joined_list = self._collect_selected_joined_templates()
        if not joined_list:
            self.log_warn("No selection to copy.")
            return

        obj = self._current_obj()
        params = obj["_params"]
        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        rendered_list = []
        for raw_joined in joined_list:
            rendered = render_keyword(raw_joined, params).strip()
            if rendered:
                rendered_list.append(rendered)

        if not rendered_list:
            self.log_warn("No valid keywords to copy.")
            return

        combined = delim.join(rendered_list)
        self.clipboard_clear()
        self.clipboard_append(combined)
        self.log_info(f"Copied Selected ({len(rendered_list)}): {combined}")

    def copy_selected_keywords_no_params(self):
        joined_list = self._collect_selected_joined_templates()
        if not joined_list:
            self.log_warn("No selection to copy.")
            return

        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        rendered_list = []
        for raw_joined in joined_list:
            np = render_keyword_without_params(raw_joined).strip()
            rendered_list.append(np)

        combined = delim.join(rendered_list)
        self.clipboard_clear()
        self.clipboard_append(combined)
        self.log_info(f"Copied Selected NP ({len(rendered_list)}): {combined}")

    # --------------------------------------------------------
    # Keyword CRUD
    # --------------------------------------------------------
    def _selected_keyword_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def add_keyword(self):
        v = self.vendor_var.get()
        if not v:
            return
        delim = self._get_vendor_delimiter(v)
        dlg = KeywordDialog(
            self,
            "Add Keyword",
            init={"parts": [""], "summary": "", "group": "", "desc": ""},
            delimiter=delim,
        )
        if dlg.result:
            obj = self._current_obj()
            obj["_keywords"].append(dlg.result)
            self._persist_db("Keyword added")

            self.refresh_keywords()
            new_iid = str(len(obj["_keywords"]) - 1)
            self._safe_tree_restore_selection([new_iid], focus_iid=new_iid)

    def edit_keyword(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return

        obj = self._current_obj()
        kw = obj["_keywords"][idx]

        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        dlg = KeywordDialog(self, "Edit Keyword", init=kw, delimiter=delim)
        if dlg.result:
            obj["_keywords"][idx] = dlg.result
            self._persist_db("Keyword edited")

            saved = self._tree_selected_iids_sorted()
            focus = str(idx)
            self.refresh_keywords()
            self._safe_tree_restore_selection(saved or [focus], focus_iid=focus)

    def remove_keyword(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return

        obj = self._current_obj()
        del obj["_keywords"][idx]
        self._persist_db("Keyword removed")

        self.refresh_keywords()

        new_len = len(obj["_keywords"])
        if new_len <= 0:
            self.clear_inline()
            return
        new_idx = min(idx, new_len - 1)
        focus = str(new_idx)
        self._safe_tree_restore_selection([focus], focus_iid=focus)
        self.clear_inline()

    # --------------------------------------------------------
    # Up / Down (single selection only)
    # --------------------------------------------------------
    def move_keyword_up(self):
        sel = self._tree_selected_iids_sorted()
        if not sel:
            return
        if len(sel) != 1:
            self.log_warn("Up/Down은 단일 선택 1개에서만 동작합니다.")
            return

        try:
            idx = int(sel[0])
        except Exception:
            return

        obj = self._current_obj()
        kws = obj.get("_keywords", [])
        if not isinstance(kws, list) or idx <= 0 or idx >= len(kws):
            return

        kws[idx - 1], kws[idx] = kws[idx], kws[idx - 1]
        obj["_keywords"] = kws
        self._persist_db("Keyword moved up")

        self.refresh_keywords()
        focus = str(idx - 1)
        self._safe_tree_restore_selection([focus], focus_iid=focus)

    def move_keyword_down(self):
        sel = self._tree_selected_iids_sorted()
        if not sel:
            return
        if len(sel) != 1:
            self.log_warn("Up/Down은 단일 선택 1개에서만 동작합니다.")
            return

        try:
            idx = int(sel[0])
        except Exception:
            return

        obj = self._current_obj()
        kws = obj.get("_keywords", [])
        if not isinstance(kws, list) or idx < 0 or idx >= len(kws) - 1:
            return

        kws[idx + 1], kws[idx] = kws[idx], kws[idx + 1]
        obj["_keywords"] = kws
        self._persist_db("Keyword moved down")

        self.refresh_keywords()
        focus = str(idx + 1)
        self._safe_tree_restore_selection([focus], focus_iid=focus)

    # --------------------------------------------------------
    # Keyword Tree click handlers
    # --------------------------------------------------------
    def on_tree_click(self, event):
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if not row:
            return

        if col == "#0":
            self._toggle_checkbox_row(row)
            return "break"

        try:
            idx = int(row)
        except Exception:
            return

        obj = self._current_obj()
        if idx < 0 or idx >= len(obj["_keywords"]):
            return

        params = obj["_params"]
        kw = obj["_keywords"][idx]

        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)
        raw_joined = keyword_joined_template(kw, delim)

        # columns: #0 checkbox, #1 summary, #2 group, #3 info, #4 copy, #5 copynp, #6 preview
        if col == "#3":  # Info
            InfoPopup(
                self,
                title="Keyword Description",
                summary=kw.get("summary", ""),
                group=kw.get("group", ""),
                keyword=raw_joined,
                desc=kw.get("desc", ""),
                desc_rich=kw.get("desc_rich", None),
            )

        elif col == "#4":  # Copy (with params)
            rendered = render_keyword(raw_joined, params)
            self.clipboard_clear()
            self.clipboard_append(rendered)
            self._show_copy_feedback(row, which="copy")
            self.log_info(f"Copied: {rendered}")

        elif col == "#5":  # CopyNP (without params)
            rendered = render_keyword_without_params(raw_joined)
            self.clipboard_clear()
            self.clipboard_append(rendered)
            self._show_copy_feedback(row, which="copynp")
            self.log_info(f"Copied NP: {rendered}")

    def on_tree_double_click(self, event):
        col = self.tree.identify_column(event.x)
        if col in ("#0", "#3", "#4", "#5"):
            return
        self.edit_keyword()

    # --------------------------------------------------------
    # Copy feedback UI
    # --------------------------------------------------------
    def _show_copy_feedback(self, row_iid: str, which: str = "copy"):
        self._clear_copy_feedback(force=True)

        try:
            vals = list(self.tree.item(row_iid, "values"))
            if len(vals) == len(KEYWORD_COLS):
                if which == "copynp":
                    vals[4] = "Copied"
                else:
                    vals[3] = "Copied"
                self.tree.item(row_iid, values=tuple(vals))
        except Exception:
            pass

        try:
            self.tree.item(row_iid, tags=("copied",))
        except Exception:
            pass

        self._copy_feedback_row = row_iid
        self._copy_feedback_after_id = self.after(COPY_FEEDBACK_MS, self._clear_copy_feedback)

    def _clear_copy_feedback(self, force=False):
        if self._copy_feedback_after_id is not None:
            try:
                self.after_cancel(self._copy_feedback_after_id)
            except Exception:
                pass
            self._copy_feedback_after_id = None

        row_iid = self._copy_feedback_row
        self._copy_feedback_row = None
        if not row_iid:
            return

        try:
            exists = row_iid in self.tree.get_children("")
        except Exception:
            exists = False
        if not exists:
            return

        try:
            vals = list(self.tree.item(row_iid, "values"))
            if len(vals) == len(KEYWORD_COLS):
                vals[3] = "Copy"
                vals[4] = "CopyNP"
                self.tree.item(row_iid, values=tuple(vals))
        except Exception:
            pass

        try:
            self.tree.item(row_iid, tags=())
        except Exception:
            pass

    # --------------------------------------------------------
    # Param panel CRUD
    # --------------------------------------------------------
    def add_param(self):
        name = simpledialog.askstring("Add Param", "Param name (e.g., CH, ABC):")
        if not name:
            return
        name = name.strip()
        if not name:
            return

        obj = self._current_obj()
        if name in obj["_params"]:
            messagebox.showwarning("Warning", "Param already exists.")
            return

        val = simpledialog.askstring("Add Param", f"Value for {name}:") or ""
        obj["_params"][name] = val
        self._persist_db("Param added")

        self.refresh_params()
        self._safe_param_restore_selection(name)
        self.refresh_keyword_previews_only()

    def remove_param(self):
        sel = self.param_tree.selection()
        if not sel:
            return
        pname = sel[0]

        obj = self._current_obj()
        if pname not in obj["_params"]:
            return

        if not messagebox.askyesno("Confirm", f"Remove param '{pname}'?"):
            return

        del obj["_params"][pname]
        self._persist_db("Param removed")

        self.refresh_params()
        self.refresh_keyword_previews_only()

    # --------------------------------------------------------
    # Param in-place edit
    # --------------------------------------------------------
    def on_param_cell_double_click(self, event):
        if self._param_editor:
            self._cancel_param_edit()

        row = self.param_tree.identify_row(event.y)
        col = self.param_tree.identify_column(event.x)
        if not row or col != "#2":
            return

        bbox = self.param_tree.bbox(row, "pval")
        if not bbox:
            return

        x, y, w, h = bbox
        obj = self._current_obj()
        val = str(obj["_params"].get(row, ""))

        e = ttk.Entry(self.param_tree)
        e.insert(0, val)
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set()
        e.selection_range(0, tk.END)

        self._param_editor = e
        self._param_editing = row

        e.bind("<Return>", lambda _: self._commit_param_edit())
        e.bind("<Escape>", lambda _: self._cancel_param_edit())

    def _commit_param_edit(self):
        if not self._param_editor or not self._param_editing:
            return

        editing_key = self._param_editing
        new_val = self._param_editor.get()

        obj = self._current_obj()
        obj["_params"][editing_key] = new_val
        self._persist_db(f"Param updated: {editing_key}={new_val}")

        self._cancel_param_edit()
        self.refresh_params()
        self._safe_param_restore_selection(editing_key)
        self.refresh_keyword_previews_only()

    def _cancel_param_edit(self):
        if self._param_editor:
            try:
                self._param_editor.destroy()
            except Exception:
                pass
        self._param_editor = None
        self._param_editing = None

    def _param_editor_cancel_on_click_elsewhere(self, _event):
        if self._param_editor:
            self._cancel_param_edit()

    # --------------------------------------------------------
    # Detail Category CRUD
    # --------------------------------------------------------
    def add_category(self):
        v, i = self.vendor_var.get(), self.issue_var.get()
        if not v or not i:
            messagebox.showinfo("Info", "Issue를 먼저 선택하세요. (Navigation에서 Issue 또는 Detail 선택)")
            return

        name = simpledialog.askstring("Add Category", "Category name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return

        self.db.setdefault(v, {})
        self.db[v].setdefault(i, self._default_issue_obj())

        if name in self.db[v][i]:
            messagebox.showwarning("Warning", "Category already exists.")
            return

        self.db[v][i][name] = {"_keywords": [], "_params": {}}
        self._persist_db("Category added")

        restore = (v, i, name)
        self.build_nav_tree(select_default=True, restore_path=restore)

    def delete_category(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if not v or not i or not d:
            return
        if d == "_COMMON":
            messagebox.showinfo("Info", "_COMMON cannot be deleted.")
            return
        if not messagebox.askyesno("Confirm", f"Delete category '{d}'?"):
            return

        try:
            del self.db[v][i][d]
        except Exception:
            return
        self._persist_db("Category deleted")

        details = list(self.db[v][i].keys())
        new_d = "_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON")
        self.build_nav_tree(select_default=True, restore_path=(v, i, new_d))

    def rename_category(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if not v or not i or not d:
            return
        if d == "_COMMON":
            messagebox.showinfo("Info", "_COMMON cannot be renamed.")
            return

        new = simpledialog.askstring("Rename Category", "New name:", initialvalue=d)
        if not new:
            return
        new = new.strip()
        if not new:
            return

        if new in self.db[v][i]:
            messagebox.showwarning("Warning", "Category already exists.")
            return

        self.db[v][i][new] = self.db[v][i].pop(d)
        self._persist_db("Category renamed")

        self.build_nav_tree(select_default=True, restore_path=(v, i, new))

    # --------------------------------------------------------
    # Issue CRUD (Vendor-scoped)
    # --------------------------------------------------------
    def add_issue(self):
        v = self.vendor_var.get()
        if not v:
            messagebox.showinfo("Info", "Vendor를 먼저 선택하세요.")
            return

        name = simpledialog.askstring("Add Issue", "Issue name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return

        issues = self._get_vendor_issues(v)
        if name in issues:
            messagebox.showwarning("Warning", "Issue already exists for this vendor.")
            return

        issues.append(name)
        self._set_vendor_issues(v, issues)
        self._persist_issues(f"Issue added for {v}")

        self.db.setdefault(v, {})
        self.db[v].setdefault(name, self._default_issue_obj())
        self._persist_db(f"DB synced (issue added) for {v}")

        self.build_nav_tree(select_default=True, restore_path=(v, name, "_COMMON"))

    def delete_issue(self):
        v = self.vendor_var.get()
        cur = self.issue_var.get()
        if not v or not cur:
            messagebox.showinfo("Info", "Issue를 먼저 선택하세요.")
            return

        issues = self._get_vendor_issues(v)
        if cur not in issues:
            return

        if len(issues) <= 1:
            messagebox.showinfo("Info", "At least one issue must remain for this vendor.")
            return

        if not messagebox.askyesno("Confirm", f"Delete issue '{cur}' for vendor '{v}'?"):
            return

        issues = [x for x in issues if x != cur]
        self._set_vendor_issues(v, issues)
        self._persist_issues(f"Issue deleted for {v}")

        if isinstance(self.db.get(v), dict) and cur in self.db[v]:
            del self.db[v][cur]
        self._persist_db(f"DB synced (issue deleted) for {v}")

        new_issue = issues[0]
        self.build_nav_tree(select_default=True, restore_path=(v, new_issue, "_COMMON"))

    def rename_issue(self):
        v = self.vendor_var.get()
        cur = self.issue_var.get()
        if not v or not cur:
            messagebox.showinfo("Info", "Issue를 먼저 선택하세요.")
            return

        issues = self._get_vendor_issues(v)
        if cur not in issues:
            return

        new = simpledialog.askstring("Rename Issue", "New issue name:", initialvalue=cur)
        if not new:
            return
        new = new.strip()
        if not new:
            return
        if new == cur:
            return
        if new in issues:
            messagebox.showwarning("Warning", "Issue already exists for this vendor.")
            return

        issues = [new if x == cur else x for x in issues]
        self._set_vendor_issues(v, issues)
        self._persist_issues(f"Issue renamed for {v}")

        self.db.setdefault(v, {})
        if cur in self.db[v]:
            self.db[v][new] = self.db[v].pop(cur)
        else:
            self.db[v][new] = self._default_issue_obj()
        self._persist_db(f"DB synced (issue renamed) for {v}")

        self.build_nav_tree(select_default=True, restore_path=(v, new, "_COMMON"))

    # --------------------------------------------------------
    # Export / Import (Replace ONLY)
    # --------------------------------------------------------
    def export_data(self):
        pkg = export_package(self.db, self.issue_cfg, self.ui_state if isinstance(self.ui_state, dict) else {})
        default_name = f"keyword_guide_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"

        path = filedialog.asksaveasfilename(
            title="Export Keyword Guide Data",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            txt = json.dumps(pkg, indent=2, ensure_ascii=False)
            _safe_atomic_write_text(Path(path), txt, encoding="utf-8")
            self.log_info(f"Exported: {path}")
        except Exception as e:
            self.log_error(f"Export Failed: {e}")
            messagebox.showerror("Export Failed", str(e))

    def _load_import_file(self) -> dict | None:
        path = filedialog.askopenfilename(
            title="Import Keyword Guide Data (Replace)",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return None

        try:
            txt = Path(path).read_text(encoding="utf-8")
            pkg = json.loads(txt)
        except Exception as e:
            self.log_error(f"Import read/parse error: {e}")
            messagebox.showerror("Import Failed", f"File read/parse error:\n{e}")
            return None

        ok, msg = validate_import_package(pkg)
        if not ok:
            self.log_error(f"Import invalid package: {msg}")
            messagebox.showerror("Import Failed", msg)
            return None

        pkg["_import_path"] = path
        return pkg

    def import_data_replace(self):
        pkg = self._load_import_file()
        if not pkg:
            return

        path = pkg.get("_import_path", "")

        notice = (
            "IMPORT (REPLACE) 안내\n\n"
            "이 작업은 현재 앱의 DB/설정 데이터를 IMPORT 파일로 '완전 교체'합니다.\n"
            "- 기존 데이터는 자동 백업되지 않습니다.\n"
            "- 필요하다면 IMPORT 전에 Export로 수동 백업을 수행하세요.\n\n"
            f"Import File:\n{path}\n\n"
            "계속 진행할까요?"
        )
        if not messagebox.askyesno("Confirm Import (Replace)", notice):
            self.log_info("Import cancelled.")
            return

        try:
            new_db = pkg.get("db", {})
            new_cfg = pkg.get("issue_cfg", {})
            new_ui = pkg.get("ui_state", {})

            vendors = list(new_db.keys()) if isinstance(new_db, dict) else []
            new_cfg = ensure_issue_config_vendor_scoped(new_cfg, vendors)

            self.db = new_db if isinstance(new_db, dict) else self._default_db()
            self.issue_cfg = new_cfg if isinstance(new_cfg, dict) else ensure_issue_config_vendor_scoped({}, list(self.db.keys()))

            if isinstance(new_ui, dict) and new_ui:
                self.ui_state = new_ui

            self._sync_vendor_scoped_config_with_db()

            save_json(DB_PATH, self.db)
            save_json(ISSUES_PATH, self.issue_cfg)
            if isinstance(self.ui_state, dict):
                save_json(UI_STATE_PATH, self.ui_state)

            self._apply_saved_widths()
            self.build_nav_tree(select_default=True, restore_path=None)
            self.refresh_all()
            self.log_info(f"Imported (Replace): {path}")
        except Exception as e:
            self.log_error(f"Import Failed: {e}")
            messagebox.showerror("Import Failed", str(e))

    # --------------------------------------------------------
    # UI State
    # --------------------------------------------------------
    def _apply_saved_widths(self):
        cols = self.ui_state.get("keyword_tree_cols", {}) if isinstance(self.ui_state, dict) else {}
        for c, w in cols.items():
            if c in DEFAULT_KEYWORD_COL_WIDTHS and c in KEYWORD_COLS:
                try:
                    self.tree.column(c, width=int(w))
                except Exception:
                    pass

        try:
            w0 = (self.ui_state.get("keyword_tree_col0_width", 34) if isinstance(self.ui_state, dict) else 34)
            self.tree.column("#0", width=int(w0))
        except Exception:
            pass

        pcols = self.ui_state.get("param_tree_cols", {}) if isinstance(self.ui_state, dict) else {}
        for c, w in pcols.items():
            if c in DEFAULT_PARAM_COL_WIDTHS:
                try:
                    self.param_tree.column(c, width=int(w))
                except Exception:
                    pass

    def reset_ui_layout(self):
        self.geometry(DEFAULT_GEOMETRY)
        for c, w in DEFAULT_KEYWORD_COL_WIDTHS.items():
            if c in KEYWORD_COLS:
                self.tree.column(c, width=w)
        try:
            self.tree.column("#0", width=34)
        except Exception:
            pass
        for c, w in DEFAULT_PARAM_COL_WIDTHS.items():
            self.param_tree.column(c, width=w)
        self.log_info("UI layout reset to defaults.")

    def on_close(self):
        nav_path = {"vendor": self.vendor_var.get(), "issue": self.issue_var.get(), "detail": self.detail_var.get()}

        state = {
            "geometry": self.geometry(),
            "nav_path": nav_path,
            "keyword_tree_col0_width": self.tree.column("#0", "width"),
            "keyword_tree_cols": {c: self.tree.column(c, "width") for c in KEYWORD_COLS},
            "param_tree_cols": {c: self.param_tree.column(c, "width") for c in DEFAULT_PARAM_COL_WIDTHS},
        }

        try:
            save_json(UI_STATE_PATH, state)
            self.log_info(f"UI state saved: {UI_STATE_PATH.name}")
        except Exception as e:
            self.log_error(f"UI state save FAILED -> {e}")

        try:
            save_json(DB_PATH, self.db)
            self.log_info(f"DB saved: {DB_PATH.name}")
        except Exception as e:
            self.log_error(f"DB save FAILED -> {e}")

        self._persist_issues("Issue config saved")
        self.destroy()


if __name__ == "__main__":
    KeywordGuideApp().mainloop()
