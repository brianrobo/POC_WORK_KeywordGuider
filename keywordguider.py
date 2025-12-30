# ============================================================
# Keyword Guide UI POC
# ------------------------------------------------------------
# Version: 0.4.5 (2025-12-30)
#
# Release Notes (v0.4.5)
# - (UI) Left navigation 변경: Vendor/Issue/Detail Combobox -> Tree Navigator
#        - Vendor -> Issue -> Detail Category 계층 구조
#        - 기본 expand/collapse(+/-) 제공 (ttk.Treeview)
# - (UX) Nav tree rebuild 시 selection 유지(restore_path) + DB/IssueCfg 정합성 보정 포함
#
# Existing (v0.4.3 유지)
# - Window title에 release version 표시
# - Param in-place edit: selection_set(None) 오류 방지 + selection 복원 공통 패턴화
# - Copy 시 parameter 치환 없이 복사하는 기능 추가
#   - Row Copy: "Copy" / "CopyNP(No Params)"
#   - Copy Selected: "Copy Selected" / "Copy Selected NP"
# - Keyword에 "Group" 필드
# - Tree selection bug fix (선택 이벤트에서 tree rebuild 제거)
# - Keyword Description Rich Text (Bold + 3 Colors) with desc_rich runs
# - Keyword List: 왼쪽(#0) 체크박스 + extended selection
# - Select All / Clear All / Copy Selected
# - Up/Down(단일 선택 1개 기준) 위치 변경
# - Vendor별 Issue 관리 + Vendor별 Delimiter 저장
# - Placeholder detect + Inline Apply (category-level params)
# - Category-level params in-place edit (Enter/Esc)
# - UI state persistence (column widths, geometry)
# - KeywordDialog: Parts fixed height + scroll, Preview wrap+scroll
# - Import Joined String split UX: confirm
# - Part 입력 중 delimiter 포함 시 split 제안: confirm(자동 아님)
# ============================================================

import json
import re
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# ------------------------------------------------------------
# Paths / Constants
# ------------------------------------------------------------
APP_VERSION = "0.4.5"

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "keywords_db.json"
UI_STATE_PATH = BASE_DIR / "ui_state.json"
ISSUES_PATH = BASE_DIR / "issues_config.json"

PLACEHOLDER_RE = re.compile(r"\{([A-Za-z0-9_]+)\}")
DEFAULT_GEOMETRY = "1250x760"

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


# ------------------------------------------------------------
# Utility
# ------------------------------------------------------------
def load_json(path: Path):
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_json(path: Path, data: dict):
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


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

        ttk.Label(frm, text=f"Import Joined String (delimiter: '{self.delimiter}')").grid(
            row=4, column=0, sticky="w"
        )

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

        self._bind_mousewheel(self.parts_canvas)
        self._bind_mousewheel(self.parts_container)

        add_btn_row = ttk.Frame(frm)
        add_btn_row.grid(row=8, column=0, sticky="w", pady=(0, 8))
        ttk.Button(add_btn_row, text="+ Add Part", command=self._add_part_row).pack(side=tk.LEFT)

        ttk.Label(frm, text="Joined Keyword Preview (auto wrap, scrollable)").grid(
            row=9, column=0, sticky="w", pady=(6, 0)
        )

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

        self.txt_desc = tk.Text(frm, height=10, wrap="word")
        self.txt_desc.grid(row=13, column=0, sticky="nsew", pady=(2, 10))

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

    def _bind_mousewheel(self, widget):
        widget.bind("<Enter>", lambda _e: self._activate_mousewheel(True))
        widget.bind("<Leave>", lambda _e: self._activate_mousewheel(False))

    def _activate_mousewheel(self, active: bool):
        if active:
            self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
            self.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
            self.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")
        else:
            try:
                self.unbind_all("<MouseWheel>")
                self.unbind_all("<Button-4>")
                self.unbind_all("<Button-5>")
            except Exception:
                pass

    def _on_mousewheel(self, event):
        try:
            self.parts_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _on_mousewheel_linux(self, event):
        try:
            if event.num == 4:
                self.parts_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.parts_canvas.yview_scroll(1, "units")
        except Exception:
            pass

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
        txt = tk.Text(frm, wrap="word")
        txt.grid(row=7, column=0, sticky="nsew")

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

        # [주의1] DB/IssueCfg 정합성 보정: nav_tree build 전에 반드시 수행
        self._sync_vendor_scoped_config_with_db()

        self.vendor_var = tk.StringVar()
        self.issue_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.delim_var = tk.StringVar(value=DEFAULT_DELIMITER)

        self._param_editor = None
        self._param_editing = None

        self._copy_feedback_after_id = None
        self._copy_feedback_row = None

        self.geometry(self.ui_state.get("geometry", DEFAULT_GEOMETRY))
        self._build_ui()
        self._make_checkbox_images()

        self._apply_saved_widths()

        # nav tree build + default selection
        self._init_vendor()

        self._ensure_current_obj_migrated()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # --------------------------------------------------------
    # Selection restore helpers (공통 패턴)
    # --------------------------------------------------------
    def _safe_tree_restore_selection(self, iids: list[str] | None, focus_iid: str | None = None):
        """Keyword tree selection 복원 (존재하는 iid만) + checkbox sync"""
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
        """param tree selection 복원"""
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
            self.status_var.set(msg)
        except Exception as e:
            self.status_var.set(f"Issue config save failed: {e}")

    def _persist_db(self, msg: str):
        try:
            save_json(DB_PATH, self.db)
            self.status_var.set(msg)
        except Exception as e:
            self.status_var.set(f"Save failed: {e}")

    def _sync_vendor_scoped_config_with_db(self):
        """
        [주의1] DB/IssueCfg 정합성 보정:
          - vendor list, issue list, delimiter 값, db issue 키 등 상호 보정
          - nav_tree build 전에 반드시 호출
        """
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
            except Exception:
                pass
        if changed_db:
            try:
                save_json(DB_PATH, self.db)
            except Exception:
                pass

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

        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky="ns")

        # ===== Left Navigator Tree =====
        ttk.Label(left, text="Navigator (Vendor > Issue > Detail)").pack(anchor="w")
        self.nav_tree = ttk.Treeview(left, show="tree", selectmode="browse", height=24)
        self.nav_tree.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        nav_sb = ttk.Scrollbar(left, orient="vertical", command=self.nav_tree.yview)
        self.nav_tree.configure(yscrollcommand=nav_sb.set)
        nav_sb.place(in_=self.nav_tree, relx=1.0, rely=0, relheight=1.0, anchor="ne")

        self.nav_tree.bind("<<TreeviewSelect>>", self.on_nav_select)

        # Vendor delimiter control
        ttk.Label(left, text="Keyword Delimiter (Vendor)").pack(anchor="w", pady=(6, 0))
        delim_row = ttk.Frame(left)
        delim_row.pack(fill=tk.X)
        self.delim_entry = ttk.Entry(delim_row, textvariable=self.delim_var, width=8)
        self.delim_entry.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(delim_row, text="Apply", command=self.apply_vendor_delimiter).pack(side=tk.LEFT)

        # CRUD controls depending on selection depth
        crud = ttk.LabelFrame(left, text="Edit (Issue/Detail)")
        crud.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(crud, text="+ Add", command=self.nav_add).pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Button(crud, text="- Delete", command=self.nav_delete).pack(fill=tk.X, padx=6, pady=2)
        ttk.Button(crud, text="R Rename", command=self.nav_rename).pack(fill=tk.X, padx=6, pady=(2, 6))

        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        ttk.Label(right, text="Keywords: [ ] / Summary / Group / Info / Copy / CopyNP / Preview(Rendered)").grid(
            row=0, column=0, sticky="w"
        )

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

        status = ttk.Frame(root)
        status.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Label(status, textvariable=self.status_var).pack(side=tk.LEFT)

        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        param_box.columnconfigure(0, weight=1)
        param_box.rowconfigure(0, weight=1)

    # --------------------------------------------------------
    # Left Navigator: Build / Select / CRUD Dispatcher
    # --------------------------------------------------------
    def _init_vendor(self):
        # [주의2] nav_tree rebuild 시 selection 유지: restore_path 사용
        self.build_nav_tree(select_default=True)

    def _nav_path(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if v and i and d:
            return (v, i, d)
        return None

    def build_nav_tree(self, select_default=False, restore_path=None):
        """
        [주의2] nav_tree rebuild 시 selection 유지:
          - restore_path = (vendor, issue, detail) 형태로 selection 복원
          - 복원 실패 시 select_default 로 fallback
        """
        # build 전에 정합성 보정(안전)
        self._sync_vendor_scoped_config_with_db()

        self.nav_tree.delete(*self.nav_tree.get_children(""))

        vendors = list(self.db.keys())
        for v in vendors:
            vid = f"V|{v}"
            self.nav_tree.insert("", "end", iid=vid, text=v, open=True)

            issues = self._get_vendor_issues(v)
            self.db.setdefault(v, {})
            for issue in issues:
                self.db[v].setdefault(issue, self._default_issue_obj())
                iid = f"I|{v}|{issue}"
                self.nav_tree.insert(vid, "end", iid=iid, text=issue, open=True)

                details = list(self.db[v][issue].keys())
                for d in details:
                    did = f"D|{v}|{issue}|{d}"
                    self.nav_tree.insert(iid, "end", iid=did, text=d, open=False)

        # restore selection
        if restore_path:
            v, i, d = restore_path
            did = f"D|{v}|{i}|{d}"
            if self.nav_tree.exists(did):
                self.nav_tree.selection_set(did)
                self.nav_tree.see(did)
                self.on_nav_select()
                return
            iid = f"I|{v}|{i}"
            if self.nav_tree.exists(iid):
                self.nav_tree.selection_set(iid)
                self.nav_tree.see(iid)
                self.on_nav_select()
                return
            vid = f"V|{v}"
            if self.nav_tree.exists(vid):
                self.nav_tree.selection_set(vid)
                self.nav_tree.see(vid)
                self.on_nav_select()
                return

        # default selection
        if select_default and vendors:
            v = vendors[0]
            issues = self._get_vendor_issues(v)
            if issues:
                i = issues[0]
                details = list(self.db[v][i].keys())
                d = "_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON")
                did = f"D|{v}|{i}|{d}"
                if self.nav_tree.exists(did):
                    self.nav_tree.selection_set(did)
                    self.nav_tree.see(did)
                    self.on_nav_select()
                    return
                iid = f"I|{v}|{i}"
                if self.nav_tree.exists(iid):
                    self.nav_tree.selection_set(iid)
                    self.nav_tree.see(iid)
                    self.on_nav_select()
                    return
            vid = f"V|{v}"
            if self.nav_tree.exists(vid):
                self.nav_tree.selection_set(vid)
                self.nav_tree.see(vid)
                self.on_nav_select()
                return

    def on_nav_select(self, _event=None):
        sel = self.nav_tree.selection()
        if not sel:
            return
        node = sel[0]
        parts = node.split("|")
        if not parts:
            return
        kind = parts[0]

        if kind == "V":
            v = parts[1]
            self.vendor_var.set(v)
            self.delim_var.set(self._get_vendor_delimiter(v))

            issues = self._get_vendor_issues(v)
            if issues:
                i = issues[0]
                self.issue_var.set(i)
                details = list(self.db[v][i].keys())
                d = "_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON")
                self.detail_var.set(d)
            self.refresh_all()
            return

        if kind == "I":
            v = parts[1]
            i = parts[2]
            self.vendor_var.set(v)
            self.issue_var.set(i)
            self.delim_var.set(self._get_vendor_delimiter(v))
            details = list(self.db[v][i].keys())
            d = "_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON")
            self.detail_var.set(d)
            self.refresh_all()
            return

        if kind == "D":
            v = parts[1]
            i = parts[2]
            d = "|".join(parts[3:])  # detail에 '|' 포함 가능성 대비
            self.vendor_var.set(v)
            self.issue_var.set(i)
            self.detail_var.set(d)
            self.delim_var.set(self._get_vendor_delimiter(v))
            self.refresh_all()
            return

    def _nav_get_context(self):
        sel = self.nav_tree.selection()
        if not sel:
            return ("", None)
        node = sel[0]
        parts = node.split("|")
        kind = parts[0] if parts else ""
        return (kind, parts)

    def nav_add(self):
        kind, _parts = self._nav_get_context()
        if kind == "V":
            self.add_issue()
        else:
            self.add_category()
        self.build_nav_tree(restore_path=self._nav_path())

    def nav_delete(self):
        kind, _parts = self._nav_get_context()
        if kind == "I":
            self.delete_issue()
        elif kind == "D":
            self.delete_category()
        else:
            messagebox.showinfo("Info", "Delete는 Issue 또는 Detail 선택 시 동작합니다.")
            return
        self.build_nav_tree(restore_path=self._nav_path())

    def nav_rename(self):
        kind, _parts = self._nav_get_context()
        if kind == "I":
            self.rename_issue()
        elif kind == "D":
            self.rename_category()
        else:
            messagebox.showinfo("Info", "Rename은 Issue 또는 Detail 선택 시 동작합니다.")
            return
        self.build_nav_tree(restore_path=self._nav_path())

    # --------------------------------------------------------
    # Delimiter
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

    # --------------------------------------------------------
    # Data helpers
    # --------------------------------------------------------
    def _current_obj(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if not v:
            v = list(self.db.keys())[0]
            self.vendor_var.set(v)
        if not i:
            issues = self._get_vendor_issues(v)
            i = issues[0] if issues else default_issues()[0]
            self.issue_var.set(i)
        if not d:
            d = "_COMMON"
            self.detail_var.set(d)

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

        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        self.status_var.set(f"Selected: {v} > {i} > {d}")

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
        ttk.Label(self.inline_box, text="Select a keyword containing placeholders like {CH}, {ABC}.").pack(
            anchor="w", padx=8, pady=8
        )

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

            ttk.Button(row, text="Apply", command=lambda k=p, e=ent: self.apply_inline_param(k, e.get())).pack(
                side=tk.LEFT
            )

        self.refresh_params()
        self.refresh_keyword_previews_only()

    def apply_inline_param(self, key, value):
        obj = self._current_obj()
        obj["_params"][key] = value
        self._persist_db(f"Param updated: {key}={value}")
        self.refresh_params()
        self.refresh_keyword_previews_only()

    # --------------------------------------------------------
    # Select All / Clear All
    # --------------------------------------------------------
    def select_all_keywords(self):
        kids = self.tree.get_children("")
        if not kids:
            return
        self.tree.selection_set(kids)
        self._sync_checkboxes_with_selection()
        self.status_var.set(f"Selected all ({len(kids)})")

    def clear_all_keywords_selection(self):
        self.tree.selection_remove(self.tree.selection())
        self._sync_checkboxes_with_selection()
        self.status_var.set("Cleared all selections")

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
            self.status_var.set("No selection to copy.")
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
            self.status_var.set("No valid keywords to copy.")
            return

        combined = delim.join(rendered_list)
        self.clipboard_clear()
        self.clipboard_append(combined)
        self.status_var.set(f"Copied Selected ({len(rendered_list)}): {combined}")

    def copy_selected_keywords_no_params(self):
        joined_list = self._collect_selected_joined_templates()
        if not joined_list:
            self.status_var.set("No selection to copy.")
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
        self.status_var.set(f"Copied Selected NP ({len(rendered_list)}): {combined}")

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
        saved = self._tree_selected_iids_sorted()
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
            self.status_var.set("Up/Down은 단일 선택 1개에서만 동작합니다.")
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
            self.status_var.set("Up/Down은 단일 선택 1개에서만 동작합니다.")
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
            self.status_var.set(f"Copied: {rendered}")
        elif col == "#5":  # CopyNP (without params)
            rendered = render_keyword_without_params(raw_joined)
            self.clipboard_clear()
            self.clipboard_append(rendered)
            self._show_copy_feedback(row, which="copynp")
            self.status_var.set(f"Copied NP: {rendered}")

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
        self.detail_var.set(name)
        self.refresh_all()

    def delete_category(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if not v or not i or not d:
            return
        if d == "_COMMON":
            messagebox.showinfo("Info", "_COMMON cannot be deleted.")
            return

        if not messagebox.askyesno("Confirm", f"Delete category '{d}'?"):
            return

        del self.db[v][i][d]
        self._persist_db("Category deleted")

        details = list(self.db[v][i].keys())
        self.detail_var.set("_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON"))
        self.refresh_all()

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

        self.detail_var.set(new)
        self.refresh_all()

    # --------------------------------------------------------
    # Issue CRUD (Vendor-scoped)
    # --------------------------------------------------------
    def add_issue(self):
        v = self.vendor_var.get()
        if not v:
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

        self.issue_var.set(name)
        details = list(self.db[v][name].keys())
        self.detail_var.set("_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON"))
        self.refresh_all()

    def delete_issue(self):
        v = self.vendor_var.get()
        cur = self.issue_var.get()
        if not v or not cur:
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

        self.issue_var.set(issues[0])
        details = list(self.db[v][issues[0]].keys())
        self.detail_var.set("_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON"))
        self.refresh_all()

    def rename_issue(self):
        v = self.vendor_var.get()
        cur = self.issue_var.get()
        if not v or not cur:
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

        self.issue_var.set(new)
        details = list(self.db[v][new].keys())
        self.detail_var.set("_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON"))
        self.refresh_all()

    # --------------------------------------------------------
    # UI State
    # --------------------------------------------------------
    def _apply_saved_widths(self):
        cols = self.ui_state.get("keyword_tree_cols", {})
        for c, w in cols.items():
            if c in DEFAULT_KEYWORD_COL_WIDTHS and c in KEYWORD_COLS:
                try:
                    self.tree.column(c, width=int(w))
                except Exception:
                    pass

        try:
            w0 = self.ui_state.get("keyword_tree_col0_width", 34)
            self.tree.column("#0", width=int(w0))
        except Exception:
            pass

        pcols = self.ui_state.get("param_tree_cols", {})
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
        self.status_var.set("UI layout reset to defaults.")

    def on_close(self):
        state = {
            "geometry": self.geometry(),
            "keyword_tree_col0_width": self.tree.column("#0", "width"),
            "keyword_tree_cols": {c: self.tree.column(c, "width") for c in KEYWORD_COLS},
            "param_tree_cols": {c: self.param_tree.column(c, "width") for c in DEFAULT_PARAM_COL_WIDTHS},
        }
        save_json(UI_STATE_PATH, state)
        save_json(DB_PATH, self.db)
        self._persist_issues("Issue config saved")
        self.destroy()


if __name__ == "__main__":
    KeywordGuideApp().mainloop()
