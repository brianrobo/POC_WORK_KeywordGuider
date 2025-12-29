# ============================================================
# Keyword Guide UI POC
# ------------------------------------------------------------
# Version: 0.4.3 (2025-12-29)
#
# Release Notes (v0.4.3)
# - (UX) Window title에 release version 표시
# - (UX) Copy 시 파라미터 미적용 옵션 추가 (Copy w/o params)
#        - {CH} 같은 placeholder는 빈 값으로 치환하여 복사
#        - 단일 Copy / Copy Selected 모두 동일 옵션 적용
#
# Existing (v0.4.2 유지)
# - Keyword에 "Group" 필드 추가
# - Tree selection bug fix (선택 이벤트에서 tree rebuild 제거)
# - Keyword Description Rich Text 지원 (Bold + 3 Colors) with desc_rich runs 저장/로드
# - Keyword List: 왼쪽(#0) 체크박스 + extended selection
# - Select All / Clear All / Copy Selected(선택된 rendered들을 vendor delimiter로 join)
# - Up/Down(단일 선택 1개 기준) 위치 변경
# - Vendor별 Issue 관리 + Vendor별 Delimiter 저장
# - Keyword CRUD / Copy feedback (Copied bold)
# - Placeholder detect + Inline Apply (category-level params)
# - Category-level params in-place edit (Enter/Esc)
# - UI state persistence
# - KeywordDialog: Parts fixed height + scroll, Preview wrap+scroll
# - Import Joined String split: 확인 후 적용
# - Part 입력 중 delimiter 포함 시 split 제안: 확인 후 적용(자동 아님)
# ============================================================

import json
import re
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# ------------------------------------------------------------
# App Version
# ------------------------------------------------------------
APP_VERSION = "0.4.3"

# ------------------------------------------------------------
# Paths / Constants
# ------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "keywords_db.json"
UI_STATE_PATH = BASE_DIR / "ui_state.json"
ISSUES_PATH = BASE_DIR / "issues_config.json"

PLACEHOLDER_RE = re.compile(r"\{([A-Za-z0-9_]+)\}")
DEFAULT_GEOMETRY = "1250x760"

# Keyword list view: Summary | Group | Info | Copy | Preview
KEYWORD_COLS = ("summary", "group", "info", "copy", "preview")

DEFAULT_KEYWORD_COL_WIDTHS = {
    "summary": 240,
    "group": 140,
    "info": 70,
    "copy": 70,
    "preview": 650,
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
    out = template or ""
    params = params or {}
    for k, v in params.items():
        out = out.replace("{" + str(k) + "}", str(v))
    return out


def strip_placeholders(template: str) -> str:
    """
    Replace all placeholders like {CH}, {ABC} with empty string.
    Example: "FreeAck;CH:{CH};" -> "FreeAck;CH:;"
    """
    return PLACEHOLDER_RE.sub("", template or "")


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
    Backward compatible:
      - string -> {"text": str, "summary":"", "group":"", "desc":""}
      - {"text","desc"} / {"text","description"}
      - {"parts":[...]} -> keep parts (trim-only, allow duplicates)
      - group optional; default ""
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
    """
    Keyword input supports multiple parts via dynamic +/- rows.

    Added:
      - Group field

    Parts 영역:
      - 고정 height + vertical scroll (Canvas)

    Preview:
      - wrap + 고정 lines + vertical scroll

    Split UX:
      - Import Joined String split: 확인 후 적용
      - Part 입력 중 delimiter 포함: 확인 후 split 제안(자동 아님)

    Description Rich:
      - Bold + 3 colors (black/red/blue) via Text tags
      - 저장: desc_rich runs + desc plain 유지
    """

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

        # ---- Description toolbar (rich text) ----
        ttk.Label(frm, text="Description (Bold + Color, Info 팝업으로 표시)").grid(row=11, column=0, sticky="w")
        desc_toolbar = ttk.Frame(frm)
        desc_toolbar.grid(row=12, column=0, sticky="w", pady=(4, 2))

        ttk.Button(desc_toolbar, text="B", width=3, command=self._toggle_bold).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(desc_toolbar, text="Black", command=lambda: self._apply_color("black")).pack(side=tk.LEFT, padx=2)
        ttk.Button(desc_toolbar, text="Red", command=lambda: self._apply_color("red")).pack(side=tk.LEFT, padx=2)
        ttk.Button(desc_toolbar, text="Blue", command=lambda: self._apply_color("blue")).pack(side=tk.LEFT, padx=2)

        self.txt_desc = tk.Text(frm, height=10, wrap="word")
        self.txt_desc.grid(row=13, column=0, sticky="nsew", pady=(2, 10))

        # tags
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

        # Build initial part rows
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

        # Load description (rich preferred)
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

    # ---------------- Parts scrolling ----------------
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

    # ---------------- Parts CRUD ----------------
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

    # ---------------- Preview ----------------
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

    # ---------------- Split UX ----------------
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

    # ---------------- Description Rich ----------------
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

    # ---------------- OK / Cancel ----------------
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
            "desc": desc_plain,          # legacy 유지
            "desc_rich": desc_rich,      # 신규 rich
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
        self.title(f"Chipset Log Keyword Guide v{APP_VERSION}")

        self.db = load_json(DB_PATH) or self._default_db()
        self.ui_state = load_json(UI_STATE_PATH)

        vendors = list(self.db.keys()) if isinstance(self.db, dict) else []
        if not vendors:
            self.db = self._default_db()
            vendors = list(self.db.keys())

        raw_cfg = load_json(ISSUES_PATH)
        self.issue_cfg = ensure_issue_config_vendor_scoped(raw_cfg, vendors)
        self._sync_vendor_scoped_config_with_db()

        self.vendor_var = tk.StringVar()
        self.issue_var = tk.StringVar()
        self.detail_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.delim_var = tk.StringVar(value=DEFAULT_DELIMITER)

        # Copy mode
        self.copy_wo_params_var = tk.BooleanVar(value=False)

        self._param_editor = None
        self._param_editing = None

        self._copy_feedback_after_id = None
        self._copy_feedback_row = None

        self.geometry(self.ui_state.get("geometry", DEFAULT_GEOMETRY))
        self._build_ui()
        self._make_checkbox_images()

        self._apply_saved_widths()
        self._init_vendor()
        self._ensure_current_obj_migrated()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

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

        ttk.Label(left, text="Vendor").pack(anchor="w")
        self.vendor_cb = ttk.Combobox(left, textvariable=self.vendor_var, state="readonly", width=26)
        self.vendor_cb.pack(fill=tk.X)
        self.vendor_cb.bind("<<ComboboxSelected>>", self.on_vendor_change)

        ttk.Label(left, text="Keyword Delimiter (Vendor)").pack(anchor="w", pady=(10, 0))
        delim_row = ttk.Frame(left)
        delim_row.pack(fill=tk.X)
        self.delim_entry = ttk.Entry(delim_row, textvariable=self.delim_var, width=8)
        self.delim_entry.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(delim_row, text="Apply", command=self.apply_vendor_delimiter).pack(side=tk.LEFT)

        ttk.Label(left, text="Issue").pack(anchor="w", pady=(10, 0))
        issue_row = ttk.Frame(left)
        issue_row.pack(fill=tk.X)

        self.issue_cb = ttk.Combobox(issue_row, textvariable=self.issue_var, state="readonly", width=18)
        self.issue_cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.issue_cb.bind("<<ComboboxSelected>>", self.on_issue_change)

        ttk.Button(issue_row, text="+", width=3, command=self.add_issue).pack(side=tk.LEFT, padx=3)
        ttk.Button(issue_row, text="-", width=3, command=self.delete_issue).pack(side=tk.LEFT, padx=3)
        ttk.Button(issue_row, text="R", width=3, command=self.rename_issue).pack(side=tk.LEFT, padx=3)

        ttk.Label(left, text="Detail Category").pack(anchor="w", pady=(10, 0))
        row = ttk.Frame(left)
        row.pack(fill=tk.X)

        self.detail_cb = ttk.Combobox(row, textvariable=self.detail_var, state="readonly", width=18)
        self.detail_cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.detail_cb.bind("<<ComboboxSelected>>", self.refresh_all)

        ttk.Button(row, text="+", width=3, command=self.add_category).pack(side=tk.LEFT, padx=3)
        ttk.Button(row, text="-", width=3, command=self.delete_category).pack(side=tk.LEFT, padx=3)
        ttk.Button(row, text="R", width=3, command=self.rename_category).pack(side=tk.LEFT, padx=3)

        right = ttk.Frame(root)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        ttk.Label(right, text="Keywords: [ ] / Summary / Group / Info / Copy / Preview(Rendered)").grid(
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
            "preview": "Preview",
        }
        for c in KEYWORD_COLS:
            self.tree.heading(c, text=headings.get(c, c))

        self.tree.column("summary", width=DEFAULT_KEYWORD_COL_WIDTHS["summary"], anchor="w", stretch=False)
        self.tree.column("group", width=DEFAULT_KEYWORD_COL_WIDTHS["group"], anchor="w", stretch=False)
        self.tree.column("info", width=DEFAULT_KEYWORD_COL_WIDTHS["info"], anchor="center", stretch=False)
        self.tree.column("copy", width=DEFAULT_KEYWORD_COL_WIDTHS["copy"], anchor="center", stretch=False)
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

        # NEW: copy w/o params toggle
        ttk.Checkbutton(btns, text="Copy w/o params", variable=self.copy_wo_params_var).pack(
            side=tk.LEFT, padx=(10, 4)
        )

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
    # Navigation
    # --------------------------------------------------------
    def _init_vendor(self):
        vendors = list(self.db.keys())
        self.vendor_cb["values"] = vendors
        self.vendor_var.set(vendors[0])
        self.on_vendor_change()

    def on_vendor_change(self, *_):
        v = self.vendor_var.get()
        if not v:
            return

        self.issue_cfg = ensure_issue_config_vendor_scoped(self.issue_cfg, list(self.db.keys()))
        if v not in self.issue_cfg.get("vendors", {}):
            self.issue_cfg["vendors"][v] = {"issues": default_issues(), "delimiter": DEFAULT_DELIMITER}
            self._persist_issues("Issue config normalized")

        self.delim_var.set(self._get_vendor_delimiter(v))

        issues = self._get_vendor_issues(v)
        if not issues:
            issues = default_issues()
            self._set_vendor_issues(v, issues)
            self._persist_issues("Issue config defaulted")

        self.db.setdefault(v, {})
        for i in issues:
            self.db[v].setdefault(i, self._default_issue_obj())

        self.issue_cb["values"] = issues
        self.issue_var.set(issues[0] if issues else "")
        self.on_issue_change()

    def on_issue_change(self, *_):
        v, i = self.vendor_var.get(), self.issue_var.get()
        if not v or not i:
            return

        self.db.setdefault(v, {})
        self.db[v].setdefault(i, self._default_issue_obj())

        details = list(self.db[v][i].keys())
        self.detail_cb["values"] = details
        self.detail_var.set("_COMMON" if "_COMMON" in details else (details[0] if details else ""))
        self.refresh_all()

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

        self._set_vendor_delimiter(v, str(delim))
        self._persist_issues(f"Delimiter updated for {v}: '{delim}'")
        self.refresh_keywords()
        self.on_keyword_select()

    # --------------------------------------------------------
    # Data helpers
    # --------------------------------------------------------
    def _current_obj(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()

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

            # Preview는 항상 rendered (기존 유지)
            preview = render_keyword(raw_joined, params)

            iid = str(idx)
            self.tree.insert("", "end", iid=iid, text="", values=(summary, group, "Info", "Copy", preview))
            self._set_checkbox_for_iid(iid, checked=False)

        self._sync_checkboxes_with_selection()

        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        self.status_var.set(f"Selected: {v} > {i} > {d}")

    def refresh_keyword_previews_only(self):
        """Treeview 전체 재생성 없이 preview 컬럼만 업데이트"""
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
        # IMPORTANT: do NOT rebuild tree here. Only update inline panel and preview cells.
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

        # update only params panel + previews (no tree rebuild)
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
    def _render_for_copy(self, raw_joined: str, params: dict) -> str:
        """
        Copy output generator.
        - If Copy w/o params: strip placeholders -> {CH} becomes empty
        - Else: render using current params
        """
        if self.copy_wo_params_var.get():
            return strip_placeholders(raw_joined)
        return render_keyword(raw_joined, params)

    def copy_selected_keywords(self, _event=None):
        sel = list(self.tree.selection())
        if not sel:
            self.status_var.set("No selection to copy.")
            return

        try:
            sel_sorted = sorted(sel, key=lambda x: int(x))
        except Exception:
            sel_sorted = sel

        obj = self._current_obj()
        params = obj["_params"]
        v = self.vendor_var.get()
        delim = self._get_vendor_delimiter(v)

        rendered_list = []
        for iid in sel_sorted:
            try:
                idx = int(iid)
            except Exception:
                continue
            if idx < 0 or idx >= len(obj["_keywords"]):
                continue
            kw = obj["_keywords"][idx]
            raw_joined = keyword_joined_template(kw, delim)

            rendered = self._render_for_copy(raw_joined, params).strip()
            if rendered:
                rendered_list.append(rendered)

        if not rendered_list:
            self.status_var.set("No valid keywords to copy.")
            return

        combined = delim.join(rendered_list)
        self.clipboard_clear()
        self.clipboard_append(combined)

        mode = "NO-PARAMS" if self.copy_wo_params_var.get() else "RENDERED"
        self.status_var.set(f"Copied Selected ({len(rendered_list)}) [{mode}]: {combined}")

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
            self.refresh_keywords()

    def remove_keyword(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return
        obj = self._current_obj()
        del obj["_keywords"][idx]
        self._persist_db("Keyword removed")
        self.refresh_keywords()
        self.clear_inline()

    # --------------------------------------------------------
    # Up / Down (single selection only)
    # --------------------------------------------------------
    def move_keyword_up(self):
        sel = list(self.tree.selection())
        if not sel:
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

        new_iid = str(idx - 1)
        try:
            self.tree.selection_set(new_iid)
            self.tree.see(new_iid)
        except Exception:
            pass

    def move_keyword_down(self):
        sel = list(self.tree.selection())
        if not sel:
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

        new_iid = str(idx + 1)
        try:
            self.tree.selection_set(new_iid)
            self.tree.see(new_iid)
        except Exception:
            pass

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

        # columns: #0 checkbox, #1 summary, #2 group, #3 info, #4 copy, #5 preview
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
        elif col == "#4":  # Copy
            rendered = self._render_for_copy(raw_joined, params)
            self.clipboard_clear()
            self.clipboard_append(rendered)
            self._show_copy_feedback(row)

            mode = "NO-PARAMS" if self.copy_wo_params_var.get() else "RENDERED"
            self.status_var.set(f"Copied [{mode}]: {rendered}")

    def on_tree_double_click(self, event):
        col = self.tree.identify_column(event.x)
        if col in ("#0", "#3", "#4"):
            return
        self.edit_keyword()

    # --------------------------------------------------------
    # Copy feedback UI
    # --------------------------------------------------------
    def _show_copy_feedback(self, row_iid: str):
        self._clear_copy_feedback(force=True)

        try:
            vals = list(self.tree.item(row_iid, "values"))
            if len(vals) == len(KEYWORD_COLS):
                # copy column index = 3 in KEYWORD_COLS ("summary","group","info","copy","preview")
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

        new_val = self._param_editor.get()
        obj = self._current_obj()
        obj["_params"][self._param_editing] = new_val

        self._persist_db(f"Param updated: {self._param_editing}={new_val}")
        self._cancel_param_edit()
        self.refresh_params()
        self.refresh_keyword_previews_only()
        self.param_tree.selection_set(self._param_editing)

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
        self.detail_cb["values"] = list(self.db[v][i].keys())
        self.detail_var.set(name)
        self.refresh_all()

    def delete_category(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
        if d == "_COMMON":
            messagebox.showinfo("Info", "_COMMON cannot be deleted.")
            return

        if not messagebox.askyesno("Confirm", f"Delete category '{d}'?"):
            return

        del self.db[v][i][d]
        self._persist_db("Category deleted")

        self.detail_cb["values"] = list(self.db[v][i].keys())
        details = list(self.db[v][i].keys())
        self.detail_var.set("_COMMON" if "_COMMON" in details else (details[0] if details else "_COMMON"))
        self.refresh_all()

    def rename_category(self):
        v, i, d = self.vendor_var.get(), self.issue_var.get(), self.detail_var.get()
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

        self.detail_cb["values"] = list(self.db[v][i].keys())
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

        self.issue_cb["values"] = issues
        self.issue_var.set(name)
        self.on_issue_change()

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

        self.issue_cb["values"] = issues
        self.issue_var.set(issues[0])
        self.on_issue_change()

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

        self.issue_cb["values"] = issues
        self.issue_var.set(new)
        self.on_issue_change()

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
