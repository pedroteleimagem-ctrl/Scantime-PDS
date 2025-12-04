import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
import os
import re
import pickle
import unicodedata
import calendar
import locale
from datetime import date, timedelta
import Assignation
from Assignation import assigner_initiales
from ConstraintsV2 import ConstraintsTable
from Export import (
    export_to_excel as export_to_excel_external,
    export_combined_to_excel as export_combined_to_excel_external
)
from eligibility import (
    AssignmentSettings,
    PlanningContext,
    candidate_is_available,
    parse_constraint_row,
)




APP_TITLE = "ScanTime PDS"
current_status_path = None
root = None  # Main Tk window, set in __main__
title_label = None  # Banner label shown at the top of the UI
UNDO_CONSTRAINT_TAG = "__constraint_widget__"


def renumber_week_tabs():
    """Ensure notebook tab labels stay contiguous: Mois 1, 2, ..."""
    nb = globals().get("notebook")
    if nb is None:
        return
    try:
        tab_ids = nb.tabs()
    except Exception:
        return
    for idx, tab_id in enumerate(tab_ids, start=1):
        try:
            nb.tab(tab_id, text=f"Mois {idx}")
        except Exception:
            pass


def _build_window_caption() -> str:
    """
    Returns the text shown in the native window title and the in-app banner.
    Mirrors the Microsoft Word convention: '<file> - <app>'.
    """
    if current_status_path:
        filename = os.path.basename(current_status_path)
        name_without_ext, _ = os.path.splitext(filename)
        display_name = name_without_ext or filename
        return f"{display_name} - {APP_TITLE}"
    return APP_TITLE


def update_window_caption():
    """
    Pushes the current file name to the OS window title + banner label.
    Safe to call before the widgets are created.
    """
    title_text = _build_window_caption()
    if root is not None:
        try:
            root.title(title_text)
        except Exception:
            pass
        renumber_week_tabs()
    if title_label is not None:
        try:
            title_label.config(text=title_text)
        except Exception:
            pass


# --- Couleurs et styles modernes (theme Word-like) ---
APP_FONT_FAMILY = "Segoe UI"
APP_WINDOW_BG = "#F3F3F3"
APP_SURFACE_BG = "#FFFFFF"
APP_PRIMARY_COLOR = "#217346"   # vert Excel
APP_PRIMARY_DARK = "#1A5A37"
APP_PRIMARY_LIGHT = "#CDE8D6"
APP_DIVIDER = "#D6E4DA"
APP_RIBBON_BG = "#E8F1E8"
CELL_EMPTY_BG = "#FFFFFF"
CELL_FILLED_BG = "#E6EDF8"
CELL_DISABLED_BG = "#DDE2EC"
CELL_DISABLED_TEXT = "#6B778D"
DAY_LABEL_BG = "#0B5E27"
WEEKEND_DAY_BG = "#FFE9D9"
WEEKEND_CELL_BG = "#FFF6F2"
EXCLUDED_CELL_BORDER = "#E0A100"
SHIFT_EVEN_ROW_BG = "#FFFFFF"
SHIFT_ODD_ROW_BG = "#F4F6FA"
RIBBON_FRAME_STYLE = "Ribbon.TFrame"
RIBBON_BUTTON_STYLE = "Ribbon.TButton"
RIBBON_ACCENT_BUTTON_STYLE = "RibbonPrimary.TButton"
RIBBON_CHECK_STYLE = "Ribbon.TCheckbutton"
NOTEBOOK_STYLE = "Office.TNotebook"
NOTEBOOK_TAB_STYLE = "Office.TNotebook.Tab"
SHIFT_TREE_STYLE = "ShiftCount.Treeview"
SHIFT_TREE_HEADING_STYLE = "ShiftCount.Treeview.Heading"
MONTH_NAMES_FR = [
    "Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre",
]
HOLIDAY_DAY_BG = "#FFD65C"
HOLIDAY_CELL_BG = "#FFF3D6"

try:
    import holidays as _holidays_lib
except Exception:
    _holidays_lib = None


def _default_holiday_country():
    """Essaie de déduire un code pays (ISO) via la locale système."""
    try:
        loc = locale.getdefaultlocale()[0] or ""
    except Exception:
        loc = ""
    if loc and "_" in loc:
        return loc.split("_", 1)[1].upper()
    return (loc or "FR").upper() or "FR"


HOLIDAY_COUNTRY = _default_holiday_country()


def month_holidays(year: int, month: int) -> set[date]:
    """
    Retourne l'ensemble des dates fériées pour un mois donné.
    Utilise la bibliothèque 'holidays' si disponible, sinon renvoie un set vide.
    """
    if _holidays_lib is None:
        return set()
    try:
        hol = _holidays_lib.CountryHoliday(HOLIDAY_COUNTRY, years=[year])
    except Exception:
        return set()
    try:
        return {d for d in hol if d.year == year and d.month == month}
    except Exception:
        return set()

def setup_modern_styles(root: tk.Tk) -> ttk.Style:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(bg=APP_WINDOW_BG)

    def _font_option(size: int, weight: str | None = None) -> str:
        family = APP_FONT_FAMILY
        if " " in family:
            family = f"{{{family}}}"
        parts = [family, str(size)]
        if weight:
            parts.append(weight)
        return " ".join(parts)

    root.option_add("*Font", _font_option(10))
    root.option_add("*TButton.Font", _font_option(10))
    root.option_add("*TLabel.Font", _font_option(10))
    root.option_add("*TEntry.Font", _font_option(10))
    root.option_add("*TCombobox*Listbox.Font", _font_option(10))

    style.configure("TFrame", background=APP_WINDOW_BG)
    style.configure("TLabel", background=APP_WINDOW_BG)

    style.configure(RIBBON_FRAME_STYLE, background=APP_RIBBON_BG, borderwidth=0)

    style.configure(
        RIBBON_BUTTON_STYLE,
        background="#FFFFFF",
        foreground="#1F2D3D",
        padding=(12, 6),
        borderwidth=1,
        relief="flat",
        focuscolor=APP_PRIMARY_LIGHT
    )
    style.map(
        RIBBON_BUTTON_STYLE,
        background=[("pressed", APP_PRIMARY_LIGHT), ("active", APP_PRIMARY_LIGHT)],
        foreground=[("disabled", CELL_DISABLED_TEXT)]
    )

    style.configure(
        RIBBON_ACCENT_BUTTON_STYLE,
        background=APP_PRIMARY_COLOR,
        foreground="#FFFFFF",
        padding=(14, 8),
        borderwidth=0,
        relief="flat"
    )
    style.map(
        RIBBON_ACCENT_BUTTON_STYLE,
        background=[("pressed", APP_PRIMARY_DARK), ("active", APP_PRIMARY_DARK)],
        foreground=[("disabled", "#E0E4EA")]
    )

    style.configure(
        RIBBON_CHECK_STYLE,
        background=APP_RIBBON_BG,
        foreground="#1F2D3D",
        focuscolor=APP_PRIMARY_LIGHT,
        padding=(6, 6, 10, 6)
    )
    style.map(
        RIBBON_CHECK_STYLE,
        background=[("active", APP_PRIMARY_LIGHT)],
        foreground=[("disabled", CELL_DISABLED_TEXT)]
    )

    style.configure(
        NOTEBOOK_STYLE,
        background=APP_WINDOW_BG,
        borderwidth=0,
        tabmargins=(18, 10, 18, 0)
    )
    style.configure(
        NOTEBOOK_TAB_STYLE,
        background=APP_WINDOW_BG,
        foreground="#1F2D3D",
        padding=(20, 10),
        font=(APP_FONT_FAMILY, 10, "normal")
    )
    style.map(
        NOTEBOOK_TAB_STYLE,
        background=[("selected", APP_PRIMARY_COLOR)],
        foreground=[("selected", "#FFFFFF")]
    )

    style.configure(
        SHIFT_TREE_STYLE,
        background=APP_SURFACE_BG,
        fieldbackground=APP_SURFACE_BG,
        borderwidth=0,
        rowheight=24,
        font=(APP_FONT_FAMILY, 10)
    )
    style.map(
        SHIFT_TREE_STYLE,
        background=[("selected", APP_PRIMARY_LIGHT)],
        foreground=[("disabled", CELL_DISABLED_TEXT)]
    )
    style.configure(
        SHIFT_TREE_HEADING_STYLE,
        background=APP_PRIMARY_COLOR,
        foreground="#FFFFFF",
        borderwidth=0,
        font=(APP_FONT_FAMILY, 10, "bold"),
        padding=(8, 6)
    )
    style.map(
        SHIFT_TREE_HEADING_STYLE,
        background=[("active", APP_PRIMARY_DARK)]
    )

    style.configure("TScrollbar", background=APP_WINDOW_BG)
    return style



_MULTI_NAME_SPLIT_RE = re.compile(r"[\n,;/&+]+")

def _normalize_initial_label(value: str) -> str:
    return " ".join(str(value or "").strip().split()).upper()

def extract_names_from_cell(raw_text: str, valid_names=None):
    return _extract_names_from_cell_impl(raw_text, valid_names)

def _extract_names_from_cell_impl(raw_text, valid_names=None):
    "Return the list of initials detected in a planning cell."
    text = str(raw_text or "").strip()
    if not text or text.lower() == "x":
        return []
    normalized = re.sub(r"\s{2,}", "\n", text)
    parts = [p.strip() for p in _MULTI_NAME_SPLIT_RE.split(normalized) if p.strip()]
    if not parts:
        parts = [text]
    if not valid_names:
        return list(dict.fromkeys(parts))
    norm_map = {_normalize_initial_label(name): name for name in valid_names}
    seen = set()
    result = []
    for part in parts:
        key = _normalize_initial_label(part)
        if key in norm_map and key not in seen:
            result.append(norm_map[key])
            seen.add(key)
    if result:
        return result
    upper_text = " ".join(text.upper().split())
    for key, original in norm_map.items():
        pattern = r"(?<!\S)" + re.escape(key) + r"(?!\S)"
        if re.search(pattern, upper_text) and key not in seen:
            result.append(original)
            seen.add(key)
    return result


# ---------- Routeur global pour la molette (singleton) ----------
class _MouseWheelManager:
    def __init__(self, root):
        self.root = root
        self.active_canvas = None
        self._registered_canvases = set()
        # Bind global une seule fois
        if not getattr(root, "_mw_bound", False):
            root.bind_all("<MouseWheel>",         self._on_wheel, add="+")
            root.bind_all("<Shift-MouseWheel>",   self._on_wheel, add="+")
            root.bind_all("<Button-4>",           self._on_wheel, add="+")  # Linux
            root.bind_all("<Button-5>",           self._on_wheel, add="+")  # Linux
            root._mw_bound = True

    def register(self, canvas, inner=None):
        self._registered_canvases.add(canvas)
        # Active ce canvas quand la souris entre ; dÃ©sactive quand elle sort
        def _activate(_e, c=canvas):   self.active_canvas = c
        def _deactivate(_e, c=canvas):
            if self.active_canvas is c:
                self.active_canvas = None
        canvas.bind("<Enter>", _activate)
        canvas.bind("<Leave>", _deactivate)
        if inner is not None:
            inner.bind("<Enter>", _activate)
            inner.bind("<Leave>", _deactivate)

    def _resolve_canvas_under_pointer(self):
        try:
            x, y = self.root.winfo_pointerxy()
            widget = self.root.winfo_containing(x, y)
        except (tk.TclError, KeyError):
            return None
        while widget is not None and widget is not self.root:
            if widget in self._registered_canvases:
                return widget
            widget = getattr(widget, "master", None)
        return None

    def _on_wheel(self, event):
        c = self.active_canvas
        if not c or not c.winfo_exists():
            c = self._resolve_canvas_under_pointer()
            if not c:
                return
            self.active_canvas = c
        # Sens du défilement
        if hasattr(event, "delta") and event.delta:
            step = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            step = -1
        elif getattr(event, "num", None) == 5:
            step = 1
        else:
            return
        # Shift = horizontal
        if event.state & 0x0001:
            c.xview_scroll(step, "units")
        else:
            c.yview_scroll(step, "units")
        return "break"


def get_mousewheel_manager(widget):
    """RÃ©cupÃ¨re (ou crÃ©e) le gestionnaire global de molette."""
    root = widget.winfo_toplevel()
    mgr = getattr(root, "_mw_manager", None)
    if mgr is None:
        mgr = _MouseWheelManager(root)
        root._mw_manager = mgr
    return mgr
# ---------------------------------------------------------------


# --- Ãvite le double chargement du module ----------------------------------
import sys
if __name__ == '__main__':
    sys.modules['Full_GUI'] = sys.modules[__name__]
# ---------------------------------------------------------------------------


# Definition des jours du mois (31 jours max pour couvrir tous les mois)
days = [str(i) for i in range(1, 32)]

# Dictionnaire des postes (astreintes) avec couleur par dÃ©faut
POST_INFO = {
    "Ligne 1": {"color": "#C6E0B4", "shifts": {}},
    "Ligne 2": {"color": "#F9E199", "shifts": {}},
    "Ligne 3": {"color": "#C5A0DD", "shifts": {}},
    "Ligne 4": {"color": "#F4B183", "shifts": {}},
}

# Liste globale des postes, utilisÃ©e Ã©galement par Constraints.py
work_posts = list(POST_INFO.keys())


def update_work_posts(new_posts):
    """
    Met Ã  jour la liste globale des postes ET notifie le reste de lâapplication
    quâun changement a eu lieu via lâÃ©vÃ©nement virtuel <<WorkPostsUpdated>>.
    """
    global work_posts
    work_posts = new_posts

    try:
        if Assignation.FORBIDDEN_POST_ASSOCIATIONS:
            valid = set(new_posts)
            filtered = {(m, a) for (m, a) in Assignation.FORBIDDEN_POST_ASSOCIATIONS if m in valid and a in valid}
            if len(filtered) != len(Assignation.FORBIDDEN_POST_ASSOCIATIONS):
                Assignation.FORBIDDEN_POST_ASSOCIATIONS.clear()
                Assignation.FORBIDDEN_POST_ASSOCIATIONS.update(filtered)
                for gui_instance in get_all_gui_instances():
                    try:
                        gui_instance.schedule_update_colors()
                    except Exception:
                        pass
    except Exception:
        pass

    # â Diffusion d'un Ã©vÃ©nement global pour avertir toutes les fenÃªtres â
    try:
        import tkinter as tk
        root = tk._default_root  # â ne crÃ©e PLUS de Tk() de secours
        if root:
            root.event_generate("<<WorkPostsUpdated>>", when="tail")
    except Exception:
        # Si l'Ã©vÃ©nement ne peut pas Ãªtre Ã©mis, on ignore silencieusement.
        pass


def get_work_posts():
    return work_posts


#####################################################################
# Gestion des fenêtres multiples et des conflits inter-fenêtres (live)
#####################################################################

_WINDOW_CONTEXTS: list = []
_LIVE_CONFLICT_JOB = None
_LIVE_CONFLICT_DEBOUNCE_MS = 220
_LIVE_CONFLICT_PAUSED = False


def register_window_context(ctx: dict):
    """Enregistre une fenêtre (root, notebook, tabs_data, is_primary)."""
    if ctx not in _WINDOW_CONTEXTS:
        _WINDOW_CONTEXTS.append(ctx)
    trigger_live_conflict_check()


def unregister_window_context(ctx: dict):
    """Supprime une fenêtre du registre et relance un nettoyage/refresh des badges."""
    try:
        _WINDOW_CONTEXTS.remove(ctx)
    except ValueError:
        pass
    trigger_live_conflict_check()


def _get_primary_context():
    for ctx in _WINDOW_CONTEXTS:
        if ctx.get("is_primary"):
            return ctx
    return _WINDOW_CONTEXTS[0] if _WINDOW_CONTEXTS else None


def _split_people_live(txt: str) -> list:
    try:
        import re as _re
        return [p.strip() for p in _MULTI_NAME_SPLIT_RE.split(txt or "") if p.strip()]
    except Exception:
        return [txt.strip()] if txt else []


def _norm_name_live(txt: str) -> str:
    try:
        return " ".join((txt or "").upper().strip().split())
    except Exception:
        return txt or ""


def _clear_conflict_marks_for_context(ctx: dict):
    for item in ctx.get("tabs_data", []):
        try:
            g_obj = item[0]
        except Exception:
            g_obj = None
        if g_obj is None:
            continue
        try:
            g_obj.clear_cross_conflict_marks()
        except Exception:
            pass


def _build_index_for_gui(g_obj):
    """
    Construit un index {norm_name: [(row, col, half, post_idx, post_name)]}
    pour un planning donn? (mode mensuel : row = jour, col = astreinte).
    """
    local_posts = getattr(g_obj, "local_work_posts", work_posts)
    index = {}
    try:
        entries = g_obj.table_entries
    except Exception:
        return index

    for r, row in enumerate(entries):
        for c, cell in enumerate(row):
            try:
                raw = cell.get().strip()
            except Exception:
                raw = ""
            if not raw:
                continue
            post_name = local_posts[c] if c < len(local_posts) else ""
            for person in _split_people_live(raw):
                norm = _norm_name_live(person)
                if not norm:
                    continue
                index.setdefault(norm, []).append((r, c, "ASTREINTE", c, post_name))
    return index

def _run_live_conflict_check():
    """Compare le planning principal avec le premier secondaire et marque les conflits en badges."""
    if _LIVE_CONFLICT_PAUSED:
        return
    global _LIVE_CONFLICT_JOB
    _LIVE_CONFLICT_JOB = None

    # Option globale d'activation (créée plus tard dans __main__)
    live_var = globals().get("live_conflict_var")
    if live_var is not None and hasattr(live_var, "get") and not live_var.get():
        # Désactivé : on nettoie et on sort
        for ctx in list(_WINDOW_CONTEXTS):
            _clear_conflict_marks_for_context(ctx)
        return

    primary_ctx = _get_primary_context()
    secondary_ctx = None
    for ctx in _WINDOW_CONTEXTS:
        if ctx is primary_ctx:
            continue
        secondary_ctx = ctx
        break

    # Rien à comparer
    if primary_ctx is None or secondary_ctx is None:
        if primary_ctx:
            _clear_conflict_marks_for_context(primary_ctx)
        return

    # Fenêtres détruites ?
    for ctx in (primary_ctx, secondary_ctx):
        root_obj = ctx.get("root")
        try:
            if root_obj is None or not root_obj.winfo_exists():
                unregister_window_context(ctx)
                return
        except Exception:
            unregister_window_context(ctx)
            return

    # Nettoyage des badges existants avant recalcul
    _clear_conflict_marks_for_context(primary_ctx)
    _clear_conflict_marks_for_context(secondary_ctx)

    primary_tabs = primary_ctx.get("tabs_data", [])
    secondary_tabs = secondary_ctx.get("tabs_data", [])
    if not primary_tabs or not secondary_tabs:
        return

    tab_pairs = zip(primary_tabs, secondary_tabs)
    for pair in tab_pairs:
        try:
            g1, _c1, _s1 = pair[0]
            g2, _c2, _s2 = pair[1]
        except Exception:
            continue
        if g1 is None or g2 is None:
            continue
        idx1 = _build_index_for_gui(g1)
        idx2 = _build_index_for_gui(g2)
        common = set(idx1.keys()) & set(idx2.keys())
        for norm in common:
            occ1 = idx1.get(norm, [])
            occ2 = idx2.get(norm, [])
            for r1, c1, half1, post_idx1, post_name1 in occ1:
                for r2, c2, half2, post_idx2, post_name2 in occ2:
                    if half1 != half2 or c1 != c2:
                        continue
                    # Conflit (poste ignor? car structures diff?rentes) : même personne, même jour/half, même poste
                    try:
                        g1.mark_cross_conflict(r1, c1)
                    except Exception:
                        pass
                    try:
                        g2.mark_cross_conflict(r2, c2)
                    except Exception:
                        pass


def trigger_live_conflict_check():
    """Planifie (débounce) un recalcul live des conflits inter-fenêtres."""
    if _LIVE_CONFLICT_PAUSED:
        return
    global _LIVE_CONFLICT_JOB
    primary_ctx = _get_primary_context()
    primary_root = primary_ctx.get("root") if primary_ctx else None
    if primary_root is None:
        return
    try:
        if _LIVE_CONFLICT_JOB is not None:
            primary_root.after_cancel(_LIVE_CONFLICT_JOB)
    except Exception:
        pass
    _LIVE_CONFLICT_JOB = primary_root.after(_LIVE_CONFLICT_DEBOUNCE_MS, _run_live_conflict_check)


def pause_live_conflict_check():
    """Suspend le job live de conflits (et annule le job en attente)."""
    global _LIVE_CONFLICT_PAUSED, _LIVE_CONFLICT_JOB
    _LIVE_CONFLICT_PAUSED = True
    try:
        primary_ctx = _get_primary_context()
        primary_root = primary_ctx.get("root") if primary_ctx else None
        if primary_root is not None and _LIVE_CONFLICT_JOB is not None:
            primary_root.after_cancel(_LIVE_CONFLICT_JOB)
    except Exception:
        pass
    _LIVE_CONFLICT_JOB = None


def resume_live_conflict_check():
    """Relance le calcul live des conflits après une pause."""
    global _LIVE_CONFLICT_PAUSED
    _LIVE_CONFLICT_PAUSED = False
    trigger_live_conflict_check()
def get_all_gui_instances():
    data = globals().get("tabs_data")
    if not data:
        return []
    instances = []
    for item in data:
        try:
            gui_instance = item[0]
        except (TypeError, IndexError):
            continue
        if gui_instance is not None:
            instances.append(gui_instance)
    return instances

# Fonction utilitaire pour une saisie personnalisÃ©e
def custom_askstring(parent, title, prompt, x, y, initial=""):
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.geometry("+{}+{}".format(x, y))
    tk.Label(dialog, text=prompt).pack(padx=10, pady=10)
    entry = tk.Entry(dialog)
    entry.insert(0, initial)
    entry.pack(padx=10, pady=10)
    entry.focus_set()
    result = []
    def on_ok():
        result.append(entry.get())
        dialog.destroy()
    ok_btn = tk.Button(dialog, text="OK", command=on_ok)
    ok_btn.pack(pady=10)
    dialog.wait_window()
    return result[0] if result else None


class MonthPickerDialog(tk.Toplevel):
    """Petit sélecteur de mois/année avec grille de calendrier."""
    def __init__(self, parent, initial_date=None, anchor_widget=None):
        super().__init__(parent)
        self.title("Choisir un mois")
        self.configure(bg=APP_WINDOW_BG)
        self.resizable(False, False)
        self.result = None
        self._anchor_widget = anchor_widget or parent

        base_date = initial_date or date.today()
        self.month_var = tk.IntVar(value=base_date.month)
        self.year_var = tk.IntVar(value=base_date.year)
        self._title_var = tk.StringVar()

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        nav = ttk.Frame(self, padding=10, style=RIBBON_FRAME_STYLE)
        nav.pack(fill="x")
        ttk.Button(nav, text="<", width=4, command=lambda: self._shift_month(-1), style=RIBBON_BUTTON_STYLE).pack(side="left")
        ttk.Label(nav, textvariable=self._title_var, anchor="center").pack(side="left", expand=True, fill="x", padx=6)
        ttk.Button(nav, text=">", width=4, command=lambda: self._shift_month(1), style=RIBBON_BUTTON_STYLE).pack(side="right")

        self.calendar_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        self.calendar_frame.pack(fill="both", expand=True)

        btns = ttk.Frame(self, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="Valider", command=self._confirm, style=RIBBON_BUTTON_STYLE).pack(side="right", padx=(0, 6))
        ttk.Button(btns, text="Annuler", command=self._on_cancel, style=RIBBON_BUTTON_STYLE).pack(side="right")

        self._build_calendar()
        self._center_over_widget(self._anchor_widget)
        self.wait_window(self)

    def _center_over_widget(self, widget):
        """Centre le popup sur le toplevel du widget (prend en compte multi-écrans)."""
        try:
            self.update_idletasks()
            target = (widget or self.master)
            if target is not None:
                try:
                    target = target.winfo_toplevel()
                except Exception:
                    pass
            if target is None:
                return
            target.update_idletasks()

            popup_w, popup_h = self.winfo_width(), self.winfo_height()
            if popup_w <= 0 or popup_h <= 0:
                return

            try:
                wx, wy = target.winfo_rootx(), target.winfo_rooty()
                ww, wh = target.winfo_width(), target.winfo_height()
            except Exception:
                wx = wy = 0
                ww, wh = self.winfo_screenwidth(), self.winfo_screenheight()

            if ww <= 0 or wh <= 0:
                ww, wh = self.winfo_screenwidth(), self.winfo_screenheight()
                wx = wy = 0

            x = wx + (ww - popup_w) // 2
            y = wy + (wh - popup_h) // 2
            self.geometry(f"+{int(x)}+{int(y)}")
        except Exception:
            pass

    def _update_title(self):
        try:
            month_label = MONTH_NAMES_FR[self.month_var.get() - 1]
        except Exception:
            month_label = f"Mois {self.month_var.get()}"
        self._title_var.set(f"{month_label} {self.year_var.get()}")

    def _shift_month(self, delta):
        month = self.month_var.get() + delta
        year = self.year_var.get()
        if month < 1:
            month = 12
            year -= 1
        elif month > 12:
            month = 1
            year += 1
        self.month_var.set(month)
        self.year_var.set(year)
        self._build_calendar()

    def _build_calendar(self):
        for child in self.calendar_frame.winfo_children():
            child.destroy()

        self._update_title()
        day_names = ["L", "M", "M", "J", "V", "S", "D"]
        for idx, name in enumerate(day_names):
            ttk.Label(self.calendar_frame, text=name, width=3, anchor="center").grid(row=0, column=idx, padx=2, pady=2)

        cal = calendar.Calendar(firstweekday=0)
        for row_idx, week in enumerate(cal.monthdayscalendar(self.year_var.get(), self.month_var.get()), start=1):
            for col_idx, day in enumerate(week):
                if day == 0:
                    ttk.Label(self.calendar_frame, text="", width=3).grid(row=row_idx, column=col_idx, padx=2, pady=2)
                else:
                    btn = ttk.Button(self.calendar_frame, text=str(day), width=3, command=lambda d=day: self._select_day(d), style=RIBBON_BUTTON_STYLE)
                    btn.grid(row=row_idx, column=col_idx, padx=2, pady=2)

    def _select_day(self, day):
        try:
            self.result = date(self.year_var.get(), self.month_var.get(), day)
        except Exception:
            self.result = None
        self.destroy()

    def _confirm(self):
        """Valide le mois affiché sans choisir un jour précis (prend le 1er)."""
        try:
            self.result = date(self.year_var.get(), self.month_var.get(), 1)
        except Exception:
            self.result = None
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()

# Zone scrollable pour le tableau
# Zone scrollable pour le tableau
class ScrollableFrame(tk.Frame):
    """
    Frame scrollable (X & Y) :
      â¢ Molette routÃ©e via un gestionnaire global (aucune duplication).
      â¢ DÃ©file uniquement quand la souris survole ce tableau.
      â¢ Compatible Windows/macOS (<MouseWheel>) et Linux (<Button-4/5>).
      â¢ Shift + molette = dÃ©filement horizontal.
    """
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        # ---- Canvas + scrollbars -------------------------------------------
        self.canvas   = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.v_scroll = tk.Scrollbar(self, orient="vertical",
                                     command=self.canvas.yview)
        self.h_scroll = tk.Scrollbar(self, orient="horizontal",
                                     command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set,
                              xscrollcommand=self.h_scroll.set)

        # Placement
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.canvas.grid(   row=0, column=0, sticky="nsew")
        self.v_scroll.grid( row=0, column=1, sticky="ns")
        self.h_scroll.grid( row=1, column=0, sticky="ew")

        # ---- Cadre intÃ©rieur (contenu rÃ©el) --------------------------------
        self.inner = tk.Frame(self.canvas)
        self._winid = self.canvas.create_window((0, 0),
                                                window=self.inner,
                                                anchor="nw")

        # Ajuste la scrollregion quand le contenu change
        self.inner.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Option : adapter la largeur du contenu Ã  la largeur visible du canvas
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        # Panoramique global : ALT + clic gauche fonctionne mÃªme sur les cellules
        root = self.canvas.winfo_toplevel()
        root.bind_all("<Alt-ButtonPress-1>",   self._pan_start, add="+")
        root.bind_all("<Alt-B1-Motion>",       self._pan_move,  add="+")
        root.bind_all("<Alt-ButtonRelease-1>", self._pan_end,   add="+")
        self._panning = False
        self._pan_canvas = None
        self._prev_cursor = ""

        root.bind_all("<ButtonPress-2>",   self._pan_start, add="+")
        root.bind_all("<B2-Motion>",       self._pan_move,  add="+")
        root.bind_all("<ButtonRelease-2>", self._pan_end,   add="+")




        # ---- Enregistre ce canvas auprÃ¨s du routeur global -----------------
        get_mousewheel_manager(self).register(self.canvas, self.inner)

    def on_canvas_configure(self, event):
        """
        Ne pas forcer la largeur du frame intÃ©rieur Ã  celle du canvas.
        Laisser la largeur naturelle pour conserver le dÃ©filement horizontal.
        (On pourrait Ã©ventuellement Ã©tirer si le canvas est plus large,
        mais surtout ne jamais rÃ©duire Ã  event.width.)
        """
        try:
            # On NE FORCE PAS la largeur quand le contenu est plus large que le canvas.
            # Si tu veux que le contenu s'Ã©tire quand le canvas est plus large,
            # dÃ©-commente la ligne suivante (optionnel) :
            # if event.width > self.inner.winfo_reqwidth():
            #     self.canvas.itemconfigure(self._winid, width=event.width)
            pass
        except Exception:
            pass

    def _canvas_under_pointer(self, event):
        """
        Retourne le canvas sous le pointeur au moment de l'Ã©vÃ©nement.
        Utile quand l'Ã©vÃ©nement vient d'un widget enfant (Label, Frame, etc.).
        """
        # Widget sous le pointeur en coord. Ã©cran
        top = event.widget.winfo_toplevel()
        w = top.winfo_containing(event.x_root, event.y_root)
        # Remonte l'arbre jusqu'Ã  trouver un Canvas (le plus proche)
        import tkinter as tk
        while w is not None and not isinstance(w, tk.Canvas):
            w = getattr(w, "master", None)
        return w

    def _pan_start(self, event):
        """
        DÃ©marre le panoramique (ALT + clic gauche ou bouton du milieu si activÃ©).
        Utilise scan_mark/scan_dragto pour un dÃ©placement fluide.
        """
        c = self._canvas_under_pointer(event)
        if c is None:
            return "break"

        self._pan_canvas = c
        self._panning = True

        # Convertit la position Ã©cran -> coords locales du canvas puis -> coords logiques
        x = int(c.canvasx(event.x_root - c.winfo_rootx()))
        y = int(c.canvasy(event.y_root - c.winfo_rooty()))
        c.scan_mark(x, y)

        # Curseur "fleur" pendant le pan
        try:
            self._prev_cursor = c["cursor"]
        except Exception:
            self._prev_cursor = ""
        c.configure(cursor="fleur")
        return "break"

    def _pan_move(self, event):
        """
        DÃ©placement pendant le panoramique.
        """
        if not self._panning or self._pan_canvas is None:
            return "break"

        c = self._pan_canvas
        x = int(c.canvasx(event.x_root - c.winfo_rootx()))
        y = int(c.canvasy(event.y_root - c.winfo_rooty()))
        c.scan_dragto(x, y, gain=1)
        return "break"

    def _pan_end(self, event):
        """
        Fin du panoramique : on restaure le curseur et on nettoie l'Ã©tat.
        """
        if self._pan_canvas is not None:
            try:
                self._pan_canvas.configure(cursor=self._prev_cursor or "")
            except Exception:
                pass
        self._panning = False
        self._pan_canvas = None
        return "break"



    



# Classe pour le tableau dynamique du dÃ©compte dÃ©taillÃ© des pÃ©riodes ("Shift Count")
class ShiftCountTable(tk.Frame):
    def __init__(self, master, planning_gui, **kwargs):
        kwargs.setdefault("bg", APP_SURFACE_BG)
        super().__init__(master, **kwargs)
        self.planning_gui = planning_gui
        self.columns = [
            "Initiales",
            "Semaine (mois)",
            "WE/Férié (mois)",
            "Total (mois)",
            "Cumul total",
            "Cumul semaine",
            "Cumul WE/Férié",
        ]
        self.tree = ttk.Treeview(
            self,
            columns=self.columns,
            show="headings",
            style=SHIFT_TREE_STYLE,
            selectmode="browse",
        )
        self.tree.heading("Initiales", text="Medecin", anchor="center")
        self.tree.column("Initiales", width=80, anchor="center")
        self.tree.heading("Semaine (mois)", text="Semaine (mois)", anchor="center")
        self.tree.column("Semaine (mois)", width=110, anchor="center")
        self.tree.heading("WE/Férié (mois)", text="WE/Férié (mois)", anchor="center")
        self.tree.column("WE/Férié (mois)", width=120, anchor="center")
        self.tree.heading("Total (mois)", text="Total (mois)", anchor="center")
        self.tree.column("Total (mois)", width=90, anchor="center")
        self.tree.heading("Cumul total", text="Cumul total", anchor="center")
        self.tree.column("Cumul total", width=90, anchor="center")
        self.tree.heading("Cumul semaine", text="Cumul semaine", anchor="center")
        self.tree.column("Cumul semaine", width=110, anchor="center")
        self.tree.heading("Cumul WE/Férié", text="Cumul WE/Férié", anchor="center")
        self.tree.column("Cumul WE/Férié", width=120, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)
        self.tree.tag_configure("highlight", background="#FFF4B5")
        self.tree.tag_configure("eligible", background="#CBE8CE")
        self.update_counts()
        self.active_day_index = None


    def update_counts(self):
        """Rebuild the shift-count table using the current planning entries."""
        if not hasattr(self.planning_gui, 'constraints_app'):
            return

        excluded = getattr(self.planning_gui, 'excluded_from_count', set())

        valid_initials = set()
        for row in self.planning_gui.constraints_app.rows:
            try:
                init = row[0].get().strip()
                if init:
                    valid_initials.add(init)
            except Exception:
                continue

        def _day_type(gui_obj, row_idx):
            """Retourne 'week' ou 'we' selon les règles weekend/férié/vendredi/veille férié."""
            try:
                day_text = gui_obj.day_labels[row_idx].cget("text").strip()
            except Exception:
                day_text = ""
            if not day_text.isdigit():
                return None
            try:
                day_num = int(day_text)
                dt = date(gui_obj.current_year, gui_obj.current_month, day_num)
            except Exception:
                return None

            weekend_rows = getattr(gui_obj, "weekend_rows", set())
            holiday_rows = getattr(gui_obj, "holiday_rows", set())
            holiday_dates = getattr(gui_obj, "holiday_dates", set())
            is_holiday = (row_idx in holiday_rows) or (dt in holiday_dates)
            is_weekend = (row_idx in weekend_rows) or dt.weekday() >= 5
            is_friday = dt.weekday() == 4
            try:
                next_day = dt + timedelta(days=1)
                next_day_holiday = next_day in holiday_dates and next_day.month == dt.month
            except Exception:
                next_day_holiday = False

            if is_holiday or is_weekend or is_friday or next_day_holiday:
                return "we"
            return "week"

        def collect_counts(gui_obj):
            """
            Retourne un d�compte par personne, en comptant 1 fois par jour (et non par ligne)
            pour �viter de sur-compter les jours multi-lignes. Les cellules exclues sont ignor�es.
            """
            excl = getattr(gui_obj, 'excluded_from_count', set())
            counts = {}
            for r, row in enumerate(gui_obj.table_entries):
                day_type = _day_type(gui_obj, r)
                if day_type is None:
                    continue
                names_in_day = set()
                for c, cell in enumerate(row):
                    if not cell or (r, c) in excl:
                        continue
                    try:
                        raw_value = cell.get()
                    except Exception:
                        raw_value = ""
                    names = extract_names_from_cell(raw_value, valid_initials)
                    if not names:
                        continue
                    for person in names:
                        if person in valid_initials:
                            names_in_day.add(person)
                if not names_in_day:
                    continue
                for person in names_in_day:
                    bucket = counts.setdefault(person, {"week": 0, "we": 0})
                    bucket[day_type] = bucket.get(day_type, 0) + 1
            return counts

        current_counts = collect_counts(self.planning_gui)

        # Cumul sur tous les onglets/mois (si disponibles)
        cumulative_counts = {}
        try:
            for (g, _c, _s) in globals().get("tabs_data", []):
                c_counts = collect_counts(g)
                for person, cnts in c_counts.items():
                    base = cumulative_counts.setdefault(person, {"week": 0, "we": 0})
                    base["week"] = base.get("week", 0) + cnts.get("week", 0)
                    base["we"] = base.get("we", 0) + cnts.get("we", 0)
        except Exception:
            pass

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree.tag_configure("evenrow", background=SHIFT_EVEN_ROW_BG)
        self.tree.tag_configure("oddrow",  background=SHIFT_ODD_ROW_BG)

        row_index = 0
        ordered_initials = []
        for c_row in self.planning_gui.constraints_app.rows:
            try:
                init = c_row[0].get().strip()
                if init and init in valid_initials:
                    ordered_initials.append(init)
            except Exception:
                continue

        for person in ordered_initials:
            month_counts = current_counts.get(person, {"week": 0, "we": 0})
            month_week = month_counts.get("week", 0)
            month_we = month_counts.get("we", 0)
            month_total = month_week + month_we

            cumul_counts = cumulative_counts.get(person, {"week": month_week, "we": month_we})
            cumul_week = cumul_counts.get("week", 0)
            cumul_we = cumul_counts.get("we", 0)
            cumul_total = cumul_week + cumul_we

            row_values = [
                person,
                month_week,
                month_we,
                month_total,
                cumul_total,
                cumul_week,
                cumul_we,
            ]

            tag = "evenrow" if row_index % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=tuple(row_values), tags=(tag,))
            row_index += 1

    def highlight_initial(self, initial):
        """
        Surligne en jaune clair la ligne du ShiftCount correspondant Ã  'initial',
        et remet le reste en evenrow/oddrow.
        """
        children = self.tree.get_children()
        for idx, item in enumerate(children):
            vals = self.tree.item(item, 'values')
            if initial and vals and vals[0] == initial:
                self.tree.item(item, tags=('highlight',))
            else:
                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
                self.tree.item(item, tags=(tag,))
    
    def highlight_states(self, current, eligibles):
        """
        Surligne en jaune la ligne correspondant Ã  'current',
        en vert clair celles dont l'initiale est dans 'eligibles',
        et remet en evenrow/oddrow les autres.
        """
        children = self.tree.get_children()
        for idx, item in enumerate(children):
            vals = self.tree.item(item, 'values')
            init = vals[0] if vals else None
            if init == current:
                tag = 'highlight'
            elif init in eligibles:
                tag = 'eligible'
            else:
                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.tree.item(item, tags=(tag,))


    def clear_highlight(self):
        """Remet toutes les lignes en evenrow/oddrow."""
        self.highlight_initial(None)

    def highlight_eligibles(self, initials):
        """
        Surligne en vert (tag 'eligible') les lignes du tableau de dÃ©compte
        dont l'initiale figure dans la liste 'initials',
        et remet le reste en evenrow/oddrow.
        """
        children = self.tree.get_children()
        for idx, item in enumerate(children):
            vals = self.tree.item(item, 'values')
            init = vals[0] if vals else None
            if init in initials:
                # ligne eligible : vert
                self.tree.item(item, tags=('eligible',))
            else:
                # retour au fond pair/impair
                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
                self.tree.item(item, tags=(tag,))


# Classe principale du planning consolidÃ©
class GUI(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        # Vous pouvez ajuster le zoom_factor selon vos prÃ©fÃ©rences
        self.zoom_factor = 1.5  
        self.table_entries = [[None for _ in range(len(work_posts))] for _ in range(len(days))]
        self.table_frames = [[None for _ in range(len(work_posts))] for _ in range(len(days))]
        self.table_labels = [[None for _ in range(len(work_posts))] for _ in range(len(days))]
        self.post_labels = {}
        self.cell_availability = {}
        self.warned_conflicts = set()
        self._update_job = None  # dÃ©bounce pour update_colors
        self.week_label = None

        self.day_labels = []
        today = date.today()
        self.current_year = today.year
        self.current_month = today.month
        self.weekend_rows = set()
        self.holiday_rows = set()
        self.holiday_dates = set()
        self._holiday_popup = None
        self.visible_day_count = len(days)
        self.hidden_rows = set()

        self._zoom_job = None
        self._zoom_accum = 1.0

        self._measure_fonts = None          # cache (f_cell, f_sched, f_head)
        self._measure_fonts_zoom = None     # zoom auquel le cache correspond

        
        # Pile dâannulation Â« planning Â» (assignations)
        self.undo_stack = []

        # ââââââââââ NOUVEAU : historique dâÃ©dition de cellule ââââââââââ
        self.cell_edit_undo_stack = []   # pile des cellules modifiÃ©es
        self.current_edit = None         # infos sur la cellule en cours
        self.swap_selection = None       # sÃ©lection pour swap (Shift-clic)
        self.copy_selection = None       # sÃ©lection pour copie (Ctrl-clic)   # <<< AJOUT
        # --- Cellules exclues du tableau de dÃ©compte ---
        self.excluded_from_count = set()   # {(row, col)} Ã  ignorer dans ShiftCountTable

        # --- Checkbox VÃ©rification (existant) ---
        self.verification_enabled = tk.BooleanVar(value=False)
        self.incompatible_cells  = set()

        # ---------------- Construction complÃ¨te de lâUI ----------------
        self.create_widgets()

    def _init_fonts(self):
        """
        CrÃ©e (ou met Ã  jour) des polices partagÃ©es pour le tableau.
        En changeant uniquement leur taille au moment du zoom, on Ã©vite
        de recrÃ©er tous les widgets (zoom bien plus fluide).
        """
        z = self.zoom_factor
        size_head  = max(6, int(8 * z))
        size_cell  = max(6, int(8 * z))
        size_sched = max(6, int(7 * z))

        if getattr(self, "f_header", None) is None:
            # CrÃ©ation la premiÃ¨re fois
            self.f_header = tkfont.Font(family=APP_FONT_FAMILY, size=size_head, weight="bold")
            self.f_cell   = tkfont.Font(family=APP_FONT_FAMILY, size=size_cell,  weight="bold")
            self.f_sched  = tkfont.Font(family=APP_FONT_FAMILY, size=size_sched)
        else:
            # Simple reconfiguration (zoom)
            self.f_header.configure(size=size_head)
            self.f_cell.configure(size=size_cell)
            self.f_sched.configure(size=size_sched)


    def create_widgets(self):
        # Polices partag?es -> zoom instantan? sans reconstruire
        self._init_fonts()
        header_font   = self.f_header
        cell_font     = self.f_cell
        pad = 1

        # Permet d'?tirer les colonnes lors du redimensionnement
        for col in range(len(work_posts) + 1):
            self.columnconfigure(col, weight=1)

        # En-t?te principale (servira d'?tiquette de mois)
        self.week_label = tk.Label(
            self,
            text="Mois",
            font=header_font,
            bg=APP_SURFACE_BG,
            fg=APP_PRIMARY_DARK,
            bd=0,
            padx=6,
            pady=4,
        )
        self.week_label.grid(row=0, column=0, padx=pad, pady=pad, sticky="nsew")
        self.week_label.bind("<Button-1>", self.edit_week_date)

        # En-t?tes des colonnes (astreintes)
        posts_source = getattr(self, "local_work_posts", work_posts)
        for col_idx, post in enumerate(posts_source):
            post_color = POST_INFO.get(post, {}).get("color", "#DDDDDD")
            lbl = tk.Label(
                self,
                text=post,
                font=header_font,
                bg=post_color,
                fg="black",
                borderwidth=0,
                padx=8,
                pady=6,
            )
            lbl.grid(row=0, column=col_idx + 1, padx=pad, pady=pad, sticky="nsew")
            lbl.bind("<Button-3>", lambda e, lbl=lbl: self.change_post_color(e, lbl.cget("text"), lbl))
            lbl.bind("<Double-Button-1>", lambda e, lbl=lbl: self.edit_post_name(e, lbl.cget("text"), lbl))
            self.post_labels[post] = lbl

        # Ligne d'actions (+) sous les en-t?tes
        day_header = tk.Label(
            self,
            text="Jours",
            font=header_font,
            bg=APP_PRIMARY_COLOR,
            fg="white",
            padx=8,
            pady=6,
            borderwidth=0,
        )
        day_header.grid(row=1, column=0, padx=pad, pady=pad, sticky="nsew")

        for col_idx, post in enumerate(posts_source):
            btn = tk.Button(
                self,
                text="+",
                font=(APP_FONT_FAMILY, int(8 * self.zoom_factor)),
                command=lambda p=post: self.open_post_action_dialog(p)
            )
            btn.grid(row=1, column=col_idx + 1, padx=pad, pady=pad, sticky="nsew")

        # Lignes de jours (1..30) et cellules d'affectation
        for day_idx, day_name in enumerate(days):
            grid_row = day_idx + 2
            day_lbl = tk.Label(
                self,
                text=day_name,
                font=cell_font,
                bg=DAY_LABEL_BG,
                fg="white",
                borderwidth=0,
                padx=8,
                pady=4,
            )
            day_lbl.grid(row=grid_row, column=0, padx=pad, pady=pad, sticky="nsew")
            day_lbl.bind("<Button-1>", lambda e, idx=day_idx: self.open_holiday_popup(idx))
            self.day_labels.append(day_lbl)

            for col_idx, post in enumerate(posts_source):
                frame = tk.Frame(
                    self,
                    borderwidth=0,
                    bg=APP_SURFACE_BG,
                    highlightthickness=1,
                    highlightbackground=APP_DIVIDER,
                )
                frame.grid(row=grid_row, column=col_idx + 1, padx=pad, pady=pad, sticky="nsew")
                self._bind_double_click(frame, day_idx, col_idx)

                entry = tk.Entry(
                    frame,
                    font=cell_font,
                    width=10,
                    justify="center",
                    relief="flat",
                    bd=0,
                    bg=CELL_EMPTY_BG,
                    disabledbackground=CELL_DISABLED_BG,
                    disabledforeground=CELL_DISABLED_TEXT,
                )

                entry.pack(side="top", fill="both", expand=True, padx=4, pady=4)
                entry.bind('<KeyRelease>', self.on_cell_key)
                entry.bind("<FocusIn>", self.on_cell_focus_in)
                entry.bind("<FocusOut>", self.on_cell_focus_out)
                entry.bind("<Control-Button-1>", lambda e, r=day_idx, c=col_idx: self.on_ctrl_click(e, r, c))
                entry.bind("<Shift-Button-1>",   lambda e, r=day_idx, c=col_idx: self.on_shift_click(e, r, c))
                self._bind_double_click(entry, day_idx, col_idx)

                self.cell_availability[(day_idx, col_idx)] = True
                self.table_entries[day_idx][col_idx] = entry
                self.table_frames[day_idx][col_idx] = frame
                self.table_labels[day_idx][col_idx] = None
                self.update_cell(day_idx, col_idx)

        self.auto_resize_all_columns()

    def is_cell_excluded_from_count(self, row_idx: int, col_idx: int) -> bool:
        """
        Indique si la cellule (row_idx, col_idx) est exclue des calculs/décomptes.
        """
        try:
            return (row_idx, col_idx) in self.excluded_from_count
        except Exception:
            return False

    def _apply_exclusion_style(self, row: int, col: int, excluded: bool) -> None:
        """
        Ajoute ou retire un liseré visuel pour les cellules exclues du décompte.
        """
        try:
            frame = self.table_frames[row][col]
            entry = self.table_entries[row][col]
            label = self.table_labels[row][col]
        except Exception:
            return

        if frame is None or entry is None:
            return

        is_disabled = not self.cell_availability.get((row, col), True)
        border_color = CELL_DISABLED_BG if is_disabled else (EXCLUDED_CELL_BORDER if excluded else APP_DIVIDER)
        try:
            frame.config(highlightbackground=border_color, highlightthickness=(2 if excluded else 1))
        except Exception:
            pass

        if label is not None:
            try:
                label.config(fg=(CELL_DISABLED_TEXT if excluded else "black"))
            except Exception:
                pass

    def refresh_exclusion_styles(self) -> None:
        """
        Réapplique le style d'exclusion pour chaque cellule (après reload/redraw).
        """
        rows = len(self.table_entries)
        cols = len(self.table_entries[0]) if self.table_entries else 0
        excluded = getattr(self, "excluded_from_count", set()) or set()
        for r in range(rows):
            for c in range(cols):
                try:
                    self._apply_exclusion_style(r, c, (r, c) in excluded)
                except Exception:
                    continue

    def undo_last_change(self):
        """
        Annule le dernier changement enregistré (copie/échange/saisie).
        """
        if not self.cell_edit_undo_stack:
            return
        action = self.cell_edit_undo_stack.pop()
        if isinstance(action, tuple) and len(action) == 3:
            actions = [action]
        elif isinstance(action, list):
            actions = action
        else:
            return
        for row, col, old_val in actions:
            try:
                cell = self.table_entries[row][col]
                if cell is None:
                    continue
                cell.delete(0, "end")
                if old_val:
                    cell.insert(0, old_val)
            except Exception:
                continue
        self.schedule_update_colors()

    def edit_week_date(self, event):
        """Modifie le libellé du mois affiché en en-tête."""
        try:
            new_date = custom_askstring(
                self,
                "Mois",
                "Entrez le mois :",
                event.x_root,
                event.y_root,
                self.week_label["text"],
            )
            if new_date:
                self.week_label.config(text=new_date)
        except Exception:
            pass

    # Nouveaux gestionnaires d'Ã©vÃ©nements pour le focus
    def on_cell_focus_in(self, event):
        widget = event.widget
        # 1) RÃ©initialiser les couleurs du planning
        self.update_colors(None)
        try:
            trigger_live_conflict_check()
        except Exception:
            pass

        # 2) RepÃ©rer le crÃ©neau cliquÃ© (row_idx, col_idx)
        row_idx = col_idx = None
        for i, row in enumerate(self.table_entries):
            for j, cell in enumerate(row):
                if cell is widget:
                    row_idx, col_idx = i, j
                    break
            if row_idx is not None:
                break
        if row_idx is None:
            return  # widget non trouvÃ© dans la grille

        # 3) Calculer tous les candidats Ã©ligibles (critÃ¨res de base seulement)
        eligibles = self.get_basic_eligible_replacements(row_idx, col_idx)
        eligibles_set = set(eligibles)

        valid_initials = set()
        constraints_app = getattr(self, "constraints_app", None)
        if constraints_app is not None:
            for cand_row in getattr(constraints_app, "rows", []):
                try:
                    init = cand_row[0].get().strip()
                except Exception:
                    continue
                if init:
                    valid_initials.add(init)
        parser_valids = valid_initials or None

        def _names_from_widget(cell_widget):
            if not cell_widget:
                return []
            try:
                raw = cell_widget.get()
            except Exception:
                raw = ""
            names = extract_names_from_cell(raw, parser_valids)
            if not names:
                raw_strip = raw.strip()
                if raw_strip:
                    names = [raw_strip]
            return names

        # 4) Surligner en vert clair toutes les cellules des Ã©ligibles
        for row in self.table_entries:
            for cell in row:
                if any(name in eligibles_set for name in _names_from_widget(cell)):
                    cell.config(bg="light green")

        # 5) Surligner en jaune la personne courante si la case n'est pas vide
        current_raw = widget.get().strip()
        current_names = _names_from_widget(widget)
        primary_current = current_names[0] if current_names else (current_raw or None)

        if current_names:
            current_set = set(current_names)
            for row in self.table_entries:
                for cell in row:
                    cell_names = _names_from_widget(cell)
                    if current_set.intersection(cell_names):
                        cell.config(bg="yellow")

        # 6) Mettre Ã  jour le ShiftCountTable (jaune et vert)
        if hasattr(self, 'shift_count_table'):
            self.shift_count_table.highlight_states(primary_current, eligibles)
        constraints_app = getattr(self, "constraints_app", None)
        if constraints_app is not None and hasattr(constraints_app, "highlight_candidate_initials"):
            try:
                constraints_app.highlight_candidate_initials(primary_current, eligibles_set)
            except Exception:
                pass

        # 7) âââ MÃ©morise la valeur dâorigine pour lâundo Â« cellule Â»
        self.current_edit = {
            'row': row_idx,
            'col': col_idx,
            'old': widget.get()
        }



    def on_cell_focus_out(self, event):
        """
        Callback when a cell loses focus.

        Ãtapes :
        1) RÃ©tablir les couleurs du planning.
        2) Nettoyer la surbrillance du ShiftCountTable.
        3) VÃ©rifier les incompatibilitÃ©s de la cellule Ã©ditÃ©e.
        4) Si la valeur a changÃ©, empiler lâancienne valeur pour Ctrl-Z (format Â« liste de changements Â»).
        """
        # 1) RÃ©tablit les couleurs du planning
        self.update_colors(None)

        # 2) EnlÃ¨ve la surbrillance du ShiftCountTable
        if hasattr(self, 'shift_count_table'):
            self.shift_count_table.clear_highlight()
        constraints_app = getattr(self, "constraints_app", None)
        if constraints_app is not None and hasattr(constraints_app, "clear_candidate_highlight"):
            try:
                constraints_app.clear_candidate_highlight()
            except Exception:
                pass

        # 3-4) Localiser la cellule qui perd le focus, vÃ©rifier contraintes
        widget = event.widget
        for row_idx, row in enumerate(self.table_entries):
            for col_idx, cell in enumerate(row):
                if cell is widget:
                    # 3) ContrÃ´le des contraintes
                    self.check_incompatibility(row_idx, col_idx)

                    # 4) Gestion de lâundo Â« cellule Â»
                    if self.current_edit and (row_idx, col_idx) == (
                        self.current_edit['row'], self.current_edit['col']):
                        old_val = self.current_edit['old']
                        new_val = widget.get()
                        if new_val != old_val:  # seulement si lâutilisateur a modifiÃ©
                            # On empile dÃ©sormais sous forme d'Â« action Â» (liste de changements)
                            # pour partager le mÃªme format que les swaps/copier.
                            self.cell_edit_undo_stack.append(
                                [(row_idx, col_idx, old_val)]
                            )
                        self.current_edit = None
                        self.auto_resize_column(col_idx)
                    return


    def schedule_update_colors(self, delay_ms: int = 120):
        """Programme un recalcul global des couleurs avec dÃ©bounce."""
        try:
            if self._update_job is not None:
                self.after_cancel(self._update_job)
        except Exception:
            pass
        self._update_job = self.after(delay_ms, self._do_update_colors)

    def _do_update_colors(self):
        """TÃ¢che diffÃ©rÃ©e : recalcul complet."""
        self._update_job = None
        self.update_colors(None)

    def on_cell_key(self, event):
        """
        Handler lÃ©ger pour les frappes : feedback immÃ©diat sur la cellule,
        puis recalcul global diffÃ©rÃ© (dÃ©bounce).
        """
        w = event.widget
        try:
            txt = w.get().strip()
            w.config(bg=(CELL_FILLED_BG if txt else CELL_EMPTY_BG), fg="black")
        except Exception:
            pass
        # Pas de push_undo_state ici : lâUNDO cellule est gÃ©rÃ© Ã  FocusOut.
        self.schedule_update_colors()

    def auto_resize_column(self, col_idx: int):
        """
        Ajuste la largeur minimale de la colonne (astreinte) en fonction :
        - des contenus des Entry (noms saisis),
        - de l'ent?te de colonne (astreinte).
        """
        try:
            f_cell, _, f_head = self._get_measure_fonts()
            posts_source = getattr(self, "local_work_posts", work_posts)
            header_text = posts_source[col_idx] if col_idx < len(posts_source) else ""
            max_px = f_head.measure(header_text)

            for i in range(len(self.table_entries)):
                ent = self.table_entries[i][col_idx]
                if ent is not None:
                    max_px = max(max_px, f_cell.measure(ent.get() or ""))

            pad = int(12 * self.zoom_factor)
            min_w = int(60 * self.zoom_factor)
            max_w = int(260 * self.zoom_factor)
            target = max(min_w, min(max_w, max_px + pad))

            # Colonne 0 = libell? des jours, on d?cale de +1 pour les astreintes
            self.grid_columnconfigure(col_idx + 1, minsize=target)
            self.update_idletasks()
        except Exception:
            pass


    def _get_measure_fonts(self):
        """
        Retourne (f_cell, f_sched, f_head) en cache.
        RecrÃ©e les tkfont.Font uniquement si le zoom a changÃ©.
        """
        z = self.zoom_factor
        if getattr(self, "_measure_fonts", None) is not None and self._measure_fonts_zoom == z:
            return self._measure_fonts

        f_cell = tkfont.Font(family=APP_FONT_FAMILY, size=int(8 * z), weight="bold")
        f_sched = tkfont.Font(family=APP_FONT_FAMILY, size=int(7 * z))
        f_head = tkfont.Font(family=APP_FONT_FAMILY, size=int(8 * z), weight="bold")

        self._measure_fonts = (f_cell, f_sched, f_head)
        self._measure_fonts_zoom = z
        return self._measure_fonts


    def auto_resize_all_columns(self):
        """Ajuste toutes les colonnes d'astreintes."""
        for j in range(len(work_posts)):
            self.auto_resize_column(j)

    def auto_resize_all_columns_fast(self, sample_rows: int = 10):
        for j in range(len(work_posts)):
            try:
                self._auto_resize_column_sampled(j, sample_rows)
            except Exception:
                pass

    def _auto_resize_column_sampled(self, col_idx: int, sample_rows: int):
        f_cell, _, f_head = self._get_measure_fonts()
        posts_source = getattr(self, "local_work_posts", work_posts)
        header_text = posts_source[col_idx] if col_idx < len(posts_source) else ""
        max_px = f_head.measure(header_text)

        total_rows = len(self.table_entries)
        limit = min(total_rows, max(1, int(sample_rows)))
        for i in range(limit):
            ent = self.table_entries[i][col_idx]
            if ent is not None:
                max_px = max(max_px, f_cell.measure(ent.get() or ""))

        pad = int(12 * self.zoom_factor)
        min_w = int(60 * self.zoom_factor)
        max_w = int(260 * self.zoom_factor)
        target = max(min_w, min(max_w, max_px + pad))
        self.grid_columnconfigure(col_idx + 1, minsize=target)

    def _bind_double_click(self, widget, row, col):
        """
        Lie un double-clic pour basculer la disponibilit? de la cellule (row, col).
        """
        widget.bind("<Double-Button-1>", lambda e, r=row, c=col: self.toggle_availability(r, c))

    def on_ctrl_click(self, event, row, col):
        """
        CTRL-clic = COPIE (Ã©crase).
        1er clic : sÃ©lectionne une cellule source (si non vide).
        2e clic : copie la valeur dans la cible et empile lâundo (format tuple).
        """
        entry = self.table_entries[row][col]
        cur_val = entry.get().strip()

        # 1) SÃ©lection (source)
        if self.copy_selection is None:
            if not cur_val:
                return  # rien Ã  copier
            # annule une sÃ©lection de swap en cours
            self.swap_selection = None
            self.copy_selection = (row, col, cur_val)
            entry.config(bg="pale green")
            return

        # 2) Application (cible)
        src_row, src_col, src_val = self.copy_selection
        src_entry = self.table_entries[src_row][src_col]
        tgt_entry = entry

        # Re-clic sur la mÃªme cellule => annulation
        if src_row == row and src_col == col:
            src_entry.config(bg="white")
            self.copy_selection = None
            self.schedule_update_colors()
            return

        old_tgt = tgt_entry.get()

        # --- Empile l'undo (format tuple compatible ancien undo) ---
        self.cell_edit_undo_stack.append((row, col, old_tgt))

        # --- Copie (Ã©crase) ---
        tgt_entry.delete(0, "end")
        tgt_entry.insert(0, src_val)

        # Nettoyage visuel + recalcul diffÃ©rÃ©
        self.copy_selection = None
        src_entry.config(bg="white")
        tgt_entry.config(bg="white")
        self.auto_resize_column(col)
        self.schedule_update_colors()



    def on_shift_click(self, event, row, col):
        """
        SHIFT-clic = SWAP (interversion).
        1er clic : sÃ©lectionne la premiÃ¨re cellule (peut Ãªtre vide).
        2e clic : Ã©change les contenus des deux cellules.
        Empile les deux anciennes valeurs pour Ctrl-Z (liste de changements).
        """
        entry = self.table_entries[row][col]
        cur_val = entry.get().strip()

        # 1) SÃ©lection
        if self.swap_selection is None:
            # annule une sÃ©lection de copie en cours
            self.copy_selection = None
            self.swap_selection = (row, col, cur_val)
            entry.config(bg="pale green")
            return

        # 2) Application
        src_row, src_col, src_val = self.swap_selection
        src_entry = self.table_entries[src_row][src_col]
        tgt_entry = entry
        tgt_val = tgt_entry.get()

        # Re-clic sur la mÃªme cellule => annulation
        if src_row == row and src_col == col:
            src_entry.config(bg="white")
            self.swap_selection = None
            self.schedule_update_colors()
            return

        # --- Empile l'undo (liste de deux tuples) ---
        undo_action = [(src_row, src_col, src_val), (row, col, tgt_val)]
        self.cell_edit_undo_stack.append(undo_action)

        # --- SWAP ---
        src_entry.delete(0, "end"); src_entry.insert(0, tgt_val)
        tgt_entry.delete(0, "end"); tgt_entry.insert(0, src_val)

        # Nettoyage visuel + recalcul diffÃ©rÃ©
        src_entry.config(bg="white")
        tgt_entry.config(bg="white")
        self.swap_selection = None
        self.auto_resize_column(src_col)
        if col != src_col:
            self.auto_resize_column(col)
        self.schedule_update_colors()



    def get_basic_eligible_replacements(self, row_idx, col_idx):
        """
        Retourne la liste des initiales ?ligibles pour la mise en surbrillance
        en appliquant toutes les contraintes individuelles connues.
        row_idx = jour, col_idx = astreinte.
        """
        if self.is_cell_excluded_from_count(row_idx, col_idx):
            return []

        constraints_app = getattr(self, "constraints_app", None)
        rows = getattr(constraints_app, "rows", []) if constraints_app is not None else []
        if not rows:
            return []

        valid_initials = set()
        for cand_row in rows:
            try:
                init = cand_row[0].get().strip()
            except Exception:
                continue
            if init:
                valid_initials.add(init)
        parser_valids = valid_initials or None

        def _names_from_cell(value):
            return extract_names_from_cell(value, parser_valids)

        context = PlanningContext(
            table_entries=self.table_entries,
            name_resolver=_names_from_cell,
            exclusion_checker=getattr(self, "is_cell_excluded_from_count", None),
            excluded_cells=getattr(self, "excluded_from_count", set()),
        )

        from Full_GUI import work_posts
        if not work_posts:
            return []
        post_idx = col_idx
        if post_idx < 0 or post_idx >= len(work_posts):
            return []
        post_name = work_posts[post_idx]
        is_morning = True  # unique cr?neau par jour/astreinte

        forbidden_morning, forbidden_afternoon = Assignation.build_forbidden_maps(work_posts)
        settings = AssignmentSettings(
            enable_max_assignments=Assignation.ENABLE_MAX_ASSIGNMENTS,
            max_assignments_per_post=Assignation.MAX_ASSIGNMENTS_PER_POST,
            enable_different_post_per_day=Assignation.ENABLE_DIFFERENT_POST_PER_DAY,
            enable_repos_securite=Assignation.ENABLE_REPOS_SECURITE,
            forbidden_morning_to_afternoon=forbidden_morning,
            forbidden_afternoon_to_morning=forbidden_afternoon,
        )

        eligibles = []
        for row in rows:
            profile = parse_constraint_row(row)
            if not profile:
                continue
            if candidate_is_available(
                profile,
                context,
                day_index=row_idx,
                post_index=post_idx,
                is_morning=is_morning,
                post_name=post_name,
                settings=settings,
            ):
                eligibles.append(profile.initial)
        return eligibles

    def toggle_availability(self, row, col):
        current_state = self.cell_availability.get((row, col), True)
        new_state = not current_state
        self.cell_availability[(row, col)] = new_state
        self.update_cell(row, col)

    def _collect_valid_initials(self):
        """
        Rassemble les initiales connues dans le tableau de contraintes.
        """
        initials = set()
        constraints_app = getattr(self, "constraints_app", None)
        rows = getattr(constraints_app, "rows", []) if constraints_app is not None else []
        for cand_row in rows:
            profile = parse_constraint_row(cand_row)
            if profile and profile.initial:
                initials.add(profile.initial.strip())
        return initials

    def choose_month(self):
        """Ouvre un calendrier pour choisir mois/année, puis met à jour l'en-tête + jours."""
        parent = self.winfo_toplevel()
        anchor = getattr(self, "scroll_canvas", self)
        try:
            initial = date(self.current_year, self.current_month, 1)
        except Exception:
            initial = date.today()
        picker = MonthPickerDialog(parent, initial_date=initial, anchor_widget=anchor)
        selected_date = getattr(picker, "result", None)
        if selected_date:
            self.apply_month_selection(selected_date.year, selected_date.month)

    def apply_month_selection(self, year: int, month: int):
        self.current_year = year
        self.current_month = month
        _, days_in_month = self._first_full_week_start(year, month)
        if 1 <= month <= 12:
            label_text = f"{MONTH_NAMES_FR[month - 1]} {year}"
        else:
            label_text = f"Mois {month} {year}"
        try:
            self.week_label.config(text=label_text)
        except Exception:
            pass

        day_numbers = list(range(1, days_in_month + 1))
        self.visible_day_count = days_in_month

        weekend_rows = set()
        holiday_rows = set()
        holiday_dates = month_holidays(year, month)
        for idx, day_value in enumerate(day_numbers):
            try:
                day_date = date(year, month, int(day_value))
                if day_date.weekday() >= 5:
                    weekend_rows.add(idx)
                if day_date in holiday_dates:
                    holiday_rows.add(idx)
            except Exception:
                continue

        self.weekend_rows = weekend_rows
        self.holiday_rows = holiday_rows
        self.holiday_dates = holiday_dates
        prev_hidden = getattr(self, "hidden_rows", set())
        new_hidden = set()
        total_cols = len(self.table_entries[0]) if self.table_entries else 0

        for idx, lbl in enumerate(self.day_labels):
            if idx >= len(self.table_entries):
                continue
            is_visible = idx < days_in_month
            try:
                if is_visible:
                    lbl.config(
                        text=str(day_numbers[idx])
                    )
                    lbl.grid()
                    self._apply_day_label_style(idx)
                else:
                    lbl.config(text="", bg=DAY_LABEL_BG, fg="white")
                    lbl.grid_remove()
                    new_hidden.add(idx)
            except Exception:
                pass

            for col_idx in range(total_cols):
                if idx >= len(self.table_frames) or col_idx >= len(self.table_frames[idx]):
                    continue
                frame = self.table_frames[idx][col_idx]
                entry = self.table_entries[idx][col_idx]
                if not frame or not entry:
                    continue
                if is_visible:
                    frame.grid()
                    if idx in prev_hidden:
                        self.cell_availability[(idx, col_idx)] = True
                    self.update_cell(idx, col_idx)
                else:
                    frame.grid_remove()
                    entry.delete(0, "end")
                    self.cell_availability[(idx, col_idx)] = False
                    self.update_cell(idx, col_idx)

        self.hidden_rows = new_hidden
        self.schedule_update_colors()

    def _row_to_date(self, row_idx: int) -> date | None:
        """
        Convertit l'index de ligne en objet date en se basant sur le libellé du jour.
        """
        try:
            if not (0 <= row_idx < len(self.day_labels)):
                return None
            label_text = self.day_labels[row_idx].cget("text").strip()
            if not label_text.isdigit():
                return None
            return date(self.current_year, self.current_month, int(label_text))
        except Exception:
            return None

    def _apply_day_label_style(self, row_idx: int) -> None:
        """
        Applique la coloration weekend/jour férié sur le libellé de jour.
        """
        try:
            lbl = self.day_labels[row_idx]
        except Exception:
            return
        is_weekend = row_idx in getattr(self, "weekend_rows", set())
        is_holiday = row_idx in getattr(self, "holiday_rows", set())
        bg_color = HOLIDAY_DAY_BG if is_holiday else (WEEKEND_DAY_BG if is_weekend else DAY_LABEL_BG)
        fg_color = "black" if (is_weekend or is_holiday) else "white"
        try:
            lbl.config(bg=bg_color, fg=fg_color)
        except Exception:
            pass

    def refresh_day_labels(self) -> None:
        """Réapplique la couleur de tous les libellés de jours."""
        for idx in range(len(self.day_labels)):
            self._apply_day_label_style(idx)

    def set_day_holiday(self, row_idx: int, is_holiday: bool) -> None:
        """
        Marque ou dé-marque une ligne comme jour férié, met à jour les couleurs et les données.
        """
        if not (0 <= row_idx < len(self.day_labels)):
            return
        if is_holiday:
            self.holiday_rows.add(row_idx)
            dt = self._row_to_date(row_idx)
            if dt:
                self.holiday_dates.add(dt)
        else:
            self.holiday_rows.discard(row_idx)
            dt = self._row_to_date(row_idx)
            if dt:
                self.holiday_dates.discard(dt)
        self._apply_day_label_style(row_idx)
        self.schedule_update_colors()

    def open_holiday_popup(self, day_idx: int) -> None:
        """
        Affiche un petit popup avec une case à cocher pour marquer/démarquer un jour férié.
        """
        try:
            if day_idx >= getattr(self, "visible_day_count", len(days)):
                return
            lbl = self.day_labels[day_idx]
            label_text = lbl.cget("text").strip()
            if not label_text.isdigit():
                return
        except Exception:
            return

        try:
            if self._holiday_popup is not None and self._holiday_popup.winfo_exists():
                self._holiday_popup.destroy()
        except Exception:
            pass

        popup = tk.Toplevel(self)
        popup.wm_overrideredirect(True)
        try:
            popup.attributes("-topmost", True)
        except Exception:
            pass
        popup.configure(bg=APP_SURFACE_BG, padx=10, pady=8, bd=1, relief="solid")

        info_lbl = tk.Label(popup, text=f"Jour {label_text}", bg=APP_SURFACE_BG, fg="black")
        info_lbl.pack(anchor="w")

        var = tk.BooleanVar(value=(day_idx in getattr(self, "holiday_rows", set())))
        cb = tk.Checkbutton(
            popup,
            text="Jour férié",
            variable=var,
            bg=APP_SURFACE_BG,
            anchor="w",
            command=lambda: self.set_day_holiday(day_idx, var.get()),
        )
        cb.pack(anchor="w", pady=(6, 2))

        popup.update_idletasks()
        try:
            x = lbl.winfo_rootx() + lbl.winfo_width() + 8
            y = lbl.winfo_rooty()
            popup.geometry(f"+{x}+{y}")
        except Exception:
            pass

        popup.bind("<FocusOut>", lambda e: popup.destroy())
        popup.bind("<Escape>",   lambda e: popup.destroy())
        popup.bind("<Destroy>",  lambda e: setattr(self, "_holiday_popup", None))
        try:
            popup.focus_force()
        except Exception:
            pass
        self._holiday_popup = popup

    @staticmethod
    def _first_full_week_start(year: int, month: int) -> tuple[int, int]:
        _, days_in_month = calendar.monthrange(year, month)
        return 1, days_in_month

    def select_and_close_posts(self):
        """
        Ouvre un popup pour Sélectionner plusieurs postes puis ferme (rend indisponibles)
        toutes les cellules des postes sÃ©lectionnÃ©s.
        """
        from Full_GUI import work_posts
        popup = MultiSelectPopup(self.master.master, work_posts)
        selected_posts = popup.selected
        if selected_posts:
            self.close_posts(selected_posts)

    def close_posts(self, post_names):
        """
        Ferme (d?sactive) toutes les cellules pour chaque poste donn?.
        """
        from Full_GUI import work_posts
        for post in post_names:
            if post in work_posts:
                col_idx = work_posts.index(post)
                for row_idx in range(len(days)):
                    self.cell_availability[(row_idx, col_idx)] = False
                    self.update_cell(row_idx, col_idx)

    def update_cell(self, row, col):
        frame = self.table_frames[row][col]
        entry = self.table_entries[row][col]
        label = self.table_labels[row][col]
        if not frame or not entry:
            return

        is_weekend = row in getattr(self, "weekend_rows", set())
        is_holiday = row in getattr(self, "holiday_rows", set())
        base_bg = HOLIDAY_CELL_BG if is_holiday else (WEEKEND_CELL_BG if is_weekend else APP_SURFACE_BG)
        empty_bg = HOLIDAY_CELL_BG if is_holiday else (WEEKEND_CELL_BG if is_weekend else CELL_EMPTY_BG)

        posts_source = getattr(self, "local_work_posts", work_posts)
        post_info_source = getattr(self, "local_post_info", POST_INFO)
        if col < len(posts_source):
            post = posts_source[col]
            post_color = post_info_source.get(post, {}).get("color", "#DDDDDD")
        else:
            post_color = "#DDDDDD"

        if self.cell_availability.get((row, col), True):
            frame.config(bg=base_bg, highlightbackground=APP_DIVIDER)
            if label:
                label.config(bg=post_color, fg="black")
            entry.config(state="normal", bg=empty_bg, fg="black")
        else:
            frame.config(bg=CELL_DISABLED_BG, highlightbackground=CELL_DISABLED_BG)
            if label:
                label.config(bg=CELL_DISABLED_BG, fg=CELL_DISABLED_BG)
            entry.config(
                state="disabled",
                disabledbackground=CELL_DISABLED_BG,
                disabledforeground=CELL_DISABLED_BG,
            )

        self._update_post_label_state(col)

        if hasattr(self, "is_cell_excluded_from_count"):
            self._apply_exclusion_style(row, col, self.is_cell_excluded_from_count(row, col))


    def _column_fully_disabled(self, col_idx: int) -> bool:
        for row_idx in range(len(days)):
            if self.cell_availability.get((row_idx, col_idx), True):
                return False
        return True

    def _update_post_label_state(self, post_index):
        posts_source = getattr(self, "local_work_posts", work_posts)
        post_info_source = getattr(self, "local_post_info", POST_INFO)
        if post_index < 0 or post_index >= len(posts_source):
            return
        post = posts_source[post_index]
        header_lbl = self.post_labels.get(post)
        if not header_lbl:
            return
        post_color = post_info_source.get(post, {}).get("color", "#DDDDDD")
        disabled_bg = CELL_DISABLED_BG
        disabled_fg = CELL_DISABLED_TEXT
        if self._column_fully_disabled(post_index):
            header_lbl.config(bg=disabled_bg, fg=disabled_fg)
        else:
            header_lbl.config(bg=post_color, fg="black")

    def update_colors(self, event):
        """
        Refresh table colors, rebuild assignments for ShiftCountTable,
        and recompute incompatibilities.
        """
        assignments = {}
        self.incompatible_cells.clear()
        weekend_rows = getattr(self, "weekend_rows", set())
        holiday_rows = getattr(self, "holiday_rows", set())

        for row_idx, row in enumerate(self.table_entries):
            for col_idx, cell in enumerate(row):
                if not cell:
                    continue
                if not self.cell_availability.get((row_idx, col_idx), True):
                    self.update_cell(row_idx, col_idx)
                    continue

                txt = cell.get().strip()
                empty_bg = HOLIDAY_CELL_BG if row_idx in holiday_rows else (WEEKEND_CELL_BG if row_idx in weekend_rows else CELL_EMPTY_BG)
                if txt:
                    cell.config(state="normal", bg=CELL_FILLED_BG, fg="black")
                    day_name = days[row_idx] if row_idx < len(days) else str(row_idx + 1)
                    local_posts = getattr(self, "local_work_posts", work_posts)
                    if col_idx < len(local_posts):
                        post_name = local_posts[col_idx]
                        assignments.setdefault(txt, {}).setdefault(day_name, []).append(post_name)
                else:
                    cell.config(state="normal", bg=empty_bg, fg="black")

        if hasattr(self, 'shift_count_table'):
            try:
                self.shift_count_table.update_counts()
            except Exception:
                pass

        for i, row in enumerate(self.table_entries):
            for j, cell in enumerate(row):
                if not self.cell_availability.get((i, j), True):
                    continue
                if cell.get().strip():
                    self.check_incompatibility(i, j)

        if getattr(self, 'verification_enabled', None) and self.verification_enabled.get():
            self.highlight_conflicts()

        self.update_idletasks()


    # === Marquage visuel des conflits inter-plannings (badge â rouge) ===
    def mark_cross_conflict(self, row_idx: int, col_idx: int, badge: str = "<->",
                            badge_bg: str = "#FF4D4F", badge_fg: str = "white", **kwargs):
        """
        Ajoute un petit badge 'â' dans le coin haut droit de la cellule (row_idx, col_idx)
        avec un fond rouge (par dÃ©faut) et du texte blanc. Ne modifie pas les couleurs de la cellule.
        CompatibilitÃ© : accepte aussi bg=/fg= via kwargs (ancien code).
        """
        import tkinter as tk

        # rÃ©tro-compat : si on vous passe bg=/fg=, on les mappe sur badge_bg/badge_fg
        if "bg" in kwargs and not badge_bg:
            badge_bg = kwargs["bg"]
        elif "bg" in kwargs:
            badge_bg = kwargs["bg"]
        if "fg" in kwargs and not badge_fg:
            badge_fg = kwargs["fg"]
        elif "fg" in kwargs:
            badge_fg = kwargs["fg"]

        if not hasattr(self, "_cross_conflict_badges"):
            self._cross_conflict_badges = {}  # (row,col) -> widget

        key = (row_idx, col_idx)

        # Si dÃ©jÃ  marquÃ© et widget toujours vivant, ne rien faire
        if key in self._cross_conflict_badges:
            w = self._cross_conflict_badges[key]
            try:
                if w and w.winfo_exists():
                    return
            except Exception:
                pass

        # On ancre le badge sur le frame de cellule si possible, sinon sur l'Entry
        target = None
        try:
            target = self.table_frames[row_idx][col_idx]
        except Exception:
            try:
                target = self.table_entries[row_idx][col_idx]
            except Exception:
                target = None

        if target is None:
            return

        # Petit label carrÃ© rouge, texte blanc, discret
        b = tk.Label(target, text=badge, font=(APP_FONT_FAMILY, 8, "bold"),
                     bg=badge_bg, fg=badge_fg, bd=0, padx=3, pady=0)
        try:
            # position coin haut droit
            b.place(relx=1.0, rely=0.0, x=-4, y=2, anchor="ne")
        except Exception:
            # fallback minimal si place() indispo : pack Ã  droite
            b.pack(anchor="ne")

        self._cross_conflict_badges[key] = b

    def clear_cross_conflict_marks(self):
        """Supprime tous les badges de conflits inter-plannings (â)."""
        if hasattr(self, "_cross_conflict_badges"):
            for key, widget in list(self._cross_conflict_badges.items()):
                try:
                    if widget and widget.winfo_exists():
                        widget.destroy()
                except Exception:
                    pass
            self._cross_conflict_badges.clear()


    def scroll_to_cell(self, row_idx: int, col_idx: int):
        """
        Fait dÃ©filer le canvas pour rendre visible la cellule et lui donner le focus.
        """
        canvas = getattr(self, "scroll_canvas", None)
        inner  = getattr(self, "scroll_inner",  None)
        entry  = None
        try:
            entry = self.table_entries[row_idx][col_idx]
        except Exception:
            entry = None

        try:
            if canvas is not None and inner is not None and entry is not None:
                inner_h   = max(1, inner.winfo_height())
                canvas_h  = max(1, canvas.winfo_height())
                target_top = entry.winfo_rooty() - inner.winfo_rooty()
                denom = max(1, inner_h - canvas_h)
                frac  = min(1.0, max(0.0, target_top / denom))
                canvas.yview_moveto(frac)
        except Exception:
            pass

        try:
            if entry is not None:
                entry.focus_set()
                old = entry.cget("bg")
                entry.config(bg="khaki1")
                self.after(250, lambda e=entry, o=old: e.config(bg=o))
        except Exception:
            pass



    def apply_verification(self):
        """
        Applique la vÃ©rification :
         1) RÃ©initialise les couleurs via update_colors
         2) Surligne vides et doublons
        """
        self.update_colors(None)
        self.highlight_conflicts()


    def highlight_conflicts(self):
        """
        Surligne en bleu les cellules vides et en rouge les doublons
        sans rÃ©initialiser toutes les couleurs.
        """
        # Vides â bleu
        for row in self.table_entries:
            for cell in row:
                if cell.get().strip() == "":
                    cell.config(bg="blue")

        if not self.table_entries:
            return
        num_days = len(self.table_entries[0])
        from collections import defaultdict

        for day in range(num_days):
            # Matin (lignes paires)
            morning_counts = defaultdict(int)
            morning_cells  = defaultdict(list)
            for i in range(0, len(self.table_entries), 2):
                cell = self.table_entries[i][day]
                val  = cell.get().strip()
                if val:
                    morning_counts[val] += 1
                    morning_cells[val].append(cell)
            for person, count in morning_counts.items():
                if count > 1:
                    for c in morning_cells[person]:
                        c.config(bg="red")

            # AprÃ¨s-midi (lignes impaires)
            afternoon_counts = defaultdict(int)
            afternoon_cells  = defaultdict(list)
            for i in range(1, len(self.table_entries), 2):
                cell = self.table_entries[i][day]
                val  = cell.get().strip()
                if val:
                    afternoon_counts[val] += 1
                    afternoon_cells[val].append(cell)
            for person, count in afternoon_counts.items():
                if count > 1:
                    for c in afternoon_cells[person]:
                        c.config(bg="red")




    def clear_schedule(self):
        """
        Efface tout le planning ET rend l'opÃ©ration annulable via CTRL-Z.
        On empile avant l'effacement la liste des cellules non vides
        sous la forme [(row, col, old_val), ...] consommÃ©e par undo_cell_edit.
        """
        # 1) Enregistrer un "gros" undo : toutes les cellules non vides
        changes = []
        for r, row_entries in enumerate(self.table_entries):
            for c, cell in enumerate(row_entries):
                if not cell:
                    continue
                old_val = cell.get()
                if old_val:  # on ne mÃ©morise que les vraies valeurs
                    changes.append((r, c, old_val))

        if changes:
            # Cette pile est lue par undo_cell_edit, qui accepte une liste de tuples
            self.cell_edit_undo_stack.append(changes)

        # 2) Effacer effectivement le contenu
        for row_entries in self.table_entries:
            for cell in row_entries:
                if cell:
                    cell.delete(0, "end")

        # 3) Nettoyer et rafraÃ®chir comme avant (inchangÃ©)
        self.warned_conflicts.clear()
        if hasattr(self, 'shift_count_table'):
            self.shift_count_table.update_counts()
        self.schedule_update_colors()

    def reset_layout_to_default(self):
        """
        RÃ©initialise le layout de la semaine en cours :
        - Restaure 5 postes (Â« Poste 1 Â» Ã  Â« Poste 5 Â») sans noms.
        - RÃ©initialise les horaires par dÃ©faut pour Poste 1 et Poste 2.
        - Efface toutes les affectations (tableau vide).
        Demande une confirmation Ã  l'utilisateur avant d'appliquer.
        """
        from tkinter import messagebox

        if not messagebox.askyesno(
            "Confirmer la rÃ©initialisation du layout",
            "Cette action va rÃ©initialiser la semaine au layout par dÃ©faut\n"
            "(5 postes vides, couleurs et horaires par dÃ©faut) et effacer les donnÃ©es.\n"
            "Voulez-vous continuer ?"
        ):
            return

        # 1) RÃ©initialiser la description des postes au niveau global.
        #    On reconstruit POST_INFO et la liste work_posts afin d'obtenir
        #    exactement le layout de dÃ©marrage.
        global POST_INFO
        POST_INFO.clear()
        base_shifts = {
            "MATIN":  {d: "08h00-13h00" for d in days},
            "AP MIDI": {d: "13h00-18h30" for d in days},
        }
        # Pas d'horaires le week-end
        for period in base_shifts.values():
            period["Samedi"] = None
            period["Dimanche"] = None

        for i in (1, 2):
            POST_INFO[f"Poste {i}"] = {
                "color": "#DDDDDD",
                "shifts": {
                    "MATIN":  base_shifts["MATIN"].copy(),
                    "AP MIDI": base_shifts["AP MIDI"].copy()
                }
            }

        # 2) RÃ©tablir la liste de 5 postes (Â« Poste 1 Â»..Â« Poste 5 Â»)
        update_work_posts([f"Poste {i}" for i in range(1, 6)])

        # 3) Reconstruire l'interface de la semaine avec un tableau vide
        self.redraw_widgets(preserve_content=False)

        # 4) Nettoyage/rafraÃ®chissement
        if hasattr(self, "warned_conflicts"):
            self.warned_conflicts.clear()
        if hasattr(self, 'shift_count_table'):
            self.shift_count_table.update_counts()
        self.schedule_update_colors()


    
    def add_work_post(self):
        # Trouve un nom "Nouveau poste N" qui n'existe pas encore (comparaison insensible Ã  la casse)
        base = "Nouveau poste"
        existing_norm = {p.strip().lower() for p in work_posts}
        n = 1
        while True:
            candidate = f"{base} {n}"
            if candidate.strip().lower() not in existing_norm:
                default_name = candidate
                break
            n += 1

        # Ajoute et notifie
        work_posts.append(default_name)
        update_work_posts(work_posts)

        # Initialise POST_INFO si le poste n'existe pas encore
        if default_name not in POST_INFO:
            POST_INFO[default_name] = {
                "color": "#DDDDDD",
                "shifts": {
                    "MATIN":  {d: "08h00-13h00"  for d in days},
                    "AP MIDI": {d: "13h00-18h30" for d in days}
                }
            }

        # Redessine le tableau principal tout en prÃ©servant le contenu
        self.redraw_widgets(preserve_content=True)

        # RafraÃ®chir Ã©galement la Combobox de suppression
        self.refresh_delete_post_combo()


  
    def _generate_duplicate_name(self, base_name: str) -> str:
        """Retourne un nom unique pour une duplication de poste."""
        normalized = {p.strip().lower() for p in work_posts}
        base_name = (base_name or '').strip()
        suffix = 2
        while True:
            candidate = f"{base_name} ({suffix})"
            if candidate.strip().lower() not in normalized:
                return candidate
            suffix += 1

    def _build_post_duplication_payload(self, post_index: int, clone_entries: bool) -> dict:
        """Prepare les donnees necessaires pour dupliquer un poste sur une GUI."""
        num_days = len(days)
        row_start = post_index * 2
        entries_snapshot: list[list[str]] = []
        labels_snapshot: list[list[tuple[str, str, str] | None]] = []
        availability_snapshot: list[list[bool]] = []
        excluded_snapshot: list[set[int]] = []
        excluded_source = getattr(self, 'excluded_from_count', set())

        for offset in range(2):
            row_idx = row_start + offset
            row_entries: list[str] = []
            row_labels: list[tuple[str, str, str] | None] = []
            row_availability: list[bool] = []
            for col_idx in range(num_days):
                cell = None
                if row_idx < len(self.table_entries):
                    row_cells = self.table_entries[row_idx]
                    if col_idx < len(row_cells):
                        cell = row_cells[col_idx]
                if clone_entries and cell is not None:
                    try:
                        entry_value = cell.get()
                    except Exception:
                        entry_value = ''
                else:
                    entry_value = ''
                row_entries.append(entry_value)
                label_widget = None
                if row_idx < len(self.table_labels):
                    row_labels_list = self.table_labels[row_idx]
                    if col_idx < len(row_labels_list):
                        label_widget = row_labels_list[col_idx]
                if label_widget is not None:
                    label_info = (
                        label_widget.cget('text'),
                        label_widget.cget('bg'),
                        label_widget.cget('fg'),
                    )
                else:
                    label_info = None
                row_labels.append(label_info)
                row_availability.append(self.cell_availability.get((row_idx, col_idx), True))
            entries_snapshot.append(row_entries)
            labels_snapshot.append(row_labels)
            excluded_cols = {c for c in range(num_days) if (row_idx, c) in excluded_source}
            availability_snapshot.append(row_availability)
            excluded_snapshot.append(excluded_cols)

        return {
            'index': post_index + 1,
            'entries': [list(row) for row in entries_snapshot],
            'labels': [list(row) for row in labels_snapshot],
            'availability': [list(row) for row in availability_snapshot],
            'excluded': [set(cols) for cols in excluded_snapshot],
        }

    def duplicate_work_post(self, post_name: str) -> str | None:
        """Cr?e une copie du poste et l'ins?re juste ? c?t?."""
        from tkinter import messagebox
        global work_posts, POST_INFO

        if post_name not in work_posts:
            messagebox.showerror("Duplication impossible", f'Le poste "{post_name}" est introuvable.')
            return None

        post_index = work_posts.index(post_name)
        new_post_name = self._generate_duplicate_name(post_name)
        color_source = POST_INFO.get(post_name, {}).get('color') or '#DDDDDD'

        work_posts.insert(post_index + 1, new_post_name)
        POST_INFO[new_post_name] = {'color': color_source, 'shifts': {}}
        update_work_posts(work_posts)

        gui_instances = get_all_gui_instances() or [self]
        for gui_instance in gui_instances:
            gui_instance.redraw_widgets(preserve_content=True)
            try:
                for row_idx in range(len(gui_instance.table_entries)):
                    src = gui_instance.table_entries[row_idx][post_index]
                    dst = gui_instance.table_entries[row_idx][post_index + 1]
                    if src and dst:
                        dst.delete(0, "end")
                        dst.insert(0, src.get())
                        gui_instance.cell_availability[(row_idx, post_index + 1)] = gui_instance.cell_availability.get((row_idx, post_index), True)
            except Exception:
                pass
            gui_instance.refresh_delete_post_combo()
            if hasattr(gui_instance, 'constraints_app'):
                try:
                    if hasattr(gui_instance.constraints_app, "refresh_work_posts"):
                        gui_instance.constraints_app.refresh_work_posts(work_posts)
                except Exception:
                    pass
            gui_instance.schedule_update_colors()
        return new_post_name

    def delete_work_post_specific(self, post_name):
        """
        Supprime le poste de travail donn? et rafra?chit l'interface.
        """
        from Full_GUI import work_posts, update_work_posts, POST_INFO

        if post_name not in work_posts:
            return

        idx = work_posts.index(post_name)
        work_posts.pop(idx)
        update_work_posts(work_posts)
        if post_name in POST_INFO:
            del POST_INFO[post_name]

        gui_instances = get_all_gui_instances() or [self]
        for gui_instance in gui_instances:
            gui_instance.redraw_widgets(preserve_content=True)
            gui_instance.refresh_delete_post_combo()
            if hasattr(gui_instance, 'constraints_app'):
                try:
                    if hasattr(gui_instance.constraints_app, "refresh_work_posts"):
                        gui_instance.constraints_app.refresh_work_posts(work_posts)
                except Exception:
                    pass
            gui_instance.schedule_update_colors()

    def open_post_action_dialog(self, post_name: str) -> None:
        """Affiche une fenêtre pour supprimer ou dupliquer un poste."""
        popup = tk.Toplevel(self)
        popup.title('Actions sur le poste')
        popup.transient(self.winfo_toplevel())
        popup.resizable(False, False)

        tk.Label(popup, text=f"Poste : {post_name}", anchor='w').pack(fill='x', padx=15, pady=(15, 5))

        delete_var = tk.BooleanVar(value=False)
        duplicate_var = tk.BooleanVar(value=False)

        tk.Checkbutton(popup, text='Delete the post?', variable=delete_var, anchor='w').pack(fill='x', padx=15, pady=2)
        tk.Checkbutton(popup, text='Duplicate the post?', variable=duplicate_var, anchor='w').pack(fill='x', padx=15, pady=2)

        button_frame = tk.Frame(popup)
        button_frame.pack(pady=15)

        def on_confirm():
            do_duplicate = duplicate_var.get()
            do_delete = delete_var.get()
            duplicate_result = None
            if do_duplicate:
                duplicate_result = self.duplicate_work_post(post_name)
            if do_delete and (duplicate_result is not None or not do_duplicate):
                self.delete_work_post_specific(post_name)
            popup.destroy()

        def on_cancel():
            popup.destroy()

        tk.Button(button_frame, text='OK', width=10, command=on_confirm).pack(side='left', padx=5)
        tk.Button(button_frame, text='Cancel', width=10, command=on_cancel).pack(side='left', padx=5)

        popup.update_idletasks()
        try:
            parent = self.winfo_toplevel()
        except Exception:
            parent = None

        placed = False
        if parent is not None:
            try:
                parent.update_idletasks()
                win_w = popup.winfo_reqwidth()
                win_h = popup.winfo_reqheight()
                parent_x = parent.winfo_rootx()
                parent_y = parent.winfo_rooty()
                parent_w = parent.winfo_width()
                parent_h = parent.winfo_height()
                pos_x = parent_x + max(0, (parent_w - win_w) // 2)
                pos_y = parent_y + max(0, (parent_h - win_h) // 2)
                popup.geometry(f'+{pos_x}+{pos_y}')
                placed = True
            except Exception:
                pass

        if not placed:
            try:
                screen_w = popup.winfo_screenwidth()
                screen_h = popup.winfo_screenheight()
                win_w = popup.winfo_reqwidth()
                win_h = popup.winfo_reqheight()
                pos_x = max(0, (screen_w - win_w) // 2)
                pos_y = max(0, (screen_h - win_h) // 2)
                popup.geometry(f'+{pos_x}+{pos_y}')
            except Exception:
                pass

        popup.protocol('WM_DELETE_WINDOW', on_cancel)
        popup.grab_set()
        popup.focus_set()
        popup.wait_window()

    def confirm_delete_work_post(self, post_name):
        """
        Affiche une boÃ®te de dialogue de confirmation avant
        de supprimer le poste de travail.
        """
        from tkinter import messagebox
        # Message de confirmation avec le nom du poste
        if messagebox.askyesno(
            "Confirmer la suppression",
            f"Voulez-vous vraiment supprimer le poste Â« {post_name} Â» ?"
        ):
            self.delete_work_post_specific(post_name)



    def change_post_color(self, event, post_name, label_widget=None):
        """
        Ouvre un sélecteur de couleur et applique la couleur à la colonne.
        """
        try:
            import tkinter.colorchooser as cc
        except Exception:
            return
        try:
            _rgb, hex_color = cc.askcolor(parent=self, title=f"Couleur pour {post_name}")
        except Exception:
            return
        if not hex_color:
            return
        POST_INFO.setdefault(post_name, {})["color"] = hex_color
        target_label = label_widget or self.post_labels.get(post_name)
        if target_label:
            try:
                target_label.config(bg=hex_color, fg="black")
            except Exception:
                pass
        if post_name in work_posts:
            col_idx = work_posts.index(post_name)
            for row_idx in range(len(self.table_entries)):
                self.update_cell(row_idx, col_idx)

    def edit_post_name(self, event, old_name, label_widget):
        """
        Permet de renommer un poste via une fenÃªtre popup.
        EmpÃªche les doublons de nom : on re-demande tant que le nom n'est pas unique.
        Met Ã  jour la liste globale work_posts, POST_INFO, puis notifie lâapplication.
        """
        from tkinter import messagebox
        # Boucle de saisie jusqu'Ã  obtenir un nom valide & unique, ou annulation
        initial = old_name
        while True:
            new_name = custom_askstring(
                self, "Modifier le nom",
                "Entrez le nouveau nom pour le poste :",
                event.x_root, event.y_root,
                initial=initial
            )
            if new_name is None:
                return  # annulÃ© par l'utilisateur

            new_name = new_name.strip()
            if not new_name:
                messagebox.showwarning("Nom invalide", "Le nom ne peut pas être vide.")
                initial = old_name
                continue

            # MÃªme nom (Ã  la casse/espaces prÃ¨s) â on ne fait rien
            norm_old = old_name.strip().lower()
            norm_new = new_name.lower()
            if norm_new == norm_old:
                return

            # Conflit : un autre poste porte dÃ©jÃ  ce nom (comparaison insensible Ã  la casse)
            if any(p.strip().lower() == norm_new for p in work_posts):
                messagebox.showerror(
                    "Nom dÃ©jÃ  utilisÃ©",
                    f"Le nom Â« {new_name} Â» est dÃ©jÃ  utilisÃ©.\nVeuillez choisir un nom unique."
                )
                initial = new_name
                continue

            # Nom valide et unique â on sort de la boucle
            break

        label_widget.config(text=new_name)

        gui_instances = get_all_gui_instances() or [self]

        try:
            index = work_posts.index(old_name)
            work_posts[index] = new_name
        except ValueError:
            pass

        if Assignation.FORBIDDEN_POST_ASSOCIATIONS:
            updated_pairs = set()
            changed = False
            for morning_name, afternoon_name in Assignation.FORBIDDEN_POST_ASSOCIATIONS:
                new_morning = new_name if morning_name == old_name else morning_name
                new_afternoon = new_name if afternoon_name == old_name else afternoon_name
                if new_morning != morning_name or new_afternoon != afternoon_name:
                    changed = True
                if new_morning in work_posts and new_afternoon in work_posts:
                    updated_pairs.add((new_morning, new_afternoon))
                else:
                    changed = True
            if changed:
                Assignation.FORBIDDEN_POST_ASSOCIATIONS.clear()
                Assignation.FORBIDDEN_POST_ASSOCIATIONS.update(updated_pairs)

        if old_name in POST_INFO:
            POST_INFO[new_name] = POST_INFO.pop(old_name)

        for gui_instance in gui_instances:
            labels_map = getattr(gui_instance, 'post_labels', {})
            if old_name in labels_map:
                header_lbl = labels_map.pop(old_name)
                labels_map[new_name] = header_lbl
                if header_lbl is not None:
                    header_lbl.config(text=new_name)
            gui_instance.refresh_delete_post_combo()
            if hasattr(gui_instance, 'constraints_app'):
                try:
                    if hasattr(gui_instance.constraints_app, "refresh_work_posts"):
                        gui_instance.constraints_app.refresh_work_posts(work_posts)
                except Exception:
                    pass

        update_work_posts(work_posts)

        for gui_instance in gui_instances:
            gui_instance.schedule_update_colors()

    def refresh_delete_post_combo(self):
        """Met Ã  jour la liste de postes dans la Combobox de suppression."""
        if hasattr(self, "del_post_combo") and self.del_post_combo is not None:
            self.del_post_combo["values"] = work_posts

    
    def redraw_widgets(self, preserve_content=False):
        saved_data = None
        saved_availability = None
        saved_excluded = set(getattr(self, "excluded_from_count", set())) if preserve_content else set()
        saved_weekend_rows = set(getattr(self, "weekend_rows", set()))
        saved_holiday_rows = set(getattr(self, "holiday_rows", set()))
        saved_holiday_dates = set(getattr(self, "holiday_dates", set()))
        saved_current_year = getattr(self, "current_year", None)
        saved_current_month = getattr(self, "current_month", None)
        saved_hidden_rows = set(getattr(self, "hidden_rows", set()))

        if preserve_content:
            saved_data = [
                [cell.get() if cell is not None else "" for cell in row]
                for row in self.table_entries
            ]
            saved_availability = dict(self.cell_availability)
        else:
            saved_excluded = set()

        for widget in self.winfo_children():
            widget.destroy()

        # Reinitialise les structures de libellés pour éviter les doublons après destruction
        self.day_labels = []
        self.table_entries = [[None for _ in range(len(work_posts))] for _ in range(len(days))]
        self.table_frames = [[None for _ in range(len(work_posts))] for _ in range(len(days))]
        self.table_labels = [[None for _ in range(len(work_posts))] for _ in range(len(days))]
        self.post_labels = {}
        if not preserve_content:
            self.cell_availability = {}

        self.create_widgets()

        if preserve_content and saved_data:
            for i in range(min(len(saved_data), len(self.table_entries))):
                for j in range(min(len(saved_data[i]), len(self.table_entries[i]))):
                    cell = self.table_entries[i][j]
                    if cell:
                        try:
                            cell.insert(0, saved_data[i][j])
                        except Exception:
                            pass

        if preserve_content and saved_availability:
            max_rows = len(self.table_entries)
            max_cols = len(self.table_entries[0]) if self.table_entries else 0
            for (r, c), avail in saved_availability.items():
                if 0 <= r < max_rows and 0 <= c < max_cols:
                    self.cell_availability[(r, c)] = avail
                    self.update_cell(r, c)

        max_rows = len(self.table_entries)
        max_cols = len(self.table_entries[0]) if self.table_entries else 0
        if preserve_content:
            filtered_excluded = {(r, c) for (r, c) in saved_excluded if 0 <= r < max_rows and 0 <= c < max_cols}
            self.excluded_from_count = filtered_excluded
        else:
            self.excluded_from_count = set()
        self.refresh_exclusion_styles()

        # Réapplique le mois courant pour conserver week-ends/jours fériés après redessin
        try:
            cur_y = getattr(self, "current_year", None)
            cur_m = getattr(self, "current_month", None)
            if cur_y and cur_m:
                self.apply_month_selection(cur_y, cur_m)
                if preserve_content:
                    self.weekend_rows = set(saved_weekend_rows)
                    self.holiday_rows = set(saved_holiday_rows)
                    self.holiday_dates = set(saved_holiday_dates)
                    self.hidden_rows = set(saved_hidden_rows)
                    self.refresh_day_labels()
        except Exception:
            # Fallback : réapplique simplement les marquages précédents
            self.weekend_rows = saved_weekend_rows
            self.holiday_rows = saved_holiday_rows
            self.holiday_dates = saved_holiday_dates
            self.hidden_rows = saved_hidden_rows
            try:
                for idx, lbl in enumerate(self.day_labels):
                    is_weekend = idx in saved_weekend_rows
                    is_holiday = idx in saved_holiday_rows
                    fg_color = "black" if (is_weekend or is_holiday) else "white"
                    bg_color = HOLIDAY_DAY_BG if is_holiday else (WEEKEND_DAY_BG if is_weekend else DAY_LABEL_BG)
                    lbl.config(bg=bg_color, fg=fg_color)
            except Exception:
                pass
        # Si current_year/month ont été perdus (peu probable), restaure les anciens
        if getattr(self, "current_year", None) is None and saved_current_year:
            self.current_year = saved_current_year
        if getattr(self, "current_month", None) is None and saved_current_month:
            self.current_month = saved_current_month

        for r in range(len(self.table_entries)):
            for c in range(len(work_posts)):
                self.update_cell(r, c)
        self.update_colors(None)
        self.auto_resize_all_columns()


    def check_incompatibility(self, row_idx, col_idx):
        """
        Colore en rouge la cellule (row_idx, col_idx) lorsqu?elle viole une
        contrainte : absence, poste non-assur?, double affectation le m?me jour,
        quota d?pass?.
        """
        if getattr(self, "constraints_app", None) is None:
            return False
        cell = self.table_entries[row_idx][col_idx]
        initial = cell.get().strip()
        parser_valids = self._collect_valid_initials() or None

        if not initial:
            self.incompatible_cells.discard((row_idx, col_idx))
            cell.config(bg="white")
            return

        if self.is_cell_excluded_from_count(row_idx, col_idx):
            cell.config(bg=CELL_FILLED_BG, fg="black")
            self.incompatible_cells.discard((row_idx, col_idx))
            return

        current_names = extract_names_from_cell(initial, parser_valids) or [initial]
        current_name_set = set(current_names)

        day_idx = row_idx
        post_idx = col_idx
        from Full_GUI import work_posts
        post_name = work_posts[post_idx] if post_idx < len(work_posts) else None

        cand_row = next((r for r in self.constraints_app.rows if r[0].get().strip() == initial), None)
        if cand_row is None:
            cell.config(bg="white")
            self.incompatible_cells.discard((row_idx, col_idx))
            return

        try:
            quota_total = int(cand_row[1].get().strip())
        except Exception:
            quota_total = None

        if quota_total is not None:
            total_assignments = 0
            for r_idx, row_entries in enumerate(self.table_entries):
                for c_idx, other_cell in enumerate(row_entries):
                    if not other_cell or self.is_cell_excluded_from_count(r_idx, c_idx):
                        continue
                    try:
                        raw_value = other_cell.get()
                    except Exception:
                        raw_value = ""
                    raw_value = (raw_value or "").strip()
                    if not raw_value:
                        continue
                    names = extract_names_from_cell(raw_value, parser_valids)
                    if names:
                        total_assignments += names.count(initial)
                    elif raw_value == initial:
                        total_assignments += 1
            if total_assignments > quota_total:
                cell.config(bg="red")
                self.incompatible_cells.add((row_idx, col_idx))
                return

        try:
            tpl = cand_row[4 + day_idx]
            state_raw = tpl[0]._var.get() if isinstance(tpl, tuple) else ""
        except Exception:
            state_raw = ""
        state = (state_raw or "").strip().upper()
        state_ascii = unicodedata.normalize("NFKD", state).encode("ASCII", "ignore").decode() if state else state
        if state_ascii in {"MATIN", "AP MIDI", "APMIDI", "JOURNEE"}:
            cell.config(bg="red")
            self.incompatible_cells.add((row_idx, col_idx))
            return

        try:
            nas = cand_row[3]._var.get() or cand_row[3].cget("text")
        except Exception:
            nas = ""
        if post_name in [p.strip() for p in str(nas).split(",") if p.strip()]:
            cell.config(bg="red")
            self.incompatible_cells.add((row_idx, col_idx))
            return

        if Assignation.ENABLE_DIFFERENT_POST_PER_DAY:
            for other_col in range(len(self.table_entries[0])):
                if other_col == col_idx:
                    continue
                other_cell = self.table_entries[row_idx][other_col]
                if not other_cell or self.is_cell_excluded_from_count(row_idx, other_col):
                    continue
                other_raw = (other_cell.get() or "").strip()
                if not other_raw:
                    continue
                other_names = extract_names_from_cell(other_raw, parser_valids) or [other_raw]
                if any(name in current_name_set for name in other_names):
                    cell.config(bg="red")
                    self.incompatible_cells.add((row_idx, col_idx))
                    return

        cell.config(bg=CELL_FILLED_BG, fg="black")
        self.incompatible_cells.discard((row_idx, col_idx))



def save_status(file_path=None, *, update_caption=True):
    """
    Sauvegarde le statut complet de l'interface.
    - Si file_path est None â ouvre une boÃ®te 'Enregistrer sousâ¦'
    - Sinon               â enregistre silencieusement dans file_path.
    Le chemin utilisÃ© est mÃ©morisÃ© dans la variable globale
    'current_status_path' pour de futurs Â« Enregistrer Â».
    """
    import pickle
    global current_status_path

    # Choix du fichier uniquement en mode 'Save As'
    if file_path is None:
        file_path = filedialog.asksaveasfilename(
            title="Sauvegarder le statut",
            defaultextension=".pkl",
            filetypes=[("Pickle Files", "*.pkl"), ("All Files", "*.*")]
        )
        if not file_path:
            return  # annulation utilisateur

    # ---------- Sauvegarde de toutes les semaines --------------------------
    all_week_status = []
    for (g, c, s) in tabs_data:
        # 1) Planning principal (texte)
        table_data = [
            [cell.get() if cell is not None else None for cell in row]
            for row in g.table_entries
        ]

        # 2) DisponibilitÃ©s
        cell_availability_data = dict(g.cell_availability)

        # 3) Tableau de contraintes
        constraints_data = []
        for row in c.rows:
            row_values = []
            for widget in row:
                if isinstance(widget, tuple) and len(widget) == 3:
                    toggle, pds_cb, pds_var = widget
                    try:
                        origin = toggle.get_origin()
                    except Exception:
                        origin = getattr(toggle, "origin", "manual")
                    try:
                        note = toggle.get_log()
                    except Exception:
                        note = getattr(toggle, "log_text", "")
                    row_values.append((toggle._var.get(), pds_var.get(), origin, note))
                elif isinstance(widget, tk.Button):
                    # Boutons avec variable associée (préférences / absences)
                    if hasattr(widget, "var"):
                        row_values.append(widget.var.get())
                    elif hasattr(widget, "_var"):
                        row_values.append(widget._var.get())
                    else:
                        row_values.append(widget.cget("text"))
                elif hasattr(widget, "_var"):
                    row_values.append(widget._var.get())
                else:
                    row_values.append(widget.get())
            constraints_data.append(row_values)

        # 4) Couleurs & texte des labels horaires
        schedule_data = []
        for row in g.table_labels:
            schedule_row = []
            for lbl in row:
                if lbl:
                    schedule_row.append((lbl.cget("text"),
                                          lbl.cget("bg"),
                                          lbl.cget("fg")))
                else:
                    schedule_row.append(None)
            schedule_data.append(schedule_row)

        # 5) LibellÃ© de la semaine
        week_label_text = g.week_label.cget("text")

        # 6) NOUVEAU : cellules exclues du dÃ©compte
        excluded_cells = sorted(list(getattr(g, "excluded_from_count", set())))

        # 7) MÃ©tadonnÃ©es (annÃ©e/mois, marquages, lignes masquÃ©es)
        meta = {
            "year": getattr(g, "current_year", None),
            "month": getattr(g, "current_month", None),
            "hidden_rows": sorted(list(getattr(g, "hidden_rows", set()))),
            "weekend_rows": sorted(list(getattr(g, "weekend_rows", set()))),
            "holiday_rows": sorted(list(getattr(g, "holiday_rows", set()))),
            "holiday_dates": sorted(list(getattr(g, "holiday_dates", set()))),
        }

        # On sauvegarde dÃ©sormais 7 Ã©lÃ©ments (compat old: le load gÃ¨re 5/6/7)
        all_week_status.append((
            table_data,
            cell_availability_data,
            constraints_data,
            schedule_data,
            week_label_text,
            excluded_cells,
            meta,
        ))

    # Options dâassignation
    valid_posts = set(work_posts)
    stored_pairs = sorted((m, a) for (m, a) in Assignation.FORBIDDEN_POST_ASSOCIATIONS if m in valid_posts and a in valid_posts)

    assignment_options = (
        Assignation.ENABLE_DIFFERENT_POST_PER_DAY,
        Assignation.ENABLE_MAX_ASSIGNMENTS,
        Assignation.MAX_ASSIGNMENTS_PER_POST,
        Assignation.ENABLE_REPOS_SECURITE,
        stored_pairs,
        getattr(Assignation, "ENABLE_MAX_WE_DAYS", False),
        getattr(Assignation, "MAX_WE_DAYS_PER_MONTH", None),
        getattr(Assignation, "ENABLE_WEEKEND_BLOCKS", False),
    )

    try:
        with open(file_path, "wb") as f:
            pickle.dump(
                (all_week_status, work_posts, POST_INFO, assignment_options),
                f
            )
        current_status_path = file_path  # mÃ©morisation
        if update_caption:
            update_window_caption()
        messagebox.showinfo("Sauvegarde", f"Statut sauvegardé dans {file_path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de sauvegarder : {e}")



def quick_save_status():
    """
    Enregistrement direct (menu Â« Enregistrer Â»).

    â¢ Si un fichier utilisateur est dÃ©jÃ  associÃ© ET quâil existe,
      on enregistre silencieusement dedans.
    â¢ Sinon (aucun fichier, ou bien câest la sauvegarde automatique),
      on bascule sur le comportement Â« Enregistrer sousâ¦ Â».
    """
    import os
    global current_status_path

    # Fichier dÃ©jÃ  dÃ©fini, existe, et nâest PAS la sauvegarde auto â save silencieux
    if (current_status_path
        and os.path.exists(current_status_path)
        and os.path.basename(current_status_path) != "sauvegarde_auto.pkl"):
        save_status(current_status_path)
    else:
        # Aucun fichier ou bien câest 'sauvegarde_auto.pkl' â forcer Save As
        save_status()



# ---------------------------------------------------------------------------
#  Utilitaire : dossier de donnÃ©es spÃ©cifique Ã  lâutilisateur
# ---------------------------------------------------------------------------
from pathlib import Path
import sys, os, datetime, traceback

def get_user_data_dir() -> Path:
    """
    Retourne un dossier du profil utilisateur oÃ¹ lâapplication peut
    librement Ã©crire :
      â¢ Windows :  %APPDATA%\PlanningRadiologie
      â¢ macOS   :  ~/Library/Application Support/PlanningRadiologie
      â¢ Linux   :  ~/.local/share/PlanningRadiologie
    Le dossier est crÃ©Ã© sâil nâexiste pas.
    """
    if sys.platform.startswith("win"):
        base = os.getenv("APPDATA") or Path.home() / "AppData" / "Roaming"
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:  # Linux & autres Unix
        base = Path.home() / ".local" / "share"

    data_dir = Path(base) / "PlanningRadiologie"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


# ---------------------------------------------------------------------------
#  Sauvegarde automatique silencieuse
# ---------------------------------------------------------------------------
def auto_save_status():
    """
    Enregistre l'Ã©tat courant dans :
        <dossier-donnÃ©es-utilisateur>/sauvegarde_auto.pkl

    â¢ Fonctionne en script, PyInstaller *onedir* et *onefile*.
    â¢ En cas dâerreur dâÃ©criture, loggue la trace dans
          .../autosave_error.log
      (aucun pop-up ne bloque l'utilisateur).
    â¢ Ne remplace plus current_status_path si l'utilisateur
      a dÃ©jÃ  choisi un fichier.
    """
    import tkinter.messagebox as _mb
    import datetime, traceback
    from pathlib import Path

    global current_status_path

    save_dir   = get_user_data_dir()
    pickle_path = save_dir / "sauvegarde_auto.pkl"
    log_path    = save_dir / "autosave_error.log"

    # Temporarily silence message-boxes
    orig_info, orig_err = _mb.showinfo, _mb.showerror
    _mb.showinfo = lambda *a, **k: None
    _mb.showerror = lambda *a, **k: None

    # --- Retient la cible actuelle de l'utilisateur ----------------------
    previous_path = current_status_path

    try:
        # save_status modifie current_status_path â on la restaure aprÃ¨s
        save_status(pickle_path, update_caption=False)
    except Exception:
        # DerniÃ¨re chance : journaliser lâexception
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}]\n")
            traceback.print_exc(file=f)
    finally:
        # Si l'utilisateur avait dÃ©jÃ  un fichier ouvert, on le remet
        if previous_path is not None:
            current_status_path = previous_path
            update_window_caption()
        _mb.showinfo, _mb.showerror = orig_info, orig_err

def open_autosave_folder():
    """
    Ouvre le dossier des donnÃ©es utilisateur contenant la sauvegarde automatique.
    - RÃ©vÃ¨le 'sauvegarde_auto.pkl' si possible (Windows/macOS).
    - Sinon ouvre simplement le dossier.
    """
    import sys, os, subprocess
    from pathlib import Path
    try:
        save_dir = get_user_data_dir()
        save_dir.mkdir(parents=True, exist_ok=True)
        pickle_path = save_dir / "sauvegarde_auto.pkl"

        if sys.platform.startswith("win"):
            # Sur Windows, si le fichier existe, on le rÃ©vÃ¨le ('/select,').
            if pickle_path.exists():
                subprocess.Popen(["explorer", "/select,", str(pickle_path)])
            else:
                # Sinon on ouvre le dossier
                os.startfile(str(save_dir))
        elif sys.platform == "darwin":
            # Sur macOS, 'open -R' rÃ©vÃ¨le le fichier, sinon 'open' ouvre le dossier
            if pickle_path.exists():
                subprocess.Popen(["open", "-R", str(pickle_path)])
            else:
                subprocess.Popen(["open", str(save_dir)])
        else:
            # Linux : ouvrir le dossier via xdg-open
            subprocess.Popen(["xdg-open", str(save_dir)])
    except Exception as e:
        # Fallback : informer du chemin exact
        try:
            messagebox.showinfo(
                "Sauvegarde automatique",
                f"Dossier: {save_dir}\n\nImpossible d'ouvrir automatiquement l'explorateur.\nDÃ©tail: {e}"
            )
        except Exception:
            pass


def open_replace_name_dialog():
    """
    Opens a simple dialog to replace every occurrence of a name across
    all weeks (planning + constraints tables).
    """
    if root is None:
        messagebox.showinfo("Remplacer nom", "Interface principale non disponible.")
        return
    data = globals().get("tabs_data")
    if not data:
        messagebox.showinfo("Remplacer nom", "Aucun planning n'est actuellement chargé.")
        return

    dialog = tk.Toplevel(root)
    dialog.title("Remplacer nom")
    dialog.transient(root)
    dialog.resizable(False, False)
    dialog.grab_set()

    ttk.Label(dialog, text="Nom à remplacer :").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
    replace_from_var = tk.StringVar()
    replace_from_entry = ttk.Entry(dialog, textvariable=replace_from_var, width=32)
    replace_from_entry.grid(row=1, column=0, sticky="ew", padx=12)

    ttk.Label(dialog, text="Nouveau nom :").grid(row=2, column=0, sticky="w", padx=12, pady=(10, 4))
    replace_to_var = tk.StringVar()
    replace_to_entry = ttk.Entry(dialog, textvariable=replace_to_var, width=32)
    replace_to_entry.grid(row=3, column=0, sticky="ew", padx=12)

    button_row = ttk.Frame(dialog)
    button_row.grid(row=4, column=0, pady=12, padx=12, sticky="e")

    def close_dialog(_event=None):
        dialog.destroy()

    def handle_apply(_event=None):
        search_text = replace_from_var.get().strip()
        replacement_text = replace_to_var.get()
        if not search_text:
            messagebox.showwarning("Remplacer nom", "Veuillez saisir le nom à rechercher.")
            return
        changes = replace_name_everywhere(search_text, replacement_text)
        if changes:
            messagebox.showinfo("Remplacer nom", f"{changes} champ(s) mis à jour.")
        else:
            messagebox.showinfo("Remplacer nom", "Aucune correspondance trouvée.")
        close_dialog()

    apply_btn = ttk.Button(button_row, text="Appliquer", command=handle_apply, style=RIBBON_ACCENT_BUTTON_STYLE)
    apply_btn.pack(side="right", padx=(8, 0))
    cancel_btn = ttk.Button(button_row, text="Annuler", command=close_dialog, style=RIBBON_BUTTON_STYLE)
    cancel_btn.pack(side="right")

    dialog.columnconfigure(0, weight=1)
    dialog.bind("<Return>", handle_apply)
    dialog.bind("<Escape>", close_dialog)

    # Position the dialog like the "+" popup: centered on the main window with a screen fallback
    placed = False
    dialog.update_idletasks()
    try:
        root.update_idletasks()
        win_w = dialog.winfo_reqwidth()
        win_h = dialog.winfo_reqheight()
        root_x = root.winfo_rootx()
        root_y = root.winfo_rooty()
        root_w = root.winfo_width()
        root_h = root.winfo_height()
        pos_x = root_x + max(0, (root_w - win_w) // 2)
        pos_y = root_y + max(0, (root_h - win_h) // 2)
        dialog.geometry(f"+{pos_x}+{pos_y}")
        placed = True
    except Exception:
        pass

    if not placed:
        try:
            screen_w = dialog.winfo_screenwidth()
            screen_h = dialog.winfo_screenheight()
            win_w = dialog.winfo_reqwidth()
            win_h = dialog.winfo_reqheight()
            pos_x = max(0, (screen_w - win_w) // 2)
            pos_y = max(0, (screen_h - win_h) // 2)
            dialog.geometry(f"+{pos_x}+{pos_y}")
        except Exception:
            pass

    replace_from_entry.focus_set()


def replace_name_everywhere(search_text: str, replacement_text: str) -> int:
    """
    Replaces every occurrence of `search_text` (literal, case-sensitive)
    across all GUI tables and constraint inputs. Returns the number of
    widgets updated.
    """
    if not search_text:
        return 0
    total_changes = 0
    data = globals().get("tabs_data") or []
    for item in data:
        try:
            gui_instance, constraints_app, _ = item
        except Exception:
            continue
        if gui_instance is None:
            continue
        total_changes += _replace_in_week(gui_instance, constraints_app, search_text, replacement_text)
    return total_changes


def _replace_in_week(gui_instance, constraints_app, search_text, replacement_text) -> int:
    week_changes = 0
    undo_payload = []
    columns_to_resize = set()

    entries = getattr(gui_instance, "table_entries", [])
    for row_idx, row in enumerate(entries):
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            try:
                current_value = cell.get()
            except Exception:
                continue
            if not current_value or search_text not in current_value:
                continue
            new_value = current_value.replace(search_text, replacement_text)
            if new_value == current_value:
                continue
            undo_payload.append((row_idx, col_idx, current_value))
            cell.delete(0, "end")
            if new_value:
                cell.insert(0, new_value)
            columns_to_resize.add(col_idx)
            week_changes += 1

    constraint_updates, constraint_payload = _replace_in_constraints(constraints_app, search_text, replacement_text)
    if constraint_payload:
        undo_payload.extend(constraint_payload)
        week_changes += constraint_updates

    if undo_payload:
        gui_instance.cell_edit_undo_stack.append(undo_payload)
        for col_idx in columns_to_resize:
            try:
                gui_instance.auto_resize_column(col_idx)
            except Exception:
                pass
        try:
            gui_instance.schedule_update_colors()
        except Exception:
            pass

    return week_changes


def _replace_in_constraints(constraints_app, search_text, replacement_text):
    if constraints_app is None or not search_text:
        return 0, []

    rows = getattr(constraints_app, "rows", None)
    if not rows:
        return 0, []

    changes = []
    updated_widgets = 0
    for row in rows:
        for widget in row:
            if not isinstance(widget, (tk.Entry, ttk.Entry)):
                continue
            try:
                current_value = widget.get()
            except Exception:
                continue
            if not current_value or search_text not in current_value:
                continue
            new_value = current_value.replace(search_text, replacement_text)
            if new_value == current_value:
                continue
            changes.append((UNDO_CONSTRAINT_TAG, widget, current_value))
            widget.delete(0, "end")
            if new_value:
                widget.insert(0, new_value)
            updated_widgets += 1

    return updated_widgets, changes


def reset_layout_from_menu():
    global gui
    current_gui = gui
    if current_gui is None:
        messagebox.showinfo("Effacer layout", "Aucun planning actif n'est disponible.")
        return
    try:
        current_gui.reset_layout_to_default()
    except Exception as exc:
        messagebox.showerror("Effacer layout", f"Impossible d'effacer le layout : {exc}")



def load_status(file_path: str | None = None):
    """
    Charge le statut complet depuis un fichier.
    Conserve la taille de la fenÃªtre et la position du sÃ©parateur (paned sash).
    DÃ©truit VRAIMENT les anciens onglets pour Ã©viter tout empilement de widgets/bindings.

    CompatibilitÃ© descendante :
    - Anciennes sauvegardes (sans Â« excluded_from_count Â») sont toujours lues.
    """
    if not file_path:
        file_path = filedialog.askopenfilename(
            title="Charger le statut",
            defaultextension=".pkl",
            filetypes=[("Pickle Files", "*.pkl"), ("All Files", "*.*")]
        )
        if not file_path:
            return

    global current_status_path
    current_status_path = file_path
    update_window_caption()

    prev_geom = root.geometry()
    prev_sashes = []
    for g_obj, _, _ in tabs_data:
        prev_sashes.append(getattr(g_obj, 'paned', None).sashpos(0) if hasattr(g_obj, 'paned') else None)

    import pickle
    try:
        with open(file_path, "rb") as f:
            all_week_status, saved_posts, saved_post_info, assignment_options = pickle.load(f)

        # Maj options globales
        update_work_posts(saved_posts)
        POST_INFO.clear()
        POST_INFO.update(saved_post_info)
        import Assignation
        assignment_len = len(assignment_options) if isinstance(assignment_options, (list, tuple)) else 0
        Assignation.ENABLE_MAX_WE_DAYS = False
        Assignation.MAX_WE_DAYS_PER_MONTH = None
        if assignment_len >= 8:
            (Assignation.ENABLE_DIFFERENT_POST_PER_DAY,
             Assignation.ENABLE_MAX_ASSIGNMENTS,
             Assignation.MAX_ASSIGNMENTS_PER_POST,
             Assignation.ENABLE_REPOS_SECURITE,
             loaded_pairs,
             Assignation.ENABLE_MAX_WE_DAYS,
             Assignation.MAX_WE_DAYS_PER_MONTH,
             Assignation.ENABLE_WEEKEND_BLOCKS) = assignment_options[:8]
        elif assignment_len >= 7:
            (Assignation.ENABLE_DIFFERENT_POST_PER_DAY,
             Assignation.ENABLE_MAX_ASSIGNMENTS,
             Assignation.MAX_ASSIGNMENTS_PER_POST,
             Assignation.ENABLE_REPOS_SECURITE,
             loaded_pairs,
             Assignation.ENABLE_MAX_WE_DAYS,
             Assignation.MAX_WE_DAYS_PER_MONTH) = assignment_options[:7]
        elif assignment_len >= 5:
            (Assignation.ENABLE_DIFFERENT_POST_PER_DAY,
             Assignation.ENABLE_MAX_ASSIGNMENTS,
             Assignation.MAX_ASSIGNMENTS_PER_POST,
             Assignation.ENABLE_REPOS_SECURITE,
             loaded_pairs) = assignment_options[:5]
        else:
            (Assignation.ENABLE_DIFFERENT_POST_PER_DAY,
             Assignation.ENABLE_MAX_ASSIGNMENTS,
             Assignation.MAX_ASSIGNMENTS_PER_POST,
             Assignation.ENABLE_REPOS_SECURITE) = assignment_options
            loaded_pairs = []

        Assignation.FORBIDDEN_POST_ASSOCIATIONS.clear()
        if loaded_pairs:
            sanitized = []
            for pair in loaded_pairs:
                try:
                    morning, afternoon = pair
                except Exception:
                    continue
                morning = str(morning)
                afternoon = str(afternoon)
                if morning in work_posts and afternoon in work_posts:
                    sanitized.append((morning, afternoon))
            Assignation.FORBIDDEN_POST_ASSOCIATIONS.update(sanitized)
        elif Assignation.ENABLE_DIFFERENT_POST_PER_DAY:
            auto_pairs = [(post, post) for post in work_posts]
            Assignation.FORBIDDEN_POST_ASSOCIATIONS.update(auto_pairs)

        different_post_var.set(Assignation.ENABLE_DIFFERENT_POST_PER_DAY)
        limitation_enabled_var.set(Assignation.ENABLE_MAX_ASSIGNMENTS)
        repos_securite_var.set(Assignation.ENABLE_REPOS_SECURITE)
        try:
            max_we_days_enabled_var.set(bool(getattr(Assignation, "ENABLE_MAX_WE_DAYS", False)))
            if getattr(Assignation, "MAX_WE_DAYS_PER_MONTH", None) is not None:
                max_we_days_value_var.set(int(Assignation.MAX_WE_DAYS_PER_MONTH))
        except Exception:
            pass
        _sync_max_we_menu_state()

        # --- NOUVEAU : prÃ©parer un hÃ©ritage des exclusions de la Mois 1 ---
        # Si les semaines >1 n'ont pas d'exclusions enregistrÃ©es, on leur appliquera
        # celles de la Mois 1 pour rester cohÃ©rent avec l'affichage que tu avais.
        first_week_excluded = None
        try:
            if all_week_status and len(all_week_status[0]) >= 6:
                first_week_excluded = all_week_status[0][5] or []
        except Exception:
            first_week_excluded = None
        # ---------------------------------------------------------------------

        # --- DÃ©truire proprement les anciens onglets -----------------------
        for tab_id in notebook.tabs():
            w = root.nametowidget(tab_id)
            notebook.forget(tab_id)
            try:
                w.destroy()
            except Exception:
                pass
        tabs_data.clear()
        root.update_idletasks()
        # -------------------------------------------------------------------

        #         # Recr?er les semaines et injecter les donn?es
        for idx, week_status in enumerate(all_week_status):
            # Compat descendante : 5 ?l?ments (ancienne sauvegarde) ou plus (nouvelle)
            table_data = cell_availability_data = constraints_data = schedule_data = week_label_text = excluded_cells = meta = None
            if len(week_status) >= 7:
                (table_data,
                 cell_availability_data,
                 constraints_data,
                 schedule_data,
                 week_label_text,
                 excluded_cells,
                 meta) = week_status
            elif len(week_status) == 6:
                (table_data,
                 cell_availability_data,
                 constraints_data,
                 schedule_data,
                 week_label_text,
                 excluded_cells) = week_status
            else:
                (table_data,
                 cell_availability_data,
                 constraints_data,
                 schedule_data,
                 week_label_text) = week_status
                excluded_cells = []  # pas d'exclusions dans les anciens fichiers

            # --- NOUVEAU : h?ritage depuis Mois 1 si vide pour cette semaine
            if idx > 0 and (not excluded_cells) and first_week_excluded:
                excluded_cells = list(first_week_excluded)
            # -----------------------------------------------------------------

            frame_for_week = tk.Frame(notebook)
            frame_for_week.pack(fill="both", expand=True)
            g, c, s = create_single_week(frame_for_week)
            tabs_data.append((g, c, s))
            notebook.add(frame_for_week, text=f"Mois {idx+1}")

            # Appliquer d'abord le mois/ann?e sauvegard?s pour restaurer weekends/feri?s
            if isinstance(meta, dict):
                meta_year = meta.get("year")
                meta_month = meta.get("month")
                if meta_year and meta_month:
                    try:
                        g.apply_month_selection(meta_year, meta_month)
                    except Exception:
                        pass

# Planning principal
            for i in range(min(len(table_data), len(g.table_entries))):
                for j in range(min(len(table_data[i]), len(g.table_entries[i]))):
                    cell = g.table_entries[i][j]
                    if cell:
                        cell.delete(0, "end")
                        val = table_data[i][j]
                        if val is None:
                            val = ""
                        cell.insert(0, val)

            # DisponibilitÃ©s
            g.cell_availability = cell_availability_data
            for (row, col), _available in cell_availability_data.items():
                if row < len(g.table_entries) and col < len(g.table_entries[row]):
                    g.update_cell(row, col)

            # RÃ©appliquer marquages/masquages si fournis
            if isinstance(meta, dict):
                try:
                    g.weekend_rows = set(meta.get("weekend_rows", []))
                    g.holiday_rows = set(meta.get("holiday_rows", []))
                    g.holiday_dates = set(meta.get("holiday_dates", []))
                    g.hidden_rows = set(meta.get("hidden_rows", []))
                    g.refresh_day_labels()
                except Exception:
                    pass

            # Tableau de contraintes
            for existing_row in c.rows:
                for widget in existing_row:
                    if hasattr(widget, "destroy"):
                        widget.destroy()
            c.rows.clear()
            for row_data in constraints_data:
                c.add_row()
                new_row = c.rows[-1]
                for idx2 in range(min(len(new_row), len(row_data))):
                    widget = new_row[idx2]
                    value  = row_data[idx2]
                    if isinstance(widget, tuple):
                        abs_state = ""
                        pds_state = 0
                        origin_state = "manual"
                        log_note = ""
                        if isinstance(value, (list, tuple)):
                            if len(value) >= 1:
                                abs_state = value[0]
                            if len(value) >= 2:
                                pds_state = value[1]
                            if len(value) >= 3 and value[2]:
                                origin_state = value[2]
                            if len(value) >= 4:
                                log_note = value[3]
                        else:
                            abs_state = value
                        toggle, pds_cb, pds_var = widget
                        # Synchroniser l'Ã©tat interne du bouton d'absence
                        if hasattr(toggle, "set_state"):
                            toggle.set_state(abs_state)
                        else:
                            toggle._var.set(abs_state)
                            toggle.config(text=abs_state)
                        try:
                            pds_var.set(int(pds_state))
                        except Exception:
                            pds_var.set(0)
                        pds_cb.config(bg=("red" if pds_var.get() == 1 else "SystemButtonFace"))
                        if hasattr(toggle, "set_origin"):
                            try:
                                toggle.set_origin(origin_state or "manual", log_text=log_note, notify=False)
                            except Exception:
                                pass
                        else:
                            try:
                                toggle.origin = origin_state
                                toggle.log_text = log_note
                            except Exception:
                                pass
                        try:
                            toggle._apply_origin_style()
                        except Exception:
                            pass
                    elif hasattr(widget, "_var") and not getattr(widget, "_is_row_action_button", False):
                        val_str = "" if value in (None, "Sélectionner") else str(value)
                        widget._var.set(val_str)
                        widget.config(text=val_str or "Sélectionner")
                    elif isinstance(widget, tk.Button):
                        if getattr(widget, "_is_row_action_button", False):
                            scope_val = str(value) if value not in (None, "+") else "all"
                            if hasattr(widget, "_var"):
                                try:
                                    widget._var.set(scope_val or "all")
                                except Exception:
                                    pass
                            widget.config(text="+")
                        else:
                            # Restaure aussi la variable sous-jacente si elle existe (préférences/absences)
                            val_str = "" if value in (None, "Sélectionner") else str(value)
                            if hasattr(widget, "var"):
                                widget.var.set(val_str)
                            if hasattr(widget, "_var"):
                                widget._var.set(val_str)
                            widget.config(text=val_str or "Sélectionner")
                    else:
                        widget.delete(0, "end")
                        widget.insert(0, value)
            
            # Rebinding menu contextuel si dispo
            if hasattr(c, "rebind_all_rows_context_menu"):
                c.rebind_all_rows_context_menu()
            
            # Horaires (labels)
            for i in range(min(len(schedule_data), len(g.table_labels))):
                for j in range(min(len(schedule_data[i]), len(g.table_labels[i]))):
                    lbl = g.table_labels[i][j]
                    if lbl and schedule_data[i][j] is not None:
                        text, bg, fg = schedule_data[i][j]
                        lbl.config(text=text, bg=bg, fg=fg)

            # Restaurer les cellules exclues du dÃ©compte (aprÃ¨s labels)
            try:
                g.excluded_from_count = set((int(r), int(c)) for (r, c) in excluded_cells)
            except Exception:
                g.excluded_from_count = set()

            # Appliquer le marquage visuel des tags horaires en fonction des exclusions
            g.refresh_exclusion_styles()

            # Label semaine + recolorations + dÃ©compte
            g.week_label.config(text=week_label_text)
            g.update_colors(None)
            s.update_counts()               # â tient compte des exclusions
            g.auto_resize_all_columns()

        # Restaurer la gÃ©omÃ©trie/sash
        for i, pos in enumerate(prev_sashes):
            try:
                g_obj, _, _ = tabs_data[i]
                if pos is not None:
                    g_obj.paned.sashpos(0, pos)
            except Exception:
                pass
        root.geometry(prev_geom)

        # Remettre les refs globales sur lâonglet actif
        notebook.bind("<<NotebookTabChanged>>", lambda e: globals().update({
            'gui': tabs_data[notebook.index(notebook.select())][0],
            'constraints_app': tabs_data[notebook.index(notebook.select())][1]
        }))
        if tabs_data:
            gui, constraints_app, _ = tabs_data[0]

        root.update_idletasks()
        messagebox.showinfo("Chargement", f"Statut chargé depuis {file_path}")

    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de charger le statut : {e}")






def _insert_names_into_constraints(names: list[str], table) -> tuple[int, int]:
    """
    Injecte une liste de noms dans le tableau de contraintes :
    - Remplit d'abord les lignes vides existantes.
    - Ajoute de nouvelles lignes si nécessaire.
    Renvoie (ajoutés, ignorés car doublons).
    """
    added = 0
    skipped = 0
    if table is None:
        return added, len(names)

    rows = getattr(table, "rows", [])
    try:
        existing = [row[0].get().strip() for row in rows if row and hasattr(row[0], "get")]
    except Exception:
        existing = []
    seen = {s.lower() for s in existing if s}

    def _set_initial(row_obj, name: str):
        try:
            entry = row_obj[0]
            entry.delete(0, "end")
            entry.insert(0, name)
        except Exception:
            pass

    empty_rows = [row for row in rows if row and hasattr(row[0], "get") and not row[0].get().strip()]

    for name in names:
        norm = name.lower()
        if norm in seen:
            skipped += 1
            continue
        target_row = None
        if empty_rows:
            target_row = empty_rows.pop(0)
        else:
            try:
                table.add_row()
                target_row = table.rows[-1]
            except Exception:
                skipped += 1
                continue
        _set_initial(target_row, name)
        seen.add(norm)
        added += 1

    return added, skipped


def insert_names_popup():
    """
    Ouvre une fenêtre pour coller des noms (depuis Excel ou texte) et les insère
    dans le tableau de contraintes. Remplit les lignes vides existantes puis en
    crée de nouvelles si besoin.
    """
    from tkinter import messagebox

    global constraints_app
    target_table = constraints_app
    if target_table is None:
        messagebox.showinfo("Insérer noms", "Aucun tableau de contraintes actif. Ouvrez un planning d'abord.")
        return

    parent = None
    try:
        parent = target_table.winfo_toplevel()
    except Exception:
        parent = None

    popup = tk.Toplevel(parent)
    popup.title("Insérer des noms")
    popup.configure(bg=APP_WINDOW_BG)
    popup.resizable(False, False)
    popup.transient(parent)
    popup.grab_set()

    instruction = tk.Label(
        popup,
        text="Collez les noms (séparés par retour à la ligne, virgule ou point-virgule) :",
        bg=APP_WINDOW_BG,
        fg="black",
        justify="left",
        wraplength=360,
        padx=8,
        pady=8,
    )
    instruction.pack(fill="x")

    text_box = tk.Text(popup, width=50, height=10, wrap="word", bg="white")
    text_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))
    text_box.focus_set()

    btns = ttk.Frame(popup, padding=(10, 0, 10, 10))
    btns.pack(fill="x")

    def _center_popup_over_widget(win, widget):
        try:
            win.update_idletasks()
            target = widget or win.master
            try:
                target = target.winfo_toplevel()
            except Exception:
                pass
            if target is None:
                return
            target.update_idletasks()
            pw, ph = win.winfo_width(), win.winfo_height()
            if pw <= 0 or ph <= 0:
                return
            try:
                wx, wy = target.winfo_rootx(), target.winfo_rooty()
                ww, wh = target.winfo_width(), target.winfo_height()
            except Exception:
                wx = wy = 0
                ww, wh = win.winfo_screenwidth(), win.winfo_screenheight()
            if ww <= 0 or wh <= 0:
                ww, wh = win.winfo_screenwidth(), win.winfo_screenheight()
                wx = wy = 0
            x = wx + (ww - pw) // 2
            y = wy + (wh - ph) // 2
            win.geometry(f"+{int(x)}+{int(y)}")
        except Exception:
            pass

    def _on_cancel():
        popup.destroy()

    def _on_ok():
        raw = text_box.get("1.0", "end")
        parts = re.split(r"[;,\n\r\t]+", raw)
        names = []
        for p in parts:
            name = p.strip()
            if name:
                names.append(name)
        if not names:
            messagebox.showinfo("Insérer noms", "Aucun nom détecté.")
            return
        added, skipped = _insert_names_into_constraints(names, target_table)
        msg = f"{added} nom(s) inséré(s)."
        if skipped:
            msg += f" {skipped} doublon(s) ignoré(s)."
        messagebox.showinfo("Insérer noms", msg)
        popup.destroy()

    ttk.Button(btns, text="Annuler", command=_on_cancel, style=RIBBON_BUTTON_STYLE).pack(side="right")
    ttk.Button(btns, text="Valider", command=_on_ok, style=RIBBON_BUTTON_STYLE).pack(side="right", padx=(0, 6))

    _center_popup_over_widget(popup, target_table)
    popup.wait_window(popup)


def import_layout():
    """
    Importe uniquement la mise en page du planning principal depuis un fichier .pkl :
    - postes de travail (noms), couleurs et horaires (labels)
    - disponibilitÃ©s des cellules
    Sans toucher au contenu du tableau (noms) ni au tableau de contraintes.
    """
    from tkinter import filedialog, messagebox
    import pickle

    # SÃ©lection du fichier
    file_path = filedialog.askopenfilename(
        title="Import Layout",
        defaultextension=".pkl",
        filetypes=[("Pickle Files", "*.pkl"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    try:
        # Chargement du pickle
        with open(file_path, "rb") as f:
            loaded = pickle.load(f)

        # Format attendu : (all_week_status, saved_posts, saved_post_info, options)
        # all_week_status[i] : (table_data, cell_av, constraints, schedule, week_label[, excluded])
        if not isinstance(loaded, tuple) or len(loaded) < 3:
            raise ValueError("Fichier incompatible (structure inattendue).")

        all_week_status, saved_posts, saved_post_info = loaded[0], loaded[1], loaded[2]

        # 1) Mettre Ã  jour les postes et info de postes AVANT de redessiner
        from Full_GUI import update_work_posts, POST_INFO
        update_work_posts(saved_posts)
        POST_INFO.clear()
        POST_INFO.update(saved_post_info)

        # 2) RÃ©cupÃ©rer 'schedule_data' et 'cell_availability' du premier onglet
        week0 = all_week_status[0]
        if len(week0) >= 6:
            _, cell_availability_data, _, schedule_data, _week_label, _excluded = week0
        else:
            # CompatibilitÃ© descendante (anciens fichiers Ã  5 items)
            _, cell_availability_data, _, schedule_data, _week_label = week0

        # 3) Redessiner proprement la grille selon les nouveaux postes
        gui.redraw_widgets(preserve_content=False)

        # 4) RÃ©tablir les horaires (labels) avec bornes de sÃ©curitÃ©
        #    â on borne sur le plus petit des deux (source & grille courante)
        max_i = min(len(schedule_data), len(gui.table_labels))
        for i in range(max_i):
            max_j = min(len(schedule_data[i]), len(gui.table_labels[i]))
            for j in range(max_j):
                cell = schedule_data[i][j]
                lbl = gui.table_labels[i][j]
                if lbl and cell is not None:
                    text, bg, fg = cell
                    lbl.config(text=text, bg=bg, fg=fg)

        # 5) RecrÃ©er l'Ã©tat des (in)disponibilitÃ©s de faÃ§on sÃ»re
        #    â on filtre les paires (r,c) hors bornes, puis on applique
        filtered_avail = {}
        row_count = len(gui.table_entries)
        for (r, c), avail in dict(cell_availability_data).items():
            if 0 <= r < row_count and 0 <= c < len(gui.table_entries[r]):
                filtered_avail[(r, c)] = bool(avail)

        gui.cell_availability = filtered_avail
        for (r, c), _avail in filtered_avail.items():
            gui.update_cell(r, c)

        # 6) Finitions visuelles
        gui.update_colors(None)
        gui.auto_resize_all_columns()

        messagebox.showinfo("Import Layout", "Mise en page importée avec succès.")

    except Exception as e:
        messagebox.showerror("Import Layout", f"Erreur lors de l'import : {e}")



def import_absences():
    """
    Handler du menu 'Import Absences...'.
    Import paresseux pour Ã©viter les erreurs si Import_absence.py n'existe pas encore.
    """
    try:
        from Import_absence import import_absences_from_excel
    except Exception as e:
        messagebox.showerror(
            "Import Absences",
            f"Le module 'Import_absence.py' est introuvable ou invalide.\n\nDÃ©tail : {e}"
        )
        return
    try:
        # 'root', 'notebook' et 'tabs_data' sont globaux dÃ©clarÃ©s plus bas dans ce fichier.
        import_absences_from_excel(root, notebook, tabs_data)
    except Exception as e:
        messagebox.showerror("Import Absences", f"échec de l'import des absences : {e}")

def check_cross_conflicts():
    """
    Handler : 'VÃ©rifier Conflits inter-plannings (.pkl)'
    """
    try:
        from Import_absence import check_cross_planning_conflicts_from_pkl
    except Exception as e:
        messagebox.showerror(
            "Conflits inter-plannings",
            f"Le module 'Import_absence.py' est introuvable ou invalide.\n\nDÃ©tail : {e}"
        )
        return
    try:
        check_cross_planning_conflicts_from_pkl(root, notebook, tabs_data)
    except Exception as e:
        messagebox.showerror("Conflits inter-plannings", f"échec de l'analyse : {e}")


def import_conflicts():
    """
    Handler du menu 'Import Conflits (.pkl)...'
    """
    try:
        from Import_absence import import_conflicts_from_pkl
    except Exception as e:
        messagebox.showerror(
            "Import Conflits",
            f"Le module 'Import_absence.py' est introuvable ou invalide.\n\nDÃ©tail : {e}"
        )
        return
    try:
        # 'root', 'notebook' et 'tabs_data' sont globaux (comme pour import_absences)
        import_conflicts_from_pkl(root, notebook, tabs_data)
    except Exception as e:
        messagebox.showerror("Import Conflits", f"échec de l'import des conflits : {e}")


# Programme principal
if __name__ == '__main__':
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    import Assignation
    from Assignation import assigner_initiales
    from tkinter import simpledialog

    # Nous conservons les variables globales pour les menus existants.
    # Elles pointeront vers l'onglet (semaine) en cours.
    gui = None
    constraints_app = None

    # --- CrÃ©ation de la fenÃªtre principale ---
    root = tk.Tk()
    style = setup_modern_styles(root)

    title_label = tk.Label(
        root,
        text=APP_TITLE,
        font=(APP_FONT_FAMILY, 20, "bold"),
        bg=APP_WINDOW_BG,
        fg=APP_PRIMARY_DARK,
        anchor="center",
    )
    title_label.pack(side="top", fill="x", padx=16, pady=(12, 6))
    update_window_caption()

    ttk.Separator(root, orient="horizontal").pack(fill="x", padx=12, pady=(0, 12))

    menu_bar = tk.Menu(root)

    # Variables globales pour les menus
    live_conflict_var = tk.BooleanVar(value=True)

    # Menu File
    file_menu = tk.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Charger Planning",    command=load_status)
    file_menu.add_command(label="Ouvrir une nouvelle fenêtre", command=lambda: open_new_window(root, notebook, tabs_data))
    file_menu.add_checkbutton(
        label="Surveiller conflits inter-fenêtres (live)",
        variable=live_conflict_var,
        command=trigger_live_conflict_check,
    )
    file_menu.add_command(label="Enregistrer",         command=quick_save_status)
    file_menu.add_command(label="Enregistrer Sous",   command=save_status)
    file_menu.add_command(label="Localiser sauvegarde automatique", command=open_autosave_folder)
    file_menu.add_separator()
    file_menu.add_command(label="Effacer layout", command=reset_layout_from_menu)
    menu_bar.add_cascade(label="File", menu=file_menu)

    # Nouveau menu Imports (inchangÃ©)
    import_menu = tk.Menu(menu_bar, tearoff=0)
    import_menu.add_command(label="Insérer noms",                          command=insert_names_popup)
    import_menu.add_command(label="Import Absences",                       command=import_absences)
    import_menu.add_command(label="Import Conflits (.pkl)",                command=import_conflicts)
    import_menu.add_command(label="Vérifier Conflits inter-plannings (.pkl)", command=check_cross_conflicts)
    menu_bar.add_cascade(label="Imports", menu=import_menu)

    # >>> Nouveau menu Export (dÃ©placement de "Export to Excel" ici)
    export_menu = tk.Menu(menu_bar, tearoff=0)
    export_menu.add_command(
        label="Export to Excel",
        command=lambda: export_to_excel_external(root, tabs_data, days, work_posts, POST_INFO)
    )
    export_menu.add_command(
        label="Export combiné (.pkl)",
        command=lambda: export_combined_to_excel_external(root, tabs_data, days, work_posts, POST_INFO)
    )

    menu_bar.add_cascade(label="Export", menu=export_menu)
    # <<< Fin du nouveau menu Export

    # --- Menu Setup (vidé pour la nouvelle configuration) ---
    setup_menu = tk.Menu(menu_bar, tearoff=0)
    different_post_var = tk.BooleanVar(value=Assignation.ENABLE_DIFFERENT_POST_PER_DAY)
    limitation_enabled_var = tk.BooleanVar(value=Assignation.ENABLE_MAX_ASSIGNMENTS)
    repos_securite_var = tk.BooleanVar(value=Assignation.ENABLE_REPOS_SECURITE)
    weekend_block_var = tk.BooleanVar(value=getattr(Assignation, "ENABLE_WEEKEND_BLOCKS", False))
    max_we_days_enabled_var = tk.BooleanVar(value=getattr(Assignation, "ENABLE_MAX_WE_DAYS", False))
    _initial_we_limit = getattr(Assignation, "MAX_WE_DAYS_PER_MONTH", None)
    max_we_days_value_var = tk.IntVar(
        value=_initial_we_limit if _initial_we_limit is not None else 4
    )
    set_max_we_entry_idx = None

    def _sync_max_we_menu_state():
        try:
            if set_max_we_entry_idx is not None:
                setup_menu.entryconfig(
                    set_max_we_entry_idx,
                    state=("normal" if max_we_days_enabled_var.get() else "disabled"),
                )
        except Exception:
            pass
        try:
            Assignation.ENABLE_MAX_WE_DAYS = bool(max_we_days_enabled_var.get())
            Assignation.MAX_WE_DAYS_PER_MONTH = int(max_we_days_value_var.get())
        except Exception:
            Assignation.ENABLE_MAX_WE_DAYS = False
            Assignation.MAX_WE_DAYS_PER_MONTH = None

    def _on_toggle_max_we_days():
        _sync_max_we_menu_state()

    def _on_toggle_weekend_block():
        try:
            Assignation.ENABLE_WEEKEND_BLOCKS = bool(weekend_block_var.get())
        except Exception:
            Assignation.ENABLE_WEEKEND_BLOCKS = False

    def open_max_we_days_dialog():
        def _center_popup(popup, widget):
            try:
                popup.update_idletasks()
                target = widget or popup.master
                try:
                    if target is not None:
                        target = target.winfo_toplevel()
                except Exception:
                    pass
                if target is None:
                    return
                target.update_idletasks()
                popup_w, popup_h = popup.winfo_width(), popup.winfo_height()
                if popup_w <= 0 or popup_h <= 0:
                    return
                try:
                    wx, wy = target.winfo_rootx(), target.winfo_rooty()
                    ww, wh = target.winfo_width(), target.winfo_height()
                except Exception:
                    wx = wy = 0
                    ww, wh = popup.winfo_screenwidth(), popup.winfo_screenheight()
                if ww <= 0 or wh <= 0:
                    ww, wh = popup.winfo_screenwidth(), popup.winfo_screenheight()
                    wx = wy = 0
                x = wx + (ww - popup_w) // 2
                y = wy + (wh - popup_h) // 2
                popup.geometry(f"+{int(x)}+{int(y)}")
            except Exception:
                pass

        try:
            current_val = int(max_we_days_value_var.get())
        except Exception:
            current_val = 4

        popup = tk.Toplevel(root)
        popup.title("Nombre Maximal de Jours de Weekend et Fériés")
        popup.resizable(False, False)
        popup.transient(root)
        popup.grab_set()

        frame = ttk.Frame(popup, padding=12)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="Fixer la limite par personne (0 = aucune garde autorisée) :",
            wraplength=320,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(0, 8))

        val_var = tk.IntVar(value=current_val)
        spin = tk.Spinbox(
            frame,
            from_=0,
            to=31,
            width=5,
            textvariable=val_var,
            justify="center",
        )
        spin.pack(pady=(0, 10))
        spin.focus_set()

        btns = ttk.Frame(frame)
        btns.pack(fill="x")

        result = {"val": None}

        def _confirm():
            try:
                result["val"] = max(0, min(31, int(val_var.get())))
            except Exception:
                result["val"] = None
            popup.destroy()

        def _cancel():
            result["val"] = None
            popup.destroy()

        ttk.Button(btns, text="OK", width=10, command=_confirm).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Annuler", width=10, command=_cancel).pack(side="right")

        popup.bind("<Return>", lambda e: _confirm())
        popup.bind("<Escape>", lambda e: _cancel())
        _center_popup(popup, root)
        popup.wait_window()

        if result["val"] is None:
            return
        max_we_days_value_var.set(int(result["val"]))
        _sync_max_we_menu_state()

    setup_menu.add_checkbutton(
        label="Nombre maximal de jours de Weekend et Fériés",
        variable=max_we_days_enabled_var,
        command=_on_toggle_max_we_days,
    )
    setup_menu.add_checkbutton(
        label="Affecter les week-ends en bloc (ven-sam-dim)",
        variable=weekend_block_var,
        command=_on_toggle_weekend_block,
    )
    setup_menu.add_command(
        label="Définir la limite (jours)",
        command=open_max_we_days_dialog,
    )
    set_max_we_entry_idx = setup_menu.index("end")
    _sync_max_we_menu_state()
    # Fonctions historiques retirées ; menu laissé vide pour éviter tout usage obsolète.
    menu_bar.add_cascade(label="Setup", menu=setup_menu)
    menu_bar.add_command(label="Remplacer nom", command=open_replace_name_dialog)

    # Application de la barre de menus Ã  la fenÃªtre principale
    root.config(menu=menu_bar)

    # Notebook principal inspirÃ© du ruban Office
    style.configure(NOTEBOOK_STYLE, background=APP_WINDOW_BG)
    notebook = ttk.Notebook(root, style=NOTEBOOK_STYLE)
    notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))


    # Liste pour stocker (gui, constraints_app, shift_count_table) de chaque onglet
    tabs_data = []

    def create_single_week(parent_frame, include_constraints: bool = True, enable_shift_counts: bool = True):
        paned = ttk.PanedWindow(parent_frame, orient=tk.VERTICAL)
        paned.pack(fill="both", expand=True)

        top_frame = tk.Frame(paned)
        bottom_frame = tk.Frame(paned)

        paned.add(top_frame, weight=3)
        if include_constraints:
            paned.add(bottom_frame, weight=1)
        else:
            paned.add(bottom_frame, weight=0)

        # --- On recrÃ©e le "planning" (GUI) ---
        content_frame = tk.Frame(top_frame)
        content_frame.pack(fill="both", expand=True)

        scroll_frame = ScrollableFrame(content_frame)
        scroll_frame.pack(side="left", fill="both", expand=True)

        gui_local = GUI(scroll_frame.inner)
        gui_local.scroll_canvas = scroll_frame.canvas   # pour scroll_to_cell(...)
        gui_local.scroll_inner  = scroll_frame.inner

        gui_local.grid(row=0, column=0, sticky="nsew")
        # On garde la rÃ©fÃ©rence au PanedWindow pour restaurer sashpos ensuite
        gui_local.paned = paned

        shift_count_frame = None
        collapsed_stub = None
        if enable_shift_counts:
            shift_count_frame = tk.Frame(content_frame, width=220, bg=APP_WINDOW_BG)
            shift_count_frame.pack(side="right", fill="y")

            # Stub button shown when the counts panel is collapsed
            collapsed_stub = ttk.Button(
                content_frame,
                text="Afficher comptes",
                command=lambda: toggle_shift_counts(True),
                style=RIBBON_BUTTON_STYLE,
                width=18,
            )
            collapsed_stub.pack_forget()

        gui_local.verification_enabled = tk.BooleanVar(value=False)

        control_frame = ttk.Frame(shift_count_frame, style=RIBBON_FRAME_STYLE, padding=(12, 10)) if shift_count_frame else None
        if control_frame is not None:
            control_frame.pack(side="top", fill="x", padx=8, pady=8)

            verification_box = ttk.Frame(control_frame, style=RIBBON_FRAME_STYLE)
            verification_box.pack(side="left", anchor="n")

            verification_chk = ttk.Checkbutton(
                verification_box,
                text="Vérification",
                variable=gui_local.verification_enabled,
                command=lambda: (
                    gui_local.apply_verification()
                    if gui_local.verification_enabled.get()
                    else gui_local.schedule_update_colors()
                ),
                style=RIBBON_CHECK_STYLE,
            )
            verification_chk.pack(anchor="w")

            ttk.Separator(control_frame, orient="vertical").pack(side="left", fill="y", padx=8, pady=4)

        action_box = ttk.Frame(control_frame, style=RIBBON_FRAME_STYLE) if control_frame else None
        if action_box is not None:
            action_box.pack(side="left", expand=True, fill="both")

        if action_box is not None:
            for col in (0, 1, 2):
                action_box.columnconfigure(col, weight=1)

        if action_box is not None:
            undo_btn = ttk.Button(
                action_box,
                text="Annuler assignation",
                command=gui_local.undo_last_change,
                style=RIBBON_BUTTON_STYLE,
            )
            undo_btn.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

            choose_month_btn = ttk.Button(
                action_box,
                text="Choisir mois",
                command=gui_local.choose_month,
                style=RIBBON_BUTTON_STYLE,
            )
            choose_month_btn.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

            clear_btn = ttk.Button(
                action_box,
                text="Effacer tout",
                command=gui_local.clear_schedule,
                style=RIBBON_BUTTON_STYLE,
            )
            clear_btn.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

            def trigger_assignation():
                # Snapshot before running auto-assign so the undo button can restore it.
                before_state = []
                for row in gui_local.table_entries:
                    row_snapshot = []
                    for cell in row:
                        if cell is None:
                            row_snapshot.append(None)
                            continue
                        try:
                            row_snapshot.append(cell.get())
                        except Exception:
                            row_snapshot.append("")
                    before_state.append(row_snapshot)
                gui_local.last_assignment_state = before_state
                current_y = getattr(gui_local, "current_year", None)
                current_m = getattr(gui_local, "current_month", None)
                assigner_initiales(constraints_app_local, gui_local)
                gui_local.schedule_update_colors()
                # Restaure l'affichage (week-ends/jours fériés) et recalcule les comptes
                if current_y and current_m:
                    try:
                        gui_local.apply_month_selection(current_y, current_m)
                    except Exception:
                        pass
                if hasattr(gui_local, "shift_count_table"):
                    try:
                        gui_local.shift_count_table.update_counts()
                    except Exception:
                        pass
                # Push the diff to the undo stack so "Annuler assignation" works.
                try:
                    changes = []
                    for r_idx, row in enumerate(gui_local.table_entries):
                        if r_idx >= len(before_state):
                            break
                        before_row = before_state[r_idx]
                        for c_idx, cell in enumerate(row):
                            if c_idx >= len(before_row):
                                break
                            if cell is None:
                                continue
                            try:
                                new_val = cell.get()
                            except Exception:
                                continue
                            old_val = before_row[c_idx]
                            if old_val is None:
                                old_val = ""
                            if new_val != old_val:
                                changes.append((r_idx, c_idx, old_val))
                    if changes:
                        gui_local.cell_edit_undo_stack.append(changes)
                except Exception:
                    pass

            assign_btn = ttk.Button(
                action_box,
                text="Assignation",
                command=trigger_assignation,
                style=RIBBON_ACCENT_BUTTON_STYLE,
            )
            assign_btn.grid(row=1, column=0, columnspan=2, padx=4, pady=(6, 4), sticky="ew")

            add_post_btn = ttk.Button(
                action_box,
                text="Ajouter poste",
                command=gui_local.add_work_post,
                style=RIBBON_BUTTON_STYLE,
            )
            add_post_btn.grid(row=1, column=2, padx=4, pady=(6, 4), sticky="ew")

        shift_count_table = None
        gui_local.shift_counts_visible = False
        if shift_count_frame is not None:
            shift_count_table = ShiftCountTable(shift_count_frame, gui_local)
            shift_count_table.pack(fill="both", expand=True, padx=8, pady=(0, 8))
            gui_local.shift_count_table = shift_count_table
            gui_local.shift_counts_visible = True

            def toggle_shift_counts(show: bool | None = None):
                current_visible = getattr(gui_local, "shift_counts_visible", True)
                if show is None:
                    show = not current_visible
                if show:
                    collapsed_stub.pack_forget()
                    shift_count_frame.pack(side="right", fill="y")
                    gui_local.shift_counts_visible = True
                    try:
                        shift_count_table.update_counts()
                    except Exception:
                        pass
                else:
                    shift_count_frame.pack_forget()
                    collapsed_stub.pack(side="right", fill="y", padx=(4, 0), pady=8)
                    gui_local.shift_counts_visible = False

            hide_counts_btn = ttk.Button(
                control_frame,
                text="Masquer comptes",
                command=lambda: toggle_shift_counts(False),
                style=RIBBON_BUTTON_STYLE,
            )
            hide_counts_btn.pack(side="right", padx=(6, 0))

        # --- On recrÃ©e la partie Contraintes ---
        constraints_app_local = None
        if include_constraints:
            constraints_app_local = ConstraintsTable(bottom_frame, work_posts=work_posts, planning_gui=gui_local)
            constraints_app_local.grid(row=0, column=0, sticky="nsew")
            bottom_frame.grid_rowconfigure(0, weight=1)
            bottom_frame.grid_columnconfigure(0, weight=1)
            constraints_app_local.set_change_callback = lambda cb=None: None
            gui_local.constraints_app = constraints_app_local
        else:
            gui_local.constraints_app = None
            bottom_frame.grid_rowconfigure(0, weight=1)
            bottom_frame.grid_columnconfigure(0, weight=1)

        return gui_local, constraints_app_local, shift_count_table

    def open_new_window(main_root, main_notebook, main_tabs_data):
        """Ouvre une fenêtre secondaire avec un planning indépendant et suivi live des conflits."""
        global current_status_path
        if not current_status_path:
            messagebox.showinfo("Fenêtre secondaire", "Charge d'abord un planning dans la fenêtre principale.")
            return
        # Si une fenêtre secondaire existe déjà, on la remet en avant au lieu d'en recréer une
        for ctx in list(_WINDOW_CONTEXTS):
            if ctx.get("is_primary"):
                continue
            win = ctx.get("root")
            try:
                if win and win.winfo_exists():
                    try:
                        win.deiconify()
                        win.lift()
                        win.focus_force()
                    except Exception:
                        pass
                    return
            except Exception:
                unregister_window_context(ctx)

        # Demande le fichier avant de créer la fenêtre
        default_path = filedialog.askopenfilename(
            title="Charger planning (fenêtre secondaire)",
            defaultextension=".pkl",
            filetypes=[("Pickle Files", "*.pkl"), ("Tous fichiers", "*.*")],
        )
        if not default_path:
            return

        new_win = tk.Toplevel(main_root)
        new_win.title(f"{APP_TITLE} - Fenêtre secondaire")
        new_win.configure(bg=APP_WINDOW_BG)

        title_var = tk.StringVar(value=f"{APP_TITLE} (fenêtre 2)")
        title = tk.Label(
            new_win,
            textvariable=title_var,
            font=(APP_FONT_FAMILY, 16, "bold"),
            bg=APP_WINDOW_BG,
            fg=APP_PRIMARY_DARK,
        )
        title.pack(side="top", fill="x", padx=12, pady=(10, 4))
        ttk.Separator(new_win, orient="horizontal").pack(fill="x", padx=12, pady=(0, 10))

        notebook2 = ttk.Notebook(new_win, style=NOTEBOOK_STYLE)
        notebook2.pack(fill="both", expand=True, padx=12, pady=12)
        tabs_data2 = []
        secondary_current_path = None

        def remove_stray_verification_widgets():
            try:
                wb = globals().get("week_buttons_frame")
                if wb:
                    for child in list(wb.winfo_children()):
                        try:
                            if isinstance(child, (tk.Checkbutton, ttk.Checkbutton)) and "Vérification" in str(child.cget("text")):
                                child.destroy()
                        except Exception:
                            continue
            except Exception:
                pass

        def load_secondary_status(file_path=None):
            if not file_path:
                file_path = filedialog.askopenfilename(
                    title="Charger planning (fenêtre secondaire)",
                    defaultextension=".pkl",
                    filetypes=[("Pickle Files", "*.pkl"), ("Tous fichiers", "*.*")],
                )
                if not file_path:
                    return
            import pickle
            try:
                with open(file_path, "rb") as f:
                    loaded = pickle.load(f)
            except Exception as e:
                messagebox.showerror("Charger planning", f"Impossible de lire le fichier : {e}")
                return

            if not isinstance(loaded, tuple) or len(loaded) < 3:
                messagebox.showerror("Charger planning", "Fichier incompatible (structure inattendue).")
                return

            all_week_status, saved_posts, saved_post_info = loaded[0], loaded[1], loaded[2]
            nonlocal secondary_current_path
            secondary_current_path = file_path
            try:
                title_name = os.path.basename(file_path)
                title_var.set(title_name)
                new_win.title(title_name)
            except Exception:
                pass

            # Efface les onglets existants
            for tab_id in notebook2.tabs():
                w = new_win.nametowidget(tab_id)
                notebook2.forget(tab_id)
                try:
                    w.destroy()
                except Exception:
                    pass
            tabs_data2.clear()
            new_win.update_idletasks()

            # Construire chaque semaine avec les postes du fichier (sans toucher aux postes globaux)
            for idx, wk in enumerate(all_week_status):
                if len(wk) >= 6:
                    (table_data,
                     cell_availability_data,
                     _constraints_data,
                     schedule_data,
                     week_label_text,
                     excluded_cells) = wk
                else:
                    (table_data,
                     cell_availability_data,
                     _constraints_data,
                     schedule_data,
                     week_label_text) = wk
                    excluded_cells = []

                # Crée une semaine dimensionnée selon les postes du fichier
                frame = tk.Frame(notebook2)
                frame.pack(fill="both", expand=True)

                # Crée la GUI secondaire en dimensionnant sur les postes du fichier sans altérer le principal
                from Full_GUI import work_posts as _wp, POST_INFO as _pi
                backup_posts_obj = _wp
                backup_info_obj = _pi
                was_paused = _LIVE_CONFLICT_PAUSED
                try:
                    if not was_paused:
                        pause_live_conflict_check()
                    # Ne plus modifier le même objet partagé : on remplace temporairement les globals
                    globals()["work_posts"] = list(saved_posts)
                    globals()["POST_INFO"] = dict(saved_post_info)
                    g_new, c_new, s_new = create_single_week(frame, include_constraints=False, enable_shift_counts=False)
                    g_new.local_work_posts = list(saved_posts)
                    g_new.local_post_info = dict(saved_post_info)
                finally:
                    globals()["work_posts"] = backup_posts_obj
                    globals()["POST_INFO"] = backup_info_obj
                    if not was_paused:
                        resume_live_conflict_check()

                tabs_data2.append((g_new, c_new, s_new))
                notebook2.add(frame, text=week_label_text or f"Mois {idx+1}")

                for i in range(min(len(table_data), len(g_new.table_entries))):
                    for j in range(min(len(table_data[i]), len(g_new.table_entries[i]))):
                        cell = g_new.table_entries[i][j]
                        cell.delete(0, "end")
                        val = table_data[i][j]
                        if val is None:
                            val = ""
                        cell.insert(0, val)

                g_new.cell_availability = dict(cell_availability_data)
                for (row, col), _available in cell_availability_data.items():
                    if row < len(g_new.table_entries) and col < len(g_new.table_entries[row]):
                        g_new.update_cell(row, col)

                try:
                    g_new.loaded_constraints_data = list(_constraints_data)
                except Exception:
                    g_new.loaded_constraints_data = []

                for i in range(min(len(schedule_data), len(g_new.table_labels))):
                    for j in range(min(len(schedule_data[i]), len(g_new.table_labels[i]))):
                        lbl = g_new.table_labels[i][j]
                        if lbl and schedule_data[i][j] is not None:
                            try:
                                text, bg, fg = schedule_data[i][j]
                            except Exception:
                                continue
                            lbl.config(text=text, bg=bg, fg=fg)

                g_new.week_label.config(text=week_label_text or f"Mois {idx+1}")
                g_new.update_colors(None)
                g_new.auto_resize_all_columns()

            try:
                notebook2.select(0)
            except Exception:
                pass
            try:
                trigger_live_conflict_check()
            except Exception:
                pass
            remove_stray_verification_widgets()

        # Toolbar for the secondary window
        toolbar = ttk.Frame(new_win, padding=(12, 6))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Charger planning (fenêtre 2)", command=load_secondary_status,
                   style=RIBBON_BUTTON_STYLE).pack(side="left")

        def swap_plannings():
            if not secondary_current_path or not current_status_path:
                messagebox.showinfo("Intervertir", "Charge un planning principal et secondaire avant d'intervertir.")
                return
            # Prévenir la perte d'éventuelles modifications non sauvegardées
            save_choice = messagebox.askyesnocancel(
                "Intervertir",
                "L'interversion recharge les deux plannings depuis leurs fichiers.\n"
                "Les modifications non sauvegardées dans la fenêtre principale seront perdues.\n"
                "Voulez-vous enregistrer avant de continuer ?"
            )
            if save_choice is None:
                return  # annulé
            if save_choice:
                try:
                    quick_save_status()
                except Exception as exc:
                    messagebox.showerror("Intervertir", f"Échec de la sauvegarde avant interversion : {exc}")
                    return
            try:
                pause_live_conflict_check()
                main_path = current_status_path
                other_path = secondary_current_path
                load_status(other_path)
                load_secondary_status(main_path)
                remove_stray_verification_widgets()
            except Exception as e:
                messagebox.showerror("Intervertir", f"Échec de l'interversion : {e}")
            finally:
                resume_live_conflict_check()

        ttk.Button(toolbar, text="Intervertir", command=swap_plannings,
                   style=RIBBON_BUTTON_STYLE).pack(side="left", padx=(8, 0))

        ctx = {
            "root": new_win,
            "notebook": notebook2,
            "tabs_data": tabs_data2,
            "is_primary": False,
        }
        register_window_context(ctx)

        def on_close():
            unregister_window_context(ctx)
            try:
                new_win.destroy()
            except Exception:
                pass

        new_win.protocol("WM_DELETE_WINDOW", on_close)
        try:
            new_win.after(100, trigger_live_conflict_check)
        except Exception:
            pass

        # Charger immédiatement le fichier sélectionné avant ouverture
        try:
            load_secondary_status(default_path)
        except Exception:
            pass

    # Fonction pour ajouter dynamiquement une nouvelle semaine

    
    def add_new_week():
        """Ajoute un nouvel onglet mois, vide, cal? sur le mois suivant le 1er onglet."""
        import tkinter as tk
        from tkinter import messagebox

        if not tabs_data:
            messagebox.showerror("Ajouter un mois", "Aucun onglet n'est disponible comme base.")
            return

        base_gui = tabs_data[0][0]
        base_year = getattr(base_gui, "current_year", date.today().year)
        base_month = getattr(base_gui, "current_month", date.today().month)

        offset = len(tabs_data)
        target_idx = (base_month - 1) + offset
        target_year = base_year + (target_idx // 12)
        target_month = (target_idx % 12) + 1

        frame_for_week = tk.Frame(notebook)
        frame_for_week.pack(fill="both", expand=True)
        new_gui, new_constraints, new_shift_table = create_single_week(frame_for_week)

        try:
            base_sash = None
            if hasattr(tabs_data[0][0], "paned"):
                base_sash = tabs_data[0][0].paned.sashpos(0)
            root.update_idletasks()
            if base_sash is not None and hasattr(new_gui, "paned"):
                new_gui.paned.sashpos(0, base_sash)
        except Exception:
            pass

        try:
            new_gui.apply_month_selection(target_year, target_month)
        except Exception:
            new_gui.schedule_update_colors()

        # Recopie du tableau de contraintes depuis le 1er onglet
        try:
            if new_constraints is not None and tabs_data and len(tabs_data[0]) >= 2:
                source_constraints = tabs_data[0][1]
                if source_constraints is not None and hasattr(source_constraints, "get_rows_data"):
                    rows_data = source_constraints.get_rows_data()
                    if hasattr(new_constraints, "set_rows_data"):
                        new_constraints.set_rows_data(rows_data)
        except Exception:
            pass

        new_gui.update_colors(None)
        try:
            new_gui.update_idletasks()
        except Exception:
            pass
        new_gui.auto_resize_all_columns()

        tabs_data.append((new_gui, new_constraints, new_shift_table))
        notebook.add(frame_for_week, text=f"Mois {len(tabs_data)}")
        notebook.select(len(tabs_data) - 1)
        renumber_week_tabs()


    def remove_current_week():
            global gui, constraints_app
    
            # Onglet courant
            try:
                current_index = notebook.index(notebook.select())
            except Exception:
                messagebox.showerror("Erreur", "Aucun onglet sélectionné.")
                return
    
            # Toujours garder au moins 1 semaine
            if len(tabs_data) <= 1:
                messagebox.showinfo("Info", "Impossible de supprimer, il doit rester au moins un onglet.")
                return
    
            # RÃ©cupÃ©rer l'identifiant et le frame AVANT de l'oublier
            try:
                tab_id = notebook.select()  # identifiant style '.!notebook.!frame'
                tab_widget = root.nametowidget(tab_id)
            except Exception:
                tab_id = None
                tab_widget = None
    
            # 1) Retirer du Notebook
            try:
                if tab_id is not None:
                    notebook.forget(tab_id)
                else:
                    notebook.forget(current_index)
            except Exception:
                pass
    
            # 2) DÃ©truire physiquement le frame pour libÃ©rer la mÃ©moire
            try:
                if tab_widget is not None:
                    tab_widget.destroy()
            except Exception:
                pass
    
            # 3) Retirer les rÃ©fÃ©rences Python
            try:
                tabs_data.pop(current_index)
            except Exception:
                pass
    
            # 4) Sélectionner un onglet valide (le prÃ©cÃ©dent si possible)
            try:
                new_index = min(current_index, len(tabs_data) - 1)
                notebook.select(new_index)
            except Exception:
                # au cas oÃ¹, on force lâindex 0 si la sÃ©lection Ã©choue
                try:
                    notebook.select(0)
                    new_index = 0
                except Exception:
                    # Situation anormale: on stoppe proprement
                    messagebox.showerror("Erreur", "Impossible de sélectionner un onglet valide aprés suppression.")
                    return
    
            # 5) Mettre Ã  jour les pointeurs actifs (important pour raccourcis/menus)
            try:
                gui, constraints_app, _ = tabs_data[new_index]
            except Exception:
                # Si quelque chose ne va pas, on informe clairement
                messagebox.showerror("Erreur", "échec de l'affectation du tableau principal aprés suppression.")
                return
    
            # 6) Laisser Tk faire le mÃ©nage des tailles/geometry
            try:
                root.update_idletasks()
            except Exception:
                pass
            renumber_week_tabs()
    
    
    
    # On crÃ©e 1 onglet par dÃ©faut
    frame_for_week = tk.Frame(notebook)
    frame_for_week.pack(fill="both", expand=True)

    g, c, s = create_single_week(frame_for_week)
    tabs_data.append((g, c, s))

    notebook.add(frame_for_week, text="Mois 1")

    main_window_ctx = {
        "root": root,
        "notebook": notebook,
        "tabs_data": tabs_data,
        "is_primary": True,
    }
    register_window_context(main_window_ctx)


    # Petit callback pour que les menus existants (File, Setup) pointent toujours
    # vers le gui/constraints_app de l'onglet actif.
    def on_tab_changed(event):
        global gui, constraints_app
        current_index = notebook.index(notebook.select())
        gui, constraints_app, _ = tabs_data[current_index]

    notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

    # ------------------------------------------------------------------
    #  Raccourcis clavier globaux (instanciÃ©s UNE seule fois)
    #  ---------------------------------------------------------------
    #  Le lambda utilise la variable globale `gui`, mise Ã  jour par
    #  on_tab_changed ; ainsi, les actions sâappliquent toujours
    #  Ã  lâonglet (semaine) actif, sans conserver dâanciennes instances.
    # ------------------------------------------------------------------
    root.bind_all("<Control-MouseWheel>", lambda e: gui.on_zoom(e))
    root.bind_all("<Control-z>",           lambda e: gui.undo_cell_edit(e))
    root.bind_all("<Control-Shift-Z>",     lambda e: gui.undo_last_change())



    # Au dÃ©marrage, on se place par dÃ©faut sur le premier onglet
    if tabs_data:
        gui, constraints_app, _ = tabs_data[0]

    # Vous pouvez choisir l'emplacement (en bas, en haut, etc.). Ici, on le met en haut sous la barre de menu.
    week_buttons_frame = ttk.Frame(root, style=RIBBON_FRAME_STYLE, padding=(12, 8))
    globals()["week_buttons_frame"] = week_buttons_frame
    week_buttons_frame.pack(side="top", fill="x", padx=12, pady=(0, 8))

    btn_add_week = ttk.Button(
        week_buttons_frame,
        text="Ajouter mois",
        command=add_new_week,
        style=RIBBON_BUTTON_STYLE,
    )
    btn_add_week.pack(side="left", padx=(0, 8))

    btn_remove_week = ttk.Button(
        week_buttons_frame,
        text="Supprimer mois",
        command=remove_current_week,
        style=RIBBON_BUTTON_STYLE,
    )
    btn_remove_week.pack(side="left", padx=8)

    def schedule_auto_save():
        auto_save_status()
        root.after(3 * 60 * 1000, schedule_auto_save)

    root.after(3 * 60 * 1000, schedule_auto_save)
    
    # Lancement de la boucle principale
    root.mainloop()
