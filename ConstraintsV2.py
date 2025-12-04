import tkinter as tk
from tkinter import ttk, messagebox
import calendar
from datetime import date

# Colonnes du tableau de contraintes (mode mensuel)
COLUMNS = [
    "Initiales",
    "Participation (%)",
    "Lignes préférentielles",
    "Lignes non assurées",
    "Associations",
    "Absences (jours du mois)",
    "Commentaire",
]


def _split_csv(text: str):
    return [p.strip() for p in str(text or "").split(",") if p.strip()]

def _center_popup_over_widget(popup: tk.Toplevel, widget) -> None:
    """
    Centre un Toplevel au-dessus du toplevel du widget donnÃ© (gÃ¨re le multi-\xc3\xa9cran).
    Reprend la logique du sÃ©lecteur de mois Â« Choisir date Â».
    """
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


class MultiSelectPopup(tk.Toplevel):
    """Popup multi-sélection générique."""

    def __init__(self, master, values, preselected=None, anchor_widget=None):
        super().__init__(master)
        self.title("Sélection")
        self.resizable(False, False)
        self.selected = []
        self._anchor_widget = anchor_widget or master
        pre = set(preselected or [])
        frame = tk.Frame(self)
        frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.listbox = tk.Listbox(frame, selectmode="multiple", exportselection=False)
        self.listbox.pack(fill="both", expand=True)
        for val in values:
            self.listbox.insert("end", val)
        for idx, val in enumerate(values):
            if val in pre:
                self.listbox.selection_set(idx)
        btns = tk.Frame(self)
        btns.pack(pady=(8, 6))
        tk.Button(btns, text="OK", width=10, command=self.on_ok).pack(side="left", padx=4)
        tk.Button(btns, text="Annuler", width=10, command=self.on_cancel).pack(side="left", padx=4)
        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_cancel())
        self.grab_set()
        _center_popup_over_widget(self, self._anchor_widget)

    def on_ok(self):
        self.selected = [self.listbox.get(i) for i in self.listbox.curselection()]
        self.destroy()

    def on_cancel(self):
        self.destroy()


class CheckListButton(tk.Button):
    """Bouton ouvrant un popup multi-sélection et stockant le résultat dans _var."""

    def __init__(self, master, values, **kwargs):
        super().__init__(master, **kwargs)
        self._values = list(values)
        self._var = tk.StringVar(value="")
        self.config(text="Sélectionner", command=self._open_popup)

    def _open_popup(self):
        preselected = _split_csv(self._var.get())
        popup = MultiSelectPopup(self, self._values, preselected=preselected, anchor_widget=self)
        self.wait_window(popup)
        if popup.selected:
            txt = ", ".join(popup.selected)
            self._var.set(txt)
            self.config(text=txt)
        else:
            self._var.set("")
            self.config(text="Sélectionner")

    def update_values(self, new_values):
        self._values = list(new_values)
        # Filtrer la sélection actuelle
        current = [v for v in _split_csv(self._var.get()) if v in self._values]
        self._var.set(", ".join(current))
        self.config(text=self._var.get() or "Sélectionner")


class MultiDayPopup(tk.Toplevel):
    """Popup avec cases 1..30 pour choisir les jours d'absence."""

    def __init__(self, master, initial=None, anchor_widget=None):
        super().__init__(master)
        self.title("Jours d'absence")
        self.resizable(False, False)
        self.selected = set(initial or [])
        self._drag_start = None
        self._drag_active = False
        self._day_widgets = {}
        self.year = None
        self.month = None
        self._anchor_widget = anchor_widget or master

        nav = tk.Frame(self)
        nav.pack(fill="x", padx=10, pady=(10, 4))
        self._title_var = tk.StringVar()
        tk.Label(nav, textvariable=self._title_var, font=("Arial", 10, "bold")).pack(side="left", expand=True, fill="x")

        self.grid_frame = tk.Frame(self)
        self.grid_frame.pack(padx=10, pady=6)

        btns = tk.Frame(self)
        btns.pack(pady=(8, 10))
        tk.Button(btns, text="OK", width=10, command=self.on_ok).pack(side="left", padx=4)
        tk.Button(btns, text="Annuler", width=10, command=self.on_cancel).pack(side="left", padx=4)

        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_cancel())
        self.grab_set()

    def load_month(self, year, month):
        self.year, self.month = year, month
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self._day_widgets.clear()
        try:
            month_label = f"{calendar.month_name[month]} {year}"
        except Exception:
            month_label = f"{month}/{year}"
        self._title_var.set(month_label)

        day_names = ["L", "M", "M", "J", "V", "S", "D"]
        for idx, dname in enumerate(day_names):
            tk.Label(self.grid_frame, text=dname, font=("Arial", 9, "bold")).grid(row=0, column=idx, padx=3, pady=3)

        cal = calendar.Calendar(firstweekday=0)
        for row_idx, week in enumerate(cal.monthdayscalendar(year, month), start=1):
            for col_idx, day in enumerate(week):
                if day == 0:
                    tk.Label(self.grid_frame, text="", width=4).grid(row=row_idx, column=col_idx, padx=2, pady=2)
                    continue
                btn = tk.Label(self.grid_frame, text=str(day), width=4, relief="raised", borderwidth=1, bg="white")
                btn.grid(row=row_idx, column=col_idx, padx=2, pady=2, sticky="nsew")
                btn.bind("<ButtonPress-1>", lambda e, d=day: self._start_drag(d, e.state))
                btn.bind("<Enter>", lambda e, d=day: self._drag_over(d))
                btn.bind("<B1-Motion>", lambda e, d=day: self._drag_over(d))
                btn.bind("<ButtonRelease-1>", lambda e, d=day: self._end_drag(d, e.state))
                self._day_widgets[day] = btn
        self._refresh_display()
        _center_popup_over_widget(self, self._anchor_widget)

    def _shift_month(self, delta):
        # Navigation désactivée : on reste sur le mois fourni par le planning
        return

    def _start_drag(self, day, state):
        self._drag_start = day
        self._drag_active = True
        self._shift_active = bool(state & 0x0001)
        self._preview = {day}
        self._refresh_display(preview=self._preview)

    def _drag_over(self, day):
        if not self._drag_active or self._drag_start is None:
            return
        start = self._drag_start
        rng = range(min(start, day), max(start, day) + 1)
        self._preview = set(rng)
        self._refresh_display(preview=self._preview)

    def _end_drag(self, day, state):
        if self._drag_start is None:
            return
        self._drag_active = False
        start = self._drag_start
        rng = set(range(min(start, day), max(start, day) + 1))
        ctrl = bool(state & 0x0004)
        shift = self._shift_active or bool(state & 0x0001)
        if shift:
            anchor = self._last_click_day if getattr(self, "_last_click_day", None) else start
            rng = set(range(min(anchor, day), max(anchor, day) + 1))
            self.selected = rng
        elif ctrl:
            if not self.selected:
                self.selected = set(rng)
            else:
                new_sel = set(self.selected)
                for d in rng:
                    if d in new_sel:
                        new_sel.remove(d)
                    else:
                        new_sel.add(d)
                self.selected = new_sel
        else:
            # Toggle simple sur clic (sans shift/ctrl) : si on clique un seul jour existant, on l'enlève s'il était déjà sélectionné
            if len(rng) == 1 and day in self.selected:
                self.selected.remove(day)
            else:
                self.selected = set(rng)
        self._drag_start = None
        self._last_click_day = day
        self._preview = set()
        self._refresh_display()

    def _refresh_display(self, preview=None):
        preview = preview or set()
        for day, widget in self._day_widgets.items():
            in_sel = day in self.selected
            in_prev = day in preview
            if in_prev:
                bg = "#CDE8FF"
            elif in_sel:
                bg = "#99D1A7"
            else:
                bg = "white"
            try:
                widget.config(bg=bg)
            except Exception:
                pass

    def on_ok(self):
        self.selected = {d for d in self.selected if d >= 1}
        self.destroy()

    def on_cancel(self):
        self.destroy()


class ConstraintsTable(tk.Frame):
    """Tableau de contraintes simplifié (mensuel)."""

    MIN_COL_WIDTHS = [120, 110, 170, 170, 170, 170, 200, 60]

    def __init__(self, master=None, work_posts=None, planning_gui=None):
        super().__init__(master)
        self.rows = []
        self.work_posts = list(work_posts or [])
        self.minimized = False
        self.table_container = None
        self.table_canvas = None
        self.table_inner = None
        self._saved_sash = None
        self.planning_gui = planning_gui
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._build_header()
        for _ in range(5):
            self.add_row()

    # UI builders ---------------------------------------------------------
    def _build_header(self):
        # Toolbar séparée pour ne pas perturber l'alignement des titres
        toolbar = tk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="e", padx=(0, 0), pady=(0, 2))
        btn_add = tk.Button(toolbar, text="Ajouter", command=self.add_row)
        btn_add.pack(side="left", padx=(0, 4))
        btn_del = tk.Button(toolbar, text="Supprimer", command=self.delete_row)
        btn_del.pack(side="left", padx=(0, 4))
        self.min_btn = tk.Button(toolbar, text="−", width=2, command=self.toggle_minimize)
        self.min_btn.pack(side="left", padx=(0, 4))

        header = tk.Frame(self)
        header.grid(row=1, column=0, sticky="nsew")
        self.header = header
        for i, col in enumerate(COLUMNS):
            tk.Label(header, text=col, font=("Arial", 10, "bold")).grid(
                row=0, column=i, padx=4, pady=4, sticky="nsew"
            )
            header.grid_columnconfigure(i, weight=1, minsize=self.MIN_COL_WIDTHS[i] if i < len(self.MIN_COL_WIDTHS) else 80)

        self.table_container = tk.Frame(self)
        self.table_container.grid(row=2, column=0, sticky="nsew")

        self.table_canvas = tk.Canvas(self.table_container, highlightthickness=0)
        vbar = tk.Scrollbar(self.table_container, orient="vertical", command=self.table_canvas.yview)
        hbar = tk.Scrollbar(self.table_container, orient="horizontal", command=self.table_canvas.xview)
        self.table_canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        self.table_canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        self.table_container.grid_rowconfigure(0, weight=1)
        self.table_container.grid_columnconfigure(0, weight=1)

        self.table_inner = tk.Frame(self.table_canvas)
        self.table = self.table_inner
        self.table_window = self.table_canvas.create_window((0, 0), window=self.table_inner, anchor="nw")

        def _on_frame_config(event):
            try:
                self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))
            except Exception:
                pass

        def _on_canvas_config(event):
            try:
                self.table_canvas.itemconfig(self.table_window, width=event.width)
            except Exception:
                pass

        self.table_inner.bind("<Configure>", _on_frame_config)
        self.table_canvas.bind("<Configure>", _on_canvas_config)
        self._apply_column_layout()
        self._setup_mousewheel()

    # Row management ------------------------------------------------------
    def add_row(self):
        idx = len(self.rows)
        entries = []

        # Initiales
        init_entry = tk.Entry(self.table, width=14)
        init_entry.grid(row=idx, column=0, padx=4, pady=2, sticky="ew")
        entries.append(init_entry)

        # Participation (%)
        part_var = tk.StringVar(value="100")
        part_spin = tk.Spinbox(
            self.table,
            from_=0,
            to=100,
            increment=5,
            width=6,
            textvariable=part_var,
            justify="center",
        )
        part_spin.grid(row=idx, column=1, padx=4, pady=2, sticky="ew")
        part_spin.bind("<MouseWheel>", lambda e, v=part_var: self._on_part_wheel(e, v))
        part_spin.bind("<Button-4>", lambda e, v=part_var: self._on_part_wheel(e, v, 1))
        part_spin.bind("<Button-5>", lambda e, v=part_var: self._on_part_wheel(e, v, -1))
        entries.append(part_spin)

        # Lignes préférentielles
        pref_btn = tk.Button(
            self.table,
            text="Sélectionner",
            font=("Arial", 9),
        )
        pref_btn.grid(row=idx, column=2, padx=4, pady=2, sticky="ew")
        pref_btn._var = tk.StringVar(value="")
        pref_btn.config(command=lambda b=pref_btn: self._open_pref_popup(b))
        entries.append(pref_btn)

        # Lignes non assurées
        non_btn = CheckListButton(self.table, values=self.work_posts, font=("Arial", 9))
        non_btn.grid(row=idx, column=3, padx=4, pady=2, sticky="ew")
        entries.append(non_btn)

        # Associations possibles (multi-sélection)
        assoc_btn = CheckListButton(self.table, values=self.work_posts, font=("Arial", 9))
        assoc_btn.grid(row=idx, column=4, padx=4, pady=2, sticky="ew")
        entries.append(assoc_btn)

        # Absences bouton + var
        abs_var = tk.StringVar(value="")
        abs_btn = tk.Button(
            self.table,
            text="Sélectionner",
            font=("Arial", 9),
        )
        # Commande configurée après création pour capturer abs_btn
        abs_btn.config(command=lambda v=abs_var, b=abs_btn: self._open_days_popup(v, b))
        abs_btn.grid(row=idx, column=5, padx=4, pady=2, sticky="ew")
        abs_btn.var = abs_var
        entries.append(abs_btn)

        # Commentaire
        comment = tk.Entry(self.table, width=20)
        comment.grid(row=idx, column=6, padx=4, pady=2, sticky="ew")
        entries.append(comment)

        # Bouton d'action (+) en fin de ligne
        action_btn = tk.Button(self.table, text="+", width=3)
        action_btn._is_row_action_button = True
        action_btn._var = tk.StringVar(master=self, value="all")
        action_btn.config(command=lambda b=action_btn: self._open_action_menu(b))
        action_btn.grid(row=idx, column=7, padx=4, pady=2, sticky="e")
        entries.append(action_btn)

        self.rows.append(entries)
        self._apply_column_layout()

    def delete_row(self):
        if not self.rows:
            return
        widgets = self.rows.pop()
        for w in widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self._apply_column_layout()

    # Helpers -------------------------------------------------------------
    def _open_pref_popup(self, btn: tk.Button):
        preselected = _split_csv(btn._var.get())
        popup = MultiSelectPopup(self, self.work_posts, preselected=preselected, anchor_widget=btn)
        self.wait_window(popup)
        if popup.selected:
            txt = ", ".join(popup.selected)
            btn._var.set(txt)
            btn.config(text=txt)
        else:
            btn._var.set("")
            btn.config(text="Sélectionner")

    def _open_days_popup(self, var: tk.StringVar, btn):
        try:
            current = {int(x.strip()) for x in var.get().split(",") if x.strip().isdigit()}
        except Exception:
            current = set()
        # Détermine le mois/année courant depuis le planning si dispo
        if hasattr(self, "planning_gui") and self.planning_gui is not None:
            year = getattr(self.planning_gui, "current_year", date.today().year)
            month = getattr(self.planning_gui, "current_month", date.today().month)
        else:
            today = date.today()
            year, month = today.year, today.month
        popup = MultiDayPopup(self, initial=current, anchor_widget=btn)
        popup.load_month(year, month)
        self.wait_window(popup)
        if popup.selected:
            txt = ",".join(str(x) for x in sorted(popup.selected))
            var.set(txt)
            btn.config(text=txt)
        else:
            var.set("")
            btn.config(text="Sélectionner")

    def _find_row_index(self, widget) -> int:
        for idx, row in enumerate(self.rows):
            if widget in row:
                return idx
        return -1

    def _open_action_menu(self, btn):
        idx = self._find_row_index(btn)
        name = ""
        try:
            if 0 <= idx < len(self.rows):
                raw = self.rows[idx][0].get()
                name = raw.strip()
        except Exception:
            name = ""
        if not name:
            name = f"Ligne {idx + 1}" if idx >= 0 else "Ligne"

        popup = tk.Toplevel(self)
        popup.title("Actions")
        popup.transient(self.winfo_toplevel())
        popup.resizable(False, False)

        tk.Label(popup, text=name, anchor="w").pack(fill="x", padx=12, pady=(12, 6))

        try:
            current_scope = btn._var.get()
        except Exception:
            current_scope = "all"
        if not current_scope or current_scope == "+":
            current_scope = "all"
        weekdays_only_var = tk.BooleanVar(value=str(current_scope).lower() == "weekdays_only")
        weekends_only_var = tk.BooleanVar(value=str(current_scope).lower() == "weekends_only")

        def _toggle(which):
            if which == "weekdays" and weekdays_only_var.get():
                weekends_only_var.set(False)
            elif which == "weekends" and weekends_only_var.get():
                weekdays_only_var.set(False)

        tk.Checkbutton(
            popup,
            text="Weekdays only (no weekends/holidays)",
            variable=weekdays_only_var,
            anchor="w",
            command=lambda: _toggle("weekdays"),
        ).pack(fill="x", padx=28, pady=2)

        tk.Checkbutton(
            popup,
            text="Weekends only (no weekdays)",
            variable=weekends_only_var,
            anchor="w",
            command=lambda: _toggle("weekends"),
        ).pack(fill="x", padx=28, pady=2)

        btn_frame = tk.Frame(popup)
        btn_frame.pack(pady=12)

        def _apply_and_close():
            scope = "weekdays_only" if weekdays_only_var.get() else "weekends_only" if weekends_only_var.get() else "all"
            try:
                btn._var.set(scope)
            except Exception:
                pass
            popup.destroy()

        tk.Button(btn_frame, text="OK", width=10, command=_apply_and_close).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Annuler", width=10, command=popup.destroy).pack(side="left", padx=4)
        popup.bind("<Return>", lambda e: _apply_and_close())
        popup.bind("<Escape>", lambda e: popup.destroy())
        popup.grab_set()
        _center_popup_over_widget(popup, btn)

    def _setup_mousewheel(self):
        """Active la molette pour faire défiler le tableau des contraintes."""
        manager = None
        try:
            from Full_GUI import get_mousewheel_manager  # import tardif pour éviter les cycles
            manager = get_mousewheel_manager
        except Exception:
            manager = None

        # Si le gestionnaire global existe, on s'en sert pour rester cohérent avec le planning.
        if manager is not None:
            try:
                manager(self).register(self.table_canvas, self.table_inner)
                return
            except Exception:
                pass

        # Fallback local (mode standalone) : route la molette vers le canvas si la souris est dessus.
        active = {"inside": False}

        def _on_wheel(event):
            if not active["inside"]:
                return
            delta = getattr(event, "delta", 0)
            if delta:
                step = -1 if delta > 0 else 1
            elif getattr(event, "num", None) == 4:
                step = -1
            elif getattr(event, "num", None) == 5:
                step = 1
            else:
                return
            try:
                self.table_canvas.yview_scroll(step, "units")
            except Exception:
                pass
            return "break"

        def _activate(_e=None):
            active["inside"] = True

        def _deactivate(_e=None):
            active["inside"] = False

        for widget in (self.table_canvas, self.table_inner):
            widget.bind("<Enter>", _activate, add="+")
            widget.bind("<Leave>", _deactivate, add="+")
        root = self.table_canvas.winfo_toplevel()
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            root.bind_all(seq, _on_wheel, add="+")

    def _on_part_wheel(self, event, var: tk.StringVar, delta_override=None):
        """Scroll facile par pas de 5% (0..100)."""
        try:
            cur = int(float(var.get()))
        except Exception:
            cur = 0
        delta_raw = delta_override if delta_override is not None else event.delta
        if delta_raw == 0:
            return "break"
        step = 5 * (1 if delta_raw > 0 else -1)
        new_val = max(0, min(100, cur + step))
        # arrondi au multiple de 5 le plus proche
        new_val = int(round(new_val / 5) * 5)
        new_val = max(0, min(100, new_val))
        try:
            var.set(str(new_val))
        except Exception:
            pass
        return "break"

    def toggle_minimize(self):
        if not self.minimized:
            try:
                self.table_container.grid_remove()
            except Exception:
                pass
            try:
                self.min_btn.config(text="+")
            except Exception:
                pass
            # Tenter de remonter la barre de séparation du paned parent
            try:
                paned = self.master.master  # bottom_frame -> paned
                if hasattr(paned, "sashpos"):
                    self._saved_sash = paned.sashpos(0)
                    paned.update_idletasks()
                    header_h = (self.header.winfo_height() if hasattr(self, "header") else 0) + (self.children.get('!frame', tk.Frame()).winfo_height() if self.children else 0)
                    total_h = paned.winfo_height()
                    new_pos = max(0, total_h - header_h - 4)
                    paned.sashpos(0, new_pos)
            except Exception:
                pass
            self.minimized = True
        else:
            try:
                self.table_container.grid()
            except Exception:
                pass
            try:
                self.min_btn.config(text="−")
            except Exception:
                pass
            try:
                paned = self.master.master
                if hasattr(paned, "sashpos") and self._saved_sash is not None:
                    paned.sashpos(0, self._saved_sash)
            except Exception:
                pass
            self.minimized = False

    # Data access ---------------------------------------------------------
    def get_rows_data(self):
        """Retourne une liste de dicts avec les valeurs saisies."""
        data = []
        for row in self.rows:
            try:
                initials = row[0].get().strip()
            except Exception:
                initials = ""
            try:
                part_val = row[1].get().strip()
            except Exception:
                part_val = ""
            pref = getattr(row[2], "_var", tk.StringVar(value="")).get()
            non = getattr(row[3], "_var", tk.StringVar(value="")).get() if hasattr(row[3], "_var") else ""
            assoc = getattr(row[4], "_var", tk.StringVar(value="")).get() if hasattr(row[4], "_var") else ""
            abs_days = getattr(row[5], "var", tk.StringVar(value="")).get()
            try:
                comment = row[6].get().strip()
            except Exception:
                comment = ""
            data.append({
                "initiales": initials,
                "participation": part_val or "100",
                "preferences": pref,
                "non_assurees": non,
                "associations": assoc,
                "absences": abs_days,
                "commentaire": comment,
            })
        return data

    def set_rows_data(self, rows_data):
        """Recharge le tableau depuis une liste de dicts."""
        # clear
        while self.rows:
            self.delete_row()
        for row_dict in rows_data or []:
            self.add_row()
            row = self.rows[-1]
            row[0].insert(0, row_dict.get("initiales", ""))
            try:
                row[1].delete(0, "end")
                row[1].insert(0, row_dict.get("participation", "100"))
            except Exception:
                pass
            row[2]._var.set(row_dict.get("preferences", ""))
            row[2].config(text=row[2]._var.get() or "Sélectionner")
            if hasattr(row[3], "_var"):
                row[3]._var.set(row_dict.get("non_assurees", ""))
                row[3].config(text=row[3]._var.get() or "Sélectionner")
            if hasattr(row[4], "_var"):
                row[4]._var.set(row_dict.get("associations", ""))
                row[4].config(text=row[4]._var.get() or "Sélectionner")
            row[5].var.set(row_dict.get("absences", ""))
            row[5].config(text=row[5].var.get() or "Sélectionner")
            row[6].insert(0, row_dict.get("commentaire", ""))

    def refresh_work_posts(self, new_posts):
        """Met à jour la liste des postes utilisable pour préf/non assurées/associations."""
        self.work_posts = list(new_posts or [])
        for row in self.rows:
            # row[3] et row[4] sont des CheckListButton
            try:
                row[3].update_values(self.work_posts)
            except Exception:
                pass
            try:
                row[4].update_values(self.work_posts)
            except Exception:
                pass

    def _apply_column_layout(self):
        """Assure l'alignement header/table en fixant minsize/weights identiques."""
        for idx in range(len(COLUMNS) + 1):  # inclut la colonne action
            minsize = self.MIN_COL_WIDTHS[idx] if idx < len(self.MIN_COL_WIDTHS) else 80
            try:
                self.header.grid_columnconfigure(idx, weight=1 if idx < len(COLUMNS) else 0, minsize=minsize)
            except Exception:
                pass
            try:
                self.table.grid_columnconfigure(idx, weight=1, minsize=minsize)
            except Exception:
                pass


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Constraints V2")
    table = ConstraintsTable(root, work_posts=["Ligne 1", "Ligne 2", "Ligne 3"])
    table.pack(fill="both", expand=True, padx=10, pady=10)
    root.mainloop()
