import tkinter as tk
from tkinter import ttk, messagebox

# Colonnes du tableau de contraintes (mode mensuel)
COLUMNS = [
    "Initiales",
    "Lignes préférentielles",
    "Lignes non assurées",
    "Absences (jours du mois)",
    "Commentaire",
]


def _split_csv(text: str):
    return [p.strip() for p in str(text or "").split(",") if p.strip()]


class MultiSelectPopup(tk.Toplevel):
    """Popup multi-sélection générique."""

    def __init__(self, master, values, preselected=None):
        super().__init__(master)
        self.title("Sélection")
        self.resizable(False, False)
        self.selected = []
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
        popup = MultiSelectPopup(self, self._values, preselected=preselected)
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

    def __init__(self, master, initial=None):
        super().__init__(master)
        self.title("Jours d'absence")
        self.resizable(False, False)
        self.selected = set(initial or [])

        grid = tk.Frame(self)
        grid.pack(padx=10, pady=10)

        self.vars = {}
        for day in range(1, 31):
            var = tk.IntVar(value=1 if day in self.selected else 0)
            self.vars[day] = var
            row = (day - 1) // 6
            col = (day - 1) % 6
            cb = tk.Checkbutton(grid, text=str(day), variable=var)
            cb.grid(row=row, column=col, sticky="w", padx=2, pady=2)

        btns = tk.Frame(self)
        btns.pack(pady=(8, 6))
        tk.Button(btns, text="OK", width=10, command=self.on_ok).pack(side="left", padx=4)
        tk.Button(btns, text="Annuler", width=10, command=self.on_cancel).pack(side="left", padx=4)

        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_cancel())
        self.grab_set()

    def on_ok(self):
        self.selected = {day for day, var in self.vars.items() if var.get() == 1}
        self.destroy()

    def on_cancel(self):
        self.destroy()


class ConstraintsTable(tk.Frame):
    """Tableau de contraintes simplifié (mensuel)."""

    def __init__(self, master=None, work_posts=None):
        super().__init__(master)
        self.rows = []
        self.work_posts = list(work_posts or [])
        self.minimized = False
        self._saved_sash = None
        self._build_header()
        for _ in range(5):
            self.add_row()

    # UI builders ---------------------------------------------------------
    def _build_header(self):
        header = tk.Frame(self)
        header.grid(row=0, column=0, sticky="nsew")
        self.header = header
        for i, col in enumerate(COLUMNS):
            tk.Label(header, text=col, font=("Arial", 10, "bold")).grid(
                row=0, column=i, padx=4, pady=4, sticky="nsew"
            )
            header.grid_columnconfigure(i, weight=1)
        # Boutons Ajouter/Supprimer + Minimiser à l'extrémité droite
        btn_add = tk.Button(header, text="Ajouter", command=self.add_row)
        btn_add.grid(row=0, column=len(COLUMNS), padx=(8, 2), pady=4, sticky="e")
        btn_del = tk.Button(header, text="Supprimer", command=self.delete_row)
        btn_del.grid(row=0, column=len(COLUMNS) + 1, padx=(2, 2), pady=4, sticky="e")
        self.min_btn = tk.Button(header, text="−", width=2, command=self.toggle_minimize)
        self.min_btn.grid(row=0, column=len(COLUMNS) + 2, padx=(2, 8), pady=4, sticky="e")
        header.grid_columnconfigure(len(COLUMNS), weight=0)
        header.grid_columnconfigure(len(COLUMNS) + 1, weight=0)
        header.grid_columnconfigure(len(COLUMNS) + 2, weight=0)

        self.table = tk.Frame(self)
        self.table.grid(row=1, column=0, sticky="nsew")

    # Row management ------------------------------------------------------
    def add_row(self):
        idx = len(self.rows)
        entries = []

        # Initiales
        init_entry = tk.Entry(self.table, width=14)
        init_entry.grid(row=idx, column=0, padx=4, pady=2, sticky="ew")
        entries.append(init_entry)

        # Lignes préférentielles
        pref_btn = tk.Button(
            self.table,
            text="Sélectionner",
            font=("Arial", 9),
        )
        pref_btn.grid(row=idx, column=1, padx=4, pady=2, sticky="ew")
        pref_btn._var = tk.StringVar(value="")
        pref_btn.config(command=lambda b=pref_btn: self._open_pref_popup(b))
        entries.append(pref_btn)

        # Lignes non assurées
        non_btn = CheckListButton(self.table, values=self.work_posts, font=("Arial", 9))
        non_btn.grid(row=idx, column=2, padx=4, pady=2, sticky="ew")
        entries.append(non_btn)

        # Absences bouton + var
        abs_var = tk.StringVar(value="")
        abs_btn = tk.Button(
            self.table,
            text="Sélectionner",
            font=("Arial", 9),
            command=lambda v=abs_var, b=None: self._open_days_popup(v, abs_btn),
        )
        abs_btn.grid(row=idx, column=3, padx=4, pady=2, sticky="ew")
        abs_btn.var = abs_var
        entries.append(abs_btn)

        # Commentaire
        comment = tk.Entry(self.table, width=20)
        comment.grid(row=idx, column=4, padx=4, pady=2, sticky="ew")
        entries.append(comment)

        # Bouton d'action (+) en fin de ligne (placeholder)
        action_btn = tk.Button(self.table, text="+", width=3)
        action_btn.grid(row=idx, column=5, padx=4, pady=2, sticky="e")
        entries.append(action_btn)

        self.rows.append(entries)
        for col in range(len(COLUMNS) + 1):
            self.table.grid_columnconfigure(col, weight=1)

    def delete_row(self):
        if not self.rows:
            return
        widgets = self.rows.pop()
        for w in widgets:
            try:
                w.destroy()
            except Exception:
                pass

    # Helpers -------------------------------------------------------------
    def _open_pref_popup(self, btn: tk.Button):
        preselected = _split_csv(btn._var.get())
        popup = MultiSelectPopup(self, self.work_posts, preselected=preselected)
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
        popup = MultiDayPopup(self, initial=current)
        self.wait_window(popup)
        if popup.selected:
            txt = ",".join(str(x) for x in sorted(popup.selected))
            var.set(txt)
            btn.config(text=txt)
        else:
            var.set("")
            btn.config(text="Sélectionner")

    def toggle_minimize(self):
        if not self.minimized:
            try:
                self.table.grid_remove()
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
                    header_h = self.header.winfo_height() if hasattr(self, "header") else 0
                    total_h = paned.winfo_height()
                    new_pos = max(0, total_h - header_h - 4)
                    paned.sashpos(0, new_pos)
            except Exception:
                pass
            self.minimized = True
        else:
            try:
                self.table.grid()
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
            pref = getattr(row[1], "_var", tk.StringVar(value="")).get()
            non = getattr(row[2], "_var", tk.StringVar(value="")).get() if hasattr(row[2], "_var") else ""
            abs_days = getattr(row[3], "var", tk.StringVar(value="")).get()
            try:
                comment = row[4].get().strip()
            except Exception:
                comment = ""
            data.append({
                "initiales": initials,
                "preferences": pref,
                "non_assurees": non,
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
            row[1]._var.set(row_dict.get("preferences", ""))
            row[1].config(text=row[1]._var.get() or "Sélectionner")
            if hasattr(row[2], "_var"):
                row[2]._var.set(row_dict.get("non_assurees", ""))
                row[2].config(text=row[2]._var.get() or "Sélectionner")
            row[3].var.set(row_dict.get("absences", ""))
            row[3].config(text=row[3].var.get() or "Sélectionner")
            row[4].insert(0, row_dict.get("commentaire", ""))

    def refresh_work_posts(self, new_posts):
        """Met à jour la liste des postes utilisable pour préf/non assurées."""
        self.work_posts = list(new_posts or [])
        for row in self.rows:
            # row[2] est le CheckListButton
            try:
                row[2].update_values(self.work_posts)
            except Exception:
                pass


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Constraints V2")
    table = ConstraintsTable(root, work_posts=["Ligne 1", "Ligne 2", "Ligne 3"])
    table.pack(fill="both", expand=True, padx=10, pady=10)
    root.mainloop()
