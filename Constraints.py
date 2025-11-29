import tkinter as tk



from tkinter import ttk, messagebox



class AbsenceToggleButton(tk.Button):



    def __init__(self, master=None, on_change=None, **kwargs):



        # DÃ©finition des Ã©tats : prÃ©sent ou absence (journÃ©e entiÃ¨re)
        self.states = ["", "Absence"]



        self.state_index = 0



        self._on_change = on_change



        self._external_command = kwargs.pop("command", None)



        super().__init__(master, text=self.states[self.state_index], command=self._handle_click, **kwargs)



        # Variable pour conserver l'Ã©tat courant (accessible lors de la collecte des contraintes)



        self._var = tk.StringVar(value=self.states[self.state_index])



        self.origin = "manual"



        self.log_text = ""



        self._default_bg = self.cget("bg")
        self._manual_log_bg = "#E9D7FF"



        try:



            self.bind("<Button-3>", self._open_log_editor, add="+")



            self.bind("<Button-2>", self._open_log_editor, add="+")



        except Exception:



            pass



    def set_state(self, value: str):



        """



        Fixe l'Ã©tat Ã  partir d'une valeur texte chargÃ©e ("" | "MATIN" | "AP MIDI" | "Journée")



        et synchronise l'index interne pour que le prochain clic poursuive le cycle.



        """



        previous_state = self.states[self.state_index]



        try:



            idx = self.states.index(value)



        except ValueError:



            idx = 0



        self.state_index = idx



        val = self.states[self.state_index]



        self._var.set(val)



        self.config(text=val)



        if not val:



            self.set_origin("manual", log_text="", notify=False)



        self._apply_origin_style()



        if val != previous_state:



            self._notify_change()



    def toggle_state(self):



        previous_state = self.states[self.state_index]



        self.state_index = (self.state_index + 1) % len(self.states)



        new_state = self.states[self.state_index]



        self.config(text=new_state)



        self._var.set(new_state)



        if not new_state:



            self.set_origin("manual", log_text="", notify=False)



        else:



            self.set_origin("manual", log_text=self.log_text, notify=False)



        self._apply_origin_style()



        if new_state != previous_state:



            self._notify_change()



    def set_on_change(self, callback):



        self._on_change = callback



    def _handle_click(self):



        self.toggle_state()



        if callable(self._external_command):



            try:



                self._external_command()



            except Exception:



                pass



    def _notify_change(self):



        if callable(self._on_change):



            try:



                self._on_change()



            except Exception:



                pass



    # --- Gestion de l'origine et du log -----------------------------------



    def _apply_origin_style(self):



        try:



            if self.origin == "import_conflict":



                self.config(bg="#FDE1C2")



            elif self.origin == "import_absence":



                self.config(bg="#FDF7C0")



            elif self.origin == "manual" and self.log_text:



                self.config(bg=self._manual_log_bg)



            else:



                self.config(bg=self._default_bg)



        except Exception:



            pass



    def set_origin(self, origin, log_text=None, notify=True):



        origin_norm = (origin or "manual").strip().lower()



        if origin_norm not in {"manual", "import_absence", "import_conflict"}:



            origin_norm = "manual"



        changed = False



        if self.origin != origin_norm:



            self.origin = origin_norm



            changed = True



        if self.states[self.state_index] == "":



            if self.log_text:



                self.log_text = ""



                changed = True



            if self.origin != "manual":



                self.origin = "manual"



                changed = True



        else:



            if log_text is not None and log_text != self.log_text:



                self.log_text = log_text



                changed = True



        self._apply_origin_style()



        if changed and notify:



            self._notify_change()



        return changed



    def get_origin(self):



        return self.origin



    def get_log(self):



        return self.log_text



    def set_log(self, text, notify=True):



        if text is None:



            text = ""



        if text == self.log_text:



            return False



        self.log_text = text



        try:



            self._apply_origin_style()



        except Exception:



            pass



        if notify:



            self._notify_change()



        return True



    def _open_log_editor(self, event=None):



        result = self._show_log_dialog()



        if result is not None:



            result = result.strip("\n")



            self.set_log(result, notify=False)



            if not self.states[self.state_index]:



                self.set_origin("manual", log_text="", notify=False)



            self._apply_origin_style()



            self._notify_change()



        return "break"



    def _show_log_dialog(self):



        try:



            parent = self.winfo_toplevel()



        except Exception:



            parent = None



        dialog = tk.Toplevel(parent or self)



        dialog.title("Informations absence")



        dialog.transient(parent or dialog)



        dialog.resizable(True, True)



        dialog.grab_set()



        try:



            x = self.winfo_rootx() + 24



            y = self.winfo_rooty() + 24



            dialog.geometry(f"+{x}+{y}")



        except Exception:



            pass



        tk.Label(dialog, text="Détails / log :", anchor="w").pack(fill="x", padx=12, pady=(10, 0))



        text_widget = tk.Text(dialog, width=40, height=6, wrap="word")



        text_widget.pack(fill="both", expand=True, padx=12, pady=8)



        text_widget.insert("1.0", self.log_text or "")



        button_frame = tk.Frame(dialog)



        button_frame.pack(fill="x", padx=12, pady=(0, 12))



        result_container = {"value": None}



        def on_ok():



            result_container["value"] = text_widget.get("1.0", "end").rstrip("\n")



            dialog.destroy()



        def on_cancel():



            dialog.destroy()



        ok_btn = tk.Button(button_frame, text="Valider", command=on_ok, width=10)



        ok_btn.pack(side="right", padx=(4, 0))



        cancel_btn = tk.Button(button_frame, text="Annuler", command=on_cancel, width=10)



        cancel_btn.pack(side="right", padx=(0, 4))



        try:



            dialog.bind("<Return>", lambda e: on_ok())



            dialog.bind("<Escape>", lambda e: on_cancel())



        except Exception:



            pass



        dialog.wait_window()



        return result_container["value"]



# Titres des colonnes



columns = [



    "Initiales",



    "PDS/semaine",



    "Lignes préférentielles",



    "Lignes non-assurées",



    "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche",



    "Commentaire"          # â† nouvelle colonne libre (index 11)



]



# Jours de la semaine



days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]



# Plages horaires (Ã  l'image du tableau principal)



shifts = ["MATIN", "AP MIDI"]



# --- Nouvelle classe utilitaire pour la sÃ©lection multiple via Listbox ---



class MultiSelectPopup(tk.Toplevel):



    def __init__(self, master, values, preselected=None, **kwargs):



        super().__init__(master, **kwargs)



        self.title("Sélectionner")



        self.geometry("+{}+{}".format(master.winfo_rootx() + 20, master.winfo_rooty() + 20))



        self.selected = []



        self.listbox = tk.Listbox(self, selectmode="multiple", width=40, height=10)



        self.listbox.pack(fill="both", expand=True, padx=10, pady=10)



        # Remplissage de la Listbox



        for val in values:



            self.listbox.insert(tk.END, val)



        # --- NOUVEAU : prÃ©-sÃ©lection des valeurs dÃ©jÃ  choisies ---



        if preselected:



            preselected_set = set(preselected)



            for i, val in enumerate(values):



                if val in preselected_set:



                    self.listbox.selection_set(i)



        ok_btn = tk.Button(self, text="OK", command=self.on_ok)



        ok_btn.pack(pady=5)



        self.grab_set()



        self.wait_window()



    def on_ok(self):



        indices = self.listbox.curselection()



        self.selected = [self.listbox.get(i) for i in indices]



        self.destroy()



class CheckListButton(tk.Button):



    def __init__(self, master=None, values=None, **kwargs):



        super().__init__(master, text="Sélectionner", command=self.show_checklist, **kwargs)



        self._var = tk.StringVar()



        self._values = values or []



    def show_checklist(self):



        """



        Ouvre un popup de sÃ©lection multiple avec la liste des postes Ã€ JOUR



        et coche dâ€™emblÃ©e les postes dÃ©jÃ  sÃ©lectionnÃ©s pour permettre



        un ajustement incrÃ©mental (ajout/retrait).



        """



        # RÃ©cupÃ©ration dynamique des postes



        try:



            from Full_GUI import get_work_posts



            updated_values = get_work_posts()



        except ImportError:



            updated_values = self._values



        # Met Ã  jour la liste de rÃ©fÃ©rence



        self._values = updated_values



        # --- NOUVEAU : construire la liste prÃ©-sÃ©lectionnÃ©e ---



        # On lit d'abord _var, sinon le texte du bouton (aprÃ¨s chargement)



        current_txt = (self._var.get() or self.cget("text") or "").strip()



        if current_txt == "Sélectionner":



            preselected = []



        else:



            preselected = [p.strip() for p in current_txt.split(",") if p.strip()]



        # Afficher le popup avec prÃ©-sÃ©lection



        popup = MultiSelectPopup(self, self._values, preselected=preselected)



        selected = popup.selected



        # Mise Ã  jour de l'Ã©tat interne + rendu



        self._var.set(", ".join(selected))



        self.config(text=(self._var.get() if selected else "Sélectionner"))



    def update(self, var, window):



        pass



class Application(tk.Frame):



    def __init__(self, master=None):



        super().__init__(master)



        self.master = master



        self.rows = []
        self._highlighted_entries = []

        self.minimized = False  # Ã‰tat de la minimisation



        self._change_callback = None



        self._change_job = None



        self.create_widgets()



                # --- Synchronisation avec les changements de work_posts ---



        root = self.winfo_toplevel()          # fenÃªtre racine



        root.bind("<<WorkPostsUpdated>>",



                  lambda e: self.refresh_available_posts())



    def set_change_callback(self, callback):



        self._change_callback = callback



    def _notify_change(self):



        callback = self._change_callback



        if not callable(callback):



            return



        try:



            if self._change_job is not None:



                self.after_cancel(self._change_job)



        except Exception:



            pass



        def _run():



            self._change_job = None



            try:



                callback()



            except Exception:



                pass



        try:



            self._change_job = self.after(60, _run)



        except Exception:



            _run()



    def create_widgets(self):



        """



        Construit lâ€™interface du tableau Contraintes :



        â€¢ en-tÃªtes + 3 lignes descriptives



        â€¢ zone scrollable avec Canvas



        â€¢ molette via routeur global (aucune duplication)



        """



        # Autoriser lâ€™Ã©tirement



        self.grid_rowconfigure(1, weight=1)



        self.grid_columnconfigure(0, weight=1)



        padx = pady = 0



        # â”€â”€â”€â”€â”€â”€â”€â”€â”€ En-tÃªte â”€â”€â”€â”€â”€â”€â”€â”€â”€



        self.header_frame = tk.Frame(self)



        self.header_frame.grid(row=0, column=0, columnspan=len(columns) + 2, sticky="nsew")



        header_lbl = tk.Label(self.header_frame, text="Contraintes", font=("Arial", 10, "bold"))



        header_lbl.pack(side="left", padx=padx, pady=pady)



        self.minimize_btn = tk.Button(

            self.header_frame, text="X", command=self.toggle_minimize,

            font=("Arial", 10, "bold"), fg="black"


        )



        self.minimize_btn.pack(side="right", padx=padx, pady=pady)



        # â”€â”€â”€â”€â”€â”€â”€â”€â”€ Zone scrollable (Canvas + Scrollbars) en GRID â”€â”€â”€â”€â”€â”€â”€â”€â”€



        self.table_container = tk.Frame(self)



        self.table_container.grid(row=1, column=0, columnspan=len(columns) + 2, sticky="nsew")



        # Le conteneur gÃ¨re lâ€™expansion du Canvas (0,0) et de la hbar (1,0)



        self.table_container.grid_rowconfigure(0, weight=1)



        self.table_container.grid_columnconfigure(0, weight=1)



        self.canvas = tk.Canvas(self.table_container, borderwidth=0, highlightthickness=0)



        self.vsb = tk.Scrollbar(self.table_container, orient="vertical", command=self.canvas.yview)



        self.hsb = tk.Scrollbar(self.table_container, orient="horizontal", command=self.canvas.xview)



        # Connexion des scrollbars



        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)



        # Placement en grid (âš  remplace lâ€™ancien pack)



        self.canvas.grid(row=0, column=0, sticky="nsew")



        self.vsb.grid(row=0, column=1, sticky="ns")



        self.hsb.grid(row=1, column=0, sticky="ew")



        # Optionnel : petit coin dâ€™angle pour un rendu propre (Ã  cÃ´tÃ© de la hbar)



        corner = tk.Frame(self.table_container, width=1, height=1)



        corner.grid(row=1, column=1, sticky="nsew")



        # Cadre rÃ©el du contenu Ã  lâ€™intÃ©rieur du Canvas



        self.content_frame = tk.Frame(self.canvas)



        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")



        # Ajuste la zone de scroll quand le contenu change (fait apparaÃ®tre hbar/vbar si besoin)



        self.content_frame.bind(



            "<Configure>",



            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))



        )



        # IMPORTANT : ne pas forcer la largeur du content_frame Ã  celle du canvas,



        # pour conserver le dÃ©filement horizontal quand le contenu est large.



        # --- Enregistrement auprÃ¨s du routeur global de molette ---



        from Full_GUI import get_mousewheel_manager  # import tardif pour Ã©viter les cycles



        get_mousewheel_manager(self).register(self.canvas, self.content_frame)



        # â”€â”€â”€â”€â”€â”€â”€â”€â”€ En-tÃªtes du tableau â”€â”€â”€â”€â”€â”€â”€â”€â”€



        for i in range(4):



            tk.Label(self.content_frame, text=columns[i], font=("Arial", 10, "bold")).grid(row=0, column=i, padx=padx, pady=pady)



        tk.Label(self.content_frame, text="Absences", font=("Arial", 10, "bold")).grid(row=0, column=4, columnspan=7, padx=padx, pady=pady, sticky="nsew")



        tk.Label(self.content_frame, text="Commentaire", font=("Arial", 10, "bold")).grid(row=0, column=11, padx=padx, pady=pady)



        for i, day in enumerate(days):



            tk.Label(self.content_frame, text=day, font=("Arial", 10, "bold")).grid(row=1, column=4 + i, padx=padx, pady=pady)



        for i in range(len(days)):



            tk.Label(self.content_frame, text="", font=("Arial", 9, "bold")).grid(row=2, column=4 + i, padx=padx, pady=pady)



        # â”€â”€â”€â”€â”€â”€â”€â”€â”€ Boutons Ajouter / Supprimer â”€â”€â”€â”€â”€â”€â”€â”€â”€



        self.add_button = tk.Button(self.content_frame, text="Ajouter", command=self.add_row, font=("Arial", 9))



        self.add_button.grid(row=0, column=len(columns), padx=padx, pady=pady)



        self.delete_button = tk.Button(self.content_frame, text="Supprimer", command=self.delete_row, font=("Arial", 9))



        self.delete_button.grid(row=0, column=len(columns) + 1, padx=padx, pady=pady)



        # â”€â”€â”€â”€â”€â”€â”€â”€â”€ Lignes initiales + menu contextuel â”€â”€â”€â”€â”€â”€â”€â”€â”€



        for _ in range(5):



            self.add_row()



        self._init_row_context_menu()



        for _idx in range(len(self.rows)):



            self._bind_row_context_menu(_idx)



    def _cleanup_bindings(self, _event=None):



        """Nettoie les bind_all (molette) enregistrÃ©s par cette instance."""



        try:



            root = self.winfo_toplevel()



        except Exception:



            return



        for seq, bind_id in getattr(self, "_wheel_bind_ids", []):



            try:



                root.unbind(seq, bind_id)  # Tk â‰¥ 8.6



            except Exception:



                # Ã€ dÃ©faut, on retire tous les handlers du sÃ©quence (filet de sÃ©curitÃ©)



                root.unbind_all(seq)



        self._wheel_bind_ids = []



    def select_preferred_posts(self, btn):



        """



        Ouvre un popup multi-sÃ©lection pour les Postes préférées avec prÃ©-cochage



        des valeurs dÃ©jÃ  prÃ©sentes sur le bouton, afin de modifier la liste



        sans repartir de zÃ©ro.



        """



        try:



            from Full_GUI import get_work_posts



            posts = get_work_posts()



        except Exception:



            posts = []



        # --- NOUVEAU : calcul de la prÃ©-sÃ©lection Ã  partir du texte du bouton ---



        current_txt = (btn.cget("text") or "").strip()



        if current_txt == "Sélectionner":



            preselected = []



        else:



            preselected = [p.strip() for p in current_txt.split(",") if p.strip()]



        # Popup avec prÃ©-sÃ©lection



        popup = MultiSelectPopup(btn, posts, preselected=preselected)



        selected = popup.selected



        # Affichage du rÃ©sultat



        btn.config(text=(", ".join(selected) if selected else "Sélectionner"))



        # Optionnel : garder une copie structurÃ©e



        btn._preferred = selected



    def toggle_minimize(self):



        # self.master est le cadre (bottom_frame) contenant le tableau de contraintes,



        # et son parent (self.master.master) est le PanedWindow.



        paned = getattr(self.master, "master", None)
        sash_getter = getattr(paned, "sashpos", None)
        can_adjust_paned = callable(sash_getter)

        if not self.minimized:

            self.content_frame.grid_remove()

            self.minimize_btn.config(text="v")

            self.minimized = True

            if can_adjust_paned:

                self.saved_sashpos = paned.sashpos(0)

                self.update_idletasks()

                total_height = paned.winfo_height()

                header_height = self.header_frame.winfo_height()

                new_sashpos = total_height - header_height

                paned.sashpos(0, new_sashpos)

        else:

            self.content_frame.grid()

            self.minimize_btn.config(text="X")

            self.minimized = False

            if can_adjust_paned and hasattr(self, 'saved_sashpos'):

                paned.sashpos(0, self.saved_sashpos)


    def _shift_rows_down_from(self, grid_row: int) -> None:



        """Shift grid rows >= grid_row down by one to make room for an insertion."""



        widgets_to_shift = []



        for widget in self.content_frame.grid_slaves():



            info = widget.grid_info()



            row = int(info.get('row', 0))



            if row >= grid_row:



                widgets_to_shift.append((row, widget))



        widgets_to_shift.sort(key=lambda item: item[0], reverse=True)



        for row, widget in widgets_to_shift:



            widget.grid_configure(row=row + 1)



    def _assign_row_index(self, widget, row_idx: int) -> None:



        """Attach the logical row index to a widget when possible."""



        if widget is None:



            return



        try:



            widget._row_index = row_idx



        except Exception:



            pass



    def _set_row_index_metadata(self, row_idx: int) -> None:



        """Ensure all widgets in the given row carry the up-to-date index."""



        if row_idx < 0 or row_idx >= len(self.rows):



            return



        for cell in self.rows[row_idx]:



            if isinstance(cell, tuple):



                for part in cell:



                    self._assign_row_index(part, row_idx)



            else:



                self._assign_row_index(cell, row_idx)



    def _update_row_index_metadata(self, start_idx: int = 0) -> None:



        """Refresh row index metadata from start_idx to the end of the table."""



        for idx in range(start_idx, len(self.rows)):



            self._set_row_index_metadata(idx)



    def _resolve_grid_widget(self, cell):



        if isinstance(cell, tuple):



            for part in cell:



                parent = getattr(part, 'master', None)



                if parent is not None and hasattr(parent, 'grid_info'):



                    return parent



            return None



        return cell if hasattr(cell, 'grid_info') else None



    def _iterate_row_grid_widgets(self, row_widgets):



        seen = set()



        for cell in row_widgets:



            widget = self._resolve_grid_widget(cell)



            if widget is None:



                continue



            key = id(widget)



            if key in seen:



                continue



            seen.add(key)



            yield widget



    def _swap_adjacent_rows(self, idx_a: int, idx_b: int) -> None:



        if idx_a == idx_b:



            return



        if idx_a < 0 or idx_b < 0:



            return



        if idx_a >= len(self.rows) or idx_b >= len(self.rows):



            return



        if idx_a > idx_b:



            idx_a, idx_b = idx_b, idx_a



        row_a_widgets = self.rows[idx_a]



        row_b_widgets = self.rows[idx_b]



        target_row_a = idx_b + 3



        target_row_b = idx_a + 3



        for widget in self._iterate_row_grid_widgets(row_a_widgets):



            try:



                widget.grid_configure(row=target_row_a)



            except Exception:



                pass



        for widget in self._iterate_row_grid_widgets(row_b_widgets):



            try:



                widget.grid_configure(row=target_row_b)



            except Exception:



                pass



        self.rows[idx_a], self.rows[idx_b] = self.rows[idx_b], self.rows[idx_a]



    def add_row(self, insert_index: int | None = None):



        """



        Ajoute une ligne complÃ¨te dans le tableau des contraintes :



          0  Initiales



          1  Vacations/semaine



          2  Postes préférées    (bouton)



          3  Postes non-assurées (CheckListButton)



          4-10 Jours Lundiâ†’Dimanche (AbsenceToggle + )



         11  Commentaire (texte libre)



        """



        from Full_GUI import work_posts



        if insert_index is None:



            insert_index = len(self.rows)



        insert_index = max(0, min(insert_index, len(self.rows)))



        target_grid_row = insert_index + 3  # dÇ¸calage des trois lignes d'en-tÇ¦tes



        if insert_index < len(self.rows):



            self._shift_rows_down_from(target_grid_row)



        padx = pady = 0          # marges minimales



        row_widgets  = []



        current_row  = target_grid_row



        for j in range(len(columns)):



            # --- 0) Initiales ------------------------------------------------



            if j == 0:



                entry = tk.Entry(self.content_frame, width=20, font=("Arial", 9))



                entry.grid(row=current_row, column=j, padx=padx, pady=pady)



                entry.bind("<FocusIn>",



                           lambda e, rw=row_widgets: self.highlight_row(rw, True))



                entry.bind("<FocusOut>",



                           lambda e, rw=row_widgets: self.highlight_row(rw, False))



                row_widgets.append(entry)



            # --- 1) Nb de vacations / semaine --------------------------------



            elif j == 1:



                entry = tk.Entry(self.content_frame, width=3, font=("Arial", 9))



                entry.grid(row=current_row, column=j, padx=padx, pady=pady)



                row_widgets.append(entry)



            # --- 2) Postes préférées ------------------------------------------



            elif j == 2:



                btn = tk.Button(self.content_frame,



                                text="Sélectionner",



                                font=("Arial", 9),



                                command=lambda b=None: self.select_preferred_posts(btn))



                btn.grid(row=current_row, column=j, padx=padx, pady=pady)



                row_widgets.append(btn)



            # --- 3) Postes non-assurées ---------------------------------------



            elif j == 3:



                checklist = CheckListButton(self.content_frame, values=work_posts)



                checklist.config(font=("Arial", 9))



                checklist.grid(row=current_row, column=j, padx=padx, pady=pady)



                row_widgets.append(checklist)



            # --- 4-10) Jours (Absence) ---------------------------------



            elif 4 <= j <= 10:



                day_frame = tk.Frame(self.content_frame)



                day_frame.grid(row=current_row, column=j, padx=padx, pady=pady)



                toggle = AbsenceToggleButton(day_frame, font=("Arial", 9), on_change=self._notify_change)
                toggle.pack(side="left")
                row_widgets.append(toggle)



            # --- 11) Commentaire ---------------------------------------------



            elif j == len(columns) - 1:



                entry = tk.Entry(self.content_frame, width=20, font=("Arial", 9))



                entry.grid(row=current_row, column=j, padx=padx, pady=pady)



                row_widgets.append(entry)



        # --- Bouton d'action de ligne ----------------------------------------



        action_btn = tk.Button(



            self.content_frame,



            text="+",



            font=("Arial", 9),



            command=lambda r=row_widgets: self.open_constraint_row_action_dialog(r)



        )



        action_btn._is_row_action_button = True



        action_btn.grid(row=current_row, column=len(columns) + 1, padx=padx, pady=pady)



        row_widgets.append(action_btn)



        self.rows.insert(insert_index, row_widgets)



        self._bind_row_context_menu(insert_index)



        self._update_row_index_metadata(insert_index)



        return row_widgets



    def delete_row_at(self, index):



        """



        Supprime la ligne dâ€™index donnÃ© dans self.rows,



        supprime tous les widgets de cette ligne, et remonte



        toutes les lignes en dessous dâ€™une position.



        """



        # Enregistrer dans la pile dâ€™undo du GUI principal



        if hasattr(self, 'undo_target'):



            self.undo_target.push_undo_state()



        removed_grid_row = index + 3  # offset des en-tÃªtes



        # Oublier les widgets de cette ligne



        for widget in self.content_frame.grid_slaves(row=removed_grid_row):



            widget.grid_forget()



        # Retirer la ligne de la liste



        del self.rows[index]



        # Descendre toutes les autres lignes dâ€™une rangÃ©e



        for widget in self.content_frame.grid_slaves():



            info = widget.grid_info()



            r = info['row']



            if r > removed_grid_row:



                widget.grid(



                    row=r-1,



                    column=info['column'],



                    padx=info.get('padx', 0),



                    pady=info.get('pady', 0),



                    sticky=info.get('sticky', '')



                )



        # Mettre Ã  jour la zone scroll



        self.content_frame.update_idletasks()



        self._update_row_index_metadata(index)



    def delete_row_obj(self, row_obj):



        """



        Delete the row associated with the per-row action button.



        """



        try:



            idx = self.rows.index(row_obj)



            self.delete_row_at(idx)



        except ValueError:



            # Si row_obj n'est plus dans self.rows, on ignore



            pass



    def highlight_row(self, row_widgets, highlight):



        """



        Colore toute la ligne en gris clair si highlight=True,



        ou restaure le fond par dÃ©faut (blanc/SystemButtonFace) si False.



        """



        # Couleurs



        row_bg = "#D3D3D3" if highlight else "white"



        default_btn = "SystemButtonFace"



        for widget in row_widgets:



            if isinstance(widget, tuple) and len(widget) >= 3:



                toggle, pds_cb, pds_var = widget



                container = getattr(toggle, "master", None)



                if container is not None:



                    try:



                        container.config(bg=row_bg)



                    except Exception:



                        pass



                try:



                    toggle._apply_origin_style()



                except Exception:



                    if not highlight:



                        try:



                            toggle.config(bg=getattr(toggle, "_default_bg", default_btn))



                        except Exception:



                            pass



                try:



                    if int(pds_var.get()) == 1:



                        pds_cb.config(bg="red")



                    else:



                        pds_cb.config(bg=row_bg if highlight else default_btn)



                except Exception:



                    pass



            elif isinstance(widget, tk.Entry):



                try:



                    widget.config(bg=row_bg)



                except Exception:



                    pass



            elif isinstance(widget, tk.Button):



                try:



                    if highlight:



                        if not hasattr(widget, "_base_bg"):



                            widget._base_bg = widget.cget("bg")



                        widget.config(bg=row_bg)



                    else:



                        base_bg = getattr(widget, "_base_bg", default_btn)



                        widget.config(bg=base_bg if base_bg is not None else default_btn)



                        if hasattr(widget, "_base_bg"):



                            delattr(widget, "_base_bg")



                except Exception:



                    pass



    def delete_row(self):



        if self.rows:



            # supprime toujours la derniÃ¨re ligne



            self.delete_row_at(len(self.rows)-1)



    def refresh_available_posts(self):



        """



        AppelÃ©e lorsquâ€™un poste est ajoutÃ© / supprimÃ© / renommÃ©.



        1. Met Ã  jour la liste interne `_values` de chaque widget.



        2. Met Ã  jour le texte affichÃ© sur les boutons pour ne conserver



           que les postes encore existants.  Si plus aucun poste nâ€™est



           valide, on rÃ©-affiche Â« Sélectionner Â».



        """



        import Full_GUI



        updated_posts = Full_GUI.get_work_posts()



        for row in self.rows:



            # Colonne 2 : "Postes préférées"



            if len(row) > 2:



                pref_btn = row[2]



                if isinstance(pref_btn, tk.Button) and getattr(pref_btn, 'winfo_exists', lambda: False)():



                    try:



                        # liste actuelle affichÃ©e sur le bouton



                        current = [p.strip() for p in pref_btn.cget("text").split(",")



                                   if p.strip() and p.strip() != "Sélectionner"]



                    except tk.TclError:



                        current = []



                    # on enlÃ¨ve les postes qui nâ€™existent plus



                    current = [p for p in current if p in updated_posts]



                    new_txt = ", ".join(current) if current else "Sélectionner"



                    try:



                        pref_btn.config(text=new_txt)



                    except tk.TclError:



                        pass



                # Colonne 3 : "Postes non-assurées" (CheckListButton)



                checklist = row[3]



                if hasattr(checklist, '_values') and getattr(checklist, 'winfo_exists', lambda: False)():



                    checklist._values = updated_posts  # nouvelle liste complÃ¨te



                    try:



                        selected = [p.strip() for p in checklist._var.get().split(",")



                                    if p.strip()]



                    except tk.TclError:



                        selected = []



                    selected = [p for p in selected if p in updated_posts]



                    checklist._var.set(", ".join(selected))



                    try:



                        checklist.config(text=checklist._var.get()



                                         if selected else "Sélectionner")



                    except tk.TclError:



                        pass



    # ======== MENU CONTEXTUEL & DUPLICATION DE LIGNE ========



    def highlight_candidate_initials(self, current_initial, eligible_set):
        """
        Met en évidence les candidats :
          - Jaune pour l'initiale courante.
          - Vert clair pour les candidats éligibles.
        """
        self.clear_candidate_highlight()
        if not self.rows:
            return
        eligible_set = set(eligible_set or [])
        for row in self.rows:
            if not row:
                continue
            widget = row[0]
            if not isinstance(widget, tk.Entry):
                continue
            try:
                value = widget.get().strip()
            except Exception:
                continue
            color = None
            if current_initial and value == current_initial:
                color = "#FFF4B5"  # jaune doux
            elif value and value in eligible_set:
                color = "#CDEDC9"  # vert clair
            if not color:
                continue
            try:
                original_bg = widget.cget("background")
            except Exception:
                original_bg = None
            try:
                original_fg = widget.cget("foreground")
            except Exception:
                original_fg = None
            self._highlighted_entries.append((widget, original_bg, original_fg))
            try:
                widget.configure(background=color)
            except Exception:
                pass

    def clear_candidate_highlight(self):
        if not self._highlighted_entries:
            return
        for widget, original_bg, original_fg in self._highlighted_entries:
            if not widget:
                continue
            try:
                if original_bg is not None:
                    widget.configure(background=original_bg)
                if original_fg is not None:
                    widget.configure(foreground=original_fg)
            except Exception:
                pass
        self._highlighted_entries.clear()

    def _init_row_context_menu(self):



        """



        CrÃ©e le menu contextuel (clic droit) pour les lignes du tableau de contraintes.



        """



        self._rclick_row_index = None



        self._row_menu = tk.Menu(self, tearoff=0)



        self._row_menu.add_command(



            label="Dupliquer cette ligne",



            command=self._do_duplicate_row_from_menu



        )



    def _bind_row_context_menu(self, row_idx: int):



        """



        Attache le clic droit Ã  tous les widgets de la ligne 'row_idx', de faÃ§on RÃ‰CURSIVE.



        Idempotent : ne double pas les bindings si dÃ©jÃ  faits.



        """



        if not hasattr(self, "rows"):



            return



        if row_idx < 0 or row_idx >= len(self.rows):



            return



        for cell in self.rows[row_idx]:



            self._attach_rclick_recursively(cell, row_idx)



        def _tag_and_bind(widget):



            # On mÃ©morise l'indice de ligne directement sur le widget



            try:



                widget._row_index = row_idx



            except Exception:



                pass



            # Propager aux sous-composants si la cellule est un tuple (ex: (toggle, cb, var))



            if isinstance(widget, tuple):



                for part in widget:



                    _tag_and_bind(part)



                return



            # Bind clic droit (Windows/Linux = <Button-3>, macOS parfois <Button-2>)



            try:



                widget.bind("<Button-3>", self.on_row_right_click, add="+")



                widget.bind("<Button-2>", self.on_row_right_click, add="+")



            except Exception:



                pass



        for cell in self.rows[row_idx]:



            _tag_and_bind(cell)



    def on_row_right_click(self, event):



        """



        DÃ©tecte la ligne visÃ©e par le clic droit et affiche le menu contextuel.



        StratÃ©gie :



        1) remonter la hiÃ©rarchie .master pour trouver _row_index ;



        2) sinon, balayage de secours : on cherche la ligne dont un des widgets



            est ancÃªtre du widget cliquÃ©.



        """



        # 1) Recherche directe par remontÃ©e d'ancÃªtres



        w = event.widget



        found = None



        hops = 0



        while w is not None and hops < 50:  # profondeur Ã©largie pour Ãªtre sÃ»r



            if hasattr(w, "_row_index"):



                found = getattr(w, "_row_index")



                break



            w = getattr(w, "master", None)



            hops += 1



        # 2) Secours : on infÃ¨re la ligne en testant l'ascendance des cellules



        if found is None and hasattr(self, "rows"):



            for idx, row in enumerate(self.rows):



                if self._widget_belongs_to_row(event.widget, row):



                    found = idx



                    break



        if found is None:



            return  # aucune ligne identifiÃ©e, on ignore



        self._rclick_row_index = int(found)



        try:



            self._row_menu.tk_popup(event.x_root, event.y_root)



        finally:



            self._row_menu.grab_release()



    def _do_duplicate_row_from_menu(self):



        """



        Callback du menu : duplique la ligne qui a reÃ§u le clic droit.



        """



        if self._rclick_row_index is None:



            return



        self.duplicate_row(self._rclick_row_index)



    def duplicate_row(self, src_idx: int, insert_index: int | None = None, placeholder_initials: str | None = "New person"):



        """



        Create a new row based on src_idx and insert it right after by default.



        Copies every column except initials, which receives a placeholder.



        """



        if not hasattr(self, "rows") or src_idx < 0 or src_idx >= len(self.rows):



            return



        if insert_index is None:



            insert_index = src_idx + 1



        insert_index = max(0, min(insert_index, len(self.rows)))



        before_count = len(self.rows)



        new_row_widgets = self.add_row(insert_index=insert_index)



        if new_row_widgets is None or len(self.rows) != before_count + 1:



            return



        new_idx = self.rows.index(new_row_widgets)



        self._copy_row_values(src_idx, new_idx)



        first_cell = new_row_widgets[0] if new_row_widgets else None



        if isinstance(first_cell, tk.Entry):



            try:



                first_cell.delete(0, tk.END)



                if placeholder_initials:



                    first_cell.insert(0, placeholder_initials)



            except Exception:



                pass



        if hasattr(self, "push_undo_state") and callable(getattr(self, "push_undo_state", None)):



            try:



                self.push_undo_state()



            except Exception:



                pass



        self._notify_change()



        return new_idx



    def move_row(self, src_idx: int, dst_idx: int) -> None:



        """Move row src_idx to dst_idx (adjacent swaps), updating grid and metadata."""



        if not hasattr(self, "rows"):



            return



        row_count = len(self.rows)



        if row_count == 0:



            return



        if src_idx < 0 or dst_idx < 0 or src_idx >= row_count or dst_idx >= row_count:



            return



        if src_idx == dst_idx:



            return



        step = 1 if dst_idx > src_idx else -1



        current = src_idx



        while current != dst_idx:



            next_idx = current + step



            if next_idx < 0 or next_idx >= row_count:



                break



            self._swap_adjacent_rows(current, next_idx)



            current = next_idx



        self._update_row_index_metadata(min(src_idx, dst_idx))



        if hasattr(self, "rebind_all_rows_context_menu"):



            try:



                self.rebind_all_rows_context_menu()



            except Exception:



                pass



        self._notify_change()



    # ---------- Helpers de copie de cellule ----------



    def _copy_row_values(self, src_idx: int, dst_idx: int):



        """



        Copie la valeur de chaque cellule de la ligne src_idx vers dst_idx



        en respectant les types de widgets. Ignore la colonne 0.



        """



        if src_idx < 0 or dst_idx < 0:



            return



        if src_idx >= len(self.rows) or dst_idx >= len(self.rows):



            return



        src_row = self.rows[src_idx]



        dst_row = self.rows[dst_idx]



        ncols = min(len(src_row), len(dst_row))



        for col in range(1, ncols):



            payload = self._get_widget_value(src_row[col])



            self._set_widget_value(dst_row[col], payload)



    def _apply_row_to_all_weeks(self, src_idx: int) -> None:



        """Copy the row at src_idx to matching initials in other weeks."""



        if not hasattr(self, 'rows') or src_idx < 0 or src_idx >= len(self.rows):



            return



        try:



            from Full_GUI import tabs_data



        except Exception:



            try:



                messagebox.showerror('Apply to all weeks', 'Unable to access other weeks.')



            except Exception:



                pass



            return



        src_row = self.rows[src_idx]



        def _extract_text(cell):



            getter = getattr(cell, 'get', None)



            if callable(getter):



                try:



                    return getter().strip()



                except Exception:



                    pass



            try:



                return cell._var.get().strip()



            except Exception:



                pass



            try:



                return str(cell.cget('text')).strip()



            except Exception:



                return ''



        src_name = _extract_text(src_row[0]) if src_row else ''



        if not src_name:



            try:



                messagebox.showinfo('Apply to all weeks', 'No initials found on this row.')



            except Exception:



                pass



            return



        row_payloads = [self._get_widget_value(cell) for cell in src_row]



        applied = 0



        for gui, constraints_app, _ in tabs_data:



            if constraints_app is self:



                continue



            for target_row in getattr(constraints_app, 'rows', []):



                if not target_row:



                    continue



                target_name = _extract_text(target_row[0])



                if target_name != src_name:



                    continue



                for col in range(1, min(len(row_payloads), len(target_row))):



                    constraints_app._set_widget_value(target_row[col], row_payloads[col])



                constraints_app._notify_change()



                applied += 1



        try:



            if applied:



                messagebox.showinfo('Apply to all weeks', f'Modifications applied to {applied} matching row(s).')



            else:



                messagebox.showinfo('Apply to all weeks', 'No matching names found in other weeks.')



        except Exception:



            pass



    def open_constraint_row_action_dialog(self, row_widgets):



        """Open a modal dialog offering to delete or duplicate the selected row."""



        def _current_index():



            try:



                return self.rows.index(row_widgets)



            except ValueError:



                return None



        row_index = _current_index()



        if row_index is None:



            return



        popup = tk.Toplevel(self)



        popup.title('Row actions')



        popup.transient(self.winfo_toplevel())



        popup.resizable(False, False)



        name = ''



        first_cell = row_widgets[0] if row_widgets else None



        if isinstance(first_cell, tk.Entry):



            try:



                name = first_cell.get().strip()



            except Exception:



                name = ''



        if not name:



            name = f'Ligne {row_index + 1}'



        tk.Label(popup, text=f"Ligne : {name}", anchor='w').pack(fill='x', padx=15, pady=(15, 5))



        delete_var = tk.BooleanVar(value=False)



        duplicate_var = tk.BooleanVar(value=False)



        move_up_var = tk.BooleanVar(value=False)



        move_down_var = tk.BooleanVar(value=False)



        apply_all_var = tk.BooleanVar(value=False)



        options = [



            ('Delete this row?', delete_var, 'delete'),



            ('Duplicate this row?', duplicate_var, 'duplicate'),



            ('Move this row up?', move_up_var, 'move_up'),



            ('Move this row down?', move_down_var, 'move_down'),



            ('Apply modifications to all weeks?', apply_all_var, 'apply_all'),



        ]



        def set_active(option_key):



            for label, var, key in options:



                var.set(1 if key == option_key else 0)



        for label, var, key in options:



            btn = tk.Checkbutton(popup, text=label, variable=var, anchor='w',



                                 command=lambda k=key: set_active(k))



            btn.pack(fill='x', padx=15, pady=2)



        button_frame = tk.Frame(popup)



        button_frame.pack(pady=15)



        def on_confirm():



            current_index = _current_index()



            if current_index is None:



                return



            do_duplicate = duplicate_var.get()



            do_delete = delete_var.get()



            do_move_up = move_up_var.get()



            do_move_down = move_down_var.get()



            do_apply_all = apply_all_var.get()



            if do_move_up and do_move_down:



                do_move_up = do_move_down = False



            if do_apply_all and (do_duplicate or do_delete or do_move_up or do_move_down):



                do_apply_all = False



            duplicate_result = None



            if do_duplicate:



                duplicate_result = self.duplicate_row(current_index)



                current_index = _current_index()



                if current_index is None:



                    return



            if do_delete and (duplicate_result is not None or not do_duplicate):



                self.delete_row_at(current_index)



                return



            if do_move_up or do_move_down:



                target_index = current_index - 1 if do_move_up else current_index + 1



                if 0 <= target_index < len(self.rows):



                    self.move_row(current_index, target_index)



                return



            if do_apply_all:



                self._apply_row_to_all_weeks(current_index)



        tk.Button(button_frame, text='OK', width=10, command=on_confirm).pack(side='left', padx=5)



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



    def _get_widget_value(self, cell):



        """



        Extrait la valeur d'une 'cellule' (widget simple, tuple composite, etc.)



        Retourne un tuple (kind, *args) pour guider la restauration.



        """



        # Cellule AbsenceToggle seule
        if isinstance(cell, AbsenceToggleButton):
            try:
                abs_val = cell._var.get()
            except Exception:
                abs_val = ""
            try:
                origin_val = cell.get_origin()
            except Exception:
                origin_val = getattr(cell, "origin", "manual")
            try:
                log_val = cell.get_log()
            except Exception:
                log_val = getattr(cell, "log_text", "")
            return ("toggle_absence", abs_val, origin_val, log_val)



        w = cell



        # Boutons/Widgets avec StringVar interne



        if hasattr(w, "_var"):



            try:



                return ("textvar", w._var.get())



            except Exception:



                pass



        # Entry / Combobox / Checkbutton / gÃ©nÃ©riques



        try:



            if isinstance(w, (tk.Entry, ttk.Entry)):



                return ("entry", w.get())



            if isinstance(w, ttk.Combobox):



                return ("combo", w.get())



            if isinstance(w, (tk.Checkbutton, ttk.Checkbutton)):



                # Cherche la variable associÃ©e



                var = getattr(w, "variable", None) or getattr(w, "_variable", None)



                if var is not None:



                    return ("check", int(var.get()))



                return ("check", 0)



            # Par dÃ©faut : on tente le texte



            return ("text", w.cget("text"))



        except Exception:



            return ("unknown", None)



    def _set_widget_value(self, cell, payload):



        """



        Applique une valeur dans une cellule selon le 'kind' fourni par _get_widget_value.



        """



        kind = payload[0] if isinstance(payload, (list, tuple)) and payload else "unknown"



        # Cellule AbsenceToggle seule
        if isinstance(cell, AbsenceToggleButton) and kind == "toggle_absence":
            abs_val = payload[1] if len(payload) > 1 else ""
            origin_val = payload[2] if len(payload) > 2 else None
            log_val = payload[3] if len(payload) > 3 else None

            try:
                cell.set_state(abs_val)
            except Exception:
                try:
                    cell._var.set(abs_val)
                    cell.config(text=abs_val)
                except Exception:
                    pass

            try:
                cell.set_origin(origin_val or "manual", log_text=log_val, notify=False)
            except Exception:
                if origin_val is not None:
                    try:
                        cell.origin = origin_val
                    except Exception:
                        pass
                if log_val is not None:
                    try:
                        cell.log_text = log_val
                    except Exception:
                        pass



            try:



                toggle._apply_origin_style()



            except Exception:



                pass



            # 



            try:



                previous_pds = int(pds_var.get())



            except Exception:



                previous_pds = None



            try:



                desired_pds = int(pds_val)



            except Exception:



                desired_pds = 0



            try:



                pds_var.set(desired_pds)



            except Exception:



                pass



            # Couleur du checkbox si ton code change le bg selon la valeur



            try:



                pds_cb.config(bg=("red" if desired_pds == 1 else "SystemButtonFace"))



            except Exception:



                # ttk : pas de bg direct ? on ignore



                pass



            if desired_pds != previous_pds:



                self._notify_change()



            return



        w = cell



        if kind in ("textvar", "text"):



            value = payload[1] if len(payload) > 1 else ""



            if hasattr(w, "_is_row_action_button"):



                value = "+"



            try:



                w._var.set(value)



                # Si c'est un bouton affichant le texte stockÃ© dans _var



                try:



                    w.config(text=value)



                except Exception:



                    pass



                return



            except Exception:



                pass



            # Fallback texte direct



            try:



                w.config(text=value)



                return



            except Exception:



                pass



        if kind == "entry":



            try:



                w.delete(0, tk.END)



                w.insert(0, payload[1])



                return



            except Exception:



                pass



        if kind == "combo":



            try:



                w.set(payload[1])



                return



            except Exception:



                pass



        if kind == "check":



            try:



                var = getattr(w, "variable", None) or getattr(w, "_variable", None)



                if var is not None:



                    var.set(int(payload[1]))



                return



            except Exception:



                pass



        # Unknown : on tente un .config(text=...)



        try:



            w.config(text=payload[1])



        except Exception:



            pass



    def _attach_rclick_recursively(self, obj, row_idx: int):



        """



        Attache <Button-3>/<Button-2> sur 'obj' ET tous ses descendants.



        Ã‰vite les doublons via un flag _ctx_bound.



        GÃ¨re les cellules composites sous forme de tuple.



        """



        def _attach_one(w):



            # Tuples (ex: (toggle, pds_cb, pds_var))



            if isinstance(w, tuple):



                for part in w:



                    _attach_one(part)



                return



            # Certains objets non-widgets peuvent circuler (variables, etc.)



            if not hasattr(w, "bind"):



                return



            # Idempotence



            try:



                if getattr(w, "_ctx_bound", False) and getattr(w, "_row_index", None) == row_idx:



                    pass



                else:



                    w._row_index = row_idx



                    try:



                        w.bind("<Button-3>", self.on_row_right_click, add="+")



                        w.bind("<Button-2>", self.on_row_right_click, add="+")  # macOS



                    except Exception:



                        pass



                    w._ctx_bound = True



            except Exception:



                pass



            # Descendants (Frames, ttk, etc.)



            try:



                for child in w.winfo_children():



                    _attach_one(child)



            except Exception:



                pass



        _attach_one(obj)



    def _widget_belongs_to_row(self, leaf_widget, row_cells) -> bool:



        """



        Retourne True si 'leaf_widget' est un descendant de l'un des widgets de la ligne.



        Supporte les cellules tuples.



        """



        def is_descendant(w, root):



            cur = w



            hops = 0



            while cur is not None and hops < 200:



                if cur is root:



                    return True



                cur = getattr(cur, "master", None)



                hops += 1



            return False



        for cell in row_cells:



            if isinstance(cell, tuple):



                for part in cell:



                    if hasattr(part, "winfo_exists") and part.winfo_exists():



                        if is_descendant(leaf_widget, part):



                            return True



            else:



                if hasattr(cell, "winfo_exists") and cell.winfo_exists():



                    if is_descendant(leaf_widget, cell):



                        return True



        return False



    def rebind_all_rows_context_menu(self):



        """



        Re-binde le menu contextuel sur toutes les lignes existantes.



        Sans effet secondaire si dÃ©jÃ  fait.



        """



        if not hasattr(self, "rows"):



            return



        for idx in range(len(self.rows)):



            self._bind_row_context_menu(idx)



if __name__ == '__main__':



    root = tk.Tk()



    app = Application(root)



    app.grid()



    root.mainloop()
