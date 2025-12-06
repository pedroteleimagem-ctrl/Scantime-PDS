import random
from pathlib import Path
import sys
from typing import Dict, Iterable, Set, Tuple

_MODULE_DIR = Path(__file__).resolve().parent
if str(_MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(_MODULE_DIR))

try:
    from eligibility import AssignmentSettings, PlanningContext, candidate_is_available, parse_constraint_row  # noqa: F401
except ModuleNotFoundError:
    trash_dir = _MODULE_DIR / "trash"
    if trash_dir.is_dir():
        sys.path.insert(0, str(trash_dir))
    from eligibility import AssignmentSettings, PlanningContext, candidate_is_available, parse_constraint_row  # type: ignore  # noqa: F401,E402

del _MODULE_DIR

# Global toggles inherited from previous version
ENABLE_MAX_ASSIGNMENTS = False
MAX_ASSIGNMENTS_PER_POST = None
ENABLE_DIFFERENT_POST_PER_DAY = False
ENABLE_REPOS_SECURITE = False
ENABLE_MAX_WE_DAYS = False
MAX_WE_DAYS_PER_MONTH = None
ENABLE_WEEKEND_BLOCKS = False
WEEKEND_BLOCK_POSTS: Set[str] = set()

FORBIDDEN_POST_ASSOCIATIONS: Set[Tuple[str, str]] = set()


def filter_weekend_block_posts(valid_posts: Iterable[str]) -> None:
    """
    Retire de WEEKEND_BLOCK_POSTS les postes qui n'existent plus
    (appelé quand la liste des postes change).
    """
    global WEEKEND_BLOCK_POSTS
    valid = set(valid_posts or [])
    WEEKEND_BLOCK_POSTS = {p for p in WEEKEND_BLOCK_POSTS if p in valid}


def build_forbidden_maps(work_posts: Iterable[str]) -> Tuple[Dict[int, frozenset[int]], Dict[int, frozenset[int]]]:
    """
    Convert the stored forbidden associations into index-based maps for fast checks.
    """
    posts = list(work_posts)
    index_map = {name: idx for idx, name in enumerate(posts)}

    morning_to_afternoon: Dict[int, Set[int]] = {}
    afternoon_to_morning: Dict[int, Set[int]] = {}

    for morning_name, afternoon_name in FORBIDDEN_POST_ASSOCIATIONS:
        morning_idx = index_map.get(morning_name)
        afternoon_idx = index_map.get(afternoon_name)
        if morning_idx is None or afternoon_idx is None:
            continue
        morning_to_afternoon.setdefault(morning_idx, set()).add(afternoon_idx)
        afternoon_to_morning.setdefault(afternoon_idx, set()).add(morning_idx)

    frozen_morning = {key: frozenset(values) for key, values in morning_to_afternoon.items()}
    frozen_afternoon = {key: frozenset(values) for key, values in afternoon_to_morning.items()}
    return frozen_morning, frozen_afternoon


def assigner_initiales(constraints_app, planning_gui):
    """
    Assigne automatiquement les cr\u00e9neaux du mois poste par poste selon le workflow demand\u00e9 :
    - parcours chronologique jour + poste,
    - priorit\u00e9 aux candidats pr\u00e9f\u00e9rentiels,
    - v\u00e9rification des contraintes (absences, postes interdits, scope semaine/WE, plafond WE global),
    - gestion des associations de postes : affectation du m\u00eame m\u00e9decin sur les postes associ\u00e9s du jour.
    Les cibles sont calcul\u00e9es \u00e0 partir du volume annuel (semaine/week-end) divis\u00e9 par 12 et pond\u00e9r\u00e9es
    par le pourcentage de participation.
    """
    import calendar
    from datetime import date, timedelta
    from Full_GUI import days, work_posts, extract_names_from_cell, POST_INFO

    try:
        from Full_GUI import month_holidays as _month_holidays
    except Exception:
        _month_holidays = None

    rows = getattr(constraints_app, "rows", []) if constraints_app is not None else []
    if not rows:
        return

    year = getattr(planning_gui, "current_year", date.today().year)
    month = getattr(planning_gui, "current_month", date.today().month)
    posts_count = len(work_posts)
    days_in_month = calendar.monthrange(year, month)[1]

    def _split_csv(text):
        return [p.strip() for p in str(text or "").replace(";", ",").split(",") if p.strip()]

    WEEKDAY_CODES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
    WEEKDAY_LABELS = {
        "lundi": "mon", "lun": "mon",
        "mardi": "tue", "mar": "tue",
        "mercredi": "wed", "mer": "wed",
        "jeudi": "thu", "jeu": "thu",
        "vendredi": "fri", "ven": "fri",
        "samedi": "sat", "sam": "sat",
        "dimanche": "sun", "dim": "sun",
    }

    def _parse_excluded_weekdays(raw):
        """
        Convertit une valeur d'exclusion (ancienne portée week/week-end ou CSV de jours)
        en ensemble de codes jour ('mon'..'sun').
        """
        if raw is None:
            return set()
        codes = set()
        if isinstance(raw, (list, tuple, set)):
            values = [str(x).strip().lower() for x in raw if str(x).strip()]
        else:
            txt = str(raw or "").strip().replace(";", ",")
            values = [p.strip().lower() for p in txt.split(",") if p.strip()]

        for val in values:
            if not val or val in {"all", "+", "aucune", "aucune exclusion", "none", "0"}:
                continue
            if val in {"weekdays_only", "weekday_only", "weekdays"}:
                codes.update(["fri", "sat", "sun"])
                continue
            if val in {"weekends_only", "weekend_only", "weekend"}:
                codes.update(["mon", "tue", "wed", "thu"])
                continue
            if val in WEEKDAY_CODES:
                codes.add(val)
                continue
            mapped = WEEKDAY_LABELS.get(val) or WEEKDAY_LABELS.get(val[:3])
            if mapped:
                codes.add(mapped)
        return set(codes)

    try:
        max_we_enabled = bool(ENABLE_MAX_WE_DAYS)
    except Exception:
        max_we_enabled = False
    try:
        max_we_limit = int(MAX_WE_DAYS_PER_MONTH) if max_we_enabled and MAX_WE_DAYS_PER_MONTH is not None else None
        if max_we_limit is not None and max_we_limit < 0:
            max_we_limit = None
    except Exception:
        max_we_limit = None

    def _holidays_for(m_year, m_month):
        if _month_holidays:
            try:
                return set(_month_holidays(m_year, m_month))
            except Exception:
                return set()
        return set()

    def _day_type(dt, hol_set):
        """week = lun-jeu non feri\u00e9, we = ven/sa/di/feri\u00e9/veille feri\u00e9."""
        is_hol = dt in hol_set
        next_hol = (dt + timedelta(days=1)) in hol_set
        wd = dt.weekday()
        if wd >= 5 or wd == 4 or is_hol or next_hol:
            return "we"
        return "week"

    def _weekday_code_for_day(day_num):
        try:
            return WEEKDAY_CODES[date(year, month, day_num).weekday()]
        except Exception:
            return None

    # Totaux annuels (m\u00eame liste de postes toute l'ann\u00e9e)
    annual_week = annual_we = 0
    for m_idx in range(1, 13):
        dim = calendar.monthrange(year, m_idx)[1]
        hol = _holidays_for(year, m_idx)
        for d_idx in range(1, dim + 1):
            dt = date(year, m_idx, d_idx)
            if _day_type(dt, hol) == "we":
                annual_we += posts_count
            else:
                annual_week += posts_count
    annual_week = annual_week or 1
    annual_we = annual_we or 1

    # Profils issus du tableau de contraintes
    profiles = []
    initials_set = set()
    for row in rows:
        try:
            init = str(row[0].get()).strip()
        except Exception:
            init = ""
        if not init:
            continue
        initials_set.add(init)
        try:
            part = float(row[1].get())
        except Exception:
            part = 100.0
        part = max(0.0, min(100.0, part)) / 100.0
        try:
            pref_txt = row[2]._var.get() if hasattr(row[2], "_var") else row[2].cget("text")
        except Exception:
            pref_txt = ""
        preferred = _split_csv(pref_txt)
        try:
            non_txt = row[3]._var.get()
        except Exception:
            try:
                non_txt = row[3].cget("text")
            except Exception:
                non_txt = ""
        non_assured = set(_split_csv(non_txt))
        try:
            abs_txt = row[5].var.get()
        except Exception:
            try:
                abs_txt = row[5].cget("text")
            except Exception:
                abs_txt = ""
        absences = set()
        for part_item in _split_csv(abs_txt):
            try:
                if "-" in part_item:
                    start, end = part_item.split("-", 1)
                    start_i, end_i = int(start), int(end)
                    if start_i <= end_i:
                        absences.update(range(start_i, end_i + 1))
                        continue
                num = int(part_item)
                absences.add(num)
            except Exception:
                continue

        scope_raw = None
        try:
            exclusion_btn = next((w for w in row if getattr(w, "_is_exclusion_button", False)), None)
        except Exception:
            exclusion_btn = None
        if exclusion_btn is not None:
            try:
                scope_raw = exclusion_btn._var.get()
            except Exception:
                try:
                    scope_raw = exclusion_btn.cget("text")
                except Exception:
                    scope_raw = None
        if scope_raw is None:
            try:
                action_btn = row[-1]
                if getattr(action_btn, "_is_row_action_button", False):
                    if hasattr(action_btn, "_var"):
                        scope_raw = action_btn._var.get()
                    else:
                        scope_raw = action_btn.cget("text")
            except Exception:
                scope_raw = None
        excluded_weekdays = _parse_excluded_weekdays(scope_raw)
        scope = ",".join([code for code in WEEKDAY_CODES if code in excluded_weekdays])

        assoc_txt = ""
        try:
            assoc_widget = row[4]
            if hasattr(assoc_widget, "_var"):
                assoc_txt = assoc_widget._var.get()
            else:
                assoc_txt = assoc_widget.cget("text")
        except Exception:
            assoc_txt = ""
        associations = [p for p in _split_csv(assoc_txt)]

        profiles.append(
            {
                "initial": init,
                "participation": part,
                "preferred": preferred,
                "non_assured": non_assured,
                "absences": absences,
                "scope": scope,
                "excluded_weekdays": excluded_weekdays,
                "associations": associations,
            }
        )

    if not profiles:
        return
    # Casse l'ordre des lignes du tableau de contraintes pour éviter un biais de sélection
    random.shuffle(profiles)

    profile_by_initial = {p["initial"]: p for p in profiles}

    parser_valids = initials_set
    exclusion_checker = getattr(planning_gui, "is_cell_excluded_from_count", None)
    excluded_cells = getattr(planning_gui, "excluded_from_count", set())
    cell_availability = getattr(planning_gui, "cell_availability", {})
    hol_month = _holidays_for(year, month)

    def _in_month(row_idx):
        try:
            day_num = int(days[row_idx])
        except Exception:
            day_num = row_idx + 1
        return 1 <= day_num <= days_in_month, day_num

    # Map jour (num) -> index de ligne dans le tableau
    day_to_row = {}
    for idx, val in enumerate(days):
        try:
            num = int(val)
        except Exception:
            num = idx + 1
        day_to_row[num] = idx

    context = PlanningContext(
        table_entries=planning_gui.table_entries,
        name_resolver=lambda raw: extract_names_from_cell(raw, parser_valids),
        exclusion_checker=exclusion_checker,
        excluded_cells=excluded_cells,
    )

    # Associations par profil et index de poste
    post_index_map = {name: idx for idx, name in enumerate(work_posts)}
    profile_assoc_map = {p["initial"]: {} for p in profiles}

    def _add_assoc_for_profile(initial, a_name, b_name):
        a_idx = post_index_map.get(a_name)
        b_idx = post_index_map.get(b_name)
        if a_idx is None or b_idx is None or a_idx == b_idx:
            return
        profile_assoc_map.setdefault(initial, {}).setdefault(a_idx, set()).add(b_idx)
        profile_assoc_map.setdefault(initial, {}).setdefault(b_idx, set()).add(a_idx)

    for p in profiles:
        assoc_list = p.get("associations", [])
        if len(assoc_list) < 2:
            continue
        for i in range(len(assoc_list)):
            for j in range(i + 1, len(assoc_list)):
                _add_assoc_for_profile(p["initial"], assoc_list[i], assoc_list[j])

    # Comptage des assignations existantes (par jour distinct)
    counts_week_days = {p["initial"]: set() for p in profiles}
    counts_we_days = {p["initial"]: set() for p in profiles}
    profile_exclusions = {p["initial"]: set(p.get("excluded_weekdays", set())) for p in profiles}

    def _is_available(profile, day_idx, day_num, post_idx, post_name, day_type, weekday_code):
        excluded_days = profile.get("excluded_weekdays", set())
        if weekday_code and weekday_code in excluded_days:
            return False
        if day_type == "we" and max_we_enabled and max_we_limit is not None:
            try:
                if len(counts_we_days.get(profile["initial"], set())) >= max_we_limit:
                    return False
            except Exception:
                return False
        if day_num in profile["absences"]:
            return False
        if post_name in profile["non_assured"]:
            return False
        if context.already_assigned_in_timeslot(profile["initial"], day_idx, True):
            return False
        return True

    for r_idx, row in enumerate(planning_gui.table_entries):
        in_month, day_num = _in_month(r_idx)
        if not in_month:
            continue
        try:
            dt = date(year, month, day_num)
        except Exception:
            continue
        dtype = _day_type(dt, hol_month)
        weekday_code = WEEKDAY_CODES[dt.weekday()] if dt else None
        day_names = set()
        for c_idx, cell in enumerate(row):
            if not cell or c_idx >= len(work_posts):
                continue
            if not cell_availability.get((r_idx, c_idx), True):
                continue
            if exclusion_checker and exclusion_checker(r_idx, c_idx):
                continue
            try:
                existing_names = context.name_resolver(cell.get())
            except Exception:
                existing_names = []
            day_names.update(nm for nm in existing_names if nm in counts_week_days)
        for nm in day_names:
            if weekday_code and weekday_code in profile_exclusions.get(nm, set()):
                continue
            if dtype == "we":
                counts_we_days[nm].add(day_num)
            else:
                counts_week_days[nm].add(day_num)

    # Cases a remplir
    cases = []
    open_week_days = set()
    open_we_days = set()
    for r_idx, row in enumerate(planning_gui.table_entries):
        in_month, day_num = _in_month(r_idx)
        if not in_month:
            continue
        try:
            dt = date(year, month, day_num)
        except Exception:
            continue
        dtype = _day_type(dt, hol_month)
        weekday_code = WEEKDAY_CODES[dt.weekday()] if dt else None
        for c_idx, cell in enumerate(row):
            if not cell or c_idx >= len(work_posts):
                continue
            if not cell_availability.get((r_idx, c_idx), True):
                continue
            if exclusion_checker and exclusion_checker(r_idx, c_idx):
                continue
            try:
                if cell.get().strip():
                    continue
            except Exception:
                continue
            cases.append((r_idx, c_idx, day_num, dtype, weekday_code))
            if dtype == "we":
                open_we_days.add(day_num)
            else:
                open_week_days.add(day_num)

    # Cibles mensuelles (volume reel du mois) mesurees en jours
    targets_week = {}
    targets_we = {}
    for p in profiles:
        excluded = set(p.get("excluded_weekdays", set()))
        effective_week = []
        for d in open_week_days:
            code = _weekday_code_for_day(d)
            if code and code in excluded:
                continue
            effective_week.append(d)
        effective_we = []
        for d in open_we_days:
            code = _weekday_code_for_day(d)
            if code and code in excluded:
                continue
            effective_we.append(d)
        targets_week[p["initial"]] = p["participation"] * len(effective_week)
        targets_we[p["initial"]] = p["participation"] * len(effective_we)

    weekend_block_posts = set()
    try:
        weekend_block_posts = set(WEEKEND_BLOCK_POSTS)
    except Exception:
        weekend_block_posts = set()
    # Compatibilité : ancien toggle global activé = appliquer à toutes les lignes
    if ENABLE_WEEKEND_BLOCKS and not weekend_block_posts:
        weekend_block_posts = set(work_posts)

    def _update_counts(profile_initial, dtype, day_num):
        if dtype == "we":
            counts_we_days.setdefault(profile_initial, set()).add(day_num)
        else:
            counts_week_days.setdefault(profile_initial, set()).add(day_num)

    def _weekend_block_days(day_num):
        """
        Retourne les jours (numériques) d'un bloc week-end (ven-sam-dim) englobant day_num.
        On reste dans le mois courant.
        """
        try:
            dt = date(year, month, day_num)
        except Exception:
            return []
        # Vendredi = 4
        offset = max(0, dt.weekday() - 4)
        start = dt - timedelta(days=offset)
        result = []
        for i in range(3):
            d = start + timedelta(days=i)
            if d.month == month:
                result.append(d.day)
        return result

    def _pick_candidate(candidate_entries, dtype):
        """
        Selection equilibree sur ratio count/target pour le type de jour.
        Tie-break : deficit le plus eleve puis un leger alea pour eviter tout biais d'ordre.
        Applique un bonus pour les preferes et un malus pour les ratios deja > 1.
        """
        if not candidate_entries:
            return None
        target_map = targets_we if dtype == "we" else targets_week
        count_map = counts_we_days if dtype == "we" else counts_week_days
        PREF_BONUS = 0.15  # reduit artificiellement le ratio pour les preferes
        OVER_PENALTY = 0.35  # augmente le ratio si deja au-dessus de la cible

        scored = []
        for p, is_pref in candidate_entries:
            cur = len(count_map.get(p["initial"], set()))
            tgt = target_map.get(p["initial"], 0.0)
            if tgt <= 0:
                continue
            ratio = cur / tgt
            over = max(0.0, ratio - 1.0)
            effective_ratio = ratio + OVER_PENALTY * over
            if is_pref:
                effective_ratio = max(0.0, effective_ratio - PREF_BONUS)
            deficit = tgt - cur
            jitter = random.random()
            scored.append((effective_ratio, deficit, jitter, p))

        if not scored:
            return None

        scored.sort(key=lambda x: (x[0], -x[1], x[2]))
        return scored[0][3]

    def _assign_profile_to_cell(profile, r_idx, c_idx, day_num, dtype, weekday_code=None, allow_weekend_block=True):
        """Affecte le profil sur (jour, poste) et sur les postes associes eligibles le meme jour."""
        try:
            cell = planning_gui.table_entries[r_idx][c_idx]
            if not cell:
                return False
        except Exception:
            return False

        try:
            cell.delete(0, "end")
            cell.insert(0, profile["initial"])
        except Exception:
            return False

        _update_counts(profile["initial"], dtype, day_num)

        if weekday_code is None:
            weekday_code = _weekday_code_for_day(day_num)

        assoc_indices = profile_assoc_map.get(profile["initial"], {}).get(c_idx, set()) or set()
        for other_idx in assoc_indices:
            try:
                other_cell = planning_gui.table_entries[r_idx][other_idx]
            except Exception:
                continue
            if not other_cell or other_idx >= len(work_posts):
                continue
            if not cell_availability.get((r_idx, other_idx), True):
                continue
            if exclusion_checker and exclusion_checker(r_idx, other_idx):
                continue
            try:
                if other_cell.get().strip():
                    continue
            except Exception:
                continue
            other_name = work_posts[other_idx]
            if not _is_available(profile, r_idx, day_num, other_idx, other_name, dtype, weekday_code):
                continue
            try:
                other_cell.delete(0, "end")
                other_cell.insert(0, profile["initial"])
            except Exception:
                continue
            _update_counts(profile["initial"], dtype, day_num)
            try:
                planning_gui.auto_resize_column(other_idx)
            except Exception:
                pass

        # Bloc week-end : remplir le m\u00eame poste sur ven/sam/dim du bloc si l'option est activ\u00e9e
        current_post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""
        if allow_weekend_block and dtype == "we" and current_post_name in weekend_block_posts:
            for other_day_num in _weekend_block_days(day_num):
                other_r_idx = day_to_row.get(other_day_num)
                if other_r_idx is None or other_r_idx == r_idx:
                    continue
                try:
                    other_dt = date(year, month, other_day_num)
                except Exception:
                    continue
                other_dtype = _day_type(other_dt, hol_month)
                if other_dtype != "we":
                    continue
                other_weekday_code = WEEKDAY_CODES[other_dt.weekday()]
                try:
                    other_cell = planning_gui.table_entries[other_r_idx][c_idx]
                    if not other_cell or other_cell.get().strip():
                        continue
                except Exception:
                    continue
                other_post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""
                if not _is_available(profile, other_r_idx, other_day_num, c_idx, other_post_name, other_dtype, other_weekday_code):
                    continue
                _assign_profile_to_cell(profile, other_r_idx, c_idx, other_day_num, other_dtype, other_weekday_code, allow_weekend_block=False)

        context.clear_caches()
        try:
            planning_gui.auto_resize_column(c_idx)
        except Exception:
            pass
        return True

    def _assign_weekend_block(r_idx, c_idx, day_num, dtype, weekday_code):
        """
        Force un bloc ven-sam-dim cohérent sur un poste marqué "bloc week-end".
        - Si une des 3 cases du bloc est déjà remplie, on étend le même profil sur les autres.
        - Sinon, on choisit un profil disponible sur les 3 jours, puis on remplit les 3.
        Retourne True si quelque chose a été affecté, False sinon.
        """
        if dtype != "we":
            return False
        post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""
        if post_name not in weekend_block_posts:
            return False

        block_cells = []
        for other_day_num in _weekend_block_days(day_num):
            other_r_idx = day_to_row.get(other_day_num)
            if other_r_idx is None:
                continue
            try:
                other_dt = date(year, month, other_day_num)
            except Exception:
                continue
            other_dtype = _day_type(other_dt, hol_month)
            if other_dtype != "we":
                continue
            other_weekday_code = WEEKDAY_CODES[other_dt.weekday()]
            block_cells.append((other_r_idx, other_day_num, other_dtype, other_weekday_code))

        # 1) Cas où un profil est déjà posé sur au moins un jour du bloc
        existing_profile = None
        for br_idx, b_day, b_dtype, b_code in block_cells:
            try:
                val = planning_gui.table_entries[br_idx][c_idx].get().strip()
            except Exception:
                val = ""
            if not val:
                continue
            existing_profile = profile_by_initial.get(val)
            if existing_profile:
                break

        if existing_profile:
            filled = False
            for br_idx, b_day, b_dtype, b_code in block_cells:
                try:
                    cell = planning_gui.table_entries[br_idx][c_idx]
                    if not cell or cell.get().strip():
                        continue
                except Exception:
                    continue
                if not _is_available(existing_profile, br_idx, b_day, c_idx, post_name, b_dtype, b_code):
                    continue
                _assign_profile_to_cell(existing_profile, br_idx, c_idx, b_day, b_dtype, b_code, allow_weekend_block=False)
                filled = True
            return filled

        # 2) Bloc vide : on cherche un profil disponible sur les 3 jours
        candidates = []
        for p in profiles:
            ok = True
            for br_idx, b_day, b_dtype, b_code in block_cells:
                if not _is_available(p, br_idx, b_day, c_idx, post_name, b_dtype, b_code):
                    ok = False
                    break
            if ok:
                candidates.append((p, post_name in p.get("preferred", [])))

        random.shuffle(candidates)
        chosen = _pick_candidate(candidates, dtype) if candidates else None
        if chosen is None:
            return False

        for br_idx, b_day, b_dtype, b_code in block_cells:
            try:
                cell = planning_gui.table_entries[br_idx][c_idx]
                if not cell or cell.get().strip():
                    continue
            except Exception:
                continue
            if not _is_available(chosen, br_idx, b_day, c_idx, post_name, b_dtype, b_code):
                continue
            _assign_profile_to_cell(chosen, br_idx, c_idx, b_day, b_dtype, b_code, allow_weekend_block=False)
        return True

    # Parcours aléatoire des cases pour limiter les biais d'ordre (jour/colonne)
    random.shuffle(cases)

    for (r_idx, c_idx, day_num, dtype, weekday_code) in cases:
        try:
            current_cell = planning_gui.table_entries[r_idx][c_idx]
            if not current_cell or current_cell.get().strip():
                continue
        except Exception:
            continue

        post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""

        # Si un autre jour du bloc a déjà été affecté, on force l'homogénéité du bloc
        if _assign_weekend_block(r_idx, c_idx, day_num, dtype, weekday_code):
            continue

        candidate_entries = []
        for p in profiles:
            if not _is_available(p, r_idx, day_num, c_idx, post_name, dtype, weekday_code):
                continue
            candidate_entries.append((p, post_name in p.get("preferred", [])))

        random.shuffle(candidate_entries)  # melange a chaque case pour casser les biais d'ordre

        chosen = _pick_candidate(candidate_entries, dtype)

        if chosen is None:
            continue

        _assign_profile_to_cell(chosen, r_idx, c_idx, day_num, dtype, weekday_code)
