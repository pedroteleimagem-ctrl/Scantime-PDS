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
ENABLE_WEEKDAY_COMPENSATION = False
WEEKDAY_COMPENSATION_MALUS = 0.8  # penalisation forte mais non bloquante
OPTIMIZE_BALANCE = False
OPTIMIZE_BALANCE = False

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

    # Fenetres de penalisation (compensation semaine autour d'un bloc WE)
    weekday_compensation_penalties: Dict[Tuple[int, int], Set[str]] = {}

    def _compensation_windows_for_block(friday_dt):
        """Retourne les jours (num) lun-jeu de la semaine precedente et suivante pour un vendredi donne."""
        before = []
        after = []
        for offset in range(-4, 0):  # lun-jeu avant le bloc
            dt = friday_dt + timedelta(days=offset)
            if dt.month == month and dt.weekday() <= 3:
                before.append(dt.day)
        for offset in range(3, 7):  # lun-jeu apres le bloc
            dt = friday_dt + timedelta(days=offset)
            if dt.month == month and dt.weekday() <= 3:
                after.append(dt.day)
        return before, after

    def _register_compensation_window(post_idx, friday_dt, profile_initial):
        if not (ENABLE_WEEKDAY_COMPENSATION and friday_dt and profile_initial):
            return
        before, after = _compensation_windows_for_block(friday_dt)
        for d in before + after:
            weekday_compensation_penalties.setdefault((post_idx, d), set()).add(profile_initial)

    def _block_assigned_initial(post_idx, block_days):
        """Retourne l'initiale si les 3 jours du bloc sont remplis par le meme profil, sinon None."""
        if not ENABLE_WEEKDAY_COMPENSATION:
            return None
        initial = None
        for d in block_days:
            r_idx = day_to_row.get(d)
            if r_idx is None:
                return None
            try:
                cell = planning_gui.table_entries[r_idx][post_idx]
            except Exception:
                return None
            if not cell:
                return None
            try:
                names = context.name_resolver(cell.get())
            except Exception:
                names = []
            if len(names) != 1:
                return None
            if initial is None:
                initial = names[0]
            elif initial != names[0]:
                return None
        return initial

    def _maybe_register_compensation_for_block(post_idx, block_days):
        """Inspecte le bloc ven-sam-dim et ajoute la penalisation si le bloc est complet."""
        if not (ENABLE_WEEKDAY_COMPENSATION and block_days and 0 <= post_idx < len(work_posts)):
            return
        try:
            friday_dt = next(
                date(year, month, d) for d in block_days
                if date(year, month, d).weekday() == 4
            )
        except Exception:
            return
        block_initial = _block_assigned_initial(post_idx, block_days)
        if block_initial:
            _register_compensation_window(post_idx, friday_dt, block_initial)

    def _recompute_compensation_from_table():
        """Recalcule toutes les fenetres de compensation depuis l'etat courant du planning."""
        weekday_compensation_penalties.clear()
        if not (ENABLE_WEEKDAY_COMPENSATION and weekend_block_posts):
            return
        for post_name in weekend_block_posts:
            post_idx = post_index_map.get(post_name)
            if post_idx is None:
                continue
            for d in range(1, days_in_month + 1):
                try:
                    dt = date(year, month, d)
                except Exception:
                    continue
                if _day_type(dt, hol_month) != "we":
                    continue
                block_days = _weekend_block_days(d)
                # Exiger un vendredi dans le mois courant pour ancrer la compensation
                if not any(date(year, month, b).weekday() == 4 for b in block_days if 1 <= b <= days_in_month):
                    continue
                block_days_in_month = [b for b in block_days if 1 <= b <= days_in_month]
                if len(block_days_in_month) < 3:
                    continue
                block_initial = _block_assigned_initial(post_idx, block_days_in_month)
                if block_initial:
                    _maybe_register_compensation_for_block(post_idx, block_days_in_month)

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

    def _pick_candidate(candidate_entries, dtype, day_num=None, post_idx=None):
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
            if ENABLE_WEEKDAY_COMPENSATION and day_num is not None and post_idx is not None:
                penalized = weekday_compensation_penalties.get((post_idx, day_num))
                if penalized and p["initial"] in penalized:
                    effective_ratio += WEEKDAY_COMPENSATION_MALUS
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
            block_days = _weekend_block_days(day_num)
            for other_day_num in block_days:
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
            if len(block_days) >= 3:
                _maybe_register_compensation_for_block(c_idx, block_days)

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
        block_day_nums = [b_day for _br, b_day, _bdt, _bc in block_cells]

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
            if filled and len(block_cells) >= 3:
                _maybe_register_compensation_for_block(c_idx, block_day_nums)
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
        if len(block_cells) >= 3:
            _maybe_register_compensation_for_block(c_idx, block_day_nums)
        return True

    # Parcours aléatoire des cases pour limiter les biais d'ordre (jour/colonne)
    _recompute_compensation_from_table()

    block_cases = [
        (r_idx, c_idx, day_num, dtype, weekday_code)
        for (r_idx, c_idx, day_num, dtype, weekday_code) in cases
        if dtype == "we" and (work_posts[c_idx] if c_idx < len(work_posts) else "") in weekend_block_posts
    ]
    random.shuffle(block_cases)
    for (r_idx, c_idx, day_num, dtype, weekday_code) in block_cases:
        try:
            current_cell = planning_gui.table_entries[r_idx][c_idx]
            if not current_cell or current_cell.get().strip():
                continue
        except Exception:
            continue
        post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""
        if _assign_weekend_block(r_idx, c_idx, day_num, dtype, weekday_code):
            continue
        candidate_entries = []
        for p in profiles:
            if not _is_available(p, r_idx, day_num, c_idx, post_name, dtype, weekday_code):
                continue
            candidate_entries.append((p, post_name in p.get("preferred", [])))
        random.shuffle(candidate_entries)
        chosen = _pick_candidate(candidate_entries, dtype, day_num=day_num, post_idx=c_idx)
        if chosen:
            _assign_profile_to_cell(chosen, r_idx, c_idx, day_num, dtype, weekday_code)

    _recompute_compensation_from_table()

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

        chosen = _pick_candidate(candidate_entries, dtype, day_num=day_num, post_idx=c_idx)

        if chosen is None:
            continue

        _assign_profile_to_cell(chosen, r_idx, c_idx, day_num, dtype, weekday_code)


def optimize_month_balance(constraints_app, planning_gui, tabs_data, current_index=None):
    """
    Passe d'optimisation locale sur le mois courant uniquement :
    cherche des swaps entre profils pour réduire l'écart cumulatif (week / we)
    sans violer les contraintes de base. Ne modifie rien si l'option n'est pas activée.
    Retourne la liste des changements effectués.
    """
    if not OPTIMIZE_BALANCE:
        return []
    try:
        if tabs_data is None or len(tabs_data) < 2:
            return []
    except Exception:
        return []

    import random
    from datetime import date, timedelta
    from Full_GUI import days, work_posts, extract_names_from_cell

    # Trouver l'index du mois courant si non fourni
    if current_index is None:
        try:
            for idx, item in enumerate(tabs_data):
                if item and item[0] is planning_gui:
                    current_index = idx
                    break
        except Exception:
            current_index = None
    if current_index is None or current_index <= 0:
        # Pas d'optimisation pour le premier mois ou index inconnu
        return []

    rows = getattr(constraints_app, "rows", []) if constraints_app is not None else []
    if not rows:
        return

    # --- Extraction des profils depuis le tableau de contraintes (copie légère) ---
    def _split_csv(text):
        return [p.strip() for p in str(text or "").replace(";", ",").split(",") if p.strip()]

    def _parse_excluded_weekdays_val(txt):
        txt_lower = str(txt or "").strip().lower()
        if txt_lower in {"weekdays_only", "weekday_only", "weekdays"}:
            return {"fri", "sat", "sun"}
        if txt_lower in {"weekends_only", "weekend_only", "weekend"}:
            return {"mon", "tue", "wed", "thu"}
        parts = [p.strip() for p in txt_lower.replace(";", ",").split(",") if p.strip()]
        return set(parts)

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
        excluded_weekdays = set(_parse_excluded_weekdays_val(scope_raw)) if scope_raw else set()

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
                "excluded_weekdays": excluded_weekdays,
                "associations": associations,
            }
        )

    if not profiles:
        return []

    profile_by_initial = {p["initial"]: p for p in profiles}

    WEEKDAY_CODES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]

    def _weekday_code(dt_obj: date | None):
        try:
            return WEEKDAY_CODES[dt_obj.weekday()]
        except Exception:
            return None

    def _day_type_for_gui(gui_obj, row_idx):
        try:
            label_text = gui_obj.day_labels[row_idx].cget("text").strip()
            day_num = int(label_text)
        except Exception:
            day_num = row_idx + 1
        try:
            year = getattr(gui_obj, "current_year", None)
            month = getattr(gui_obj, "current_month", None)
            dt = date(year, month, day_num) if year and month else None
        except Exception:
            dt = None
        weekend_rows = getattr(gui_obj, "weekend_rows", set())
        holiday_rows = getattr(gui_obj, "holiday_rows", set())
        holiday_dates = getattr(gui_obj, "holiday_dates", set())
        is_weekend = row_idx in weekend_rows
        is_holiday = row_idx in holiday_rows
        if dt:
            try:
                wd = dt.weekday()
                is_weekend = is_weekend or wd >= 5 or wd == 4
                is_holiday = is_holiday or dt in holiday_dates or (dt + timedelta(days=1) in holiday_dates)
            except Exception:
                pass
        dtype = "we" if (is_weekend or is_holiday) else "week"
        return dtype, day_num, dt

    # Comptage cumulatif sur tous les mois jusqu'au courant
    counts_week = {p["initial"]: 0 for p in profiles}
    counts_we = {p["initial"]: 0 for p in profiles}
    assigned_week_total = 0
    assigned_we_total = 0

    def _accumulate_tab(gui_obj, upto_only=False):
        nonlocal assigned_week_total, assigned_we_total
        entries = getattr(gui_obj, "table_entries", [])
        cell_availability = getattr(gui_obj, "cell_availability", {})
        for r_idx, row in enumerate(entries):
            dtype, day_num, dt = _day_type_for_gui(gui_obj, r_idx)
            for c_idx, cell in enumerate(row):
                try:
                    if not cell or not cell_availability.get((r_idx, c_idx), True):
                        continue
                except Exception:
                    continue
                try:
                    val = cell.get().strip()
                except Exception:
                    val = ""
                if not val:
                    continue
                names = extract_names_from_cell(val, initials_set)
                for nm in names:
                    if nm not in counts_week:
                        continue
                    if dtype == "we":
                        counts_we[nm] += 1
                        assigned_we_total += 1
                    else:
                        counts_week[nm] += 1
                        assigned_week_total += 1

    try:
        for idx, (g_obj, _c, _s) in enumerate(tabs_data):
            if g_obj is None:
                continue
            _accumulate_tab(g_obj)
            if idx >= current_index:
                break
    except Exception:
        return

    # Cibles : moyenne des 100% comme référence (sinon fallback proportionnel au total assigné)
    full_time_inits = [p["initial"] for p in profiles if p.get("participation", 1.0) >= 0.99]
    if full_time_inits:
        avg_week_full = sum(counts_week.get(init, 0) for init in full_time_inits) / max(1, len(full_time_inits))
        avg_we_full = sum(counts_we.get(init, 0) for init in full_time_inits) / max(1, len(full_time_inits))
        targets_week = {p["initial"]: p["participation"] * avg_week_full for p in profiles}
        targets_we = {p["initial"]: p["participation"] * avg_we_full for p in profiles}
    else:
        total_part = sum(p.get("participation", 1.0) for p in profiles) or 1.0
        targets_week = {p["initial"]: (p.get("participation", 1.0) / total_part) * assigned_week_total for p in profiles}
        targets_we = {p["initial"]: (p.get("participation", 1.0) / total_part) * assigned_we_total for p in profiles}

    # Préparer les cellules du mois courant
    current_gui = planning_gui
    entries = getattr(current_gui, "table_entries", [])
    cell_availability = getattr(current_gui, "cell_availability", {})
    weekend_block_posts = set()
    try:
        weekend_block_posts = set(WEEKEND_BLOCK_POSTS)
    except Exception:
        weekend_block_posts = set()

    def _is_eligible(profile, r_idx, c_idx, dtype, day_num, dt_obj):
        post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""
        # Cellule active
        try:
            if not cell_availability.get((r_idx, c_idx), True):
                return False
        except Exception:
            return False
        # Absence
        if day_num in profile.get("absences", set()):
            return False
        # Exclusion jour
        code = _weekday_code(dt_obj)
        if code and code in profile.get("excluded_weekdays", set()):
            return False
        # Poste non assuré
        if post_name and post_name in profile.get("non_assured", set()):
            return False
        # Pas de doublon non autorisé : on autorise si le profil est déjà présent ET que le poste est dans ses associations
        row_cells = entries[r_idx] if r_idx < len(entries) else []
        try:
            existing_names = set()
            for cell in row_cells:
                if not cell:
                    continue
                val = cell.get().strip()
                if not val:
                    continue
                existing_names.update(extract_names_from_cell(val, initials_set))
            if profile["initial"] in existing_names:
                if profile.get("associations"):
                    return True
                else:
                    return False
        except Exception:
            pass
        return True

    # Liste des cellules remplies (mois courant) par type/poste
    filled_cells = []
    for r_idx, row in enumerate(entries):
        dtype, day_num, dt_obj = _day_type_for_gui(current_gui, r_idx)
        for c_idx, cell in enumerate(row):
            try:
                if not cell:
                    continue
                val = cell.get().strip()
            except Exception:
                continue
            if not val:
                continue
            names = extract_names_from_cell(val, initials_set)
            for nm in names:
                filled_cells.append({
                    "r": r_idx,
                    "c": c_idx,
                    "dtype": dtype,
                    "day_num": day_num,
                    "dt": dt_obj,
                    "initial": nm,
                })

    def _diff(initial, dtype):
        if dtype == "we":
            return counts_we.get(initial, 0) - targets_we.get(initial, 0.0)
        return counts_week.get(initial, 0) - targets_week.get(initial, 0.0)

    def _apply_swap(cell_a, cell_b):
        r1, c1, init1 = cell_a["r"], cell_a["c"], cell_a["initial"]
        r2, c2, init2 = cell_b["r"], cell_b["c"], cell_b["initial"]
        try:
            entries[r1][c1].delete(0, "end")
            entries[r1][c1].insert(0, init2)
            entries[r2][c2].delete(0, "end")
            entries[r2][c2].insert(0, init1)
        except Exception:
            return False
        if cell_a["dtype"] == "we":
            counts_we[init1] = counts_we.get(init1, 0) - 1
            counts_we[init2] = counts_we.get(init2, 0) + 1
        else:
            counts_week[init1] = counts_week.get(init1, 0) - 1
            counts_week[init2] = counts_week.get(init2, 0) + 1
        cell_a["initial"], cell_b["initial"] = init2, init1
        return True

    # Optimisation : swaps locaux
    improved = True
    iterations = 0
    MAX_ITERS = 100
    changes_log = []
    while improved and iterations < MAX_ITERS:
        improved = False
        iterations += 1
        best_gain = 0
        best_pair = None
        random.shuffle(filled_cells)
        for i in range(len(filled_cells)):
            a = filled_cells[i]
            dtype = a["dtype"]
            # skip bloc week-end (trop complexe à swaper proprement)
            if dtype == "we":
                post_name_a = work_posts[a["c"]] if a["c"] < len(work_posts) else ""
                if post_name_a in weekend_block_posts:
                    continue
            for j in range(i + 1, len(filled_cells)):
                b = filled_cells[j]
                if b["dtype"] != dtype:
                    continue
                if dtype == "we":
                    post_name_b = work_posts[b["c"]] if b["c"] < len(work_posts) else ""
                    if post_name_b in weekend_block_posts:
                        continue
                if a["initial"] == b["initial"]:
                    continue
                prof_a = profile_by_initial.get(a["initial"])
                prof_b = profile_by_initial.get(b["initial"])
                if not prof_a or not prof_b:
                    continue
                # Eligibilité croisée
                if not _is_eligible(prof_a, b["r"], b["c"], dtype, b["day_num"], b["dt"]):
                    continue
                if not _is_eligible(prof_b, a["r"], a["c"], dtype, a["day_num"], a["dt"]):
                    continue
                # Gain potentiel
                delta = 0
                if dtype == "we":
                    da_before = counts_we.get(prof_a["initial"], 0)
                    db_before = counts_we.get(prof_b["initial"], 0)
                    da_after = da_before - 1
                    db_after = db_before + 1
                    delta = (abs(da_before - targets_we.get(prof_a["initial"], 0.0)) +
                             abs(db_before - targets_we.get(prof_b["initial"], 0.0)) -
                             abs(da_after - targets_we.get(prof_a["initial"], 0.0)) -
                             abs(db_after - targets_we.get(prof_b["initial"], 0.0)))
                else:
                    da_before = counts_week.get(prof_a["initial"], 0)
                    db_before = counts_week.get(prof_b["initial"], 0)
                    da_after = da_before - 1
                    db_after = db_before + 1
                    delta = (abs(da_before - targets_week.get(prof_a["initial"], 0.0)) +
                             abs(db_before - targets_week.get(prof_b["initial"], 0.0)) -
                             abs(da_after - targets_week.get(prof_a["initial"], 0.0)) -
                             abs(db_after - targets_week.get(prof_b["initial"], 0.0)))
                if delta > best_gain:
                    best_gain = delta
                    best_pair = (i, j)
        if best_pair is not None and best_gain > 0:
            a = filled_cells[best_pair[0]]
            b = filled_cells[best_pair[1]]
            if _apply_swap(a, b):
                # Ajout au log
                try:
                    day_label = str(days[a["r"]]) if a["r"] < len(days) else str(a["r"] + 1)
                except Exception:
                    day_label = str(a["r"] + 1)
                post_name = work_posts[a["c"]] if a["c"] < len(work_posts) else f"Poste {a['c']+1}"
                changes_log.append({
                    "day": day_label,
                    "post": post_name,
                    "from": b["initial"],
                    "to": a["initial"],
                    "dtype": a["dtype"],
                })
                improved = True
    return changes_log
