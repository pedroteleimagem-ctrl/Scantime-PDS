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

FORBIDDEN_POST_ASSOCIATIONS: Set[Tuple[str, str]] = set()


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

    def _normalize_scope(raw):
        val = str(raw or "").strip().lower()
        if val in {"weekdays_only", "weekday_only", "weekdays"}:
            return "weekdays_only"
        if val in {"weekends_only", "weekend_only", "weekend"}:
            return "weekends_only"
        return "all"

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
            action_btn = row[-1]
            if getattr(action_btn, "_is_row_action_button", False):
                if hasattr(action_btn, "_var"):
                    scope_raw = action_btn._var.get()
                else:
                    scope_raw = action_btn.cget("text")
        except Exception:
            scope_raw = None
        scope = _normalize_scope(scope_raw)

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
                "associations": associations,
            }
        )

    if not profiles:
        return

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

    def _is_available(profile, day_idx, day_num, post_idx, post_name, day_type):
        scope = profile.get("scope", "all")
        if scope == "weekdays_only" and day_type == "we":
            return False
        if scope == "weekends_only" and day_type == "week":
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
        dtype = _day_type(date(year, month, day_num), hol_month)
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
            if dtype == "we":
                counts_we_days[nm].add(r_idx)
            else:
                counts_week_days[nm].add(r_idx)

    # Cases a remplir
    cases = []
    week_slots = 0
    we_slots = 0
    for r_idx, row in enumerate(planning_gui.table_entries):
        in_month, day_num = _in_month(r_idx)
        if not in_month:
            continue
        dtype = _day_type(date(year, month, day_num), hol_month)
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
            cases.append((r_idx, c_idx, day_num, dtype))
            if dtype == "we":
                we_slots += 1
            else:
                week_slots += 1

    # Cibles mensuelles (volume r\u00e9el du mois)
    targets_week = {}
    targets_we = {}
    for p in profiles:
        scope = p.get("scope", "all")
        targets_week[p["initial"]] = (p["participation"] * week_slots) if scope != "weekends_only" else 0.0
        targets_we[p["initial"]] = (p["participation"] * we_slots) if scope != "weekdays_only" else 0.0

    def _update_counts(profile_initial, dtype, day_idx):
        if dtype == "we":
            counts_we_days.setdefault(profile_initial, set()).add(day_idx)
        else:
            counts_week_days.setdefault(profile_initial, set()).add(day_idx)

    def _pick_candidate(candidate_profiles, dtype):
        """Tirage pondere en fonction du deficit par rapport a la cible du type de jour."""
        if not candidate_profiles:
            return None
        target_map = targets_we if dtype == "we" else targets_week
        count_map = counts_we_days if dtype == "we" else counts_week_days
        weighted = []
        for p in candidate_profiles:
            cur = len(count_map.get(p["initial"], set()))
            tgt = target_map.get(p["initial"], 0.0)
            deficit = tgt - cur
            weight = deficit if deficit > 0 else 0.1
            weighted.append((p, max(weight, 1e-6), deficit))
        positives = [(p, w) for (p, w, d) in weighted if d > 0]
        if positives:
            pool = positives
            total = sum(w for _, w in pool)
            if total <= 0:
                return random.choice([p for p, _ in pool])
            pick = random.random() * total
            acc = 0.0
            for p, w in pool:
                acc += w
                if pick <= acc:
                    return p
            return pool[-1][0]

        # Aucun d\u00e9ficit positif : on prend ceux avec le plus faible ratio count/target (ou count si target=0)
        ratios = []
        for p in candidate_profiles:
            cur = len(count_map.get(p["initial"], set()))
            tgt = target_map.get(p["initial"], 0.0)
            ratio = (cur / tgt) if tgt > 0 else float(cur)
            ratios.append((p, ratio))
        min_ratio = min(r for _, r in ratios) if ratios else 0
        pool = [p for (p, r) in ratios if r == min_ratio]
        return random.choice(pool) if pool else None

    def _assign_profile_to_cell(profile, r_idx, c_idx, day_num, dtype):
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

        _update_counts(profile["initial"], dtype, r_idx)

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
            if not _is_available(profile, r_idx, day_num, other_idx, other_name, dtype):
                continue
            try:
                other_cell.delete(0, "end")
                other_cell.insert(0, profile["initial"])
            except Exception:
                continue
            _update_counts(profile["initial"], dtype, r_idx)
            try:
                planning_gui.auto_resize_column(other_idx)
            except Exception:
                pass

        context.clear_caches()
        try:
            planning_gui.auto_resize_column(c_idx)
        except Exception:
            pass
        return True

    # Parcours chronologique : lignes (jours) puis colonnes (postes)
    cases.sort(key=lambda tpl: (tpl[2], tpl[1]))

    for (r_idx, c_idx, day_num, dtype) in cases:
        try:
            current_cell = planning_gui.table_entries[r_idx][c_idx]
            if not current_cell or current_cell.get().strip():
                continue
        except Exception:
            continue

        post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""

        pref_candidates = []
        for p in profiles:
            if post_name not in p.get("preferred", []):
                continue
            if not _is_available(p, r_idx, day_num, c_idx, post_name, dtype):
                continue
            pref_candidates.append(p)
        random.shuffle(pref_candidates)

        chosen = _pick_candidate(pref_candidates, dtype)
        if chosen is None:
            other_candidates = []
            for p in profiles:
                if post_name in p.get("preferred", []):
                    continue
                if not _is_available(p, r_idx, day_num, c_idx, post_name, dtype):
                    continue
                other_candidates.append(p)
            random.shuffle(other_candidates)
            chosen = _pick_candidate(other_candidates, dtype)

        if chosen is None:
            continue

        _assign_profile_to_cell(chosen, r_idx, c_idx, day_num, dtype)
