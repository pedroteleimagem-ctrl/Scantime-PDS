import random
from pathlib import Path
import sys
from typing import Dict, Iterable, Set, Tuple

_MODULE_DIR = Path(__file__).resolve().parent
if str(_MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(_MODULE_DIR))

try:
    from eligibility import AssignmentSettings, PlanningContext, candidate_is_available, parse_constraint_row
except ModuleNotFoundError:
    trash_dir = _MODULE_DIR / "trash"
    if trash_dir.is_dir():
        sys.path.insert(0, str(trash_dir))
    from eligibility import AssignmentSettings, PlanningContext, candidate_is_available, parse_constraint_row  # noqa: E402

del _MODULE_DIR

# Variables globales pour la limitation d'affectation et l'option de postes diffÃ©rents dans la mÃªme Journée
ENABLE_MAX_ASSIGNMENTS = True              # Active/dÃ©sactive la limitation par poste (nombre max par semaine)
MAX_ASSIGNMENTS_PER_POST = 2               # Nombre maximal d'affectations par poste par semaine
ENABLE_DIFFERENT_POST_PER_DAY = False      # Si True, empÃªche la mÃªme affectation (matin & aprÃ¨s-midi) dans un mÃªme poste
ENABLE_REPOS_SECURITE = False              # Nouvelle option pour activer le repos de sÃ©curitÃ©


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
    Assigne automatiquement les crénaux du mois en respectant uniquement les contraintes
    du tableau (participation %, absences, postes non assurés, doublons sur la même journée),
    avec cibles séparées semaine vs week-end/jours fériés et tirage pondéré.
    """
    import calendar
    from datetime import date, timedelta
    from Full_GUI import days, work_posts, extract_names_from_cell

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
        return [p.strip() for p in str(text or "").split(",") if p.strip()]

    def _holidays_for(m_year, m_month):
        if _month_holidays:
            try:
                return set(_month_holidays(m_year, m_month))
            except Exception:
                return set()
        return set()

    def _day_type(dt, hol_set):
        """week = lun-jeu non férié, we = ven/sa/di/férié/veille férié."""
        is_hol = dt in hol_set
        next_hol = (dt + timedelta(days=1)) in hol_set
        wd = dt.weekday()
        if wd >= 5 or wd == 4 or is_hol or next_hol:
            return "we"
        return "week"

    # Totaux annuels (même liste de postes toute l'année)
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
            abs_txt = row[4].var.get()
        except Exception:
            try:
                abs_txt = row[4].cget("text")
            except Exception:
                abs_txt = ""
        absences = set()
        for part_item in _split_csv(abs_txt):
            try:
                num = int(part_item)
                absences.add(num)
            except Exception:
                continue

        profiles.append(
            {
                "initial": init,
                "participation": part,
                "preferred": preferred,
                "non_assured": non_assured,
                "absences": absences,
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

    from eligibility import PlanningContext, AssignmentSettings

    context = PlanningContext(
        table_entries=planning_gui.table_entries,
        name_resolver=lambda raw: extract_names_from_cell(raw, parser_valids),
        exclusion_checker=exclusion_checker,
        excluded_cells=excluded_cells,
    )

    settings = AssignmentSettings(
        enable_max_assignments=False,
        max_assignments_per_post=None,
        enable_different_post_per_day=False,
        enable_repos_securite=False,
        forbidden_morning_to_afternoon={},
        forbidden_afternoon_to_morning={},
    )

    def _is_available(profile, day_idx, day_num, post_idx, post_name):
        if day_num in profile["absences"]:
            return False
        if post_name in profile["non_assured"]:
            return False
        if context.already_assigned_in_timeslot(profile["initial"], day_idx, True):
            return False
        return True

    # Comptage des assignations existantes (pour pondérer)
    counts_week = {p["initial"]: 0 for p in profiles}
    counts_we = {p["initial"]: 0 for p in profiles}

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
                existing_names = context.name_resolver(cell.get())
            except Exception:
                existing_names = []
            for nm in existing_names:
                if nm not in counts_week:
                    continue
                if dtype == "we":
                    counts_we[nm] += 1
                else:
                    counts_week[nm] += 1

    # Cases à remplir (on ignore les cellules déjà renseignées ou exclues)
    cases_we = []
    cases_week = []
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
            slot = (r_idx, c_idx, day_num, dtype)
            if dtype == "we":
                cases_we.append(slot)
            else:
                cases_week.append(slot)

    random.shuffle(cases_we)
    random.shuffle(cases_week)

    month_week_total = len(cases_week)
    month_we_total = len(cases_we)

    targets_week = {p["initial"]: p["participation"] * month_week_total for p in profiles}
    targets_we = {p["initial"]: p["participation"] * month_we_total for p in profiles}

    def _assign_slots(slots, target_map, count_map):
        for (r_idx, c_idx, day_num, dtype) in slots:
            post_name = work_posts[c_idx] if c_idx < len(work_posts) else ""
            eligible = []
            for p in profiles:
                if not _is_available(p, r_idx, day_num, c_idx, post_name):
                    continue
                eligible.append(p)
            if not eligible:
                continue

            weights = []
            for p in eligible:
                cur = count_map[p["initial"]]
                tgt = target_map.get(p["initial"], 0.0)
                weights.append(max(1e-6, tgt - cur))

            total_w = sum(weights)
            if total_w <= 0:
                chosen = random.choice(eligible)
            else:
                pick = random.random() * total_w
                acc = 0.0
                chosen = eligible[-1]
                for p, w in zip(eligible, weights):
                    acc += w
                    if pick <= acc:
                        chosen = p
                        break

            try:
                cell = planning_gui.table_entries[r_idx][c_idx]
                cell.delete(0, "end")
                cell.insert(0, chosen["initial"])
            except Exception:
                continue

            if dtype == "we":
                counts_we[chosen["initial"]] += 1
            else:
                counts_week[chosen["initial"]] += 1

            context.clear_caches()
            try:
                planning_gui.auto_resize_column(c_idx)
            except Exception:
                pass

    # Ordre : week-ends / fériés d'abord, puis semaine
    _assign_slots(cases_we, targets_we, counts_we)
    _assign_slots(cases_week, targets_week, counts_week)
