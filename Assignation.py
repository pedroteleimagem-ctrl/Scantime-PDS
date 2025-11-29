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
    Assigne automatiquement les cr?neaux (mode mensuel : 1 cellule par jour/astreinte).
    """
    from Full_GUI import days, work_posts, extract_names_from_cell

    rows = getattr(constraints_app, "rows", []) if constraints_app is not None else []
    if not rows:
        return

    profiles = []
    valid_initials = set()
    pref_count = {}

    for cand_row in rows:
        profile = parse_constraint_row(cand_row)
        if not profile:
            continue
        profiles.append(profile)
        valid_initials.add(profile.initial)
        pref_count.setdefault(profile.initial, 0)

    if not profiles:
        return

    parser_valids = valid_initials or None

    def _names_from_cell(raw_value):
        return extract_names_from_cell(raw_value, parser_valids)

    exclusion_checker = getattr(planning_gui, "is_cell_excluded_from_count", None)

    def _is_excluded(row, col):
        if exclusion_checker is None:
            return False
        try:
            return exclusion_checker(row, col)
        except Exception:
            return False

    context = PlanningContext(
        table_entries=planning_gui.table_entries,
        name_resolver=_names_from_cell,
        exclusion_checker=exclusion_checker,
        excluded_cells=getattr(planning_gui, "excluded_from_count", set()),
    )

    forbidden_morning, forbidden_afternoon = build_forbidden_maps(work_posts)

    settings = AssignmentSettings(
        enable_max_assignments=ENABLE_MAX_ASSIGNMENTS,
        max_assignments_per_post=MAX_ASSIGNMENTS_PER_POST,
        enable_different_post_per_day=ENABLE_DIFFERENT_POST_PER_DAY,
        enable_repos_securite=ENABLE_REPOS_SECURITE,
        forbidden_morning_to_afternoon=forbidden_morning,
        forbidden_afternoon_to_morning=forbidden_afternoon,
    )

    total_rows = len(planning_gui.table_entries)

    def assign_slot(day_index, post_index):
        row_index = day_index
        col_index = post_index
        if row_index >= total_rows:
            return
        if _is_excluded(row_index, col_index):
            return
        cell = planning_gui.table_entries[row_index][col_index]
        if not cell:
            return
        try:
            if cell.get().strip():
                return
        except Exception:
            return

        context.clear_caches()
        try:
            post_name = work_posts[post_index] if post_index < len(work_posts) else ""
            eligible = []
            for profile in profiles:
                if not candidate_is_available(
                    profile,
                    context,
                    day_index=day_index,
                    post_index=post_index,
                    is_morning=True,
                    post_name=post_name,
                    settings=settings,
                ):
                    continue
                priority = post_name in profile.preferred_posts and pref_count.get(profile.initial, 0) < 2
                eligible.append((profile, priority))

            if not eligible:
                return

            prioritized = [item for item in eligible if item[1]]
            chosen_profile = prioritized[0][0] if prioritized else eligible[0][0]

            try:
                cell.delete(0, "end")
            except Exception:
                pass
            try:
                cell.insert(0, chosen_profile.initial)
            except Exception:
                return

            try:
                planning_gui.auto_resize_column(col_index)
            except Exception:
                pass

            if post_name in chosen_profile.preferred_posts and pref_count.get(chosen_profile.initial, 0) < 2:
                pref_count[chosen_profile.initial] = pref_count.get(chosen_profile.initial, 0) + 1
        finally:
            context.clear_caches()

    for day_index, _day in enumerate(days):
        for post_index in range(len(work_posts)):
            assign_slot(day_index, post_index)


    rows = getattr(constraints_app, "rows", []) if constraints_app is not None else []
    if not rows:
        return

    profiles = []
    valid_initials = set()
    pref_count = {}

    for cand_row in rows:
        profile = parse_constraint_row(cand_row)
        if not profile:
            continue
        profiles.append(profile)
        valid_initials.add(profile.initial)
        pref_count.setdefault(profile.initial, 0)

    if not profiles:
        return

    parser_valids = valid_initials or None

    def _names_from_cell(raw_value):
        return extract_names_from_cell(raw_value, parser_valids)

    exclusion_checker = getattr(planning_gui, "is_cell_excluded_from_count", None)

    def _is_excluded(row, col):
        if exclusion_checker is None:
            return False
        try:
            return exclusion_checker(row, col)
        except Exception:
            return False

    context = PlanningContext(
        table_entries=planning_gui.table_entries,
        name_resolver=_names_from_cell,
        exclusion_checker=exclusion_checker,
        excluded_cells=getattr(planning_gui, "excluded_from_count", set()),
    )

    forbidden_morning, forbidden_afternoon = build_forbidden_maps(work_posts)

    settings = AssignmentSettings(
        enable_max_assignments=ENABLE_MAX_ASSIGNMENTS,
        max_assignments_per_post=MAX_ASSIGNMENTS_PER_POST,
        enable_different_post_per_day=ENABLE_DIFFERENT_POST_PER_DAY,
        enable_repos_securite=ENABLE_REPOS_SECURITE,
        forbidden_morning_to_afternoon=forbidden_morning,
        forbidden_afternoon_to_morning=forbidden_afternoon,
    )

    num_days = len(planning_gui.table_entries)
    num_posts = len(work_posts)

    def assign_slot(day_index, post_index):
        if day_index >= num_days or post_index >= num_posts:
            return
        if _is_excluded(day_index, post_index):
            return
        cell = planning_gui.table_entries[day_index][post_index]
        if not cell:
            return
        try:
            if cell.get().strip():
                return
        except Exception:
            return

        context.clear_caches()
        try:
            post_name = work_posts[post_index] if post_index < len(work_posts) else ""
            eligible = []
            for profile in profiles:
                if not candidate_is_available(
                    profile,
                    context,
                    day_index=day_index,
                    post_index=post_index,
                    is_morning=True,
                    post_name=post_name,
                    settings=settings,
                ):
                    continue
                priority = post_name in profile.preferred_posts and pref_count.get(profile.initial, 0) < 2
                eligible.append((profile, priority))

            if not eligible:
                return

            prioritized = [item for item in eligible if item[1]]
            chosen_profile = random.choice(prioritized)[0] if prioritized else random.choice(eligible)[0]

            try:
                cell.delete(0, "end")
            except Exception:
                pass
            try:
                cell.insert(0, chosen_profile.initial)
            except Exception:
                return

            try:
                planning_gui.auto_resize_column(post_index)
            except Exception:
                pass

            if post_name in chosen_profile.preferred_posts and pref_count.get(chosen_profile.initial, 0) < 2:
                pref_count[chosen_profile.initial] = pref_count.get(chosen_profile.initial, 0) + 1
        finally:
            context.clear_caches()

    for day_index, _day in enumerate(days):
        for post_index in range(num_posts):
            assign_slot(day_index, post_index)

