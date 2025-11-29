from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from typing import Callable, Dict, FrozenSet, List, Optional, Sequence, Set, Tuple


def _normalize_absence_state(state: str) -> str:
    """Return an uppercase ASCII version of the absence state."""
    if not state:
        return ""
    text = unicodedata.normalize("NFKD", str(state))
    ascii_text = text.encode("ascii", "ignore").decode("ascii")
    return ascii_text.upper().strip()


def _split_csv(text: str) -> List[str]:
    return [part.strip() for part in str(text or "").split(",") if part.strip()]


@dataclass(frozen=True)
class ConstraintProfile:
    initial: str
    quota_total: Optional[int]
    preferred_posts: List[str]
    non_assured_posts: List[str]
    absences: List[str]
    pds_flags: List[bool]

    def is_absent(self, day_index: int, is_morning: bool) -> bool:
        state = ""
        if 0 <= day_index < len(self.absences):
            state = self.absences[day_index]
        if not state:
            return False
        if state == "JOURNEE":
            return True
        if is_morning and state == "MATIN":
            return True
        if not is_morning and state in {"AP MIDI", "APMIDI"}:
            return True
        return False

    def has_pds_previous_day(self, day_index: int) -> bool:
        if day_index <= 0:
            return False
        prev = day_index - 1
        if 0 <= prev < len(self.pds_flags):
            return self.pds_flags[prev]
        return False


def parse_constraint_row(row_widgets: Sequence) -> Optional[ConstraintProfile]:
    try:
        initial_widget = row_widgets[0]
    except (IndexError, TypeError):
        return None

    try:
        initial = str(initial_widget.get()).strip()
    except Exception:
        initial = ""

    if not initial:
        return None

    # Quota (column 1)
    try:
        raw_quota = str(row_widgets[1].get()).strip()
        quota_total = int(raw_quota) if raw_quota else None
    except Exception:
        quota_total = None

    # Preferred posts button (column 2)
    try:
        preferred_text = row_widgets[2].cget("text")
    except Exception:
        preferred_text = ""
    preferred_posts: List[str] = []
    if preferred_text and preferred_text.strip().lower() != "selectionner":
        preferred_posts = _split_csv(preferred_text)

    # Non-assured posts (column 3)
    non_assured_text = ""
    try:
        widget = row_widgets[3]
        var = getattr(widget, "_var", None)
        if var is not None:
            non_assured_text = var.get()
        else:
            raise AttributeError
    except Exception:
        try:
            non_assured_text = row_widgets[3].cget("text")
        except Exception:
            non_assured_text = ""
    non_assured_posts = _split_csv(non_assured_text)

    absences: List[str] = []
    pds_flags: List[bool] = []
    for offset in range(7):
        try:
            tpl = row_widgets[4 + offset]
        except Exception:
            absences.append("")
            pds_flags.append(False)
            continue

        state = ""
        pds_flag = False
        if isinstance(tpl, tuple) and tpl:
            toggle = tpl[0]
            try:
                state = toggle._var.get()
            except Exception:
                try:
                    state = toggle.cget("text")
                except Exception:
                    state = ""
            try:
                pds_flag = bool(tpl[2].get())
            except Exception:
                pds_flag = False
        absences.append(_normalize_absence_state(state))
        pds_flags.append(pds_flag)

    if preferred_posts and non_assured_posts:
        blocked = {p.lower() for p in non_assured_posts}
        preferred_posts = [p for p in preferred_posts if p.lower() not in blocked]

    return ConstraintProfile(
        initial=initial,
        quota_total=quota_total,
        preferred_posts=preferred_posts,
        non_assured_posts=non_assured_posts,
        absences=absences,
        pds_flags=pds_flags,
    )


@dataclass(frozen=True)
class AssignmentSettings:
    enable_max_assignments: bool
    max_assignments_per_post: Optional[int]
    enable_different_post_per_day: bool
    enable_repos_securite: bool
    forbidden_morning_to_afternoon: Dict[int, FrozenSet[int]] = field(default_factory=dict)
    forbidden_afternoon_to_morning: Dict[int, FrozenSet[int]] = field(default_factory=dict)


@dataclass
class PlanningContext:
    table_entries: Sequence[Sequence]
    name_resolver: Callable[[str], Sequence[str]]
    exclusion_checker: Optional[Callable[[int, int], bool]] = None
    excluded_cells: Optional[Set[Tuple[int, int]]] = None
    _total_cache: dict = field(default_factory=dict, init=False, repr=False)
    _post_cache: dict = field(default_factory=dict, init=False, repr=False)
    _slot_cache: dict = field(default_factory=dict, init=False, repr=False)

    def clear_caches(self) -> None:
        self._total_cache.clear()
        self._post_cache.clear()
        self._slot_cache.clear()

    def _should_skip(self, row: int, col: int) -> bool:
        if self.exclusion_checker is not None:
            try:
                if self.exclusion_checker(row, col):
                    return True
            except Exception:
                pass
        if self.excluded_cells and (row, col) in self.excluded_cells:
            return True
        return False

    def names_at(self, row: int, col: int) -> List[str]:
        try:
            cell = self.table_entries[row][col]
        except Exception:
            return []
        if not cell or self._should_skip(row, col):
            return []
        try:
            raw = cell.get()
        except Exception:
            raw = ""
        return list(self.name_resolver(raw))

    def count_total_assignments(self, initial: str) -> int:
        if not initial:
            return 0
        cached = self._total_cache.get(initial)
        if cached is not None:
            return cached
        total = 0
        for r_idx, row in enumerate(self.table_entries):
            for c_idx, _cell in enumerate(row):
                if self._should_skip(r_idx, c_idx):
                    continue
                if initial in self.names_at(r_idx, c_idx):
                    total += 1
        self._total_cache[initial] = total
        return total

    def count_assignments_for_post(self, initial: str, post_index: int) -> int:
        if not initial:
            return 0
        key = (initial, post_index)
        cached = self._post_cache.get(key)
        if cached is not None:
            return cached
        morning_row = post_index * 2
        afternoon_row = morning_row + 1
        total = 0
        for r_idx in (morning_row, afternoon_row):
            if r_idx >= len(self.table_entries):
                continue
            for c_idx, _cell in enumerate(self.table_entries[r_idx]):
                if self._should_skip(r_idx, c_idx):
                    continue
                if initial in self.names_at(r_idx, c_idx):
                    total += 1
        self._post_cache[key] = total
        return total

    def already_assigned_in_timeslot(self, initial: str, day_index: int, is_morning: bool) -> bool:
        if not initial:
            return False
        key = (initial, day_index, is_morning)
        cached = self._slot_cache.get(key)
        if cached is not None:
            return cached
        start = 0 if is_morning else 1
        for r_idx in range(start, len(self.table_entries), 2):
            if self._should_skip(r_idx, day_index):
                continue
            if initial in self.names_at(r_idx, day_index):
                self._slot_cache[key] = True
                return True
        self._slot_cache[key] = False
        return False


def candidate_is_available(
    profile: ConstraintProfile,
    context: PlanningContext,
    *,
    day_index: int,
    post_index: int,
    is_morning: bool,
    post_name: str,
    settings: AssignmentSettings,
) -> bool:
    if not profile.initial:
        return False

    if profile.quota_total is None:
        return False
    if context.count_total_assignments(profile.initial) >= profile.quota_total:
        return False

    if profile.is_absent(day_index, is_morning):
        return False

    if settings.enable_repos_securite and is_morning and profile.has_pds_previous_day(day_index):
        return False

    if profile.non_assured_posts and post_name in profile.non_assured_posts:
        return False

    if context.already_assigned_in_timeslot(profile.initial, day_index, is_morning):
        return False

    if settings.enable_different_post_per_day:
        fallback_same_post = (
            not settings.forbidden_morning_to_afternoon
            and not settings.forbidden_afternoon_to_morning
        )

        if is_morning:
            blocked_afternoon = settings.forbidden_morning_to_afternoon.get(post_index)
            if blocked_afternoon is None and fallback_same_post:
                blocked_afternoon = frozenset({post_index})
            if blocked_afternoon:
                for blocked_idx in blocked_afternoon:
                    row = blocked_idx * 2 + 1
                    if profile.initial in context.names_at(row, day_index):
                        return False
        else:
            blocked_morning = settings.forbidden_afternoon_to_morning.get(post_index)
            if blocked_morning is None and fallback_same_post:
                blocked_morning = frozenset({post_index})
            if blocked_morning:
                for blocked_idx in blocked_morning:
                    row = blocked_idx * 2
                    if profile.initial in context.names_at(row, day_index):
                        return False

    if settings.enable_max_assignments:
        limit = settings.max_assignments_per_post
        if limit is not None and context.count_assignments_for_post(profile.initial, post_index) >= limit:
            return False

    return True
