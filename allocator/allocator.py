from collections import defaultdict
from django.db import transaction
from django.db.models import Count, Q
from .models import Faculty, FacultySlot, Exam, Room, Duty


def overlap(a_start, a_end, b_start, b_end):
    return a_start < b_end and b_start < a_end


def get_free_candidates(exam, exclude_faculty_ids=None):
    """
    Return eligible faculty sorted by lowest active duty count:
    1. Prefer faculty with NO overlapping BUSY slot and NOT already assigned in this exam.
    2. Fallback to any faculty NOT already assigned in this exam (balanced by lowest duty count)
       so rooms never fail with 'NO FREE FACULTY'.
    """
    if exclude_faculty_ids is None:
        exclude_faculty_ids = set()
    else:
        exclude_faculty_ids = set(exclude_faculty_ids)

    already_assigned = set(
        Duty.objects.filter(exam=exam, status="ALLOCATED").values_list("faculty_id", flat=True)
    )
    blocked_ids = exclude_faculty_ids | already_assigned

    weekday = exam.date.strftime("%a").upper()[:3]
    busy_slots = FacultySlot.objects.filter(weekday=weekday, status__iexact="BUSY")
    busy_by_fac = defaultdict(list)
    for s in busy_slots:
        busy_by_fac[s.faculty_id].append(s)

    faculties = list(
        Faculty.objects.exclude(id__in=blocked_ids)
        .annotate(active_count=Count("duty", filter=Q(duty__status="ALLOCATED")))
        .order_by("active_count", "name")
    )

    free_candidates = []
    reserve_candidates = []
    for person in faculties:
        has_conflict = any(
            overlap(s.start_time, s.end_time, exam.start_time, exam.end_time)
            for s in busy_by_fac.get(person.id, [])
        )
        if not has_conflict:
            free_candidates.append(person)
        else:
            reserve_candidates.append(person)

    free_candidates.sort(key=lambda p: (p.active_count, p.name.lower()))
    reserve_candidates.sort(key=lambda p: (p.active_count, p.name.lower()))
    return free_candidates + reserve_candidates


def allocate_duties(exams, rooms, faculty, faculty_slots):
    """
    Allocate a faculty member to every room of every exam with equal workload balancing.
    - Tier 1: Strictly conflict-free faculty (no teaching overlap & not already assigned in this exam).
    - Tier 2: Reliever / Reserve faculty (not already assigned in this exam) when rooms exceed free faculty,
      ensuring every room is covered with zero 'NO FREE FACULTY' alerts.
    """
    assigned = []
    warnings = []
    if not faculty:
        return assigned, warnings

    load = defaultdict(int)
    slot_map = defaultdict(list)
    for slot in faculty_slots:
        slot_map[slot.faculty_id].append(slot)

    for exam in sorted(exams, key=lambda x: (x.date, x.start_time, x.id)):
        weekday = exam.date.strftime("%a").upper()[:3]
        used_in_exam = set()
        exam_rooms = rooms.get(exam.id, [])

        for room in exam_rooms:
            free_cands = []
            reserve_cands = []

            for person in faculty:
                if person.id in used_in_exam:
                    continue
                busy = any(
                    s.weekday == weekday and s.status.upper() == "BUSY"
                    and overlap(s.start_time, s.end_time, exam.start_time, exam.end_time)
                    for s in slot_map.get(person.id, [])
                )
                if not busy:
                    free_cands.append(person)
                else:
                    reserve_cands.append(person)

            if free_cands:
                free_cands.sort(key=lambda p: (load[p.id], p.name.lower()))
                chosen = free_cands[0]
            elif reserve_cands:
                reserve_cands.sort(key=lambda p: (load[p.id], p.name.lower()))
                chosen = reserve_cands[0]
            else:
                # If total rooms in a single session exceed total faculty count, pick lowest-load faculty
                pool = sorted(faculty, key=lambda p: (load[p.id], p.name.lower()))
                chosen = pool[0]

            assigned.append((exam, room, chosen))
            used_in_exam.add(chosen.id)
            load[chosen.id] += 1

    return assigned, warnings


@transaction.atomic
def reallocate_absent_duty(duty_id):
    """
    Find an available replacement with the lowest workload.
    Only modifies the original allocation after a valid replacement is found.
    """
    duty = Duty.objects.select_for_update().select_related("exam", "room", "faculty").get(id=duty_id)
    absent_person = duty.faculty

    candidates = get_free_candidates(duty.exam, exclude_faculty_ids={absent_person.id})
    if not candidates:
        return False, (
            f"No replacement faculty available for {absent_person.name} "
            f"on {duty.exam.date} (Room {duty.room.room_number}). Original allocation unchanged."
        )

    replacement = candidates[0]
    duty.replaced_faculty = absent_person
    duty.faculty = replacement
    duty.status = "ALLOCATED"
    duty.save()

    new_total = Duty.objects.filter(faculty=replacement, status="ALLOCATED").count()
    return True, (
        f"Reallocated Room {duty.room.room_number} ({duty.exam.date}) from "
        f"{absent_person.name} to {replacement.name} (Total duties: {new_total})."
    )
