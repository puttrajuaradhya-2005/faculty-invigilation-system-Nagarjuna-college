from pathlib import Path
import csv
from django.conf import settings
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.db.models import Count, Q
from django.http import HttpResponse
from django.shortcuts import render, redirect, get_object_or_404

from .models import Faculty, FacultySlot, Exam, Room, Duty
from .parser import preview, parse_faculty_file, parse_exam_file, parse_room_file
from .allocator import allocate_duties, reallocate_absent_duty, get_free_candidates


def login_view(request):
    if request.user.is_authenticated:
        return redirect("dashboard")
    if request.method == "POST":
        user = authenticate(
            request,
            username=request.POST.get("username", "").strip(),
            password=request.POST.get("password", "")
        )
        if user is not None:
            login(request, user)
            return redirect("dashboard")
        messages.error(request, "Invalid username or password.")
    return render(request, "login.html")


def logout_view(request):
    logout(request)
    return redirect("login")


@login_required
def dashboard(request):
    duties = Duty.objects.select_related("faculty", "exam", "room", "replaced_faculty").order_by(
        "exam__date", "exam__start_time", "room__room_number"
    )
    exams = Exam.objects.all().order_by("date", "start_time")
    faculty_summary = Faculty.objects.annotate(
        duty_total=Count("duty", filter=Q(duty__status="ALLOCATED"))
    ).order_by("-duty_total", "name")
    all_faculty = Faculty.objects.all().order_by("name")

    raw_warnings = request.session.get("warnings", [])
    clean_warnings = [w for w in raw_warnings if not str(w).startswith("NO FREE FACULTY")]

    return render(request, "dashboard.html", {
        "faculty_count": Faculty.objects.count(),
        "slot_count": FacultySlot.objects.count(),
        "exam_count": Exam.objects.count(),
        "room_count": Room.objects.count(),
        "duty_count": Duty.objects.filter(status="ALLOCATED").count(),
        "duties": duties,
        "exams": exams,
        "faculty_summary": faculty_summary,
        "all_faculty": all_faculty,
        "previews": request.session.get("previews", []),
        "warnings": clean_warnings,
        "process_log": request.session.get("process_log", []),
    })


@login_required
@transaction.atomic
def add_faculty(request):
    if request.method == "POST":
        name = request.POST.get("faculty_name", "").strip()
        department = request.POST.get("department", "").strip()
        employee_id = request.POST.get("employee_id", "").strip()
        if not name:
            messages.error(request, "Please enter a faculty name.")
            return redirect("dashboard")

        fac, created = Faculty.objects.get_or_create(
            name__iexact=name,
            defaults={"name": name, "department": department, "employee_id": employee_id}
        )
        if created:
            messages.success(request, f"Added extra faculty '{fac.name}' (free across all slots). Click 'GENERATE DUTIES' or use Edit/Reallocate.")
        else:
            messages.info(request, f"Faculty '{fac.name}' already exists in the system.")
    return redirect("dashboard")


@login_required
@transaction.atomic
def upload_files(request):
    if request.method != "POST":
        return redirect("dashboard")

    faculty_files = request.FILES.getlist("faculty_file")
    exam_files = request.FILES.getlist("exam_file")
    room_files = request.FILES.getlist("room_file")

    if not faculty_files or not exam_files or not room_files:
        messages.error(request, "Please upload files for all 3 categories: Faculty/Class Timetable, Exam Timetable and Room Details.")
        return redirect("dashboard")

    upload_dir = Path(settings.MEDIA_ROOT) / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    previews = []
    extraction_warnings = []

    def _save_list(flist):
        paths = []
        for f in flist:
            name = Path(f.name).name
            path = upload_dir / name
            with open(path, "wb+") as out:
                for chunk in f.chunks():
                    out.write(chunk)
            paths.append(path)
            previews.append(preview(path))
        return paths

    try:
        saved_fac = _save_list(faculty_files)
        saved_ex = _save_list(exam_files)
        saved_rm = _save_list(room_files)

        Duty.objects.all().delete()
        Room.objects.all().delete()
        Exam.objects.all().delete()
        FacultySlot.objects.all().delete()
        Faculty.objects.all().delete()

        busy_records = []
        faculty_names_set = set()
        for p in saved_fac:
            b_recs, f_names = parse_faculty_file(p, warnings=extraction_warnings)
            busy_records.extend(b_recs)
            faculty_names_set.update(f_names)
        faculty_names = sorted(faculty_names_set)

        exam_records = []
        seen_exams = set()
        for p in saved_ex:
            for e in parse_exam_file(p, warnings=extraction_warnings):
                ek = (e["date"], e["start"], e["end"], e["slot"])
                if ek not in seen_exams:
                    seen_exams.add(ek)
                    exam_records.append(e)

        room_numbers_list = []
        for p in saved_rm:
            room_numbers_list.extend(parse_room_file(p, warnings=extraction_warnings))
        room_numbers = list(dict.fromkeys(room_numbers_list))

        faculty_by_name = {}
        for name in faculty_names:
            faculty_by_name[name.lower()] = Faculty.objects.create(name=name)
        for rec in busy_records:
            person = faculty_by_name.get(rec["faculty"].lower())
            if not person:
                continue
            FacultySlot.objects.create(
                faculty=person, weekday=rec["weekday"], start_time=rec["start"], end_time=rec["end"],
                status="BUSY", source_slot=rec["source_slot"], course_code=rec["course_code"], course_title=rec["course_title"]
            )

        exams = []
        for rec in exam_records:
            e = Exam.objects.create(
                date=rec["date"], start_time=rec["start"], end_time=rec["end"],
                slot=rec.get("slot", ""), course=rec.get("course", ""),
                course_code=rec.get("course_code", ""), class_name=rec.get("class_name", "")
            )
            exams.append(e)

        for e in exams:
            for room_no in room_numbers:
                Room.objects.create(exam=e, room_number=room_no)

        process_log = [
            f"Faculty records imported: {Faculty.objects.count()}",
            f"Recurring occupied timetable slots imported: {FacultySlot.objects.count()}",
            f"Exam slots imported: {Exam.objects.count()}",
            f"Rooms imported per exam: {len(room_numbers)}",
            "Faculty availability will be calculated from weekday + time overlap during allocation.",
        ]
        request.session["previews"] = previews
        request.session["process_log"] = process_log
        request.session["warnings"] = extraction_warnings

        if extraction_warnings:
            messages.warning(
                request,
                f"Files imported with {len(extraction_warnings)} extraction warning(s). Review the alerts below before generating duties."
            )
        else:
            messages.success(
                request,
                "All 3 files were parsed and converted into exam slots, faculty timetable records and rooms. You can now generate duties."
            )
    except Exception as exc:
        transaction.set_rollback(True)
        messages.error(request, f"Automatic processing failed: {exc}")
    return redirect("dashboard")


@login_required
def import_preview(request):
    return redirect("dashboard")


@login_required
@transaction.atomic
def allocate(request):
    if request.method != "POST":
        return redirect("dashboard")
    exams = list(Exam.objects.all())
    faculty = list(Faculty.objects.all())
    slots = list(FacultySlot.objects.all())
    rooms = {}
    for room in Room.objects.select_related("exam"):
        rooms.setdefault(room.exam_id, []).append(room)

    Duty.objects.all().delete()
    assigned, warnings = allocate_duties(exams, rooms, faculty, slots)
    for exam, room, person in assigned:
        Duty.objects.create(exam=exam, room=room, faculty=person, status="ALLOCATED")

    request.session["warnings"] = []
    messages.success(request, f"Allocation complete: {len(assigned)} duties generated equally across faculty.")
    return redirect("dashboard")


@login_required
@transaction.atomic
def edit_duty(request, duty_id):
    duty = get_object_or_404(Duty, id=duty_id)
    if request.method == "POST":
        new_fac_id = request.POST.get("faculty_id")
        new_fac = get_object_or_404(Faculty, id=new_fac_id)

        if new_fac.id == duty.faculty.id:
            messages.info(request, "No changes made; same faculty selected.")
            return redirect("dashboard")

        already_in_exam = Duty.objects.filter(
            exam=duty.exam, faculty=new_fac, status="ALLOCATED"
        ).exclude(id=duty.id).exists()
        if already_in_exam:
            messages.error(
                request,
                f"Conflict prevented: {new_fac.name} is already allocated to another room during this exam session."
            )
            return redirect("dashboard")

        old_fac = duty.faculty
        duty.replaced_faculty = old_fac
        duty.faculty = new_fac
        duty.status = "ALLOCATED"
        duty.save()
        messages.success(request, f"Manually updated Room {duty.room.room_number} from {old_fac.name} to {new_fac.name}.")
    return redirect("dashboard")


@login_required
def absent_reallocate(request, duty_id):
    if request.method == "POST":
        ok, msg = reallocate_absent_duty(duty_id)
        if ok:
            messages.success(request, msg)
        else:
            messages.error(request, msg)
    return redirect("dashboard")


@login_required
def export_csv(request):
    response = HttpResponse(content_type="text/csv")
    response["Content-Disposition"] = 'attachment; filename="invigilation_duties.csv"'
    writer = csv.writer(response)
    writer.writerow(["Date", "Day", "Slot", "Start", "End", "Course", "Course Code", "Class", "Room", "Faculty ID", "Faculty Name", "Replaced Faculty"])
    for d in Duty.objects.select_related("faculty", "exam", "room", "replaced_faculty").order_by("exam__date", "exam__start_time", "room__room_number"):
        writer.writerow([
            d.exam.date, d.exam.date.strftime("%A"), d.exam.slot, d.exam.start_time, d.exam.end_time,
            d.exam.course, d.exam.course_code, d.exam.class_name, d.room.room_number,
            d.faculty.employee_id, d.faculty.name, d.replaced_faculty.name if d.replaced_faculty else ""
        ])
    return response


@login_required
def clear_data(request):
    if request.method == "POST":
        Duty.objects.all().delete()
        Room.objects.all().delete()
        Exam.objects.all().delete()
        FacultySlot.objects.all().delete()
        Faculty.objects.all().delete()
        request.session["warnings"] = []
        request.session["process_log"] = []
        request.session["previews"] = []
        messages.success(request, "Imported data and generated duties cleared.")
    return redirect("dashboard")
