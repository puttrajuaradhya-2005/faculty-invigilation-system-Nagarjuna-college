from django.db import models


class Faculty(models.Model):
    name = models.CharField(max_length=200)
    employee_id = models.CharField(max_length=100, blank=True)
    department = models.CharField(max_length=150, blank=True)

    class Meta:
        ordering = ["name"]

    @property
    def active_duty_count(self):
        return self.duty_set.filter(status="ALLOCATED").count()

    def __str__(self):
        return self.name


class FacultySlot(models.Model):
    """A recurring weekly occupied/free slot derived from the uploaded faculty timetable."""
    WEEKDAYS = [(x, x) for x in ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]]
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name="slots")
    weekday = models.CharField(max_length=3, choices=WEEKDAYS)
    start_time = models.TimeField()
    end_time = models.TimeField()
    status = models.CharField(max_length=10, default="BUSY")
    source_slot = models.CharField(max_length=80, blank=True)
    course_code = models.CharField(max_length=80, blank=True)
    course_title = models.CharField(max_length=250, blank=True)

    class Meta:
        ordering = ["weekday", "start_time", "faculty__name"]


class Exam(models.Model):
    date = models.DateField()
    start_time = models.TimeField()
    end_time = models.TimeField()
    slot = models.CharField(max_length=20, blank=True)
    course = models.CharField(max_length=250, blank=True)
    course_code = models.CharField(max_length=100, blank=True)
    class_name = models.CharField(max_length=150, blank=True)

    class Meta:
        ordering = ["date", "start_time"]


class Room(models.Model):
    exam = models.ForeignKey(Exam, on_delete=models.CASCADE, related_name="rooms", null=True, blank=True)
    room_number = models.CharField(max_length=100)
    capacity = models.PositiveIntegerField(null=True, blank=True)

    class Meta:
        ordering = ["room_number"]


class Duty(models.Model):
    exam = models.ForeignKey(Exam, on_delete=models.CASCADE)
    room = models.ForeignKey(Room, on_delete=models.CASCADE)
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE)
    status = models.CharField(max_length=20, default="ALLOCATED")
    replaced_faculty = models.ForeignKey(
        Faculty, null=True, blank=True, on_delete=models.SET_NULL, related_name="replaced_duties"
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ("exam", "room")
        ordering = ["exam__date", "exam__start_time", "room__room_number"]
