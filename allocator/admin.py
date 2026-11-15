from django.contrib import admin
from .models import Faculty, FacultySlot, Exam, Room, Duty
admin.site.register([Faculty, FacultySlot, Exam, Room, Duty])
