from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):
    initial = True
    dependencies = []
    operations = [
        migrations.CreateModel(name="Faculty", fields=[
            ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
            ("name", models.CharField(max_length=200)), ("employee_id", models.CharField(blank=True, max_length=100)), ("department", models.CharField(blank=True, max_length=150)),
        ]),
        migrations.CreateModel(name="Exam", fields=[
            ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")), ("date", models.DateField()), ("start_time", models.TimeField()), ("end_time", models.TimeField()), ("slot", models.CharField(blank=True, max_length=20)), ("course", models.CharField(blank=True, max_length=250)), ("course_code", models.CharField(blank=True, max_length=100)), ("class_name", models.CharField(blank=True, max_length=150)),
        ]),
        migrations.CreateModel(name="FacultySlot", fields=[
            ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")), ("weekday", models.CharField(choices=[("MON","MON"),("TUE","TUE"),("WED","WED"),("THU","THU"),("FRI","FRI"),("SAT","SAT"),("SUN","SUN")], max_length=3)), ("start_time", models.TimeField()), ("end_time", models.TimeField()), ("status", models.CharField(default="BUSY", max_length=10)), ("source_slot", models.CharField(blank=True, max_length=80)), ("course_code", models.CharField(blank=True, max_length=80)), ("course_title", models.CharField(blank=True, max_length=250)), ("faculty", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="slots", to="allocator.faculty")),
        ], options={"ordering":["weekday","start_time","faculty__name"]}),
        migrations.CreateModel(name="Room", fields=[
            ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")), ("room_number", models.CharField(max_length=100)), ("capacity", models.PositiveIntegerField(blank=True, null=True)), ("exam", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name="rooms", to="allocator.exam")),
        ]),
        migrations.CreateModel(name="Duty", fields=[
            ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")), ("created_at", models.DateTimeField(auto_now_add=True)), ("exam", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to="allocator.exam")), ("faculty", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to="allocator.faculty")), ("room", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to="allocator.room")),
        ], options={"unique_together":{("exam","room")},"ordering":["exam__date","exam__start_time","room__room_number"]}),
    ]
