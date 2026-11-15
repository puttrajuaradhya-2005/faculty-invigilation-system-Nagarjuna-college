from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("allocator", "0002_alter_exam_options_alter_faculty_options_and_more"),
    ]

    operations = [
        migrations.AddField(
            model_name="duty",
            name="status",
            field=models.CharField(default="ALLOCATED", max_length=20),
        ),
        migrations.AddField(
            model_name="duty",
            name="replaced_faculty",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="replaced_duties",
                to="allocator.faculty",
            ),
        ),
    ]
