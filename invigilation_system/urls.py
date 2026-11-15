from django.urls import path
from allocator import views

urlpatterns = [
    path("", views.login_view, name="login"),
    path("dashboard/", views.dashboard, name="dashboard"),
    path("upload/", views.upload_files, name="upload"),
    path("faculty/add/", views.add_faculty, name="add_faculty"),
    path("allocate/", views.allocate, name="allocate"),
    path("duty/<int:duty_id>/edit/", views.edit_duty, name="edit_duty"),
    path("duty/<int:duty_id>/absent/", views.absent_reallocate, name="absent_reallocate"),
    path("export/", views.export_csv, name="export"),
    path("clear/", views.clear_data, name="clear"),
    path("logout/", views.logout_view, name="logout"),
]
