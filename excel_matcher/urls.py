from django.urls import path
from . import views

app_name = "excel_matcher"

urlpatterns = [
    path("", views.index, name="index"),
    path("upload/", views.upload_file, name="upload_file"),
    path("process/", views.process_file, name="process_file"),
    path("preview/", views.preview_matching, name="preview_matching"),
    path("download/", views.download_file, name="download_file"),
]
