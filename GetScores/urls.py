from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="upload"),
    path("success", views.upload, name="success")
]