from django.urls import path
from . import views

urlpatterns = [
    path('excel/', views.ExcelAPI.as_view()),
]
