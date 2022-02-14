from django.urls import path

from . import views

app_name = "excel_table"

urlpatterns = [
    path('', views.index, name='index'),
]