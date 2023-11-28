from django.urls import path, include
from excel_app import views

urlpatterns = [
    path('', views.index)
]