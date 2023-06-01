from django.urls import path
from . import views

urlpatterns = [
	path('', views.index),
    path('check', views.check, name='check'),
    path('index', views.index, name='index'),
    path('excel', views.excel, name='excel'),
    path('excel2', views.excel2, name='excel2'),
    path('aggiorna', views.aggiorna, name='aggiorna'),
    path('insert', views.insert, name='insert'),
]