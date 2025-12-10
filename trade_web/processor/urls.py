from django.urls import path
from . import views

urlpatterns = [
    path('', views.process_document, name='process_document'),
]
