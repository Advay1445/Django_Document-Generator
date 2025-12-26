# doc_gen/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),  # The homepage with the form
    path('generate/', views.generate_docx, name='generate_docx'), # The action URL
]