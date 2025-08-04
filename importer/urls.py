from django.urls import path
from .views import *

urlpatterns = [
    path('import-livraisons/', trigger_import_ajax, name='import_livraisons'),
    path("", livraison_list, name="livraison_list"),
    path('edit/<int:pk>/', edit_livraison, name='edit_livraison'),
    path('generate-invoice/', generate_invoice, name='generate_invoice'),
    path('generate-word-invoice/', generate_word_invoice, name='generate_word_invoice'),
]
