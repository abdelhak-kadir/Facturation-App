from django import forms
from .models import Livraison

class LivraisonForm(forms.ModelForm):
    class Meta:
        model = Livraison
        fields = '__all__'
        widgets = {
            'timestamp': forms.DateTimeInput(attrs={'type': 'datetime-local', 'class': 'form-control'}),
            'nom_chauffeur': forms.TextInput(attrs={'class': 'form-control'}),
            'client': forms.TextInput(attrs={'class': 'form-control'}),
            'chargements': forms.TextInput(attrs={'class': 'form-control'}),
            'date_chargement': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'dechargement': forms.TextInput(attrs={'class': 'form-control'}),
            'bon_livraison': forms.TextInput(attrs={'class': 'form-control'}),
            'tarif': forms.TextInput(attrs={'class': 'form-control'}),
            'deplacement': forms.NumberInput(attrs={'class': 'form-control'}),
            'avance': forms.NumberInput(attrs={'class': 'form-control'}),
            'charge_variable': forms.NumberInput(attrs={'class': 'form-control'}),
            'prix_cv': forms.NumberInput(attrs={'class': 'form-control'}),
            'operateur': forms.TextInput(attrs={'class': 'form-control'}),
            'commercial': forms.TextInput(attrs={'class': 'form-control'}),
            'ice': forms.TextInput(attrs={'class': 'form-control'}),
            'qte': forms.NumberInput(attrs={'class': 'form-control'}),
            'nom_destinataire': forms.TextInput(attrs={'class': 'form-control'}),
            'facturation': forms.TextInput(attrs={'class': 'form-control'}),
        }
