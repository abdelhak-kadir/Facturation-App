from django.db import models

class Livraison(models.Model):
    timestamp = models.DateTimeField(unique=True)
    nom_chauffeur = models.CharField(max_length=255,null=True, blank=True)
    client = models.CharField(max_length=255,null=True, blank=True)
    chargements = models.CharField(max_length=255,null=True, blank=True)
    date_chargement = models.DateField(null=True, blank=True)
    dechargement = models.CharField(max_length=255,null=True, blank=True)
    bon_livraison = models.CharField(max_length=255,null=True, blank=True)
    tarif = models.CharField(max_length=255,null=True, blank=True)
    deplacement = models.DecimalField(max_digits=10, decimal_places=2,null=True, blank=True)
    avance = models.DecimalField(max_digits=10, decimal_places=2,null=True, blank=True)
    charge_variable = models.DecimalField(max_digits=10, decimal_places=2,null=True, blank=True)
    prix_cv = models.DecimalField(max_digits=10, decimal_places=2,null=True, blank=True)
    operateur = models.CharField(max_length=255,null=True, blank=True)
    commercial = models.CharField(max_length=255,null=True, blank=True)
    ice = models.CharField(max_length=255,null=True, blank=True)
    qte = models.PositiveIntegerField(null=True, blank=True)
    nom_destinataire = models.CharField(max_length=255,null=True, blank=True)
    facturation = models.CharField(max_length=100,null=True, blank=True)

    def __str__(self):
        return f"{self.client} - {self.nom_chauffeur} ({self.date_chargement})"
