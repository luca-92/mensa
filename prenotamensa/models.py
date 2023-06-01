from django.db import models

class Personale(models.Model):
    nominativo = models.CharField(max_length=50)
    numero_CMD = models.CharField(max_length=10)

    def __str__(self):
        return self.nominativo + " " + self.numero_CMD

class AssociazioneAtt(models.Model):
    datapasto = models.DateField()
    numero_CMD = models.CharField(max_length=10)
    Colazione = models.CharField(max_length=1, null=True)
    Pranzo = models.CharField(max_length=1, null=True)
    Cena = models.CharField(max_length=1, null=True)
    Sandwich = models.CharField(max_length=1, null=True)
    Lavanderia = models.CharField(max_length=1, null=True)

    def __str__(self):
        return self.numero_CMD
    