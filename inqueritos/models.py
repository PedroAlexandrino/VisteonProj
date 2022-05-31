from django.db import models

class LinguagemCafetaria(models.Model):
    nome = models.CharField(max_length=50, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Linguagem Cafetaria'
        verbose_name_plural = 'Linguagem Cafetaria'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.nome}"


class LinguagemPortaria(models.Model):
    nome = models.CharField(max_length=50, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Linguagem Portaria'
        verbose_name_plural = 'Linguagem Portaria'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.nome}"


class HistoricoPortaria2(models.Model):
    data = models.CharField(max_length=50, null=True)
    seguro = models.CharField(max_length=50, null=True)
    comentarios = models.CharField(max_length=3500, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Historico Portaria'
        verbose_name_plural = 'Historico Portaria'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.data}"


class ChaveSecretaRH(models.Model):
    nome = models.CharField(max_length=50, null=True)


class HistoricoCafetaria2(models.Model):
    data = models.CharField(max_length=50, null=True)
    comentarios = models.CharField(max_length=3500, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Historico Cafetaria'
        verbose_name_plural = 'Historico Cafetaria'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.data}"
