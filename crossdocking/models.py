from django.db import models

class ProdlineTable(models.Model):
    day = models.CharField(max_length=50, null=True)
    line = models.CharField(max_length=50, null=True)
    site = models.CharField(max_length=50, null=True)
    due_date = models.CharField(max_length=50, null=True)
    item_number = models.CharField(max_length=50, null=True)
    description = models.CharField(max_length=50, null=True)
    to_complete = models.CharField(max_length=50, null=True)
    receiving = models.BooleanField(default=None, blank='true', null=True)
    comentarioReceiving = models.CharField(max_length=350, null=True)
    shipping = models.BooleanField(default=None, blank='true', null=True)
    comentarioShipping = models.CharField(max_length=350, null=True)
    allOkReceiving = models.BooleanField(default=None, blank='true', null=True)
    allOkShipping = models.BooleanField(default=None, blank='true', null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Warehouse - Production'
        verbose_name_plural = 'Warehouse - Production'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.day + ' ' + self.line}"


class Prodlines(models.Model):
    nome = models.CharField(max_length=100, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Prodlines'
        verbose_name_plural = 'Prodlines'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.nome}"


class ProdlineTableHistory(models.Model):
    day = models.CharField(max_length=50, null=True)
    line = models.CharField(max_length=50, null=True)
    site = models.CharField(max_length=50, null=True)
    due_date = models.CharField(max_length=50, null=True)
    item_number = models.CharField(max_length=50, null=True)
    description = models.CharField(max_length=50, null=True)
    to_complete = models.CharField(max_length=50, null=True)
    receiving = models.BooleanField(default=None, blank='true', null=True)
    allOkReceiving = models.BooleanField(default=None, blank='true', null=True)
    comentarioReceiving = models.CharField(max_length=350, null=True)
    shipping = models.BooleanField(default=None, blank='true', null=True)
    allOkShipping = models.BooleanField(default=None, blank='true', null=True)
    comentarioShipping = models.CharField(max_length=350, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Warehouse - Production - History'
        verbose_name_plural = 'Warehouse - Production - History'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.day + ' ' + self.line}"


class ProdlineHistoryAllTime(models.Model):
    day = models.CharField(max_length=50, null=True)
    line = models.CharField(max_length=50, null=True)
    site = models.CharField(max_length=50, null=True)
    due_date = models.CharField(max_length=50, null=True)
    item_number = models.CharField(max_length=50, null=True)
    description = models.CharField(max_length=50, null=True)
    to_complete = models.CharField(max_length=50, null=True)
    receiving = models.BooleanField(default=None, blank='true', null=True)
    allOkReceiving = models.BooleanField(default=None, blank='true', null=True)
    comentarioReceiving = models.CharField(max_length=350, null=True)
    shipping = models.BooleanField(default=None, blank='true', null=True)
    allOkShipping = models.BooleanField(default=None, blank='true', null=True)
    comentarioShipping = models.CharField(max_length=350, null=True)

    # META CLASS
    class Meta:
        verbose_name = 'Warehouse - Production - History All Time'
        verbose_name_plural = 'Warehouse - Production - History All Time'

    # TO STRING METHOD
    def __str__(self):
        return f"{self.day + ' ' + self.line}"