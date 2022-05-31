from django.contrib import admin
from import_export.admin import ImportExportModelAdmin
from .models import *

@admin.register(LinguagemPortaria)
@admin.register(LinguagemCafetaria)
@admin.register(HistoricoPortaria2)
@admin.register(ChaveSecretaRH)
@admin.register(HistoricoCafetaria2)

class ViewAdmin(ImportExportModelAdmin):
    pass