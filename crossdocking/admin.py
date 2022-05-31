from django.contrib import admin
from import_export.admin import ImportExportModelAdmin
from .models import *


@admin.register(Prodlines)
@admin.register(ProdlineTable)
@admin.register(ProdlineTableHistory)
@admin.register(ProdlineHistoryAllTime)

class ViewAdmin(ImportExportModelAdmin):
    pass