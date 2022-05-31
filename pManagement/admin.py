from django.contrib import admin
from import_export.admin import ImportExportModelAdmin

from pManagement.models import *


@admin.register(PartNumbersExpendable)
@admin.register(PartNumbersReturnable)
@admin.register(extraProductTable)
@admin.register(extraProductHistoric)
@admin.register(buttonAuthorHistory)
@admin.register(StockElements)

class ViewAdmin(ImportExportModelAdmin):
    pass