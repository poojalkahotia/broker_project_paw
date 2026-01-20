from django.contrib import admin
from .models import SaleMaster, SaleDetails

@admin.register(SaleMaster)
class SaleMasterAdmin(admin.ModelAdmin):
    list_display = ('invno', 'invdate', 'party', 'broker', 'firm', 'netamt')
    list_filter = ('invdate', 'party', 'broker', 'firm')
    search_fields = ('invno', 'party__partyname', 'broker__brokername')

@admin.register(SaleDetails)
class SaleDetailsAdmin(admin.ModelAdmin):
    list_display = ('salemaster', 'item', 'qty', 'rate', 'amount')
    list_filter = ('item',)
