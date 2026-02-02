from django.db import models
from django.contrib.auth.models import User
from django.utils.translation import gettext_lazy as _
from decimal import Decimal

# Create your models here.
class HeadParty(models.Model):
    partyname = models.CharField(max_length=100, primary_key=True)
    org = models.ForeignKey('Organization', on_delete=models.CASCADE,null=True, blank=True)
    add1 = models.CharField(max_length=200, blank=True)
    add2 = models.CharField(max_length=200, blank=True)
    city = models.CharField(max_length=100, blank=True)
    state = models.CharField(max_length=100, blank=True)
    mobile = models.CharField(max_length=15, blank=True)
    otherno = models.CharField(max_length=15, blank=True)
    email = models.EmailField(blank=True)
    remark = models.TextField(blank=True)
    openingdebit = models.DecimalField(max_digits=10, decimal_places=2, default=0, blank=True, null=True)
    openingcredit = models.DecimalField(max_digits=10, decimal_places=2, default=0, blank=True, null=True)
    def __str__(self):
        return self.partyname

class Broker(models.Model):
    brokername = models.CharField(max_length=100, primary_key=True)
    org = models.ForeignKey('Organization', on_delete=models.CASCADE,null=True, blank=True)
    mobileno = models.CharField(max_length=15, blank=True)
    email = models.EmailField(blank=True)
    openingdebit = models.DecimalField(max_digits=10, decimal_places=2, default=0, blank=True, null=True)
    openingcredit = models.DecimalField(max_digits=10, decimal_places=2, default=0, blank=True, null=True)
    remark = models.TextField(blank=True)

    def __str__(self):
        return self.brokername

class Firm(models.Model):
    # single field, primary key as requested
    firmname = models.CharField(max_length=100, primary_key=True)

    def __str__(self):
        return self.firmname
    
class HeadItem(models.Model):
    item_name = models.CharField(max_length=100, primary_key=True)
    org = models.ForeignKey('Organization', on_delete=models.CASCADE,null=True, blank=True)
    def __str__(self):
        return self.item_name

SIGN_CHOICES = [
    ('PLUS', '+'),
    ('MINUS', '-'),
]

class SaleMaster(models.Model):
    invno = models.AutoField(primary_key=True)
    org = models.ForeignKey('Organization', on_delete=models.CASCADE,null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    invdate = models.DateField()
    awakno = models.CharField(max_length=50, blank=True, null=True)

    party = models.ForeignKey("HeadParty", on_delete=models.CASCADE)
    broker = models.ForeignKey("Broker", on_delete=models.CASCADE)
    firm = models.ForeignKey("Firm", on_delete=models.CASCADE, null=True, blank=True)
    # NEW field after party, broker
    extra = models.CharField(max_length=255, blank=True, null=True)

    vehicleno = models.CharField(max_length=50, blank=True, null=True)

    totalamt = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    # RENAMED fields: borapercent -> batavpercent, boraamt -> batavamt
    # (since DB is empty it's safe to replace names)
    batavpercent = models.DecimalField(max_digits=5, decimal_places=2, default=0)   # e.g. %
    batavamt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    batav_sign = models.CharField(
       max_length=5,
       choices=SIGN_CHOICES,
       default='PLUS'
   )
    # Dalali (dr/dramt) kept as-is; after Dalali add QI
    dr = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    dramt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    dr_sign = models.CharField(
       max_length=5,
       choices=SIGN_CHOICES,
       default='PLUS'
   )
    # NEW QI field after Dalali
    qi = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    other = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    advance = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    netamt = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    # remark stays CharField (single-line input in template)
    remark = models.CharField(max_length=255, blank=True, null=True)

    # amt_in_words REMOVED (deleted from model)

    def __str__(self):
        return f"Invoice {self.invno} - {self.party}"


class SaleDetails(models.Model):
    salemaster = models.ForeignKey("SaleMaster", on_delete=models.CASCADE, related_name='details')
    item = models.ForeignKey("HeadItem", on_delete=models.CASCADE)

    bora = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    # NEW 4 fields after bora
    bn = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    bnwt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    bo = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    bowt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    # NEW: TBWt (shown before PWt in UI)
    tbwt = models.DecimalField("TBWt", max_digits=12, decimal_places=2, default=0)
    qty = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    rate = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    amount = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    partywt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    millwt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    diffwt = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    # NEW: FRK weight (shown in UI as FrkWt) — store it explicitly
    frkwt = models.DecimalField("FrkWt", max_digits=12, decimal_places=2, default=0)
    
    # NEW Lot number after diffwt
    lotno = models.CharField(max_length=50, blank=True, null=True)

    # 👇 YAHI paste karo
    def save(self, *args, **kwargs):
        if self.rate and self.qty:
            self.amount = (self.rate * self.qty) / Decimal('100')
        else:
            self.amount = Decimal('0.00')
        super().save(*args, **kwargs)


    def __str__(self):
        return f"{self.item} - {self.qty}"




class PurchaseMaster(models.Model):
    invno = models.AutoField(primary_key=True)
    org = models.ForeignKey('Organization', on_delete=models.CASCADE,null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    invdate = models.DateField()
    awakno = models.CharField(max_length=50, blank=True, null=True)

    party = models.ForeignKey("HeadParty", on_delete=models.CASCADE)
    broker = models.ForeignKey("Broker", on_delete=models.CASCADE)
    firm = models.ForeignKey("Firm", on_delete=models.CASCADE, null=True, blank=True)
    # NEW field after party, broker
    extra = models.CharField(max_length=255, blank=True, null=True)

    vehicleno = models.CharField(max_length=50, blank=True, null=True)
    totalamt = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    # RENAMED fields
    batavpercent = models.DecimalField(max_digits=5, decimal_places=2, default=0)
    batavamt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    batav_sign = models.CharField(
       max_length=5,
       choices=SIGN_CHOICES,
       default='PLUS'
   )

    dr = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    dramt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    dr_sign = models.CharField(
       max_length=5,
       choices=SIGN_CHOICES,
       default='PLUS'
   )
    # NEW QI field
    qi = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    other = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    advance = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    netamt = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    remark = models.CharField(max_length=255, blank=True, null=True)

    def __str__(self):
        return f"Purchase Invoice {self.invno} - {self.party}"

class PurchaseDetails(models.Model):
    purchasemaster = models.ForeignKey("PurchaseMaster", on_delete=models.CASCADE, related_name='details')
    item = models.ForeignKey("HeadItem", on_delete=models.CASCADE)

    bora = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    # NEW 4 fields after bora
    bn = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    bnwt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    bo = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    bowt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    # ✅ NEW: TBWt (same as SaleDetails)
    tbwt = models.DecimalField("TBWt", max_digits=12, decimal_places=2, default=0)
    qty = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    rate = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    amount = models.DecimalField(max_digits=12, decimal_places=2, default=0)

    partywt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    millwt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    diffwt = models.DecimalField(max_digits=12, decimal_places=2, default=0)
    # ✅ NEW: FrkWt (same as SaleDetails)
    frkwt = models.DecimalField("FrkWt", max_digits=12, decimal_places=2, default=0)
    # NEW Lot number after diffwt
    lotno = models.CharField(max_length=50, blank=True, null=True)

    def save(self, *args, **kwargs):
        if self.rate and self.qty:
            self.amount = (self.rate * self.qty) / Decimal('100')
        else:
            self.amount = Decimal('0.00')
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.item} - {self.qty}"


class DailyPage(models.Model):
    org = models.ForeignKey('Organization', on_delete=models.CASCADE,null=True, blank=True)
    date = models.DateField()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(fields=['org', 'date'], name='uniq_dailypage_per_org_date')
        ]

    def __str__(self):
        return f"DailyPage {self.date}"

class JamaEntry(models.Model):
    entry_no = models.AutoField(primary_key=True)
    daily_page = models.ForeignKey(DailyPage, on_delete=models.CASCADE, related_name='jama_entries')

    # make FK nullable, keep column name partyname, add textual fallback
    party = models.ForeignKey('HeadParty', on_delete=models.PROTECT, db_column='partyname',
                              blank=True, null=True)
    party_name = models.CharField(max_length=200, blank=True, default='')

    broker = models.ForeignKey(Broker, on_delete=models.PROTECT, blank=True, null=True)

    # firm: store FK + textual fallback
    firm = models.ForeignKey('Firm', on_delete=models.PROTECT, blank=True, null=True)
    firm_name = models.CharField(max_length=100, blank=True, default='')

    amount = models.DecimalField(max_digits=14, decimal_places=2)
    remark = models.CharField(max_length=255, blank=True, default='')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['entry_no']

    def __str__(self):
        party_display = self.party_name or (getattr(self.party, 'partyname', '') if self.party else '')
        return f"Jama #{self.entry_no} - {party_display} - {self.broker} - {self.amount}"
 

class NaameEntry(models.Model):
    entry_no = models.AutoField(primary_key=True)
    daily_page = models.ForeignKey(DailyPage, on_delete=models.CASCADE, related_name='naame_entries')

    party = models.ForeignKey('HeadParty', on_delete=models.PROTECT, db_column='partyname',
                              blank=True, null=True)
    party_name = models.CharField(max_length=200, blank=True, default='')

    broker = models.ForeignKey(Broker, on_delete=models.PROTECT, blank=True, null=True)

    firm = models.ForeignKey('Firm', on_delete=models.PROTECT, blank=True, null=True)
    firm_name = models.CharField(max_length=100, blank=True, default='')

    amount = models.DecimalField(max_digits=14, decimal_places=2)
    remark = models.CharField(max_length=255, blank=True, default='')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['entry_no']

    def __str__(self):
        party_display = self.party_name or (getattr(self.party, 'partyname', '') if self.party else '')
        return f"Naame #{self.entry_no} - {party_display} - {self.broker} - {self.amount}"

class Organization(models.Model):

    name = models.CharField(max_length=150, unique=True)
    # Owner is optional in single-company mode (handy for seeding via command)
    owner = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="owned_orgs")
    created_at = models.DateTimeField(auto_now_add=True)
    def __str__(self): return self.name
    # Alias for older templates/code that expect .orgname
    @property
    def orgname(self):
        return self.name


# Inherit this in your business models to auto-get org + created_by
class OrgScopedModel(models.Model):
    org = models.ForeignKey(Organization, on_delete=models.CASCADE)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    class Meta:
        abstract = True 