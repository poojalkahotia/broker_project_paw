from django import forms
from .models import HeadParty, Broker, HeadItem, SaleMaster, SaleDetails, PurchaseMaster, PurchaseDetails,Firm
from django.core.exceptions import ValidationError
from django.contrib.auth.models import User
from .models import Organization

# ----------------- Base: BlankZeroModelForm -----------------
from django.forms import ModelForm
from django.forms.fields import DecimalField, IntegerField, FloatField

class BlankZeroModelForm(ModelForm):
    """
    Base ModelForm: when rendering a new (unbound) form (no instance),
    remove initial/default zeros from numeric fields so inputs render blank.
    This is minimal and safe: it only affects display for new forms.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        is_new_display = (not self.is_bound) and (kwargs.get('instance') is None)
        if is_new_display:
            for name, field in self.fields.items():
                if isinstance(field, (DecimalField, IntegerField, FloatField)):
                    # remove any initial so widget won't render a value attribute
                    field.initial = None
                    # remove any hard-coded 'value' in widget attrs (defensive)
                    try:
                        if hasattr(field, 'widget') and getattr(field.widget, 'attrs', None):
                            field.widget.attrs.pop('value', None)
                    except Exception:
                        pass

# ----------------- Party / Broker / Item Forms -----------------
class PartyForm(BlankZeroModelForm):
    class Meta:
        model = HeadParty
        exclude = ['org']
        widgets = {
            'partyname': forms.TextInput(attrs={'class': 'form-control'}),
            'add1': forms.TextInput(attrs={'class': 'form-control'}),
            'add2': forms.TextInput(attrs={'class': 'form-control'}),
            'city': forms.TextInput(attrs={'class': 'form-control'}),
            'state': forms.TextInput(attrs={'class': 'form-control'}),
            'mobile': forms.TextInput(attrs={'class': 'form-control'}),
            'otherno': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 2, 'data-next': '#saveBtn'}),
            'openingdebit': forms.NumberInput(attrs={'class': 'form-control'}),
            'openingcredit': forms.NumberInput(attrs={'class': 'form-control'}),
        }

    def __init__(self, *args, **kwargs):
        self.current_org = kwargs.pop('current_org', None)
        super(PartyForm, self).__init__(*args, **kwargs)
        for field in self.fields:
            self.fields[field].label = field.capitalize()
        if self.instance and self.instance.pk:
            self.fields['partyname'].widget.attrs['readonly'] = True

    def clean_partyname(self):
        name = self.cleaned_data.get('partyname')
        if not name:
            return name
        qs = HeadParty.objects.filter(partyname__iexact=name)
        if self.current_org:
            qs = qs.filter(org=self.current_org)
        if self.instance and self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            raise ValidationError("⚠️ This party already exists.")
        return name

class BrokerForm(BlankZeroModelForm):
    class Meta:
        model = Broker
        exclude = ['org']
        widgets = {
            'brokername': forms.TextInput(attrs={'class': 'form-control'}),
            'mobileno': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'openingdebit': forms.NumberInput(attrs={'class': 'form-control'}),
            'openingcredit': forms.NumberInput(attrs={'class': 'form-control'}),
            'remark': forms.Textarea(attrs={'class': 'form-control', 'rows': 2, 'data-next': '#saveBtn'}),
        }

    def __init__(self, *args, **kwargs):
        self.current_org = kwargs.pop('current_org', None)
        super(BrokerForm, self).__init__(*args, **kwargs)
        for field in self.fields:
            self.fields[field].label = field.capitalize()
        if self.instance and self.instance.pk:
            self.fields['brokername'].disabled = True

    def clean_brokername(self):
        brokername = self.cleaned_data.get('brokername')
        if not brokername:
            return brokername
        qs = Broker.objects.filter(brokername__iexact=brokername)
        if self.current_org:
            qs = qs.filter(org=self.current_org)
        if self.instance and self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            raise ValidationError("⚠️ This broker already exists.")
        return brokername

class FirmForm(forms.ModelForm):
    class Meta:
        model = Firm
        fields = ['firmname']
        widgets = {
            'firmname': forms.TextInput(attrs={'class': 'form-control'})
        }


class ItemForm(BlankZeroModelForm):
    class Meta:
        model = HeadItem
        fields = ['item_name']
        widgets = {
            'item_name': forms.TextInput(attrs={'class': 'form-control'}),
        }

    def __init__(self, *args, **kwargs):
        self.current_org = kwargs.pop('current_org', None)
        super().__init__(*args, **kwargs)
        self.fields['item_name'].label = "Item Name"

    def clean_item_name(self):
        name = self.cleaned_data.get('item_name')
        if not name:
            return name
        qs = HeadItem.objects.filter(item_name__iexact=name)
        if self.current_org:
            qs = qs.filter(org=self.current_org)
        if self.instance and self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            raise ValidationError("⚠️ This item already exists.")
        return name

# --------------------- SALE MASTER FORM ---------------------
class SaleMasterForm(BlankZeroModelForm):
    party = forms.ModelChoiceField(
        queryset=HeadParty.objects.all().order_by('partyname'),
        empty_label="Select Party",
        widget=forms.Select(attrs={'class': 'form-control'})
    )

    broker = forms.ModelChoiceField(
        queryset=Broker.objects.all().order_by('brokername'),
        empty_label="Select Broker",
        widget=forms.Select(attrs={'class': 'form-control'})
    )
    firm = forms.ModelChoiceField(
        queryset=Firm.objects.all().order_by('firmname'),
        empty_label="Select Firm",
        widget=forms.Select(attrs={'class': 'form-control'})
    )
    
    
    class Meta:
        model = SaleMaster
        fields = [
            'invno', 'invdate', 'awakno',
            'party', 'broker', 'firm','extra', 'vehicleno',
            'totalamt', 'batavpercent', 'batavamt',
            'dr', 'dramt', 'qi', 'other', 'total',
            'advance', 'netamt', 'remark'
        ]
        widgets = {
            'invno': forms.NumberInput(attrs={'class': 'form-control', 'readonly': 'readonly'}),
            'invdate': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'awakno': forms.TextInput(attrs={'class': 'form-control'}),
            'extra': forms.TextInput(attrs={'class': 'form-control'}),
            'vehicleno': forms.TextInput(attrs={'class': 'form-control'}),
            'totalamt': forms.NumberInput(attrs={'class': 'form-control'}),
            'batavpercent': forms.NumberInput(attrs={'class': 'form-control'}),
            'batavamt': forms.NumberInput(attrs={'class': 'form-control'}),
            
            'dr': forms.NumberInput(attrs={'class': 'form-control'}),
            'dramt': forms.NumberInput(attrs={'class': 'form-control'}),
            
            'qi': forms.NumberInput(attrs={'class': 'form-control'}),
            'other': forms.NumberInput(attrs={'class': 'form-control'}),
            'total': forms.NumberInput(attrs={'class': 'form-control'}),
            'advance': forms.NumberInput(attrs={'class': 'form-control'}),
            'netamt': forms.NumberInput(attrs={'class': 'form-control'}),
            'remark': forms.TextInput(attrs={'class': 'form-control', 'data-next': '#saveBtn'}),
        }

    def clean(self):
        cleaned_data = super().clean()
        return cleaned_data


class SaleDetailsForm(BlankZeroModelForm):
    item = forms.ModelChoiceField(
        queryset=HeadItem.objects.all().order_by('item_name'),
        empty_label="Select Item",
        widget=forms.Select(attrs={'class': 'form-control'})
    )

    class Meta:
        model = SaleDetails
        fields = [
            'item', 'bora', 'bn', 'bnwt', 'bo', 'bowt','tbwt',
            'qty', 'rate', 'amount', 'partywt', 'millwt', 'diffwt', 'frkwt','lotno'
        ]
        widgets = {
            'bora': forms.NumberInput(attrs={'class': 'form-control'}),
            'bn': forms.NumberInput(attrs={'class': 'form-control'}),
            'bnwt': forms.NumberInput(attrs={'class': 'form-control'}),
            'bo': forms.NumberInput(attrs={'class': 'form-control'}),
            'bowt': forms.NumberInput(attrs={'class': 'form-control'}),
            'tbwt': forms.NumberInput(attrs={'class': 'form-control'}),
            'qty': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateAmount()'}),
            'rate': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateAmount()'}),
            'amount': forms.NumberInput(attrs={'class': 'form-control', 'readonly': 'readonly'}),
            'partywt': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateDiffWt()'}),
            'millwt': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateDiffWt()'}),
            'diffwt': forms.NumberInput(attrs={'class': 'form-control', 'readonly': 'readonly'}),
            'frkwt': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01', 'oninput': 'calculateDiffWt()'}),
            'lotno': forms.TextInput(attrs={'class': 'form-control'}),
        }

    def clean(self):
        cleaned_data = super().clean()
        qty = cleaned_data.get('qty') or 0
        rate = cleaned_data.get('rate') or 0
        amount = cleaned_data.get('amount')
        try:
            if (not amount or float(amount) == 0) and qty and rate:
                cleaned_data['amount'] = float(qty) * float(rate)
        except Exception:
            pass
        return cleaned_data

# --------------------- PURCHASE MASTER FORM ---------------------
class PurchaseMasterForm(BlankZeroModelForm):
    party = forms.ModelChoiceField(
        queryset=HeadParty.objects.none(),
        empty_label="Select Party",
        widget=forms.Select(attrs={'class': 'form-control'})
    )
    broker = forms.ModelChoiceField(
        queryset=Broker.objects.none(),
        empty_label="Select Broker",
        widget=forms.Select(attrs={'class': 'form-control'})
    )

    class Meta:
        model = PurchaseMaster
        fields = [
            'invno', 'invdate', 'awakno',
            'party', 'broker', 'extra', 'vehicleno',
            'totalamt', 'batavpercent', 'batavamt', 
            'dr', 'dramt', 'qi', 'other', 'total',
            'advance', 'netamt', 'remark'
        ]
        widgets = {
            'invno': forms.NumberInput(attrs={'class': 'form-control', 'readonly': 'readonly'}),
            'invdate': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'awakno': forms.TextInput(attrs={'class': 'form-control'}),
            'extra': forms.TextInput(attrs={'class': 'form-control'}),
            'vehicleno': forms.TextInput(attrs={'class': 'form-control'}),
            'totalamt': forms.NumberInput(attrs={'class': 'form-control'}),
            'batavpercent': forms.NumberInput(attrs={'class': 'form-control'}),
            'batavamt': forms.NumberInput(attrs={'class': 'form-control'}),
            

            'dr': forms.NumberInput(attrs={'class': 'form-control'}),
            'dramt': forms.NumberInput(attrs={'class': 'form-control'}),
            
            'qi': forms.NumberInput(attrs={'class': 'form-control'}),
            'other': forms.NumberInput(attrs={'class': 'form-control'}),
            'total': forms.NumberInput(attrs={'class': 'form-control'}),
            'advance': forms.NumberInput(attrs={'class': 'form-control'}),
            'netamt': forms.NumberInput(attrs={'class': 'form-control'}),
            'remark': forms.TextInput(attrs={'class': 'form-control', 'data-next': '#saveBtn'}),
        }

    def clean(self):
        cleaned_data = super().clean()
        return cleaned_data

    def __init__(self, *args, **kwargs):
        self.current_org = kwargs.pop('current_org', None)
        super().__init__(*args, **kwargs)
        if self.current_org:
            self.fields['party'].queryset = HeadParty.objects.filter(org=self.current_org).order_by('partyname')
            self.fields['broker'].queryset = Broker.objects.filter(org=self.current_org).order_by('brokername')

# --------------------- PURCHASE DETAILS FORM ---------------------
class PurchaseDetailsForm(BlankZeroModelForm):
    item = forms.ModelChoiceField(
        queryset=HeadItem.objects.none(),
        empty_label="Select Item",
        widget=forms.Select(attrs={'class': 'form-control'})
    )

    class Meta:
        model = PurchaseDetails
        fields = [
            'item', 'bora', 'bn', 'bnwt', 'bo', 'bowt',
            'qty', 'rate', 'amount', 'partywt', 'millwt', 'diffwt', 'lotno'
        ]
        widgets = {
            'bora': forms.NumberInput(attrs={'class': 'form-control'}),
            'bn': forms.NumberInput(attrs={'class': 'form-control'}),
            'bnwt': forms.NumberInput(attrs={'class': 'form-control'}),
            'bo': forms.NumberInput(attrs={'class': 'form-control'}),
            'bowt': forms.NumberInput(attrs={'class': 'form-control'}),
            'qty': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateAmount()'}),
            'rate': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateAmount()'}),
            'amount': forms.NumberInput(attrs={'class': 'form-control', 'readonly': 'readonly'}),
            'partywt': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateDiffWt()'}),
            'millwt': forms.NumberInput(attrs={'class': 'form-control', 'oninput': 'calculateDiffWt()'}),
            'diffwt': forms.NumberInput(attrs={'class': 'form-control', 'readonly': 'readonly'}),
            'lotno': forms.TextInput(attrs={'class': 'form-control'}),
        }

    def clean(self):
        cleaned_data = super().clean()
        qty = cleaned_data.get('qty') or 0
        rate = cleaned_data.get('rate') or 0
        amount = cleaned_data.get('amount')
        try:
            if (not amount or float(amount) == 0) and qty and rate:
                cleaned_data['amount'] = float(qty) * float(rate)
        except Exception:
            pass
        return cleaned_data

    def __init__(self, *args, **kwargs):
        self.current_org = kwargs.pop('current_org', None)
        super().__init__(*args, **kwargs)
        if self.current_org:
            self.fields['item'].queryset = HeadItem.objects.filter(org=self.current_org).order_by('item_name')

class AllPartyBalanceForm(forms.Form):
    start_date = forms.DateField(widget=forms.DateInput(attrs={"type":"date"}))
    end_date   = forms.DateField(widget=forms.DateInput(attrs={"type":"date"}))
    party      = forms.ModelChoiceField(queryset=HeadParty.objects.none(),
                                        required=False, empty_label="All")
    report_type = forms.ChoiceField(choices=[("date", "Date Wise")], required=False)
