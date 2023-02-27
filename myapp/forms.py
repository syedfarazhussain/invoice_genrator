from django.forms import ModelForm
from .models import Invoice


class InvoiceForm(ModelForm):
    class Meta:
        model = Invoice
        fields =  ['gbp_rate', 'smtp_user', 'smtp_server','smtp_port', 'smtp_pass','smtp_sender']