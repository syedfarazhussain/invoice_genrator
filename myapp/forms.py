from django.forms import ModelForm
from .models import UserData


class UserDataForm(ModelForm):
    class Meta:
        model = UserData
        fields =  ['gbp_rate', 'smtp_user', 'smtp_server','smtp_port', 'smtp_pass','smtp_sender']
