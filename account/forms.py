from django import forms
from django.contrib.auth.forms import UserCreationForm
from django.db import transaction
from django.forms.utils import ValidationError

from .models import User

class AdminForm(forms.Form):
    class Meta(UserCreationForm.Meta):
        model = User
    
    username = forms.CharField(widget=forms.TextInput(attrs={'class':'form-control'}))
    password = forms.CharField(widget=forms.PasswordInput(attrs={'class':'form-control'}))
    password_repeat = forms.CharField(widget=forms.PasswordInput(attrs={'class':'form-control'}))

    @transaction.atomic
    def save(self):
        user = super().save(commit=False)
        user.is_admin = True
        user.save()
        admin = User.objects.create(user=user)
        return admin

class UserForm(forms.Form):
    class Meta(UserCreationForm.Meta):
        model = User

    username = forms.CharField(widget=forms.TextInput(attrs={'class':'form-control'}))
    password = forms.CharField(widget=forms.PasswordInput(attrs={'class':'form-control'}))
    password_repeat = forms.CharField(widget=forms.PasswordInput(attrs={'class':'form-control'}))

    def save(self, commit=True):
        user = super().save(commit=False)
        user.is_user= True
        if commit:
            user.save()
        return user


