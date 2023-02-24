from django.shortcuts import render
from django.contrib.auth import get_user_model


# Create your views here.

def home(request):
    if request.user.is_authenticated:
        if request.user.is_admin:
            return render(request, 'admin/content.html')
        else:
            return render(request, 'user/test.html')
    else:
        return render(request, 'base.html')

def invoice_form(request):
    return render(request, 'admin/content.html')