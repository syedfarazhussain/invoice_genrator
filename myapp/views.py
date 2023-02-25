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

def user_page(request):
    is_user = []
    User = get_user_model()
    users = User.objects.all()
    for user in users:
        # Where user is whatever foreign key you specified that relates to user.
        if user.is_user:
            is_user.append(user)
        else:
            pass
    context = {
        "get_all_is_user" : is_user,
    }
    return render(request, 'admin/user_page.html', context)

def invoice_form(request):
    
    return render(request, 'admin/content.html')
