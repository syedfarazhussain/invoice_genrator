from django.shortcuts import render
from django.contrib.auth import get_user_model


# Create your views here.

def dashboard(request):
    template_admin = 'dashboard.html'
    template_user = 'user/test.html'

    context = {
        'title': 'Dashboard'
    }
    return render(request, template_admin, context)


def user_page(request):
    template = 'admin/users.html'
    is_user = []
    User = get_user_model()
    users = User.objects.filter(is_user=True)
    # bhai idher for loop ki zarorat nae thi just filter karna tha record agar users hi dikhany thy
    # for user in users:
    #     # Where user is whatever foreign key you specified that relates to user.
    #     if user.is_user:
    #         is_user.append(user)
    #     else:
    #         pass
    context = {
        "get_all_is_user": users,
        'title': 'Users'
    }
    return render(request, template, context)


def invoice_form(request):

    return render(request, 'admin/content.html')
