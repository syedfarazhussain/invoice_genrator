from django.contrib.auth.models import User
from django.shortcuts import render, redirect
from django.http import HttpResponseRedirect
from .forms import AdminForm, UserForm
from django.contrib.auth import login as dj_login
from django.contrib import auth
from django.contrib import messages
from django.contrib.auth import authenticate

from django.contrib.auth import get_user_model
User = get_user_model()


def admin_signup(request):

    if request.user.is_authenticated:
    # Do something for authenticated users.
        return render(request, 'base.html', {'error_message': 'please logout to the website to signup'})
    
    # if this is a POST request we need to process the form data
    template = 'admin_register.html'

    if request.method == 'POST':
        # create a form instance and populate it with data from the request:
        form = AdminForm(request.POST)
        # check whether it's valid:
        if form.is_valid():
            if User.objects.filter(username=form.cleaned_data['username']).exists():
                return render(request, template, {
                    'form': form,
                    'error_message': 'Username already exists.'
                })
            elif form.cleaned_data['password'] != form.cleaned_data['password_repeat']:
                return render(request, template, {
                    'form': form,
                    'error_message': 'Passwords do not match.'
                })
            else:
                # Create the user:
                user = User.objects.create_user(
                    form.cleaned_data['username'],
                )
                user.set_password(form.cleaned_data['password'])
                user.email = ''
                user.is_admin = True
                user.save()
                registred = True

                # redirect to accounts page:
                messages.success(
                    request, ('User has been registerd successfully '))
                return redirect('login')

    # No post data availabe, let's just show the page.
    else:
        form = AdminForm()

        return render(request, template, {'form': form})


def create_user(request):
    # if this is a POST request we need to process the form data
    template = 'admin/users.html'

    if request.method == 'POST':
        # create a form instance and populate it with data from the request:
        form = UserForm(request.POST)
        # check whether it's valid:
        if form.is_valid():
            if User.objects.filter(username=form.cleaned_data['username']).exists():
                return render(request, template, {
                    'form'          : form,
                    messages.error  : 'Username already exists.'
                })
            elif form.cleaned_data['password'] != form.cleaned_data['password_repeat']:
                return render(request, template, {
                    'form'          : form,
                    messages.error  : 'Passwords do not match.'
                })
            else:
                # Create the user:
                user = User.objects.create_user(
                    form.cleaned_data['username'],
                )
                user.set_password(form.cleaned_data['password'])
                user.email = ''
                user.is_user = True
                user.save()
                registred = True

                # redirect to accounts page:
                messages.success(
                    request, ('User has been added successfully '))
                return redirect('users')

   # No post data availabe, let's just show the page.
    else:
        form = UserForm()

    return render(request, template, {'form': form})


def login(request):

    if request.user.is_authenticated:
        username = request.user
    # Do something for authenticated users.
        return render(request, 'base.html', {'error_message': 'user has been already loggedin'})

    if request.method == 'POST':
        # Process the request if posted data are available
        username = request.POST['username']
        password = request.POST['password']
        # Check username and password combination if correct
        user = authenticate(username=username, password=password)

        if user is not None:
            # Save session as cookie to login the user
            dj_login(request, user)
            # Success, now let's login the user.
            messages.success(request, ('User has been login successfully '))
            if user.is_superuser == True:
                return redirect('/')
            if user.is_admin == True:
                return redirect('/')
            if user.is_user == True:
                return redirect('/')
        else:
            # Incorrect credentials, let's throw an error to the screen.
            messages.error(request, 'Incorrect username and / or password.')
            return redirect('login')
    else:
        # Do something for anonymous users.
        if request.method == 'POST':
            # Process the request if posted data are available
            username = request.POST['username']
            password = request.POST['password']
            # Check username and password combination if correct
            user = authenticate(username=username, password=password)
            
            if user is not None:
                # Save session as cookie to login the user
                dj_login(request, user)
                # Success, now let's login the user.
                messages.success(request, ('User has been login successfully '))
                if user.is_superuser == True:
                    return redirect('home')
                if user.is_admin == True:
                    return redirect('home')
                if user.is_user == True:
                    return redirect('home')
            else:
                # Incorrect credentials, let's throw an error to the screen.
                messages.error(request, 'Incorrect username and / or password.')
                return redirect('login')
        else:
            # No post data availabe, let's just show the page to the user.
            return render(request, 'login.html')


def logout(request):
    auth.logout(request)
    return redirect('login')
