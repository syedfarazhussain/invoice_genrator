from django.shortcuts import render
from django.urls import reverse_lazy 
from django.contrib.auth import get_user_model
from django.contrib.messages.views import SuccessMessageMixin
from django.views.generic import CreateView
from .models import Invoice
from .forms import InvoiceForm

# Create your views here.
class InvoiceCreateView(SuccessMessageMixin, CreateView):
    model = Invoice
    form_class = InvoiceForm
    template_name = 'contact.html'
    success_message = 'your message Successful send'
    success_url = reverse_lazy('home')


    def form_valid(self, form):
        form.instance.user = self.request.user
        return super().form_valid(form)
    
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
