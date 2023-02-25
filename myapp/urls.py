from django.urls import path
from .import views
from .views import (
    InvoiceCreateView, 
)

urlpatterns = [
    path('', views.home, name='home'),
    path('user_page/', views.user_page, name='user_page'),
    path('CreateInvoice/', InvoiceCreateView.as_view(), name='CreateInvoice'),
]