from django.urls import path
from .import views
from .views import (
    InvoiceCreateView, 
)

urlpatterns = [
    path('', views.dashboard, name=''),
    path('users/', views.user_page, name='users'),
    path('dashboard/', views.dashboard, name='dashboard'),
]