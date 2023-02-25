from django.urls import path
from .import views


urlpatterns = [
    path('', views.home, name='home'),
    path('user_page/', views.user_page, name='user_page'),
]