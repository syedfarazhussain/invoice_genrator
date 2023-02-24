from django.urls import path
from .import views


urlpatterns = [
    path('admin_signup/', views.admin_signup, name='admin_signup'),
    path('create_user/', views.create_user, name='create_user'),
    path('login/', views.login, name='login'),
    path('logout/', views.logout, name='logout'),
]