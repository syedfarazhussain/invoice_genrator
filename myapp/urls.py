from django.urls import path
from .import views

urlpatterns = [
    path('', views.dashboard, name=''),
    path('users/', views.user_page, name='users'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('insert_user_data/', views.insert_user_data, name='insert_user_data'),
    path('upload_master_data/', views.upload_master_data, name='upload_master_data'),
    path('add_email_condition/', views.add_email_condition, name='add_email_condition'),
    path('add_group_condition/', views.add_group_condition, name='add_group_condition'),
]