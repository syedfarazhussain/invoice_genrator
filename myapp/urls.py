from django.urls import path
from .import views

urlpatterns = [
    path('', views.dashboard, name=''),
    path('users/', views.user_page, name='users'),
    path('settings/', views.settings, name='settings'),
    path('upload_files/', views.upload_files, name='upload_files'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('cp_desk_view/', views.cp_desk_view, name='cp_desk_view'),
    path('insert_user_data/', views.insert_user_data, name='insert_user_data'),
    path('update_user_data/<int:id>/update/', views.update_user_data, name='update_user_data'),
    path('upload_master_data/', views.upload_master_data, name='upload_master_data'),
    path('add_email_condition/', views.add_email_condition, name='add_email_condition'),
    path('add_group_condition/', views.add_group_condition, name='add_group_condition'),
    path('create_account_list/', views.create_account_list, name='create_account_list'),
    path('browse_lifting_file/', views.browse_lifting_file, name='browse_lifting_file'),
    path('split_lifting_file/', views.split_lifting_file, name='split_lifting_file'),
    path('get_Invoice_data/', views.get_Invoice_data, name='get_Invoice_data'),
    path('browse_cp_desk_file/', views.browse_cp_desk_file, name='browse_cp_desk_file'),
    path('split_cp_desk_file/', views.split_cp_desk_file, name='split_cp_desk_file'),
    path('get_cp_desk_Invoice_data/', views.get_cp_desk_Invoice_data, name='get_cp_desk_Invoice_data'),
]