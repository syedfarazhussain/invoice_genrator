from django.db import IntegrityError
from django.shortcuts import render, redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.contrib.auth import get_user_model
from .forms import UserDataForm
import pandas as pd
import openpyxl
from .models import (UserData,
                    EmailCondition,
                    GroupComp,
                    MasterData,
                    Invoices,
                    AccountList,
                    CpDeskInvoices,
                    )

# Create your views here.
@login_required
def dashboard(request):
    template_admin = 'dashboard.html'
    template_user = 'user/test.html'

    if request.user.is_authenticated:
        if request.user.is_admin:
            context = {
                'title': 'Dashboard'
            }
            return render(request, template_admin, context)
        else:
            return render(request, template_user)
    else:
        return redirect('account/login')


@login_required
def user_page(request):
    template = 'admin/users.html'
    is_user = []
    User = get_user_model()
    users = User.objects.filter(is_user=True)
    
    context = {
        "get_all_is_user": users,
        'title': 'Users'
    }
    return render(request, template, context)


@login_required
def insert_user_data(request):
    user_name = request.user
    
    if request.method == 'POST':
        form = UserDataForm(request.POST)
        if form.is_valid():
            gbp_rate = form.cleaned_data['gbp_rate']
            smtp_server = form.cleaned_data['smtp_server']
            smtp_port = form.cleaned_data['smtp_port']
            smtp_user = form.cleaned_data['smtp_user']
            smtp_sender = form.cleaned_data['smtp_sender']
            smtp_pass = form.cleaned_data['smtp_pass']

            user = UserData.objects.filter(user=user_name).first()

            if user:
                user.gbp_rate = gbp_rate
                user.smtp_server = smtp_server
                user.smtp_port = smtp_port
                user.smtp_user = smtp_user
                user.smtp_sender = smtp_sender
                user.smtp_pass = smtp_pass
                user.save()
                messages.success(request, 'User Data Updated Successfully')
                return redirect('dashboard')

            else:
                UserData.objects.create(
                    gbp_rate=gbp_rate,
                    smtp_server=smtp_server,
                    smtp_port=smtp_port,
                    smtp_user=smtp_user,
                    smtp_sender=smtp_sender,
                    smtp_pass=smtp_pass,
                    user = request.user
                )

            messages.success(request, 'User Data Uploaded Successfully')
            return redirect('dashboard')

    else:
        form = UserDataForm()

    return render(request, 'dashboard.html', {'form': form})


@login_required
def upload_master_data(request):
    # Open the excel file and get the active sheet
    try:
        if request.method == 'POST' and request.FILES['master_data_file']:
            file = request.FILES['master_data_file']
            wb = openpyxl.load_workbook(file)
            ws = wb.active
    except:
        return render(request, 'dashboard.html', {'error_message': '"Invalid Excel File"'})

    # Validate the columns
    expected_columns = ["*ContactName", "Customer Address", "Email id", "CC email id", "*Description", "VAT Type", "Status"]
    header = [cell.value for cell in ws[1]]
    if header != expected_columns:
        return render(request, 'dashboard.html', {'error_message': '"Invalid Data"'})

    # Map the columns to the corresponding fields in the SQL table
    column_mapping = {
        "*ContactName": "contact_name",
        "Customer Address": "customer_address",
        "Email id": "email_id",
        "CC email id": "cc_email_id",
        "*Description": "description",
        "VAT Type": "vat_type",
        "Status": "status",
    }

    # Iterate through the rows and add the data to the SQL table
    for row in ws.iter_rows(min_row=2):
        contact_name = row[0].value
        data = {
            column_mapping[header[i]]: row[i].value
            for i in range(1, len(header))
        }

        # Check if a row with the same contact_name and user already exists
        existing_data = MasterData.objects.filter(contact_name=contact_name, user=request.user).first()

        if existing_data:
            # Update the existing row with the new data
            existing_data.customer_address = data["customer_address"]
            existing_data.email_id = data["email_id"]
            existing_data.cc_email_id = data["cc_email_id"]
            existing_data.description = data["description"]
            existing_data.vat_type = data["vat_type"]
            existing_data.status = data["status"]
            existing_data.save()
        else:
            # Create a new row with the user's data
            master_data = MasterData.objects.create(
                contact_name=contact_name,
                customer_address=data["customer_address"],
                email_id=data["email_id"],
                cc_email_id=data["cc_email_id"],
                description=data["description"],
                vat_type=data["vat_type"],
                status=data["status"],
                user=request.user
            )

    messages.success(request, ('Master Data Uploaded Successfully'))
    return redirect('dashboard')


@login_required
def add_email_condition(request):
    try:
        # Open the excel file and get the active sheet
        if request.method == 'POST' and request.FILES['email_data_file']:
            file = request.FILES['email_data_file']
            wb = openpyxl.load_workbook(file)
            ws = wb.active
    except:
        return render(request, 'dashboard.html', {'error_message': 'Invalid Excel File'})

    # Validate the columns
    expected_columns = ["Name", "Type"]
    header = [cell.value for cell in ws[1]]
    if 'Sr.No.' in header:
        header.remove('Sr.No.')
    if header != expected_columns:
        return render(request, 'dashboard.html', {'error_message': 'Invalid Data'})

    # Map the columns to the corresponding fields in the SQL table
    column_mapping = {
        "Name": "contact_name",
        "Type": "type",
    }

    # Initialize the variable for success message display
    success_message_displayed = False

    # Iterate through the rows and add the data to the SQL table
    for row in ws.iter_rows(min_row=2):
        contact_name = row[1].value
        data = {
            column_mapping[header[i]]: row[i+1].value
            for i in range(len(header))
        }

        try:
            # Check if a row with the same contact_name and user already exists
            existing_data = EmailCondition.objects.get(contact_name=contact_name, user=request.user)

            # Update the existing row with the new data
            existing_data.type = data["type"]
            existing_data.save()

            if not success_message_displayed:
                messages.success(request, 'Email Data Updated Successfully')
                success_message_displayed = True
        except EmailCondition.DoesNotExist:
            # Create a new row with the user's data
            email_data = EmailCondition.objects.create(
                contact_name=contact_name,
                type=data["type"],
                user=request.user
            )

            if not success_message_displayed:
                messages.success(request, 'Email Data Uploaded Successfully')
                success_message_displayed = True
    return redirect('dashboard')


@login_required
def add_group_condition(request):
    
    try:
        # Open the excel file and get the active sheet
        if request.method == 'POST' and request.FILES['parent_child_file']:
            file = request.FILES['parent_child_file']
            wb = openpyxl.load_workbook(file)
            ws = wb.active
    except:
        return render(request, 'dashboard.html', {'error_message': 'Invalid Excel File'})

    # Validate the columns
    expected_columns = ["Parent", "Child","Parent_Company_Name"]
    header = [cell.value for cell in ws[1]]
    if header != expected_columns:
        return render(request, 'dashboard.html', {'error_message': 'Invalid Data'})

    # Map the columns to the corresponding fields in the SQL table
    column_mapping = {
        "Parent": "Parent_Comp",
        "Child": "Child_Comp",
        "Parent_Company_Name" : "Parent_Company_Name"
    }
    
    # Initialize the variable for success message display
    success_message_displayed = False

    # Iterate through the rows and add the data to the SQL table
    for row in ws.iter_rows(min_row=2):
        Parent_Comp = row[0].value
        data = {
            column_mapping[header[i+1]]: row[i+1].value if row[i+1].value else None
            for i in range(len(header)-1)
        }

        if data["Child_Comp"] is not None:
            try:
                # Check if a row with the same Parent_Comp and Child_Comp already exists
                existing_data = GroupComp.objects.get(Parent_Comp=Parent_Comp, Child_Comp=data["Child_Comp"], user=request.user)

                # Update the existing row with the new data
                existing_data.Parent_Company_Name = data["Parent_Company_Name"]
                existing_data.save()
                if not success_message_displayed:
                    messages.success(request, 'Group Parent-child Data Updated Successfully')
                    success_message_displayed = True
                    
            except GroupComp.DoesNotExist:
                # Create a new row with the user's data
                try:
                    group_parent_child_data = GroupComp.objects.create(
                        Parent_Comp=Parent_Comp,
                        Child_Comp=data["Child_Comp"],
                        Parent_Company_Name=data["Parent_Company_Name"],
                        user=request.user
                    )
                except IntegrityError as e:
                    messages.error(request, 'Error: {}'.format(str(e)))
                    return redirect('dashboard')
                    
                if not success_message_displayed:
                    messages.success(request, 'Group Parent-child Data Uploaded Successfully')
                    success_message_displayed = True
                    
            except GroupComp.MultipleObjectsReturned:
                # Handle the case when multiple rows with the same Parent_Comp and Child_Comp values exist
                messages.error(request, 'Multiple rows with the same Parent_Comp and Child_Comp values exist')
                return redirect('dashboard')
                    
        else:
            # Delete the row with null values
            GroupComp.objects.filter(Parent_Comp=Parent_Comp, Child_Comp__isnull=True, user=request.user).delete()
            if not success_message_displayed:
                messages.error(request, 'Group Parent-child Data null row has been deleted')
                success_message_displayed = True
    return redirect('dashboard')

