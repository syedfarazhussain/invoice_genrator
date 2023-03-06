import traceback
from django.db import IntegrityError
from django.shortcuts import render, redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.contrib.auth import get_user_model
from .forms import UserDataForm
import pandas as pd
from django.conf import settings
from django.core.mail import EmailMultiAlternatives
from django.template.loader import render_to_string
from django.utils.html import strip_tags
import openpyxl
import json
from django.core import serializers
from .utils import process_de_data
from time import sleep
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
    template_user = 'user/user_dashboard.html'

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


@login_required
def create_account_list(request):
    try:
        # Open the excel file and get the active sheet
        if request.method == 'POST' and request.FILES['account_data_file']:
            file = request.FILES['account_data_file']
            wb = openpyxl.load_workbook(file)
            ws = wb.active
    except:
        return render(request, 'user/user_dashboard.html', {'error_message': 'Invalid Excel File'})

    expected_columns = ["Account Name", "isConsidered"]
    header = [cell.value for cell in ws[1]]
    header = [col_name for col_name in header if col_name != 'SrN0.']
    if header != expected_columns:
        return render(request, 'user/user_dashboard.html', {'error_message': 'Invalid Data'})
    
    column_mapping = {
        "Account Name": "account_name",
        "isConsidered": "is_considered",
    }
    success_message_displayed = False

    for row in ws.iter_rows(min_row=2):
        account_name = row[1].value
        is_considered = row[2].value

        try:
            if is_considered == True:
                # Check if a row with the same contact_name and user already exists
                existing_data = AccountList.objects.get(account_name=account_name, user=request.user)

                # Update the existing row with the new data
                existing_data.is_considered = is_considered
                existing_data.save()

                if not success_message_displayed:
                    messages.success(request, 'Account_list Updated Successfully')
                    success_message_displayed = True
        except AccountList.DoesNotExist:
            # Create a new row with the user's data
            if is_considered == True:
                account = AccountList.objects.create(account_name=account_name, user=request.user, is_considered=True)
                account.save()
                if not success_message_displayed:
                    messages.success(request, 'Account_list uploaded successfully.')
                    success_message_displayed = True
    return render(request, "user/user_dashboard.html")


def browse_lifting_file(request):
    template_name = 'user/user_dashboard.html'
    confirm_continue_result = False

    if request.method == 'POST':
        if 'lift_fee_file' in request.FILES:
            file = request.FILES['lift_fee_file']

            lifting_file = file

            lift_fee_file = pd.read_excel(file)
            required_columns = ['Buy', 'Buy Amount', 'Sell', 'Sell Amount', 'Fee', 'Total Settlement', 'Beneficiary', 'When Booked', 'When Created', 'Created By', 'Delivery Method', 'Reference', 'Delivery Country ISO Code', 'Delivery Country Name', 'Your Reference', 'Cheque Number', 'Beneficiary ID', 'Execution Date', 'Payment Line Status', 'Payment Number', 'Processing Date', 'Bank Reference Number', 'AccountName', 'Bank Value Date', 'Exchange Rate', 'Our Reference', 'Charges Type', 'USD AMOUNT']

            if not set(required_columns).issubset(lift_fee_file.columns):
                messages.error(request, '!Invalid Lifting fee file.')
                return render(request, template_name)

            user_data = UserData.objects.all()
            if len(user_data) == 0:
                messages.error(request, '!No Email configuration found, please contact with admin')
                return render(request, template_name)
            else:
                user_data = serializers.serialize('json', user_data)  # serialize the UserData object to JSON
                user_data = json.loads(user_data)[0]['fields']  # deserialize the JSON to a dictionary and get the fields

            request.session['lifting_file'] = lifting_file
            request.session['user_data'] = user_data

            template_data = {
                'lifting_file': lifting_file,
                'user_data': user_data,
            }
            return render(request, 'user/confirm_continue.html', template_data)

        elif 'confirm_continue' in request.POST:
            confirm_continue_result = request.POST.get('confirm_continue')
            if confirm_continue_result == 'true':
                confirm_continue_result = True
            else:
                confirm_continue_result = False

            if not confirm_continue_result:
                return redirect('browse_lifting_file')

            lifting_file = request.session.get('lifting_file')
            user_data = request.session.get('user_data')

        try:
            process_de_data(request, lifting_file, user_data)
            messages.success(request, '!Process is completed successfully')

            subject = 'Process is completed successfully'
            html_content = render_to_string('user/user_dashboard.html', {'context': 'All invoices are generated and sent over recipient email(s) address'})
            text_content = strip_tags(html_content)
            msg = EmailMultiAlternatives(subject, text_content, settings.EMAIL_HOST_USER, ['recipient@example.com'])
            msg.attach_alternative(html_content, "text/html")
            msg.send()

        except Exception as e:
            traceback.print_exc()
            messages.error(request, f'!Error Occurred: {e}')
        finally:
            messages.success(request, '')

        return render(request, 'user/user_dashboard.html')
    else:
        messages.error(request, '!No file was uploaded.')
        return render(request, 'user/user_dashboard.html')




def cp_desk_view(request):
    return render(request, "user/cp_desk.html")


# def browse_cpdesk_file(request):
#     if request.method == 'POST' and request.FILES['cp_desk_file']:
#         file = request.FILES['cp_desk_file']
#         filename = file.name
#         DefaultPath = settings.MEDIA_ROOT

#         cp_desk_file = pd.read_excel(file)
#         required_columns = ['TLA', 'Principals', 'Inv no', 'To', 'CC']

#         if not set(required_columns).issubset(cp_desk_file.columns):
#             messages.error(request, '!Invalid CP Desk file.')
#             return redirect("cp_desk_view")

#         user_data = get_user_model().objects.filter(is_user=True)
#         if len(user_data) == 0:
#             messages.error(request, '!No Email configuration found, please contact with admin')
#             return render(request, 'user/user_dashboard.html')

#         messages.info(request, '!Checking File')
#         sleep(1)

#         try:
#             confirm_template = 'user/confirm_continue.html'
#             if request.method == 'POST' and 'confirm' in request.POST:
#                 confirm_continue_result = True
#             elif request.method == 'POST' and 'cancel' in request.POST:
#                 confirm_continue_result = False
#             else:
#                 confirm_continue_result = confirm_continue(request, confirm_template)
#             if not confirm_continue_result:
#                 return redirect('confirm_continue')

#             process_de_data_cp_desk(filename, user_data[0])
#             messages.success(request, '!Process is completed successfully')

#             subject = 'Process is completed successfully'
#             html_content = render_to_string('user/user_dashboard.html',
#                                             {'context': 'All invoices are generated and sent over recipient email(s) address'})
#             text_content = strip_tags(html_content)
#             msg = EmailMultiAlternatives(subject, text_content, settings.EMAIL_HOST_USER, ['recipient@example.com'])
#             msg.attach_alternative(html_content, "text/html")
#             msg.send()

#         except:
#             messages.error(request, '!Error Occurred.')
#         finally:
#             sleep(2)
#             messages.success(request, '')
#         return render(request, 'user/confirm_continue.html')

#     return render('cp_desk_view')