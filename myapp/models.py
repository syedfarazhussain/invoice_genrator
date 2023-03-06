from django.db import models
from django.conf import settings
# Create your models here.


class UserData(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    gbp_rate = models.CharField(max_length=200)
    smtp_user = models.CharField(max_length=200)
    smtp_server = models.CharField(max_length=200)
    smtp_port = models.CharField(max_length=200)
    smtp_pass = models.CharField(max_length=200)
    smtp_sender = models.CharField(max_length=200)
  
    def __str__(self):
        return str(self.user)


class EmailCondition(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    contact_name = models.CharField(max_length=200)
    type = models.CharField(max_length=200)

    def __str__(self):
        return str(self.user)


class GroupComp(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    Child_Comp = models.CharField(max_length=200)
    Parent_Comp = models.CharField(max_length=200)
    Parent_Company_Name = models.CharField(max_length=200)

    def __str__(self):
        return str(self.user)
    

class MasterData(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    contact_name = models.CharField(max_length=200)
    customer_address = models.CharField(max_length=200)
    email_id = models.CharField(max_length=200)
    cc_email_id = models.CharField(max_length=200)
    description = models.CharField(max_length=200)
    vat_type = models.CharField(max_length=200)
    status = models.CharField(max_length=200)
  
    def __str__(self):
        return str(self.user)


class Invoices(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    account_name = models.CharField(max_length=200)
    invoice_file_name = models.CharField(max_length=200)
    invoice_number = models.CharField(max_length=200)
    previous_month = models.CharField(max_length=200)
    previous_year = models.CharField(max_length=200)
    current_year = models.CharField(max_length=200)
    total_cost = models.CharField(max_length=200)
    user_name = models.CharField(max_length=200)
    invoice_Type = models.CharField(max_length=200)
    
    def __str__(self):
        return str(self.account_name)
    

class AccountList(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    account_name = models.CharField(max_length=200)
    user_name = models.CharField(max_length=200)
    is_considered = models.CharField(max_length=200)

    def __str__(self):
        return str(self.user)
    

class CpDeskInvoices(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    account_name = models.CharField(max_length=200)
    invoice_file_pdf = models.CharField(max_length=200)
    invoice_file_excel = models.CharField(max_length=200)
    invoice_number = models.CharField(max_length=200)
    previous_month = models.CharField(max_length=200)
    previous_year = models.CharField(max_length=200)
    current_year = models.CharField(max_length=200)
    user_name = models.CharField(max_length=200)
    
    def __str__(self):
        return str(self.account_name)