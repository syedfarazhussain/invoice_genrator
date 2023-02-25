from django.db import models
from django.conf import settings
# Create your models here.

class Invoice(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, null=True, blank=True)
    gbp_rate = models.CharField(max_length=200)
    smtp_user = models.CharField(max_length=200)
    smtp_server = models.CharField(max_length=200)
    smtp_port = models.DateTimeField()
    smtp_pass = models.CharField(max_length=50)
    smtp_sender = models.EmailField()
  
    def __str__(self):
        return str(self.user)