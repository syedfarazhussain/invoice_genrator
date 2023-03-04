from django.contrib import admin
from django.apps import apps

# Register your models here.

app_config = apps.get_app_config('myapp')
models = app_config.get_models()

for model in models:
    admin.site.register(model)