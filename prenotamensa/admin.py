from django.contrib import admin

from .models import Personale
from .models import AssociazioneAtt

# Register your models here.

admin.site.register(Personale)
admin.site.register(AssociazioneAtt)
