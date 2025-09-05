from django.urls import path
from .views import gatepass,visitorpass,uniformslip

urlpatterns = [
    path("gatepass/", gatepass, name="gatepass"),
    path("visitorpass/", visitorpass, name="visitorpass"),
    path("uniformslip/", uniformslip, name="uniformslip"),
]
