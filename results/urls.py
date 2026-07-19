
from django.urls import path
from . import views

urlpatterns = [
    path('', views.home_view, name='home'),
    path('stats/', views.stats_view, name='stats'),  # هـــذا هو السطر الناقص الذي سيحل المشكلة!
    path('control-panel/upload/', views.upload_excel_view, name='upload_excel'),
    path('control-panel/import/', views.import_mapped_data_view, name='import_mapped_data'),
]