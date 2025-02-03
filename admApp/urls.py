
from django.contrib import admin
from django.urls import path
from  .views import ExcelNCRView,ExcelPlantillaBCPView
urlpatterns = [
    path('excel_def/', ExcelNCRView.as_view(),name="def-excel"),
    path('excel_plantilla_bcp/', ExcelPlantillaBCPView.as_view(),name="plantilla-bcp-excel"),
] 
