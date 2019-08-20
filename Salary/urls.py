from django.conf.urls import url
from . import views

urlpatterns = [
    url(r'^$',views.SalaryCalculation,name='SalaryCalculation'), # default Calculator
    url(r"^calculation/$",views.renderTable,name='tableData'), #calculator/calculation
    url(r"^parameter/$",views.paremeters,name="Parameters"), #calculator/parameter
    url(r'^DownloadFile/$',views.DownloadFile,name="export")
]
