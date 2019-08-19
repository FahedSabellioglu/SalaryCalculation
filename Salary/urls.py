from django.conf.urls import url
from . import views

urlpatterns = [
    url(r'^$',views.SalaryCalculation,name='SalaryCalculation'), # default Salary
    url(r"^salary/calculation/$",views.renderTable,name='tableData'),
    url(r"^salary/paremeter/$",views.paremeters,name="Parameters"),
    url(r'^DownloadFile/$',views.DownloadFile,name="export")
]
