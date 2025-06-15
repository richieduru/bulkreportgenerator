from django.urls import path
from . import views


app_name = 'bulkrep'
 
urlpatterns = [
    path('', views.home, name='home'),
    path('single-report/', views.single_report, name='single_report'),
    path('bulk-report/', views.bulk_report, name='bulk_report'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('dashboard-api/', views.dashboard_api, name='dashboard_api'),
]