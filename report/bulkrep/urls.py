from django.urls import path
from . import views


app_name = 'bulkrep'
 
urlpatterns = [
    path('', views.home, name='home'),
    path('single-report/', views.single_report, name='single_report'),
    path('bulk-report/', views.bulk_report, name='bulk_report'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('dashboard-api/', views.dashboard_api, name='dashboard_api'),
    path('download-churned-subscribers/', views.download_churned_subscribers, name='download_churned_subscribers'),
    path('download-new-subscribers/', views.download_new_subscribers, name='download_new_subscribers'),
    path('api/new-subscribers-trend/', views.new_subscribers_trend_api, name='new_subscribers_trend_api'),
    path('api/usage-trends/', views.usage_trends_api, name='usage_trends_api'),
]