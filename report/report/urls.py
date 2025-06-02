"""
URL configuration for report project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from django.urls import re_path
from django.views.generic import RedirectView
from django.views.static import serve

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('bulkrep.urls')),
    path('accounts/', include('accounts.urls', namespace='accounts')),
    # Impersonation URLs
    path('impersonate/', include('impersonate.urls')),
    # Redirect to admin after impersonation
    path('impersonate/', include('impersonate.urls'), {'next': '/admin/'}),
    # Add a simple view to stop impersonation
    path('stop-impersonate/', RedirectView.as_view(pattern_name='impersonate-stop'), name='stop_impersonate'),
]
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)