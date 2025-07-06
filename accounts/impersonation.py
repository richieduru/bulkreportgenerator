"""
Impersonation utilities for the accounts app.
"""
from django.contrib.auth import get_user_model

def get_impersonatable_users(request):
    """
    Returns a queryset of users that can be impersonated.
    By default, returns all non-superusers.
    
    Args:
        request: The current request object
    """
    User = get_user_model()
    return User.objects.filter(is_superuser=False)
