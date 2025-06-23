"""
Custom middleware for handling impersonation context in templates.
"""
from django.utils.functional import SimpleLazyObject
from django.contrib.auth import get_user_model

def get_impersonation_context(request):
    """
    Returns a dictionary with impersonation context.
    """
    context = {
        'impersonate_is_active': False,
        'impersonate_user': None,
        'is_impersonated': False
    }
    
    # Check Django's built-in impersonate first
    if hasattr(request, 'impersonate') and request.impersonate:
        context.update({
            'impersonate_is_active': True,
            'impersonate_user': request.impersonate,
            'is_impersonated': True
        })
    # Fallback to session-based check
    elif '_impersonate' in request.session:
        User = get_user_model()
        try:
            user = User.objects.get(pk=request.session['_impersonate'])
            context.update({
                'impersonate_is_active': True,
                'impersonate_user': user,
                'is_impersonated': True
            })
        except (User.DoesNotExist, KeyError, ValueError):
            pass
    
    return context

class ImpersonationMiddleware:
    """
    Middleware that adds impersonation context to the request object.
    This makes it easy to check impersonation status in templates.
    """
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        # Get impersonation context
        impersonation_context = get_impersonation_context(request)
        
        # Add to request for direct access
        request.impersonate_is_active = impersonation_context['impersonate_is_active']
        request.impersonate_user = impersonation_context['impersonate_user']
        
        # Add to request for context processor
        request.impersonation_context = SimpleLazyObject(lambda: impersonation_context)
        
        # Add impersonation class to body if active
        response = self.get_response(request)
        
        if hasattr(response, 'content') and impersonation_context['is_impersonated']:
            content = response.content.decode('utf-8')
            if '<body' in content and 'impersonate-active' not in content:
                content = content.replace('<body', '<body class="impersonate-active"')
                response.content = content.encode('utf-8')
                
        return response

def impersonation_context_processor(request):
    """
    Context processor that adds impersonation context to all templates.
    """
    return get_impersonation_context(request)
