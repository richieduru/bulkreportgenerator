"""
Middleware to prevent caching of authenticated pages and API responses.
"""
from django.utils.cache import add_never_cache_headers

class PreventCacheMiddleware:
    """
    Middleware to prevent caching of authenticated pages and API responses.
    This ensures that users can't access cached versions of pages after logging out.
    """
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        
        # Don't cache authenticated pages
        if hasattr(request, 'user') and request.user.is_authenticated:
            add_never_cache_headers(response)
            
            # Additional headers to prevent caching
            response['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0, private'
            response['Pragma'] = 'no-cache'
            response['Expires'] = '0'
            
        return response
