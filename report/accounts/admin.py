from django.contrib import admin
from django.contrib.auth import get_user_model
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.utils.html import format_html
from django.urls import reverse
from django.utils.safestring import mark_safe

User = get_user_model()

class UserAdmin(BaseUserAdmin):
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_staff', 'impersonate_button')
    list_filter = ('is_staff', 'is_superuser', 'is_active', 'groups')
    search_fields = ('username', 'first_name', 'last_name', 'email')
    
    def impersonate_button(self, obj):
        return format_html(
            '<a class="button" href="{}" title="Impersonate this user">' +
            '<i class="fas fa-user-secret"></i> Impersonate' +
            '</a>',
            reverse('impersonate-start', args=[obj.id])
        )
    impersonate_button.short_description = 'Actions'
    impersonate_button.allow_tags = True

# Unregister the default User admin and register our custom one
admin.site.unregister(User)
admin.site.register(User, UserAdmin)
