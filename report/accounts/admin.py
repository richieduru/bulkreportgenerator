from django.contrib import admin
from django.contrib.auth import get_user_model
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth.forms import UserCreationForm, UserChangeForm
from django.utils.html import format_html
from django.urls import reverse
from django import forms

User = get_user_model()

class UserCreationForm(UserCreationForm):
    class Meta(UserCreationForm.Meta):
        model = User
        fields = ('username', 'email', 'first_name', 'last_name')

class UserChangeForm(UserChangeForm):
    class Meta(UserChangeForm.Meta):
        model = User
        fields = '__all__'

class UserAdmin(BaseUserAdmin):
    add_form = UserCreationForm
    form = UserChangeForm
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_active', 'is_staff', 'last_login', 'impersonate_button')
    list_filter = ('is_staff', 'is_superuser', 'is_active', 'groups')
    search_fields = ('username', 'first_name', 'last_name', 'email')
    ordering = ('-date_joined',)
    
    fieldsets = (
        (None, {'fields': ('username', 'password')}),
        ('Personal Info', {'fields': ('first_name', 'last_name', 'email')}),
        ('Permissions', {
            'fields': ('is_active', 'is_staff', 'is_superuser', 'groups', 'user_permissions'),
        }),
        ('Important dates', {'fields': ('last_login', 'date_joined')}),
    )
    
    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': ('username', 'email', 'password1', 'password2', 'is_active', 'is_staff'),
        }),
    )
    
    def impersonate_button(self, obj):
        return format_html(
            '<a class="button" href="{}" title="Impersonate this user" style="background: #4CAF50; color: white; padding: 5px 10px; border-radius: 4px; text-decoration: none;">' +
            '<i class="fas fa-user-secret"></i> Impersonate' +
            '</a>',
            reverse('impersonate-start', args=[obj.id])
        ) if obj.is_active else ''
    
    impersonate_button.short_description = 'Actions'
    impersonate_button.allow_tags = True

# Unregister the default User admin and register our custom one
admin.site.unregister(User)
admin.site.register(User, UserAdmin)
