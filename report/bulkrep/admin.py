from django.contrib import admin
from .models import Usagereport, ReportGeneration
from django.utils.html import format_html
from django.urls import reverse
from django.utils.safestring import mark_safe
from django.utils import timezone


class UsagereportAdmin(admin.ModelAdmin):
    list_display = ('SubscriberName', 'DetailsViewedDate', 'ProductName', 'SystemUser')
    list_filter = ('DetailsViewedDate', 'ProductName')
    search_fields = ('SubscriberName', 'ProductName', 'SystemUser', 'SearchIdentity')
    date_hierarchy = 'DetailsViewedDate'
    ordering = ('-DetailsViewedDate', 'SubscriberName')
    list_per_page = 50


class ReportGenerationAdmin(admin.ModelAdmin):
    list_display = ('user', 'report_type', 'status', 'formatted_generated_at', 'formatted_completed_at', 'subscriber_name', 'duration_display')
    list_filter = ('report_type', 'status', 'generated_at')
    search_fields = ('user__username', 'subscriber_name', 'generator')
    date_hierarchy = 'generated_at'
    readonly_fields = ('generated_at', 'completed_at', 'duration_display')
    list_per_page = 50
    
    def formatted_generated_at(self, obj):
        if obj.generated_at:
            local_time = timezone.localtime(obj.generated_at)
            return local_time.strftime('%Y-%m-%d %H:%M:%S')
        return "-"
    formatted_generated_at.short_description = 'Generated At (Local)'
    formatted_generated_at.admin_order_field = 'generated_at'
    
    def formatted_completed_at(self, obj):
        if obj.completed_at:
            local_time = timezone.localtime(obj.completed_at)
            return local_time.strftime('%Y-%m-%d %H:%M:%S')
        return "-"
    formatted_completed_at.short_description = 'Completed At (Local)'
    formatted_completed_at.admin_order_field = 'completed_at'
    
    def duration_display(self, obj):
        if obj.duration is not None:
            return f"{obj.duration:.2f} seconds"
        return "-"
    duration_display.short_description = 'Duration'


# Register your models with the custom admin classes
admin.site.register(Usagereport, UsagereportAdmin)
admin.site.register(ReportGeneration, ReportGenerationAdmin)
