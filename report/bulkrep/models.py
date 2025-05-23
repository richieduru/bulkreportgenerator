from django.db import models
from django.contrib.auth import get_user_model
from django.utils import timezone

# Create your models here.

class Usagereport(models.Model):
    # Define a primary key field that Django will use internally but won't be part of SQL queries
    # This solves the "Invalid column name 'id'" error
    id = models.CharField(primary_key=True, max_length=255, editable=False)
    
    # Actual fields from the database
    SubscriberName = models.CharField(max_length=255, db_column='SubscriberName')
    DetailsViewedDate = models.DateField(db_column='DetailsViewedDate')
    ProductName = models.CharField(max_length=255, db_column='ProductName')
    SystemUser = models.CharField(max_length=255, db_column='SystemUser', null=True, blank=True)
    SearchIdentity = models.CharField(max_length=255, db_column='SearchIdentity', null=True, blank=True)
    SubscriberEnquiryDate = models.DateField(db_column='SubscriberEnquiryDate', null=True, blank=True)
    SearchOutput = models.TextField(db_column='SearchOutput', null=True, blank=True)
    ProductInputed = models.CharField(max_length=255, db_column='ProductInputed', null=True, blank=True)

    class Meta:
        managed = False  # Tell Django not to manage this table
        db_table = 'usagereport' # Specify the existing table name
    
    def __str__(self):
        return f"{self.SubscriberName} - {self.ProductName}"
    
    # Override the save method to set the id field based on other fields
    def save(self, *args, **kwargs):
        # Create a unique identifier from the field values
        self.id = f"{self.SubscriberName}_{self.DetailsViewedDate}_{self.ProductName}"
        super().save(*args, **kwargs)


class ReportGeneration(models.Model):
    """
    Tracks report generation events by users.
    """
    REPORT_TYPES = [
        ('single', 'Single Report'),
        ('bulk', 'Bulk Report'),
        ('both', 'Both Single and Bulk')
    ]
    
    STATUS_CHOICES = [
        ('success', 'Success'),
        ('failed', 'Failed'),
        ('in_progress', 'In Progress'),
    ]
    
    user = models.ForeignKey(
        get_user_model(),
        on_delete=models.CASCADE,
        help_text='User who generated the report',
        db_index=True,
        related_name='generated_reports'
    )
    
    generator = models.CharField(
        max_length=255,
        help_text='Name of the user who generated the report',
        db_index=True,
        null=True,
        blank=True,
        default='Unknown'
    )
    
    report_type = models.CharField(
        max_length=10,
        choices=REPORT_TYPES,
        help_text='Type of report generated',
        db_index=True
    )
    
    status = models.CharField(
        max_length=15,
        choices=STATUS_CHOICES,
        default='success',
        help_text='Status of the report generation'
    )
    
    generated_at = models.DateTimeField(
        auto_now_add=True,
        help_text='When the report was generated',
        db_index=True
    )
    
    completed_at = models.DateTimeField(
        null=True,
        blank=True,
        help_text='When the report generation was completed'
    )
    
    def save(self, *args, **kwargs):
        # Remove microseconds from datetime fields
        if self.generated_at:
            self.generated_at = self.generated_at.replace(microsecond=0)
        if self.completed_at:
            self.completed_at = self.completed_at.replace(microsecond=0)
            
        # Set generator name if not set
        if not self.generator and self.user:
            self.generator = self.user.get_full_name() or self.user.username
            
        # Update completed_at when status changes to success or failed
        if self.pk:
            old_instance = ReportGeneration.objects.get(pk=self.pk)
            if (old_instance.status != self.status and 
                self.status in ['success', 'failed'] and 
                not self.completed_at):
                self.completed_at = timezone.now().replace(microsecond=0)
        super().save(*args, **kwargs)
    
    subscriber_name = models.CharField(
        max_length=255,
        null=True,
        blank=True,
        help_text='Name of the subscriber for single reports',
        db_index=True
    )
    
    from_date = models.DateField(
        null=True,
        blank=True,
        help_text='Start date of the report period',
        db_index=True
    )
    
    to_date = models.DateField(
        null=True,
        blank=True,
        help_text='End date of the report period',
        db_index=True
    )
    
    error_message = models.TextField(
        null=True,
        blank=True,
        help_text='Error message if the report generation failed'
    )

    class Meta:
        ordering = ['-generated_at']
        verbose_name = 'Report Generation'
        verbose_name_plural = 'Report Generations'
        indexes = [
            models.Index(fields=['user', 'report_type']),
            models.Index(fields=['generated_at', 'status']),
        ]

    def __str__(self):
        username = self.user.get_full_name() or self.user.username
        return f"{username} - {self.get_report_type_display()} - {self.generated_at.strftime('%Y-%m-%d %H:%M')}"

    def save(self, *args, **kwargs):
        # Set generator name if not set
        if not self.generator and self.user:
            self.generator = self.user.get_full_name() or self.user.username
            
        # Update completed_at when status changes to success or failed
        if self.pk:
            old_instance = ReportGeneration.objects.get(pk=self.pk)
            if (old_instance.status != self.status and 
                self.status in ['success', 'failed'] and 
                not self.completed_at):
                self.completed_at = timezone.now()
        super().save(*args, **kwargs)
    
    @property
    def duration(self):
        """Calculate the duration of report generation in seconds."""
        if self.completed_at and self.generated_at:
            return (self.completed_at - self.generated_at).total_seconds()
        return None