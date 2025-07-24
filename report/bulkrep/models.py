from django.db import models
from django.contrib.auth import get_user_model
from django.utils import timezone
from decimal import Decimal

# Create your models here.

class SubscriberProductRate(models.Model):
    id = models.AutoField(primary_key=True)
    subscriber_name = models.CharField(max_length=255, db_column='SubscriberName')
    product_name = models.CharField(max_length=255, db_column='ProductName')
    rate = models.DecimalField(max_digits=10, decimal_places=2, db_column='rate')

    class Meta:
        managed = False
        db_table = 'SubscriberProductRate'
        unique_together = (('subscriber_name', 'product_name'),)


    def __str__(self):
        return f"{self.subscriber_name} - {self.product_name} - {self.rate}"


class Usagereport(models.Model):
    # Actual fields from the database
    SubscriberName = models.CharField(max_length=255, db_column='SubscriberName')
    DetailsViewedDate = models.DateField(db_column='DetailsViewedDate')
    ProductName = models.CharField(max_length=255, db_column='ProductName')
    SystemUser = models.CharField(max_length=255, db_column='SystemUser', null=True, blank=True)
    SearchIdentity = models.CharField(max_length=255, db_column='SearchIdentity', primary_key=True) 
    SubscriberEnquiryDate = models.DateField(db_column='SubscriberEnquiryDate', null=True, blank=True)
    SearchOutput = models.TextField(db_column='SearchOutput', null=True, blank=True)
    ProductInputed = models.CharField(max_length=255, db_column='ProductInputed', null=True, blank=True)

    class Meta:
        managed = False  # Tell Django not to manage this table
        db_table = 'usagereport' # Specify the existing table name
        unique_together = (('SubscriberName', 'ProductName', 'SearchIdentity'),)

        indexes = [
            models.Index(fields=['DetailsViewedDate']),
            models.Index(fields=['SubscriberName']),
        ]
    
    def __str__(self):
        return f"{self.SubscriberName} - {self.ProductName} - {self.SearchIdentity}"


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

# Add this after the imports at the top of the file
ENQUIRY_RATES = {
    'consumer_snap_check': Decimal('500.00'),
    'consumer_basic_trace': Decimal('170.00'),
    'consumer_basic_credit': Decimal('170.00'),
    'consumer_detailed_credit': Decimal('240.00'),
    'xscore_consumer_detailed_credit': Decimal('500.00'),
    'commercial_basic_trace': Decimal('275.00'),
    'commercial_detailed_credit': Decimal('500.00'),
    'enquiry_report': Decimal('50.00'),
    'consumer_dud_cheque': Decimal('0.00'),
    'commercial_dud_cheque': Decimal('0.00'),
    'director_basic_report': Decimal('0.00'),
    'director_detailed_report': Decimal('0.00'),
}
