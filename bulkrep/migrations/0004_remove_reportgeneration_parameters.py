# Generated by Django 5.0.14 on 2025-05-23 09:39

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('bulkrep', '0003_reportgeneration_completed_at_and_more'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='reportgeneration',
            name='parameters',
        ),
    ]
