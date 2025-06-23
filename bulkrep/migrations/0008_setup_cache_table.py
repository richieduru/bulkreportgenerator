# Generated migration for setting up Django cache table

from django.db import migrations
from django.core.management import call_command


def create_cache_table(apps, schema_editor):
    """Create the cache table for Django's database cache backend"""
    # This will create the cache table using Django's createcachetable command
    # We'll use raw SQL since we can't call management commands in migrations directly
    db_alias = schema_editor.connection.alias
    
    # SQL to create cache table for SQL Server
    sql = """
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='dashboard_cache_table' AND xtype='U')
    CREATE TABLE dashboard_cache_table (
        cache_key NVARCHAR(255) NOT NULL PRIMARY KEY,
        value NTEXT NOT NULL,
        expires DATETIME2 NOT NULL
    )
    """
    
    with schema_editor.connection.cursor() as cursor:
        cursor.execute(sql)


def drop_cache_table(apps, schema_editor):
    """Drop the cache table"""
    sql = "DROP TABLE IF EXISTS dashboard_cache_table"
    
    with schema_editor.connection.cursor() as cursor:
        cursor.execute(sql)


class Migration(migrations.Migration):

    dependencies = [
        ('bulkrep', '0007_add_usagereport_indexes'),
    ]

    operations = [
        migrations.RunPython(
            create_cache_table,
            reverse_code=drop_cache_table,
            atomic=False,
        ),
    ]