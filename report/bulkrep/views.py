from django.shortcuts import render
from django.contrib import messages
from django.http import HttpResponse, FileResponse
from django.db.models import Q, Count, Case, When, IntegerField, Sum
from django.conf import settings
from django.utils import timezone
# Remove unused import
# from django.db import connection
from .models import Usagereport
from datetime import date, timedelta, datetime
import calendar
import io
import os
import re
import zipfile
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image
import os.path
import uuid
from django.urls import reverse
from django.contrib.auth.decorators import login_required
from .models import ReportGeneration

def write_to_cell(ws, row, col, value):
    """
    Safely writes a value to a cell, handling merged cells and preserving formatting.
    
    When using the direct cell assignment approach like in VBA, we need to 
    handle merged cells specially as they are read-only in openpyxl.
    """
    coordinate = ws.cell(row=row, column=col).coordinate
    is_merged = False
    merged_range_to_restore = None
    
    # Store the original cell style before unmerging
    original_cell = ws.cell(row=row, column=col)
    original_style = original_cell._style
    original_number_format = original_cell.number_format
    
    # Special handling for product name header (row 32)
    if row == 32 and col == 4:  # D32
        # Try additional columns for row 32 as it might be a merged cell
        for try_col in range(1, 10):  # Try columns A through I
            try:
                ws.cell(row=row, column=try_col).value = value
            except:
                pass
    
    # Check if the cell is in a merged range
    for merged_range in list(ws.merged_cells.ranges):
        if coordinate in merged_range:
            is_merged = True
            merged_range_to_restore = str(merged_range)
            ws.unmerge_cells(merged_range_to_restore)
            break
    
    # Now we can safely set the value
    target_cell = ws.cell(row=row, column=col)
    target_cell.value = value
    
    # Restore the original style
    target_cell._style = original_style
    if original_number_format != 'General':
        target_cell.number_format = original_number_format
    
    # Restore the merge if needed
    if is_merged and merged_range_to_restore:
        ws.merge_cells(merged_range_to_restore)

# Function to safely populate direct assignments by checking for merged cells first
def safe_cell_assignment(ws, row, col, value):
    """Helper function to safely assign values to cells, handling merged cells."""
    write_to_cell(ws, row, col, value)

# Remove execute_raw_sql function as we're using ORM now
# def execute_raw_sql(query, params=None):
#     """
#     Execute a raw SQL query and return the results as a list of dictionaries
#     """
#     with connection.cursor() as cursor:
#         cursor.execute(query, params or ())
#         columns = [col[0] for col in cursor.description]
#         return [dict(zip(columns, row)) for row in cursor.fetchall()]

@login_required
def home(request):
    """Home page view with subscriber list and date range."""
    today = date.today()
    first_day_of_month = today.replace(day=1)
    
    # Calculate first day of next month
    if today.month == 12:
        first_day_next_month = date(today.year + 1, 1, 1)
    else:
        first_day_next_month = date(today.year, today.month + 1, 1)
    
    # Use Django ORM to fetch distinct subscriber names
    start_date = first_day_of_month
    end_date = first_day_next_month
    
    subscribers = Usagereport.objects.filter(
        DetailsViewedDate__gte=start_date,
        DetailsViewedDate__lt=end_date
    ).values_list('SubscriberName', flat=True).distinct().order_by('SubscriberName')

    # Format dates as strings for the template
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')

    context = {
        'subscribers': subscribers,
        'start_date': start_date_str,
        'end_date': end_date_str,
    }
    return render(request, 'bulkrep/home.html', context)

def clean_filename(filename):
    """Clean subscriber name for a valid filename, similar to VBA script."""
    # Replace invalid characters with hyphens
    invalid_chars = r'[\/\\\:\*\?"<>\|]'
    return re.sub(invalid_chars, '-', filename)

@login_required
def single_report(request):
    """View for generating a single report."""
    # Get the current date range (first day of current month to first day of next month)
    today = date.today()
    first_day_of_month = today.replace(day=1)
    
    # Calculate first day of next month
    if today.month == 12:
        first_day_next_month = date(today.year + 1, 1, 1)
    else:
        first_day_next_month = date(today.year, today.month + 1, 1)
    
    # Initialize report generation tracking
    report_gen = None
    
    # Use Django ORM to fetch ALL distinct subscriber names without date filtering
    subscribers = Usagereport.objects.values_list('SubscriberName', flat=True).distinct().order_by('SubscriberName')
    
    start_date_str = first_day_of_month.strftime('%Y-%m-%d')
    end_date_str = first_day_next_month.strftime('%Y-%m-%d')
    
    # Initial context with date range and subscribers
    context = {
        'subscribers': subscribers,
        'start_date': start_date_str,
        'end_date': end_date_str,
    }
    
    if request.method == 'POST':
        subscriber_id = request.POST.get('subscriber_id')
        start_date_str = request.POST.get('start_date')
        end_date_str = request.POST.get('end_date')
        include_bills = request.POST.get('include_bills') == 'on'
        include_products = request.POST.get('include_products') == 'on'
        
        # Get subscriber name for tracking
        subscriber_name = next((sub for sub in subscribers if sub == subscriber_id), None)
        
        # Create report generation record at the start
        if subscriber_name:
            report_gen = ReportGeneration.objects.create(
                user=request.user,
                generator=request.user.username,  # Add the generator field
                report_type='single',
                status='in_progress',
                subscriber_name=subscriber_name,
                from_date=start_date_str if start_date_str else None,
                to_date=end_date_str if end_date_str else None
            )
            print(f"Started tracking single report generation for {subscriber_name} by {request.user.username}")
        
        if not subscriber_id:
            messages.error(request, "Please select a subscriber.")
            if report_gen:
                report_gen.status = 'failed'
                report_gen.error_message = 'No subscriber selected'
                report_gen.save()
            return render(request, 'bulkrep/single_report.html', context)
            
        # Continue with report generation...
    
    if request.method == 'POST':
        subscriber_id = request.POST.get('subscriber_id')
        start_date_str = request.POST.get('start_date')
        end_date_str = request.POST.get('end_date')
        include_bills = request.POST.get('include_bills') == 'on'
        include_products = request.POST.get('include_products') == 'on'

        if not subscriber_id:
            messages.error(request, "Please select a subscriber.")
            return render(request, 'bulkrep/single_report.html', context)

        # Convert date strings to date objects - this is critical for display formatting later
        try:
            # Parse the input date strings
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
            
            # Format the dates for display in the report (DD/MM/YYYY)
            start_date_display = start_date.strftime('%d/%m/%Y')
            end_date_display = end_date.strftime('%d/%m/%Y')
            
            # Log the date processing for debugging
            print(f"Date processing: {start_date_str} -> {start_date} -> {start_date_display}")
            # print(f"Date processing: {end_date_str} -> {end_date} -> {end_date_display}")
        except (ValueError, TypeError) as e:
            messages.error(request, f"Invalid date format: {str(e)}")
            return render(request, 'bulkrep/single_report.html', context)

        # Fetch data using Django ORM
        if include_bills:
            # Query for summary bills using Django ORM
            queryset = Usagereport.objects.filter(
                DetailsViewedDate__gte=start_date,
                DetailsViewedDate__lte=end_date,
                SubscriberName=subscriber_id
            )
            
            # Initialize summary dictionary with all possible keys
            summary_bills = {
                'consumer_snap_check': 0,
                'consumer_basic_trace': 0,
                'consumer_basic_credit': 0,
                'consumer_detailed_credit': 0,
                'xscore_consumer_credit': 0,
                'commercial_basic_trace': 0,
                'commercial_detailed_credit': 0,
                'enquiry_report': 0,
                'consumer_dud_cheque': 0,
                'commercial_dud_cheque': 0,
                'director_basic_report': 0,
                'director_detailed_report': 0
            }
            
            # Count each product type using Django ORM's Q objects for case-insensitive contains
            summary_bills['consumer_snap_check'] = queryset.filter(ProductName__icontains='Snap Check').count()
            summary_bills['consumer_basic_trace'] = queryset.filter(ProductName__icontains='Basic Trace').count()
            summary_bills['consumer_basic_credit'] = queryset.filter(ProductName__icontains='Basic Credit').count()
            summary_bills['consumer_detailed_credit'] = queryset.filter(
                ProductName__icontains='Detailed Credit'
            ).exclude(ProductName__icontains='X-SCore').count()
            summary_bills['xscore_consumer_credit'] = queryset.filter(
                ProductName__icontains='X-SCore Consumer Detailed Credit'
            ).count()
            summary_bills['commercial_basic_trace'] = queryset.filter(
                ProductName__icontains='Commercial Basic Trace'
            ).count()
            summary_bills['commercial_detailed_credit'] = queryset.filter(
                ProductName__icontains='Commercial detailed Credit'
            ).count()
            summary_bills['enquiry_report'] = queryset.filter(
                ProductName__icontains='Enquiry Report'
            ).count()
            summary_bills['consumer_dud_cheque'] = queryset.filter(
                ProductName__icontains='Consumer Dud Cheque'
            ).count()
            summary_bills['commercial_dud_cheque'] = queryset.filter(
                ProductName__icontains='Commercial Dud Cheque'
            ).count()
            summary_bills['director_basic_report'] = queryset.filter(
                ProductName__icontains='Director Basic Report'
            ).count()
            summary_bills['director_detailed_report'] = queryset.filter(
                ProductName__icontains='Director Detailed Report'
            ).count()
        else:
            summary_bills = {}

        # Query for product details using Django ORM
        if include_products:
            product_data = Usagereport.objects.filter(
                DetailsViewedDate__gte=start_date,
                DetailsViewedDate__lte=end_date,
                SubscriberName=subscriber_id
            ).order_by('ProductName', 'DetailsViewedDate').values(
                'SubscriberName', 'SystemUser', 'SearchIdentity', 'SubscriberEnquiryDate',
                'SearchOutput', 'DetailsViewedDate', 'ProductInputed', 'ProductName'
            )
            
            # Group by ProductName
            product_sections = {}
            for record in product_data:
                product_name = record['ProductName']
                if product_name not in product_sections:
                    product_sections[product_name] = []
                product_sections[product_name].append(record)
            
            product_data = list(product_data)  # Convert to list for compatibility
        else:
            product_sections = {}
            product_data = []

        if not product_data and include_products:
            messages.warning(request, f"No data found for subscriber {subscriber_id} between {start_date_display} and {end_date_display}.")
            return render(request, 'bulkrep/single_report.html', context)

        # Generate Excel report
        try:
            template_path = os.path.join(settings.MEDIA_ROOT, 'Templateuse.xlsx')
            buffer = io.BytesIO()
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            
            # Look for "Productname" cell to identify where to put the dynamic product name
            product_name_cell = None
            for row in range(30, 35):  # Search rows 30-34
                for col in range(1, 10):  # Search columns A-I
                    cell_value = ws.cell(row=row, column=col).value
                    if cell_value and "product" in str(cell_value).lower():
                        product_name_cell = (row, col)
                        break
                if product_name_cell:
                    break
            
            # Write subscriber info directly to cells using the safe method
            # For row 2, which is merged from H to P, we need to set the value to the first cell in the merged range
            # First, find the merged range that contains row 2
            row2_merged_range = None
            for merged_range in list(ws.merged_cells.ranges):
                if merged_range.min_row == 2 and merged_range.max_row == 2:
                    row2_merged_range = merged_range
                    break
            
            # If we found a merged range for row 2, unmerge it, set the value with the required format, and remerge it
            if row2_merged_range:
                merged_range_str = str(row2_merged_range)
                ws.unmerge_cells(merged_range_str)
                # Set the value to the first cell in the merged range with the required format
                original_cell = ws.cell(row=2, column=row2_merged_range.min_col)
                new_content = f"FirstCentral NIGERIA - BILLING DETAILS - {subscriber_id}"
                original_cell.value = new_content
                
                # Remerge the cells
                ws.merge_cells(merged_range_str)
            else:
                # Fallback to the original method if no merged range is found
                safe_cell_assignment(ws, 2, 5, subscriber_id)  # E2
            
            safe_cell_assignment(ws, 5, 4, f"BILLING DETAILS - {subscriber_id}")  # D5
            
            # For row 6, which is merged from B to Q, we need to set the value to the first cell in the merged range
            # First, find the merged range that contains row 6
            row6_merged_range = None
            for merged_range in list(ws.merged_cells.ranges):
                if merged_range.min_row == 6 and merged_range.max_row == 6:
                    row6_merged_range = merged_range
                    break
            
            # If we found a merged range for row 6, unmerge it, set the value, and remerge it
            date_range_text = f"REPORT GENERATED FOR RECORDS BETWEEN {start_date_display} and {end_date_display}"
            if row6_merged_range:
                merged_range_str = str(row6_merged_range)
                ws.unmerge_cells(merged_range_str)
                # Set the value to the first cell in the merged range (column B = 2)
                ws.cell(row=6, column=2).value = date_range_text
                # Remerge the cells
                ws.merge_cells(merged_range_str)
            else:
                # Fallback to the original method if no merged range is found
                safe_cell_assignment(ws, 6, 4, date_range_text)  # D6
                
               
            
            if include_bills:
                # Match exact VBA cell addressing approach but safely handle merged cells
                safe_cell_assignment(ws, 12, 9, summary_bills.get('consumer_snap_check', 0) or 0)  # I12
                safe_cell_assignment(ws, 13, 9, summary_bills.get('consumer_basic_trace', 0) or 0)  # I13
                safe_cell_assignment(ws, 14, 9, summary_bills.get('consumer_basic_credit', 0) or 0)  # I14
                safe_cell_assignment(ws, 15, 9, summary_bills.get('consumer_detailed_credit', 0) or 0)  # I15
                safe_cell_assignment(ws, 16, 9, summary_bills.get('xscore_consumer_credit', 0) or 0)  # I16
                safe_cell_assignment(ws, 17, 9, summary_bills.get('commercial_basic_trace', 0) or 0)  # I17
                safe_cell_assignment(ws, 18, 9, summary_bills.get('commercial_detailed_credit', 0) or 0)  # I18
                safe_cell_assignment(ws, 20, 9, summary_bills.get('enquiry_report', 0) or 0)  # I20
                safe_cell_assignment(ws, 22, 9, summary_bills.get('consumer_dud_cheque', 0) or 0)  # I22
                safe_cell_assignment(ws, 23, 9, summary_bills.get('commercial_dud_cheque', 0) or 0)  # I23
                safe_cell_assignment(ws, 25, 9, summary_bills.get('director_basic_report', 0) or 0)  # I25
                safe_cell_assignment(ws, 26, 9, summary_bills.get('director_detailed_report', 0) or 0)  # I26
            
            if include_products:
                start_row_offset = 36  # Initial start row for product sections on Sheet1
                current_sheet = ws
                sheet2 = wb["Sheet2"] if "Sheet2" in wb.sheetnames else None
                
                # Sort product_sections by product_name
                sorted_product_names = sorted(product_sections.keys())
                
                # First, find the original product name cell in the template (rows 32-35)
                product_name_cell = None
                for row in range(32, 36):  # Rows 32-35 (inclusive)
                    for col in range(1, 16):  # Assuming columns A-O are important
                        cell_value = ws.cell(row=row, column=col).value
                        if cell_value and "product" in str(cell_value).lower():
                            product_name_cell = (row, col)
                            break
                    if product_name_cell:
                        break
                        
                # Save the template header structure (rows 32-35) for subsequent products
                header_template = []
                header_rows = (32, 35)  # Range of rows to copy for the header template
                for row in range(header_rows[0], header_rows[1] + 1):
                    row_data = []
                    for col in range(1, 16):  # Assuming columns A-O are important
                        cell = ws.cell(row=row, column=col)
                        # Store cell value and position only - we'll copy styles directly later
                        cell_info = {
                            'value': cell.value,
                            'position': (row, col)
                        }
                        # Check if cell is part of merged range
                        for m_range in ws.merged_cells.ranges:
                            if (row, col) == (m_range.min_row, m_range.min_col):
                                cell_info['merged'] = (m_range.max_row - m_range.min_row + 1, 
                                                      m_range.max_col - m_range.min_col + 1)
                                break
                        row_data.append(cell_info)
                    header_template.append(row_data)
                
                # Set data row start - will be used for first product                
                data_start_row = 36  # Initial start for data rows after template header
                lastProduct = ""
                
                # For each product, use appropriate header section
                for product_idx, product_name in enumerate(sorted_product_names):
                    product_records = product_sections[product_name]
                    
                    if product_idx == 0:
                        # For first product, use the existing template header
                        if product_name_cell:
                            row, col = product_name_cell
                            # Replace "Product Name" with actual first product name in template
                            safe_cell_assignment(ws, row, col, product_name)
                        current_sheet = ws
                        current_row_offset = data_start_row  # Start data at row 36
                        serial_number_base = data_start_row - 1
                    else:
                        # Add space between different products (extra spacing)
                        current_row_offset += 4  # Add significant spacing between products
                            
                        # Create new header for subsequent products
                        header_start_row = current_row_offset
                        
                        # Clone the header section for this product
                        for template_row_idx, template_row in enumerate(header_template):
                            target_row = header_start_row + template_row_idx
                        
                            # Ensure we're not past row limits
                            if current_sheet == ws and target_row > 1000000 and sheet2:
                                current_sheet = sheet2
                                # Reset for Sheet2
                                header_start_row = 13
                                target_row = header_start_row + template_row_idx
                        
                            # Unmerge any existing merged cells in the target area
                            for m_range in list(current_sheet.merged_cells.ranges):
                                if m_range.min_row <= target_row <= m_range.max_row:
                                    current_sheet.unmerge_cells(
                                        start_row=m_range.min_row, 
                                        start_column=m_range.min_col,
                                        end_row=m_range.max_row, 
                                        end_column=m_range.max_col
                                    )
                        
                            # Copy each cell from the template to the target area
                            for col_idx, cell_info in enumerate(template_row):
                                target_col = col_idx + 1
                                target_cell = current_sheet.cell(row=target_row, column=target_col)
                                
                                # Copy styles directly from the original cell
                                original_row, original_col = cell_info['position']
                                original_cell = ws.cell(row=original_row, column=original_col)
                                
                                # Copy cell format using openpyxl's built-in method
                                target_cell._style = copy(original_cell._style)
                                
                                # Set value using safe_cell_assignment to handle merged cells properly
                                if template_row_idx == product_name_cell[0] - header_rows[0] and \
                                   col_idx == product_name_cell[1] - 1:
                                    # For the product name cell, use the safe assignment method
                                    safe_cell_assignment(current_sheet, target_row, target_col, product_name)
                                elif cell_info['value'] is not None:
                                    # For other cells with values, use the safe assignment method
                                    safe_cell_assignment(current_sheet, target_row, target_col, cell_info['value'])
                                
                                # Recreate merged cells
                                if 'merged' in cell_info:
                                    rows, cols = cell_info['merged']
                                    current_sheet.merge_cells(
                                        start_row=target_row, 
                                        start_column=target_col,
                                        end_row=target_row + rows - 1, 
                                        end_column=target_col + cols - 1
                                    )
                        
                        # Update data start row to be after this new header
                        current_row_offset = header_start_row + (header_rows[1] - header_rows[0] + 1)
                        serial_number_base = current_row_offset - 1
                        
                        # Add header for the data section
                        safe_cell_assignment(current_sheet, current_row_offset - 1, 4, "Unique Tracking Number")

                    # Process data records for this product
                    for record_idx, record in enumerate(product_records):
                        current_row = current_row_offset + record_idx
                        
                        # Switch to Sheet2 when reaching max row (like VBA)
                        if current_row > 1000000 and sheet2 and current_sheet != sheet2:
                            current_sheet = sheet2
                            current_row_offset = 13  # Reset row offset for Sheet2 as per VBA
                            serial_number_base = 12  # Reset serial number base for Sheet2
                            current_row = current_row_offset + record_idx  # Recalculate current_row for Sheet2
                        
                        # Copy row format from template data row
                        template_data_row = 36
                        max_col = 17  # Extend to include all columns that need formatting
                        copy_row_format(current_sheet, template_data_row, current_row, max_col)
                        
                        # Now assign values as before
                        safe_cell_assignment(current_sheet, current_row, 2, current_row - serial_number_base)  # Serial Number
                        safe_cell_assignment(current_sheet, current_row, 3, "")  # Branch ID
                        safe_cell_assignment(current_sheet, current_row, 4, "")  # Unique Tracking Number column left blank
                        safe_cell_assignment(current_sheet, current_row, 5, record['SubscriberName'])
                        safe_cell_assignment(current_sheet, current_row, 7, record['SystemUser'] if record['SystemUser'] else "")
                        safe_cell_assignment(current_sheet, current_row, 10, record['SubscriberEnquiryDate'])
                        safe_cell_assignment(current_sheet, current_row, 11, record['ProductName'])
                        safe_cell_assignment(current_sheet, current_row, 12, record['DetailsViewedDate'])
                        safe_cell_assignment(current_sheet, current_row, 15, record['SearchOutput'] if record['SearchOutput'] else "")
                        
                        # Apply merge and center formatting to this data row
                        merge_and_center_data_row(current_sheet, current_row)
                    
                    # Move offset to after this product's data for next product
                    current_row_offset += len(product_records)
                    
                    # Set lastProduct for the next iteration
                    lastProduct = product_name

            # Auto-size columns for better readability
            for sheet in wb.worksheets:
                auto_size_columns(sheet)

            # Add 'Generated by: <username>' to the first available merged O-Q row, or row 8 if none, with bold and italic formatting
            add_generated_by(ws, request.user.username, current_row_offset - 1)

            wb.save(buffer)
            buffer.seek(0)
            month_year = start_date.strftime('%B%Y')
            clean_subscriber = clean_filename(subscriber_id)
            filename = f"{clean_subscriber}_{month_year}_{uuid.uuid4().hex[:8]}.xlsx"
            single_reports_dir = os.path.join(settings.MEDIA_ROOT, 'reports', 'single')
            os.makedirs(single_reports_dir, exist_ok=True)
            file_path = os.path.join(single_reports_dir, filename)
            with open(file_path, 'wb') as f:
                f.write(buffer.read())
            download_url = settings.MEDIA_URL + f'reports/single/{filename}'
            
            # Update report generation status to success
            if 'report_gen' in locals() and report_gen:
                report_gen.status = 'success'
                report_gen.completed_at = timezone.now()
                report_gen.save()
                
            return render(request, 'bulkrep/download_ready.html', {
                'download_url': download_url
            })
            
        except Exception as e:
            error_msg = f"Error generating report: {str(e)}"
            messages.error(request, error_msg)
            if 'report_gen' in locals() and report_gen:
                report_gen.status = 'failed'
                report_gen.error_message = error_msg[:500]  # Truncate to fit in the field
                report_gen.completed_at = timezone.now()
                report_gen.save()
            return render(request, 'bulkrep/single_report.html', context)

    return render(request, 'bulkrep/single_report.html', context)

# Helper function to auto-size columns for better readability
def auto_size_columns(worksheet, min_col=1, max_col=17):
    """
    Auto-size columns in the worksheet to fit their contents with reasonable widths.
    Applies modest adjustments to prevent columns from being too wide.
    Handles merged cells properly, especially for rows 2 and 6.
    """
    # Define minimum and maximum widths for specific columns
    min_widths = {
        # Set minimum widths to prevent columns from being too narrow
        5: 10,  # SubscriberName (E)
        6: 10,  # SubscriberName (F)
        7: 9,  # SystemUser (G)
        8: 10,  # SystemUser (H)
        9: 9,  # SystemUser (I)
        10: 15, # SubscriberEnquiryDate (J)
        11: 20, # ProductName (K)
        12: 5, # DetailsViewedDate (L)
        13: 10, # DetailsViewedDate (M)
        14: 5, # DetailsViewedDate (N)
        # Default minimum width for other columns
        'default': 10
    }


    
    # Define maximum widths for specific columns that tend to get too wide
    max_widths = {
        # DetailsViewedDate (columns L-N, 12-14)
        12: 5, 13: 10, 14: 5,
        # SearchOutput (columns O-Q, 15-17) - increased to allow more content to be visible
        15: 40, 16: 40, 17: 40,
        # Default max width for other columns
        'default': 30
    }
    
    # Store merged ranges for special handling
    merged_ranges = list(worksheet.merged_cells.ranges)
    
    # Special handling for rows 2 and 6 which have specific merged ranges
    row2_merged_range = None
    row6_merged_range = None
    
    for merged_range in merged_ranges:
        if merged_range.min_row == 2 and merged_range.max_row == 2:
            row2_merged_range = merged_range
        elif merged_range.min_row == 6 and merged_range.max_row == 6:
            row6_merged_range = merged_range
    
    for col_idx in range(min_col, max_col + 1):
        # Get the maximum content width in the column
        max_length = 0
        column = worksheet.column_dimensions[get_column_letter(col_idx)]
        
        # Skip columns that are part of the merged ranges for rows 2 and 6
        # This prevents these merged cells from affecting column widths
        skip_column = False
        if row2_merged_range and col_idx >= row2_merged_range.min_col and col_idx <= row2_merged_range.max_col:
            # For row 2 merged range, only consider the first column in the range
            if col_idx > row2_merged_range.min_col:
                skip_column = True
        
        if row6_merged_range and col_idx >= row6_merged_range.min_col and col_idx <= row6_merged_range.max_col:
            # For row 6 merged range, only consider the first column in the range
            if col_idx > row6_merged_range.min_col:
                skip_column = True
        
        if not skip_column:
            # Check all cells in this column
            for row_idx, row in enumerate(worksheet.rows, 1):
                # Skip rows 2 and 6 for all columns except the first column in their merged ranges
                if (row_idx == 2 and row2_merged_range and col_idx != row2_merged_range.min_col) or \
                   (row_idx == 6 and row6_merged_range and col_idx != row6_merged_range.min_col):
                    continue
                
                if len(row) >= col_idx:
                    cell = row[col_idx-1]  # 0-based index
                    if cell.value:
                        # Calculate the approximate width based on content
                        try:
                            cell_length = len(str(cell.value))
                            # Adjust for merged cells
                            is_in_merge = False
                            for merged_range in merged_ranges:
                                if cell.coordinate in merged_range:
                                    # Special handling for different types of merged cells
                                    merge_width = merged_range.max_col - merged_range.min_col + 1
                                    
                                    # For rows 2 and 6, handle specially
                                    if (row_idx == 2 and merged_range == row2_merged_range) or \
                                       (row_idx == 6 and merged_range == row6_merged_range):
                                        # For the first column in the merge, allocate more space
                                        if col_idx == merged_range.min_col:
                                            # Allocate more space to first column but not all
                                            cell_length = cell_length * 0.4  # 40% to first column
                                        else:
                                            # Distribute remaining 60% evenly among other columns
                                            cell_length = (cell_length * 0.6) / (merge_width - 1)
                                    
                                    # Special handling for SearchOutput columns (O-Q, 15-17)
                                    elif 15 <= merged_range.min_col <= 17 and 15 <= merged_range.max_col <= 17:
                                        # For SearchOutput with wrap_text, distribute more evenly
                                        # Give more width to first column for better readability
                                        if col_idx == merged_range.min_col:
                                            cell_length = cell_length * 0.5  # 50% to first column
                                        else:
                                            # Distribute remaining 50% evenly
                                            cell_length = (cell_length * 0.5) / (merge_width - 1)
                                    else:
                                        # For other merged cells, divide by the number of columns
                                        # but ensure a minimum reasonable length
                                        cell_length = max(10, cell_length / merge_width)
                                    
                                    is_in_merge = True
                                    break
                            
                            # Update max_length regardless of merge status
                            # This ensures all cells contribute to column width calculation
                            max_length = max(max_length, cell_length)
                        except:
                            pass
        
        # Set column width with appropriate padding and respect minimum/maximum widths
        if max_length > 0:
            # Add padding based on content type - more padding for text columns
            if col_idx >= 15 and col_idx <= 17:  # SearchOutput columns
                # For SearchOutput columns with wrap_text, use more generous padding
                calculated_width = max_length + 3
            else:
                # For other columns, use standard padding
                calculated_width = max_length + 2
            
            # Apply column-specific minimum and maximum width constraints
            column_min_width = min_widths.get(col_idx, min_widths['default'])
            column_max_width = max_widths.get(col_idx, max_widths['default'])
            
            # Set the width to be between minimum and maximum constraints
            column.width = max(column_min_width, min(calculated_width, column_max_width))
        else:
            # If no content, use the minimum width
            column.width = min_widths.get(col_idx, min_widths['default'])

def copy_row_format(ws, template_row_idx, target_row_idx, max_col=15):
    """
    Copy cell styles (and optionally values) from a template row to a target row.
    Also handles merged cells and maintains their formatting.
    """
    # First, check and unmerge any existing merged cells in the target row
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row <= target_row_idx <= merged_range.max_row:
            ws.unmerge_cells(str(merged_range))
    
    # Copy format and check for merged cells in template
    template_merged_ranges = []
    for col in range(1, max_col + 1):
        template_cell = ws.cell(row=template_row_idx, column=col)
        target_cell = ws.cell(row=target_row_idx, column=col)
        
        # Copy style and number format
        if hasattr(template_cell, '_style'):
            target_cell._style = copy(template_cell._style)
        if hasattr(template_cell, 'number_format'):
            target_cell.number_format = template_cell.number_format
        if hasattr(template_cell, 'alignment') and template_cell.alignment:
            # Create a new alignment object instead of copying the StyleProxy
            h_align = template_cell.alignment.horizontal if template_cell.alignment.horizontal else 'general'
            v_align = template_cell.alignment.vertical if template_cell.alignment.vertical else 'bottom'
            target_cell.alignment = Alignment(horizontal=h_align, vertical=v_align)
        
        # Check if this cell is part of a merged range in the template
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == template_row_idx and merged_range.min_col <= col <= merged_range.max_col:
                template_merged_ranges.append((
                    merged_range.min_col,
                    merged_range.max_col,
                    merged_range.max_col - merged_range.min_col + 1
                ))
                break
    
    # Recreate merged ranges in the target row
    for min_col, max_col, span in set(template_merged_ranges):
        # Set alignment for all cells in the range before merging
        template_main_cell = ws.cell(row=template_row_idx, column=min_col)
        
        # Get alignment values from template cell
        h_align = 'center'  # Default to center
        v_align = 'center'  # Default to center
        
        if hasattr(template_main_cell, 'alignment') and template_main_cell.alignment:
            if template_main_cell.alignment.horizontal:
                h_align = template_main_cell.alignment.horizontal
            if template_main_cell.alignment.vertical:
                v_align = template_main_cell.alignment.vertical
        
        # Apply alignment to all cells in the range
        for col in range(min_col, max_col + 1):
            target_cell = ws.cell(row=target_row_idx, column=col)
            target_cell.alignment = Alignment(horizontal=h_align, vertical=v_align)
        
        # Now merge the cells
        ws.merge_cells(
            start_row=target_row_idx,
            start_column=min_col,
            end_row=target_row_idx,
            end_column=max_col
        )
        
        # Ensure the merged cell has proper alignment
        merged_cell = ws.cell(row=target_row_idx, column=min_col)
        merged_cell.alignment = Alignment(horizontal=h_align, vertical=v_align)

def copy_merged_and_center(ws, template_ws, template_row_start, template_row_end, target_row_start):
    """
    Copy merged cell structure and center alignment from template header rows to target rows.
    """
    # 1. Copy merged cells
    for m_range in template_ws.merged_cells.ranges:
        if template_row_start <= m_range.min_row <= template_row_end:
            row_offset = target_row_start - template_row_start
            new_min_row = m_range.min_row + row_offset
            new_max_row = m_range.max_row + row_offset
            ws.merge_cells(start_row=new_min_row, start_column=m_range.min_col,
                           end_row=new_max_row, end_column=m_range.max_col)
    # 2. Copy center alignment for header cells
    for row in range(template_row_start, template_row_end + 1):
        for col in range(1, ws.max_column + 1):
            template_cell = template_ws.cell(row=row, column=col)
            if hasattr(template_cell, 'alignment') and template_cell.alignment and template_cell.alignment.horizontal == 'center':
                target_cell = ws.cell(row=target_row_start + (row - template_row_start), column=col)
                # Set center alignment
                from openpyxl.styles import Alignment
                target_cell.alignment = Alignment(horizontal='center', vertical=template_cell.alignment.vertical)

# Helper function to merge and center header columns
def merge_and_center_header_columns(sheet, start_row, end_row):
    for row in range(start_row, end_row + 1):
        # Set row height to 23.5 for better readability of wrapped text in headers
        sheet.row_dimensions[row].height = 23.5
        
        # Unmerge any existing merged cells in this row
        for merged_range in list(sheet.merged_cells.ranges):
            if merged_range.min_row <= row <= merged_range.max_row:
                if ((5 <= merged_range.min_col <= 6) or  # SubscriberName (E-F)
                    (7 <= merged_range.min_col <= 9) or  # SystemUser (G-I)
                    (12 <= merged_range.min_col <= 14) or  # DetailsViewedDate (L-N)
                    (15 <= merged_range.min_col <= 17)):  # SearchOutput (O-Q)
                    sheet.unmerge_cells(str(merged_range))
        
        # Merge SubscriberName (E-F)
        sheet.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
        sheet.cell(row=row, column=5).alignment = Alignment(horizontal='center', vertical='center',wrap_text=True)
        
        # Merge SystemUser (G-I)
        sheet.merge_cells(start_row=row, start_column=7, end_row=row, end_column=9)
        sheet.cell(row=row, column=7).alignment = Alignment(horizontal='center', vertical='center')
        
        # Merge DetailsViewedDate (L-N)
        sheet.merge_cells(start_row=row, start_column=12, end_row=row, end_column=14)
        sheet.cell(row=row, column=12).alignment = Alignment(horizontal='center', vertical='center')
        
        # Merge SearchOutput (O-Q) with wrap text
        sheet.merge_cells(start_row=row, start_column=15, end_row=row, end_column=17)
        sheet.cell(row=row, column=15).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Helper function to merge and center data row columns
def merge_and_center_data_row(sheet, row):
    """Merge and center columns for a specific row.
    Handles SubscriberName (E-F), SystemUser (G-I), DetailsViewedDate (L-N), and SearchOutput (O-Q).
    Also sets row height to 23.5 for better readability of wrapped text.
    """
    # Set row height to 23.5 for better readability of wrapped text
    sheet.row_dimensions[row].height = 23.5
    
    # Check for existing merged cells in the target row and unmerge them
    for merged_range in list(sheet.merged_cells.ranges):
        if merged_range.min_row <= row <= merged_range.max_row:
            # Columns for DetailsViewedDate (L-N), SearchOutput (O-Q), SubscriberName (E-F), SystemUser (G-I)
            if (12 <= merged_range.min_col <= 14) or \
                (15 <= merged_range.min_col <= 17) or \
                (5 <= merged_range.min_col <= 6) or \
                (7 <= merged_range.min_col <= 9):
                sheet.unmerge_cells(str(merged_range))
    
    # Set alignment for SubscriberName (E-F)
    for col in range(5, 7):
        cell = sheet.cell(row=row, column=col)
        cell.alignment = Alignment(horizontal='center', vertical='center',wrap_text=True)
    
    # Set alignment for SystemUser (G-I)
    for col in range(7, 10):
        cell = sheet.cell(row=row, column=col)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Set alignment for DetailsViewedDate (L-N)
    for col in range(12, 15):
        cell = sheet.cell(row=row, column=col)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Set alignment for SearchOutput (O-Q) with wrap text
    for col in range(15, 18):
        cell = sheet.cell(row=row, column=col)
        # Use top vertical alignment for better readability with wrapped text
        # Center horizontally but align top vertically for multi-line text
        cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    
    # Now merge the cells
    # Merge SubscriberName (E-F)
    sheet.merge_cells(start_row=row, start_column=5, end_row=row, end_column=6)
    # Merge SystemUser (G-I)
    sheet.merge_cells(start_row=row, start_column=7, end_row=row, end_column=9)
    # Merge DetailsViewedDate (L-N)
    sheet.merge_cells(start_row=row, start_column=12, end_row=row, end_column=14)
    # Merge SearchOutput (O-Q)
    sheet.merge_cells(start_row=row, start_column=15, end_row=row, end_column=17)
    
    # Ensure the merged cells have proper alignment
    # SubscriberName
    cell_E = sheet.cell(row=row, column=5)
    cell_E.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # SystemUser
    cell_G = sheet.cell(row=row, column=7)
    cell_G.alignment = Alignment(horizontal='center', vertical='center')
    
    # DetailsViewedDate
    cell_L = sheet.cell(row=row, column=12)
    cell_L.alignment = Alignment(horizontal='center', vertical='center')
    
    # SearchOutput with wrap text
    cell_O = sheet.cell(row=row, column=15)
    # Use top vertical alignment for better readability with wrapped text
    cell_O.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

def add_generated_by(ws, username, last_data_row=None):
    """
    Add 'Report Generated by: <username>' two rows below the last data row (or at row 10 if no data), merging O-Q, Trebuchet MS, bold, italic, centered.
    """
    from openpyxl.styles import Font
    from openpyxl.styles import Alignment as OpenpyxlAlignment
    if not last_data_row or last_data_row < 1:
        target_row = 10
    else:
        target_row = last_data_row + 5
    ws.merge_cells(start_row=target_row, start_column=15, end_row=target_row, end_column=17)
    cell = ws.cell(row=target_row, column=15)
    cell.value = f"Report Generated by: {username}"
    cell.font = Font(name='Trebuchet MS', bold=True, italic=True, color='FF7F7F7F')
    cell.alignment = OpenpyxlAlignment(horizontal='center', vertical='center')

@login_required
def bulk_report(request):
    """View for generating bulk reports."""
    # Get the current date range (first day of current month to first day of next month)
    today = date.today()
    first_day_of_month = today.replace(day=1)
    
    # Calculate first day of next month
    if today.month == 12:
        first_day_next_month = date(today.year + 1, 1, 1)
    else:
        first_day_next_month = date(today.year, today.month + 1, 1)
    
    # Initialize report generation tracking
    report_gen = None
    
    # Initial context with date range
    context = {
        'start_date': first_day_of_month.strftime('%Y-%m-%d'),
        'end_date': first_day_next_month.strftime('%Y-%m-%d'),
    }
    
    if request.method == 'POST':
        generation_method = request.POST.get('generation_method')
        subscriber_file = request.FILES.get('subscriber_file')
        start_date_str = request.POST.get('start_date')
        end_date_str = request.POST.get('end_date')
        include_bills = request.POST.get('include_bills') == 'on'
        include_products = request.POST.get('include_products') == 'on'

        # Create report generation record at the start
        report_gen = ReportGeneration.objects.create(
            user=request.user,
            generator=request.user.username,  # Add the generator field
            report_type='bulk',
            status='in_progress',
            subscriber_name='Multiple Subscribers',
            from_date=start_date_str if start_date_str else None,
            to_date=end_date_str if end_date_str else None
        )
        print(f"Started tracking bulk report generation by {request.user.username}")
        
        # Continue with report generation...
    
    # Initial context with date range
    context = {
        'start_date': first_day_of_month.strftime('%Y-%m-%d'),
        'end_date': first_day_next_month.strftime('%Y-%m-%d'),
    }
    
    if request.method == 'POST':
        generation_method = request.POST.get('generation_method')
        subscriber_file = request.FILES.get('subscriber_file')
        start_date_str = request.POST.get('start_date')
        end_date_str = request.POST.get('end_date')
        include_bills = request.POST.get('include_bills') == 'on'
        include_products = request.POST.get('include_products') == 'on'

        # Convert date strings to date objects - this is critical for display formatting later
        try:
            # Parse the input date strings
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
            
            # Format the dates for display in the report (DD/MM/YYYY)
            start_date_display = start_date.strftime('%d/%m/%Y')
            end_date_display = end_date.strftime('%d/%m/%Y')
            
            # Log the date processing for debugging
            print(f"Bulk report date processing: {start_date_str} -> {start_date} -> {start_date_display}")
            print(f"Bulk report date processing: {end_date_str} -> {end_date} -> {end_date_display}")
        except (ValueError, TypeError) as e:
            messages.error(request, f"Invalid date format: {str(e)}")
            return render(request, 'bulkrep/bulk_report.html', context)

        # Determine the list of subscribers
        subscribers_list = []
        if generation_method == 'file':
            if not subscriber_file:
                messages.error(request, "Subscriber list file is required for file upload method.")
                return render(request, 'bulkrep/bulk_report.html', context)
            
            # Process the uploaded file
            try:
                # For CSV files
                if subscriber_file.name.endswith('.csv'):
                    import csv
                    decoded_file = subscriber_file.read().decode('utf-8').splitlines()
                    reader = csv.reader(decoded_file)
                    for row in reader:
                        if row:  # Check if row is not empty
                            subscribers_list.append(row[0].strip())  # Assume first column contains subscriber names
                
                # For Excel files
                elif subscriber_file.name.endswith(('.xlsx', '.xls')):
                    wb = openpyxl.load_workbook(subscriber_file)
                    ws = wb.active
                    for row in ws.iter_rows(values_only=True):
                        if row and row[0]:  # Check if the first cell in the row has a value
                            subscribers_list.append(str(row[0]).strip())
                else:
                    messages.error(request, "Unsupported file format. Please upload a CSV or Excel file.")
                    return render(request, 'bulkrep/bulk_report.html', context)
                
                # Remove duplicates and empty strings
                subscribers_list = list(filter(None, set(subscribers_list)))
                
            except Exception as e:
                messages.error(request, f"Error processing file: {str(e)}")
                return render(request, 'bulkrep/bulk_report.html', context)

        elif generation_method == 'all':
            # Use Django ORM to get all distinct subscribers
            subscribers_list = list(Usagereport.objects.filter(
                DetailsViewedDate__gte=start_date,
                DetailsViewedDate__lte=end_date
            ).values_list('SubscriberName', flat=True).distinct())

        if not subscribers_list:
            messages.warning(request, f"No subscribers found for the selected criteria between {start_date_display} and {end_date_display}.")
            return render(request, 'bulkrep/bulk_report.html', context)

        # Generate reports for all subscribers using a zip file
        try:
            template_path = os.path.join(settings.MEDIA_ROOT, 'Templateuse.xlsx')
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                processed_subscribers = []
                
                for subscriber_id in subscribers_list:
                    try:
                        # Fetch data using raw SQL - summary bills
                        
                        # Fetch data using Django ORM
                        if include_bills:
                            # Query for summary bills using Django ORM
                            queryset = Usagereport.objects.filter(
                                DetailsViewedDate__gte=start_date,
                                DetailsViewedDate__lte=end_date,
                                SubscriberName=subscriber_id
                            )
                            
                            # Initialize summary dictionary with all possible keys
                            summary_bills = {
                                'consumer_snap_check': 0,
                                'consumer_basic_trace': 0,
                                'consumer_basic_credit': 0,
                                'consumer_detailed_credit': 0,
                                'xscore_consumer_credit': 0,
                                'commercial_basic_trace': 0,
                                'commercial_detailed_credit': 0,
                                'enquiry_report': 0,
                                'consumer_dud_cheque': 0,
                                'commercial_dud_cheque': 0,
                                'director_basic_report': 0,
                                'director_detailed_report': 0
                            }
                            
                            # Count each product type using Django ORM's Q objects for case-insensitive contains
                            summary_bills['consumer_snap_check'] = queryset.filter(ProductName__icontains='Snap Check').count()
                            summary_bills['consumer_basic_trace'] = queryset.filter(ProductName__icontains='Basic Trace').count()
                            summary_bills['consumer_basic_credit'] = queryset.filter(ProductName__icontains='Basic Credit').count()
                            summary_bills['consumer_detailed_credit'] = queryset.filter(
                                ProductName__icontains='Detailed Credit'
                            ).exclude(ProductName__icontains='X-SCore').count()
                            summary_bills['xscore_consumer_credit'] = queryset.filter(
                                ProductName__icontains='X-SCore Consumer Detailed Credit'
                            ).count()
                            summary_bills['commercial_basic_trace'] = queryset.filter(
                                ProductName__icontains='Commercial Basic Trace'
                            ).count()
                            summary_bills['commercial_detailed_credit'] = queryset.filter(
                                ProductName__icontains='Commercial detailed Credit'
                            ).count()
                            summary_bills['enquiry_report'] = queryset.filter(
                                ProductName__icontains='Enquiry Report'
                            ).count()
                            summary_bills['consumer_dud_cheque'] = queryset.filter(
                                ProductName__icontains='Consumer Dud Cheque'
                            ).count()
                            summary_bills['commercial_dud_cheque'] = queryset.filter(
                                ProductName__icontains='Commercial Dud Cheque'
                            ).count()
                            summary_bills['director_basic_report'] = queryset.filter(
                                ProductName__icontains='Director Basic Report'
                            ).count()
                            summary_bills['director_detailed_report'] = queryset.filter(
                                ProductName__icontains='Director Detailed Report'
                            ).count()
                        else:
                            summary_bills = {}
                        
                        # Fetch product details using Django ORM
                        if include_products:
                            product_data = Usagereport.objects.filter(
                                DetailsViewedDate__gte=start_date,
                                DetailsViewedDate__lte=end_date,
                                SubscriberName=subscriber_id
                            ).order_by('ProductName', 'DetailsViewedDate').values(
                                'SubscriberName', 'SystemUser', 'SearchIdentity', 'SubscriberEnquiryDate',
                                'SearchOutput', 'DetailsViewedDate', 'ProductInputed', 'ProductName'
                            )
                            
                            # Skip if no data
                            if not product_data.exists():
                                continue
                            
                            # Group by ProductName
                            product_sections = {}
                            for record in product_data:
                                product_name = record['ProductName']
                                if product_name not in product_sections:
                                    product_sections[product_name] = []
                                product_sections[product_name].append(record)
                        else:
                            product_sections = {}
                            product_data = []
                        
                        # Skip if no data for this subscriber
                        if not product_data and include_products:
                            continue
                        
                        # Create Excel report for this subscriber
                        excel_buffer = io.BytesIO()
                        wb = openpyxl.load_workbook(template_path)
                        ws = wb.active
                        
                        
                        # Look for "Productname" cell to identify where to put the dynamic product name
                        product_name_cell = None
                        for row in range(30, 35):  # Search rows 30-34
                            for col in range(1, 10):  # Search columns A-I
                                cell_value = ws.cell(row=row, column=col).value
                                if cell_value and "product" in str(cell_value).lower():
                                    product_name_cell = (row, col)
                                    break
                            if product_name_cell:
                                break
                        
                        row2_merged_range = None
                        for merged_range in list(ws.merged_cells.ranges):
                            if merged_range.min_row == 2 and merged_range.max_row == 2:
                                row2_merged_range = merged_range
                                break
                        
                        # If we found a merged range for row 2, unmerge it, set the value with the required format, and remerge it
                        if row2_merged_range:
                            merged_range_str = str(row2_merged_range)
                            ws.unmerge_cells(merged_range_str)
                            # Set the value to the first cell in the merged range with the required format
                            original_cell = ws.cell(row=2, column=row2_merged_range.min_col)
                            new_content = f"FirstCentral NIGERIA - BILLING DETAILS - {subscriber_id}"
                            original_cell.value = new_content
                            
                            # Remerge the cells
                            ws.merge_cells(merged_range_str)
                        else:
                            # Fallback to the original method if no merged range is found
                            safe_cell_assignment(ws, 2, 5, subscriber_id)  # E2
                        
                        safe_cell_assignment(ws, 5, 4, f"BILLING DETAILS - {subscriber_id}")
                        
                        # For row 6, which is merged from B to Q, we need to set the value to the first cell in the merged range
                        # First, find the merged range that contains row 6
                        row6_merged_range = None
                        for merged_range in list(ws.merged_cells.ranges):
                            if merged_range.min_row == 6 and merged_range.max_row == 6:
                                row6_merged_range = merged_range
                                break
                        
                        # If we found a merged range for row 6, unmerge it, set the value, and remerge it
                        date_range_text = f"REPORT GENERATED FOR RECORDS BETWEEN {start_date_display} and {end_date_display}"
                        if row6_merged_range:
                            merged_range_str = str(row6_merged_range)
                            ws.unmerge_cells(merged_range_str)
                            # Set the value to the first cell in the merged range (column B = 2)
                            ws.cell(row=6, column=2).value = date_range_text
                            # Remerge the cells
                            ws.merge_cells(merged_range_str)
                        else:
                            # Fallback to the original method if no merged range is found
                            safe_cell_assignment(ws, 6, 4, date_range_text)  # D6
                        
                        if include_bills:
                            safe_cell_assignment(ws, 12, 9, summary_bills.get('consumer_snap_check', 0) or 0)
                            safe_cell_assignment(ws, 13, 9, summary_bills.get('consumer_basic_trace', 0) or 0)
                            safe_cell_assignment(ws, 14, 9, summary_bills.get('consumer_basic_credit', 0) or 0)
                            safe_cell_assignment(ws, 15, 9, summary_bills.get('consumer_detailed_credit', 0) or 0)
                            safe_cell_assignment(ws, 16, 9, summary_bills.get('xscore_consumer_credit', 0) or 0)
                            safe_cell_assignment(ws, 17, 9, summary_bills.get('commercial_basic_trace', 0) or 0)
                            safe_cell_assignment(ws, 18, 9, summary_bills.get('commercial_detailed_credit', 0) or 0)
                            safe_cell_assignment(ws, 20, 9, summary_bills.get('enquiry_report', 0) or 0)
                            safe_cell_assignment(ws, 22, 9, summary_bills.get('consumer_dud_cheque', 0) or 0)
                            safe_cell_assignment(ws, 23, 9, summary_bills.get('commercial_dud_cheque', 0) or 0)
                            safe_cell_assignment(ws, 25, 9, summary_bills.get('director_basic_report', 0) or 0)
                            safe_cell_assignment(ws, 26, 9, summary_bills.get('director_detailed_report', 0) or 0)
                        
                        if include_products:
                            row_offset = 36
                            current_sheet = ws
                            sheet2 = wb["Sheet2"] if "Sheet2" in wb.sheetnames else None
                            serial_number_base = 35

                            sorted_product_names = sorted(product_sections.keys())
                            
                            # Set the first product name in the "Productname" header
                            if sorted_product_names and product_name_cell:
                                first_product = sorted_product_names[0]
                                row, col = product_name_cell
                                safe_cell_assignment(ws, row, col, first_product)

                            lastProduct = ""
                            for product_idx, product_name in enumerate(sorted_product_names):
                                product_records = product_sections[product_name]
                                # Add space between different products (like VBA)
                                if lastProduct != "" and lastProduct != product_name:
                                    row_offset += 2
                                # Write the Product Name header dynamically, ONE row above the details start
                                header_row = row_offset - 1
                                # Ensure header_row is valid (e.g., > 0)
                                if header_row > 0:
                                    safe_cell_assignment(current_sheet, header_row, 4, "Unique Tracking Number") # Unique Tracking Number header remains
                                # Initialize current_row_offset for this product
                                current_row_offset = row_offset
                                for record_idx, record in enumerate(product_records):
                                    current_row = current_row_offset + record_idx
                                    if current_row > 1000000 and sheet2 and current_sheet != sheet2:
                                        current_sheet = sheet2
                                        row_offset = 13 
                                        serial_number_base = 12
                                        current_row = row_offset + record_idx
                                    # Copy row format from template data row
                                    template_data_row = 36
                                    max_col = 17  # Extend to include all columns that need formatting
                                    copy_row_format(current_sheet, template_data_row, current_row, max_col)
                                    safe_cell_assignment(current_sheet, current_row, 2, current_row - serial_number_base)
                                    safe_cell_assignment(current_sheet, current_row, 3, "")
                                    safe_cell_assignment(current_sheet, current_row, 4, "")
                                    safe_cell_assignment(current_sheet, current_row, 5, record['SubscriberName'])
                                    safe_cell_assignment(current_sheet, current_row, 7, record['SystemUser'] if record['SystemUser'] else "")
                                    safe_cell_assignment(current_sheet, current_row, 10, record['SubscriberEnquiryDate'])
                                    safe_cell_assignment(current_sheet, current_row, 11, record['ProductName'])
                                    safe_cell_assignment(current_sheet, current_row, 12, record['DetailsViewedDate'])
                                    safe_cell_assignment(current_sheet, current_row, 15, record['SearchOutput'] if record['SearchOutput'] else "")
                                    # Merge and center columns L:N and O:Q
                                    merge_and_center_data_row(current_sheet, current_row)
                                row_offset += len(product_records)
                                lastProduct = product_name
                        
                        # Auto-size columns for better readability
                        for sheet in wb.worksheets:
                            auto_size_columns(sheet)
                        
                        # Add 'Generated by: <username>' to the first available merged O-Q row, or row 8 if none, with bold and italic formatting
                        add_generated_by(ws, request.user.username, current_row_offset - 1)
                        
                        # Save to buffer
                        wb.save(excel_buffer)
                        excel_buffer.seek(0)
                        
                        # Prepare filename
                        month_year = start_date.strftime('%B%Y')
                        clean_subscriber = clean_filename(subscriber_id)
                        filename = f"{clean_subscriber}_{month_year}.xlsx"
                        
                        # Add to zip file
                        zip_file.writestr(filename, excel_buffer.getvalue())
                        
                        # Add to processed subscribers list
                        processed_subscribers.append(subscriber_id)
                        
                    except Exception as e:
                        print(f"Error processing subscriber {subscriber_id}: {str(e)}")
                        error_msg = f"Skipped report for {subscriber_id} due to error: {str(e)}"
                        messages.warning(request, error_msg)
                        if 'report_gen' in locals() and report_gen:
                            report_gen.status = 'failed'
                            report_gen.error_message = error_msg[:500]  # Truncate to fit in the field
                            report_gen.completed_at = timezone.now()
                            report_gen.save()
                        continue
            
            # Return the zip file if any reports were generated
            if processed_subscribers:
                zip_buffer.seek(0)
                month_year = start_date.strftime('%B%Y')
                zip_filename = f"all_subscriber_reports_{month_year}_{uuid.uuid4().hex[:8]}.zip"
                bulk_reports_dir = os.path.join(settings.MEDIA_ROOT, 'reports', 'bulk')
                os.makedirs(bulk_reports_dir, exist_ok=True)
                zip_path = os.path.join(bulk_reports_dir, zip_filename)
                with open(zip_path, 'wb') as f:
                    f.write(zip_buffer.read())
                download_url = settings.MEDIA_URL + f'reports/bulk/{zip_filename}'
                
                # Update report generation status to success
                if 'report_gen' in locals() and report_gen:
                    report_gen.status = 'success'
                    report_gen.completed_at = timezone.now()
                    report_gen.save()
                
                success_msg = f"Successfully generated reports for {len(processed_subscribers)} out of {len(subscribers_list)} subscribers for the period {start_date_display} to {end_date_display}."
                messages.success(request, success_msg)
                return render(request, 'bulkrep/download_ready.html', {
                    'download_url': download_url
                })
            else:
                messages.warning(request, "No reports were generated. Please check the data or try different criteria.")
                return render(request, 'bulkrep/bulk_report.html', context)
            
        except Exception as e:
            messages.error(request, f"Error generating bulk reports: {str(e)}")
            return render(request, 'bulkrep/bulk_report.html', context)

    return render(request, 'bulkrep/bulk_report.html', context)

