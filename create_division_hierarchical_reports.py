from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy as copy_style
import pandas as pd
import os
from datetime import datetime

def create_division_excel_report(div_code, affiliate, div_name, summary_df, output_dir):
    """Create Excel report for a specific Division with perfect formatting based on Excel template"""
    
    try:
        # Load Excel template file (not CSV)
        template_file = 'division summary.xlsx'
        
        if not os.path.exists(template_file):
            print(f"   ❌ Template file not found: {template_file}")
            return
        
        # Load the Excel template to preserve formatting
        wb = load_workbook(template_file)
        ws = wb.active

        def get_cell_value_handling_merged(row, col):
            """Get cell value even if it's part of a merged cell"""
            cell = ws.cell(row=row, column=col)
            
            # Check if this cell is part of a merged range
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Get the top-left cell of the merged range
                    top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                    return top_left_cell.value
            
            return cell.value
        
        # Search for header row containing "Affiliate"
        header_row = None
        for row_idx in range(1, 15):
            for col_idx in range(1, min(30, ws.max_column + 1)):
                cell_value = get_cell_value_handling_merged(row_idx, col_idx)
                if cell_value and 'Affiliate' in str(cell_value):
                    header_row = row_idx
                    break
            if header_row:
                break
        
        if header_row is None:
            header_row = 3  # Default based on CSV template structure (row 3)
        
        # Find "Total" row
        total_row = None
        for row_idx in range(header_row + 1, min(header_row + 20, ws.max_row + 1)):
            cell_value = get_cell_value_handling_merged(row_idx, 1)
            if cell_value and 'Total' in str(cell_value):
                total_row = row_idx
                break
        
        if total_row is None:
            total_row = header_row + 6  # Default position
        
        data_start_row = header_row + 1
        
        # Read actual column positions from template header row
        column_mapping = {}
        for col_idx in range(1, min(30, ws.max_column + 1)):
            header_val = get_cell_value_handling_merged(header_row, col_idx)
            if header_val:
                header_str = str(header_val).strip()
                
                # Map columns based on template
                if 'Affiliate' in header_str:
                    column_mapping['Affiliate'] = col_idx
                elif 'Division' in header_str and 'Name' not in header_str:
                    column_mapping['Division'] = col_idx
                elif 'Division Name' in header_str:
                    column_mapping['Division Name'] = col_idx
                elif 'TBMs' in header_str or '# Unique TBMs' in header_str:
                    column_mapping['# Unique TBMs'] = col_idx
                elif 'HCPs' in header_str or '# Unique HCPs' in header_str:
                    column_mapping['# Unique HCPs'] = col_idx
                elif 'Requests Raised' in header_str or 'Requests raised' in header_str:
                    column_mapping['# Requests Raised\n(A+B+C)'] = col_idx
                elif 'Out of Stock' in header_str or 'Out of stock' in header_str:
                    column_mapping['Request Cancelled / Out of Stock (A)'] = col_idx
                elif 'Action pending' in header_str and 'HO' in header_str:
                    column_mapping['Action pending / In Process At HO (B)'] = col_idx
                elif 'Sent to HUB' in header_str:
                    column_mapping["Sent to HUB ('C)\n(D+E+F)"] = col_idx
                elif 'Pending for Invoicing' in header_str:
                    column_mapping['Pending for Invoicing (D)'] = col_idx
                elif 'Pending for Dispatch' in header_str:
                    column_mapping['Pending for Dispatch (E)'] = col_idx
                elif 'Requests Dispatched' in header_str and 'In Transit' not in header_str:
                    column_mapping['# Requests Dispatched (F)\n(G+H+I)'] = col_idx
                elif 'Delivered' in header_str and '(' in header_str:
                    column_mapping['Delivered (G)'] = col_idx
                elif 'Dispatched & In Transit' in header_str or 'Dispatched &amp; In Transit' in header_str:
                    column_mapping['Dispatched & In Transit (H)'] = col_idx
                elif 'RTO' in header_str and '(' in header_str and 'Hold' not in header_str:
                    column_mapping['RTO (I)'] = col_idx
                elif 'Incomplete Address' in header_str:
                    column_mapping['Incomplete Address'] = col_idx
                elif 'Doctor Non Contactable' in header_str or 'Non Contactable' in header_str:
                    column_mapping['Doctor Non Contactable'] = col_idx
                elif 'Refused to Accept' in header_str or 'refused to accept' in header_str:
                    column_mapping['Doctor Refused to Accept'] = col_idx
                elif 'Hold Delivery' in header_str:
                    column_mapping['Hold Delivery'] = col_idx

        # Clear existing data rows (between header and total)
        for r in range(data_start_row, total_row):
            for c in range(1, ws.max_column + 1):
                try:
                    cell = ws.cell(row=r, column=c)
                    cell.value = None
                except:
                    pass

        def copy_row_style(src_row_idx, dst_row_idx):
            """Copy formatting from source row to destination row"""
            for c in range(1, ws.max_column + 1):
                try:
                    src = ws.cell(row=src_row_idx, column=c)
                    dst = ws.cell(row=dst_row_idx, column=c)
                    
                    if src.font:
                        dst.font = copy_style(src.font)
                    if src.alignment:
                        dst.alignment = copy_style(src.alignment)
                    if src.border:
                        dst.border = copy_style(src.border)
                    if src.fill:
                        dst.fill = copy_style(src.fill)
                    dst.number_format = src.number_format
                except:
                    pass

        # Write data to Total row with values
        copy_row_style(total_row, total_row)
        
        # Set "Total" text in first column
        ws.cell(row=total_row, column=1, value="Total")
        
        for col_name, col_idx in column_mapping.items():
            if col_name in summary_df.columns:
                value = summary_df.iloc[0][col_name]
                
                try:
                    cell = ws.cell(row=total_row, column=col_idx)
                    cell.value = value
                    
                    if isinstance(value, (int, float)) and not pd.isna(value):
                        cell.number_format = '0'
                        cell.font = Font(bold=True, name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.font = Font(bold=True, name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                except Exception as e:
                    print(f"   Warning: Could not set value for column {col_name}: {e}")

        # Save file
        safe_div_name = str(div_name).replace(' ', '_').replace('/', '_').replace('\\', '_')
        filename = f"Division_Summary_{div_code}_{safe_div_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        wb.save(filepath)
        print(f"   ✅ Created: {filename}")
        
    except Exception as e:
        print(f"   ❌ Error creating Excel report for Division {div_code}: {e}")
        import traceback
        traceback.print_exc()
