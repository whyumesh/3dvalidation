# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy as copy_style
import warnings
import sys

# Set UTF-8 encoding for console output (fixes Windows encoding issues)
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# Suppress FutureWarning for groupby operations
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def create_division_hierarchical_reports():
    """
    Create separate Division reports showing ZBM hierarchy with perfect tallies
    Each Division gets a report showing all ZBMs under them
    """
    
    print("🔄 Starting Division Hierarchical Reports Creation...")
    
    # Read master tracker data from Excel file
    print("📖 Reading Sample Master Tracker.xlsx...")
    try:
        df = pd.read_excel('Sample Master Tracker.xlsx')
        print(f"✅ Successfully loaded {len(df)} records from Sample Master Tracker.xlsx")
    except Exception as e:
        print(f"❌ Error reading Sample Master Tracker.xlsx: {e}")
        return
    
    # Using Final Answer field computed from Request Status using business rules
    print("📊 Using Final Answer field computed from Request Status using business rules for accurate counts...")
    
    # Clean and prepare data
    print("🧹 Cleaning and preparing data...")
    
    # Ensure required columns exist
    required_columns = ['TBM Division', 'ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID',
                        'ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID',
                        'TBM HQ', 'TBM EMAIL_ID',
                        'Doctor: Customer Code', 'Assigned Request Ids', 'Request Status', 'Rto Reason']
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing required columns in Sample Master Tracker.xlsx: {missing}")
        return

    # Remove rows where key fields are null or empty
    df = df.dropna(subset=['TBM Division', 'ZBM Terr Code', 'ZBM Name', 'ABM Terr Code', 'ABM Name', 'TBM HQ'])
    df = df[df['TBM Division'].astype(str).str.strip() != '']
    df = df[df['ZBM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ABM Terr Code'].astype(str).str.strip() != '']
    df = df[df['TBM HQ'].astype(str).str.strip() != '']

    # Filter for ZBM codes that start with "ZN" (only restriction needed)
    df = df[df['ZBM Terr Code'].astype(str).str.startswith('ZN')]
    print(f"📊 After cleaning and ZBM filtering: {len(df)} records remaining")

    # Compute Final Answer per unique request id using rules from logic.xlsx
    print("🧠 Computing final status per unique Request Id using rules...")
    try:
        xls_rules = pd.ExcelFile('logic.xlsx')
        sheet2 = pd.read_excel(xls_rules, 'Sheet2')

        def normalize(text):
            return str(text).strip().casefold()

        rules = {}
        for _, row in sheet2.iterrows():
            statuses = [normalize(s) for s in row.drop('Final Answer').dropna().tolist()]
            statuses = tuple(sorted(set(statuses)))
            rules[statuses] = row['Final Answer']

        # Group statuses by request id from master data
        grouped = df.groupby('Assigned Request Ids')['Request Status'].apply(list).reset_index()

        def get_final_answer(status_list):
            key = tuple(sorted(set(normalize(s) for s in status_list)))
            return rules.get(key, '❌ No matching rule')

        grouped['Request Status'] = grouped['Request Status'].apply(lambda lst: sorted(set(lst), key=str))
        grouped['Final Answer'] = grouped['Request Status'].apply(get_final_answer)

        def has_action_pending(status_list):
            target = 'action pending / in process'
            return any(normalize(s) == target for s in status_list)
        grouped['Has D Pending'] = grouped['Request Status'].apply(has_action_pending)

        # Merge Final Answer back to main dataframe
        df = df.merge(grouped[['Assigned Request Ids', 'Final Answer', 'Has D Pending']], on='Assigned Request Ids', how='left')
    except Exception as e:
        print(f"❌ Error computing final status from logic.xlsx: {e}")
        return
    
    # Get unique Divisions
    divisions = df[['TBM Division']].drop_duplicates().sort_values('TBM Division')
    print(f"📋 Found {len(divisions)} unique Divisions")
    
    # Debug: Show all Divisions and their ZBMs
    print("\n🔍 Division-ZBM Mapping:")
    for _, div_row in divisions.iterrows():
        div_code = div_row['TBM Division']
        div_data_temp = df[df['TBM Division'] == div_code]
        zbms_temp = div_data_temp[['ZBM Terr Code', 'ZBM Name']].drop_duplicates()
        print(f"   Division {div_code}: {len(zbms_temp)} ZBMs")
        for _, zbm_row in zbms_temp.iterrows():
            print(f"      - {zbm_row['ZBM Terr Code']}: {zbm_row['ZBM Name']}")
    
    # Create output directory
    timestamp = datetime.now().strftime('%Y%m%d')
    output_dir = f"Division_Reports_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 Created output directory: {output_dir}")
    
    # Process each Division
    for _, div_row in divisions.iterrows():
        div_code = div_row['TBM Division']
        
        print(f"\n🔄 Processing Division: {div_code}")
        
        # Filter data for this Division
        div_data = df[df['TBM Division'] == div_code].copy()
        
        if len(div_data) == 0:
            print(f"⚠️ No data found for Division: {div_code}")
            continue
        
        # Get unique ZBMs under this Division
        zbms = div_data.groupby(['ZBM Terr Code', 'ZBM Name']).agg({
            'ZBM EMAIL_ID': 'first'
        }).reset_index()
        
        zbms = zbms.sort_values('ZBM Terr Code')
        print(f"   📊 Found {len(zbms)} ZBMs under this Division")
        
        # Create summary data for this Division
        summary_data = []
        
        for _, zbm_row in zbms.iterrows():
            zbm_code = zbm_row['ZBM Terr Code']
            zbm_name = zbm_row['ZBM Name']
            zbm_email = zbm_row['ZBM EMAIL_ID']
            
            # Filter data for this specific ZBM
            zbm_data = div_data[(div_data['ZBM Terr Code'] == zbm_code) & (div_data['ZBM Name'] == zbm_name)]
            
            print(f"      Processing {zbm_name} ({zbm_code}): {len(zbm_data)} records")
            
            # Calculate metrics for this ZBM
            unique_tbms = zbm_data['TBM EMAIL_ID'].nunique() if 'TBM EMAIL_ID' in zbm_data.columns else 0
            unique_hcps = zbm_data['Doctor: Customer Code'].nunique()
            unique_requests = zbm_data['Assigned Request Ids'].nunique()
            
            # HO Section (A + B) - Using Final Answer instead of Request Status
            # Count unique request IDs for each status category
            request_cancelled_out_of_stock = zbm_data[zbm_data['Final Answer'].isin(['Out of stock', 'On hold', 'Not permitted'])]['Assigned Request Ids'].nunique()
            action_pending_at_ho = zbm_data[zbm_data['Final Answer'].isin(['Request Raised', 'Action pending / In Process At HO'])]['Assigned Request Ids'].nunique()
            
            # HUB Section (D + E) - Using Final Answer instead of Request Status
            pending_for_invoicing = zbm_data[zbm_data['Final Answer'].isin(['Action pending / In Process At Hub'])]['Assigned Request Ids'].nunique()
            pending_for_dispatch = zbm_data[zbm_data['Final Answer'].isin(['Dispatch  Pending'])]['Assigned Request Ids'].nunique()
            
            # Delivery Status (G + H) - Using Final Answer instead of Request Status
            delivered = zbm_data[zbm_data['Final Answer'].isin(['Delivered'])]['Assigned Request Ids'].nunique()
            dispatched_in_transit = zbm_data[zbm_data['Final Answer'].isin(['Dispatched & In Transit'])]['Assigned Request Ids'].nunique()
            
            # RTO Reasons - Calculate FIRST before using in formulas
            # Note: RTO reasons are based on Rto Reason field, not Final Answer
            # Count unique request IDs for each RTO reason
            incomplete_address = zbm_data[zbm_data['Rto Reason'].str.contains('Incomplete Address', na=False, case=False)]['Assigned Request Ids'].nunique()
            doctor_non_contactable = zbm_data[zbm_data['Rto Reason'].str.contains('Dr. Non contactable', na=False, case=False)]['Assigned Request Ids'].nunique()
            doctor_refused_to_accept = zbm_data[zbm_data['Rto Reason'].str.contains('Doctor Refused to Accept', na=False, case=False)]['Assigned Request Ids'].nunique()
            
            # Calculate RTO as sum of RTO reasons
            rto_total = incomplete_address + doctor_non_contactable + doctor_refused_to_accept
            
            # Calculated fields using the RTO total
            requests_dispatched = delivered + dispatched_in_transit + rto_total  # F = G + H + I
            sent_to_hub = pending_for_invoicing + pending_for_dispatch + requests_dispatched  # C = D + E + F
            requests_raised = request_cancelled_out_of_stock + action_pending_at_ho + sent_to_hub  # A + B + C
            hold_delivery = 0
            
            summary_data.append({
                'ZBM Code': zbm_code,
                'ZBM Name': zbm_name,
                'Unique TBMs': unique_tbms,
                'Unique HCPs': unique_hcps,
                'Requests Raised': requests_raised,
                'Request Cancelled Out of Stock': request_cancelled_out_of_stock,
                'Action Pending at HO': action_pending_at_ho,
                'Sent to HUB': sent_to_hub,
                'Pending for Invoicing': pending_for_invoicing,
                'Pending for Dispatch': pending_for_dispatch,
                'Requests Dispatched': requests_dispatched,
                'Delivered': delivered,
                'Dispatched In Transit': dispatched_in_transit,
                'RTO': rto_total,  # Use rto_total instead of rto
                'Incomplete Address': incomplete_address,
                'Doctor Non Contactable': doctor_non_contactable,
                'Doctor Refused to Accept': doctor_refused_to_accept,
                'Hold Delivery': hold_delivery
            })
        
        # Create DataFrame for this Division
        div_summary_df = pd.DataFrame(summary_data)
        
        # Create Excel file for this Division
        create_division_excel_report(div_code, div_summary_df, output_dir)
    
    print(f"\n🎉 Successfully created {len(divisions)} Division reports in directory: {output_dir}")

def create_division_excel_report(div_code, summary_df, output_dir):
    """Create Excel report for a specific Division with perfect formatting"""
    
    try:
        # Load template (use zbm_summary.xlsx as template)
        wb = load_workbook('zbm_summary.xlsx')
        ws = wb['ZBM']

        print(f"   📋 Creating Excel report for Division {div_code}...")

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
        
        # Search for header row
        header_row = None
        for row_idx in range(1, 15):  # Check first 15 rows
            for col_idx in range(1, min(30, ws.max_column + 1)):  # Check first 30 columns
                cell_value = get_cell_value_handling_merged(row_idx, col_idx)
                if cell_value and 'Area Name' in str(cell_value):
                    header_row = row_idx
                    break
            if header_row:
                break
        
        if header_row is None:
            print(f"   ⚠️ Could not find header row in template, using row 7 as default")
            header_row = 7
        
        print(f"   ℹ️ Detected header row: {header_row}")
        data_start_row = header_row + 1
        
        # Read actual column positions from template header row, handling merged cells
        column_mapping = {}
        for col_idx in range(1, min(30, ws.max_column + 1)):
            header_val = get_cell_value_handling_merged(header_row, col_idx)
            if header_val:
                header_str = str(header_val).strip()
                
                # Map template headers to our data columns - SKIP Area Name and ABM Name columns
                if 'Area Name' in header_str:
                    continue  # Skip Area Name column for division reports
                elif 'ABM Name' in header_str or ('Name' in header_str and 'ZBM' not in header_str):
                    continue  # Skip ABM Name column for division reports
                elif 'ZBM Name' in header_str or 'ZBM Code' in header_str:
                    # Add ZBM Name/Code to the first available column after skipping Area/ABM Name
                    if 'ZBM Name' not in column_mapping:
                        column_mapping['ZBM Name'] = col_idx
                    elif 'ZBM Code' not in column_mapping:
                        column_mapping['ZBM Code'] = col_idx
                elif 'Unique TBMs' in header_str or '# Unique TBMs' in header_str:
                    column_mapping['Unique TBMs'] = col_idx
                elif 'Unique HCPs' in header_str or '# Unique HCPs' in header_str:
                    column_mapping['Unique HCPs'] = col_idx
                elif 'Requests Raised' in header_str or '# Requests Raised' in header_str:
                    column_mapping['Requests Raised'] = col_idx
                elif 'Request Cancelled' in header_str or 'Out of Stock' in header_str:
                    column_mapping['Request Cancelled Out of Stock'] = col_idx
                elif 'Action pending' in header_str and 'HO' in header_str:
                    column_mapping['Action Pending at HO'] = col_idx
                elif 'Sent to HUB' in header_str:
                    column_mapping['Sent to HUB'] = col_idx
                elif 'Pending for Invoicing' in header_str:
                    column_mapping['Pending for Invoicing'] = col_idx
                elif 'Pending for Dispatch' in header_str:
                    column_mapping['Pending for Dispatch'] = col_idx
                elif 'Requests Dispatched' in header_str or '# Requests Dispatched' in header_str:
                    column_mapping['Requests Dispatched'] = col_idx
                elif header_str == 'Delivered' or 'Delivered (G)' in header_str:
                    column_mapping['Delivered'] = col_idx
                elif 'Dispatched & In Transit' in header_str or 'Dispatched In Transit' in header_str:
                    column_mapping['Dispatched In Transit'] = col_idx
                elif header_str == 'RTO' or 'RTO (I)' in header_str:
                    column_mapping['RTO'] = col_idx
                elif 'Incomplete Address' in header_str:
                    column_mapping['Incomplete Address'] = col_idx
                elif 'Doctor Non Contactable' in header_str or 'Dr. Non contactable' in header_str:
                    column_mapping['Doctor Non Contactable'] = col_idx
                elif 'Doctor Refused' in header_str or 'Refused to Accept' in header_str:
                    column_mapping['Doctor Refused to Accept'] = col_idx
                elif 'Hold Delivery' in header_str:
                    column_mapping['Hold Delivery'] = col_idx
        
        print(f"   ℹ️ Detected {len(column_mapping)} columns: {list(column_mapping.keys())}")
        
        # Verify we have the essential columns
        essential_cols = ['Unique TBMs', 'Unique HCPs', 'Requests Raised']
        missing_essential = [col for col in essential_cols if col not in column_mapping]
        if missing_essential:
            print(f"   ⚠️ WARNING: Missing essential columns in template: {missing_essential}")
        
        # Print summary_df to verify data exists
        print(f"   ℹ️ Summary DataFrame shape: {summary_df.shape}")
        print(f"   ℹ️ Summary DataFrame columns: {list(summary_df.columns)}")
        if len(summary_df) > 0:
            print(f"   ℹ️ First row sample: TBMs={summary_df.iloc[0]['Unique TBMs']}, HCPs={summary_df.iloc[0]['Unique HCPs']}")
        else:
            print(f"   ⚠️ WARNING: Summary DataFrame is empty!")
            return
        
        # Debug: Print all headers with their values from template
        print(f"   🔍 Template headers found:")
        for col_idx in range(1, min(30, ws.max_column + 1)):
            val = get_cell_value_handling_merged(header_row, col_idx)
            if val:
                print(f"      Column {col_idx}: '{val}'")
        
        # Find and delete Area Name and ABM Name columns (columns 5 and 6 typically)
        columns_to_delete = []
        for col_idx in range(1, min(30, ws.max_column + 1)):
            header_val = get_cell_value_handling_merged(header_row, col_idx)
            if header_val:
                header_str = str(header_val).strip()
                if 'Area Name' in header_str or 'ABM Name' in header_str:
                    columns_to_delete.append(col_idx)
        
        # Delete columns in reverse order to maintain correct indices
        for col_idx in sorted(columns_to_delete, reverse=True):
            ws.delete_cols(col_idx)
            # Adjust column mappings after deletion
            column_mapping = {k: (v - 1 if v > col_idx else v) for k, v in column_mapping.items()}
        
        # Clear existing data rows (preserve header)
        max_clear_rows = max(len(summary_df) + 10, 50)
        for r in range(data_start_row, data_start_row + max_clear_rows):
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

        # Write data rows
        template_data_row = data_start_row  # Use first data row as template
        for i in range(len(summary_df)):
            target_row = data_start_row + i
            
            # Copy formatting from template
            copy_row_style(template_data_row, target_row)
            
            # Write data to mapped columns
            for col_name, col_idx in column_mapping.items():
                if col_name in summary_df.columns:
                    value = summary_df.iloc[i][col_name]  # Use iloc for safer access
                    
                    # Debug print for first row
                    if i == 0:
                        print(f"      Writing '{col_name}' = {value} to column {col_idx}")
                    
                    try:
                        cell = ws.cell(row=target_row, column=col_idx)
                        cell.value = value
                        
                        # Apply number formatting for numeric columns
                        if isinstance(value, (int, float)) and not pd.isna(value):
                            cell.number_format = '0'
                    except Exception as e:
                        print(f"      Warning: Could not write to cell ({target_row}, {col_idx}): {e}")

        # Add total row
        total_row = data_start_row + len(summary_df)
        copy_row_style(template_data_row, total_row)
        
        # Write "Total" label in the first data column (skip Area Name and ABM Name if present)
        # Find the first numeric column after skipping Area/ABM Name columns
        first_data_col = None
        for col_name in ['ZBM Code', 'ZBM Name', 'Unique TBMs', 'Unique HCPs']:
            if col_name in column_mapping:
                first_data_col = column_mapping[col_name]
                break
        
        if first_data_col:
            try:
                cell = ws.cell(row=total_row, column=first_data_col)
                cell.value = "Total"
                cell.font = Font(bold=True, name='Arial', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            except:
                pass
        
        # Calculate and write totals
        for col_name, col_idx in column_mapping.items():
            if col_name in summary_df.columns and col_name not in ['ZBM Code', 'ZBM Name']:
                total_value = int(summary_df[col_name].sum())  # Ensure it's an integer
                
                print(f"      Writing Total '{col_name}' = {total_value} to column {col_idx}")
                
                try:
                    cell = ws.cell(row=total_row, column=col_idx)
                    cell.value = total_value
                    cell.font = Font(bold=True, name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.number_format = '0'
                except Exception as e:
                    print(f"      Warning: Could not write total to cell ({total_row}, {col_idx}): {e}")

        # Save file
        filename = f"Division_Summary_{div_code}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        wb.save(filepath)
        print(f"   ✅ Created: {filename}")
        
        # Print summary statistics
        print(f"   📊 Summary for Division {div_code}:")
        print(f"      Total ZBMs: {len(summary_df)}")
        print(f"      Total Unique TBMs: {summary_df['Unique TBMs'].sum()}")
        print(f"      Total Unique HCPs: {summary_df['Unique HCPs'].sum()}")
        print(f"      Total Requests Raised: {summary_df['Requests Raised'].sum()}")
        print(f"      Total Delivered: {summary_df['Delivered'].sum()}")
        print(f"      Total RTO: {summary_df['RTO'].sum()}")
        
    except Exception as e:
        print(f"   ❌ Error creating Excel report for Division {div_code}: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    create_division_hierarchical_reports()