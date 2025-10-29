import pandas as pd
import numpy as np
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy as copy_style
import warnings

# Suppress FutureWarning for groupby operations
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def create_division_hierarchical_reports():
    """
    Create separate Division reports showing TBM Division hierarchy with perfect tallies
    Each TBM Division gets a report showing all TBMs under them
    """
    
    print("📄 Starting Division Hierarchical Reports Creation...")
    
    # Read master tracker data from Excel file
    print("📖 Reading ZBM Automation Email 2410252.xlsx...")
    try:
        df = pd.read_excel('ZBM Automation Email 2410252.xlsx')
        print(f"✅ Successfully loaded {len(df)} records")
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return
    
    print(f"📋 Columns in file: {list(df.columns)}")
    
    # Basic data preparation
    print("🧹 Preparing data...")
    
    # Find the correct column name for TBM/Created By
    tbm_created_by_col = None
    for col in df.columns:
        if 'created by' in col.lower() or 'created_by' in col.lower():
            tbm_created_by_col = col
            print(f"✅ Found TBM Created By column: '{col}'")
            break
    
    if tbm_created_by_col is None:
        print("Warning: Could not find 'Created By' column, will use 'TBM EMAIL_ID' instead")
        tbm_created_by_col = 'TBM EMAIL_ID'
    
    # Ensure required columns exist
    required_columns = ['TBM Division', 'AFFILIATE', 'DIV_NAME',
                        'ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID',
                        'ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID',
                        'Doctor: Customer Code', 'Assigned Request Ids', 'Request Status', 'Rto Reason']
    
    # Add the TBM created by column if it's different from TBM EMAIL_ID
    if tbm_created_by_col != 'TBM EMAIL_ID' and tbm_created_by_col not in required_columns:
        required_columns.append(tbm_created_by_col)
    
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing required columns: {missing}")
        return

    print(f"📊 Total rows in file: {len(df)}")
    print(f"📊 Unique Request IDs in raw data: {df['Assigned Request Ids'].nunique()}")
    print(f"📊 Unique TBM Divisions in raw data: {df['TBM Division'].nunique()}")

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

        # Merge Final Answer back to main dataframe
        df = df.merge(grouped[['Assigned Request Ids', 'Final Answer']], on='Assigned Request Ids', how='left')
        
        # Check for unmapped requests
        unmapped_count = (df['Final Answer'] == '❌ No matching rule').sum()
        if unmapped_count > 0:
            print(f"   WARNING: {unmapped_count} rows have no matching rule in logic.xlsx")
            print(f"   Unique Request IDs with no rule: {df[df['Final Answer'] == '❌ No matching rule']['Assigned Request Ids'].nunique()}")
            
    except Exception as e:
        print(f"❌ Error computing final status from logic.xlsx: {e}")
        return
    
    # Deduplicate at Request ID + TBM Division + ABM level to get correct counts
    print("🔧 Deduplicating data at Request ID + TBM Division + ABM level...")
    
    # Store original data for validation
    original_request_count = df['Assigned Request Ids'].nunique()
    
    # Deduplicate: Each unique (Request ID + TBM Division + ABM) combination should appear once
    agg_dict = {
        'AFFILIATE': 'first',
        'DIV_NAME': 'first',
        'ABM Name': 'first',
        'ABM EMAIL_ID': 'first',
        'ZBM Terr Code': 'first',
        'ZBM Name': 'first',
        'ZBM EMAIL_ID': 'first',
        'Doctor: Customer Code': 'first',
        'Final Answer': 'first',
        'Rto Reason': 'first',
    }
    
    # Add TBM created by column if it exists and is different
    if tbm_created_by_col and tbm_created_by_col != 'TBM EMAIL_ID':
        agg_dict[tbm_created_by_col] = 'first'
    
    # Add TBM HQ if it exists
    if 'TBM HQ' in df.columns:
        agg_dict['TBM HQ'] = 'first'
    
    # Add ABM HQ if it exists
    if 'ABM HQ' in df.columns:
        agg_dict['ABM HQ'] = 'first'
    
    df_dedup = df.groupby(['Assigned Request Ids', 'TBM Division', 'ABM Terr Code']).agg(agg_dict).reset_index()
    
    print(f"📊 Deduplicated from {len(df)} rows to {len(df_dedup)} unique (Request ID + TBM Division + ABM) combinations")
    print(f"📊 Unique Request IDs after dedup: {df_dedup['Assigned Request Ids'].nunique()}")
    
    # Get unique TBM Divisions
    divisions = df_dedup.groupby('TBM Division').agg({
        'AFFILIATE': lambda x: x.mode()[0] if len(x.mode()) > 0 else x.iloc[0],
        'DIV_NAME': lambda x: x.mode()[0] if len(x.mode()) > 0 else x.iloc[0]
    }).reset_index().sort_values('TBM Division')
    
    print(f"📋 Found {len(divisions)} unique TBM Divisions")
    
    # Debug: Check for any duplicates
    duplicate_codes = divisions['TBM Division'].value_counts()
    if len(duplicate_codes[duplicate_codes > 1]) > 0:
        print(f"WARNING: Found duplicate TBM Division codes after deduplication!")
        print(duplicate_codes[duplicate_codes > 1])
    
    # Debug: Show first few Divisions and their ABMs
    print("\n🔍 Division-ABM Mapping (first 5):")
    for idx, (_, div_row) in enumerate(divisions.head(5).iterrows()):
        div_code = div_row['TBM Division']
        affiliate = div_row['AFFILIATE']
        div_name = div_row['DIV_NAME']
        div_data_temp = df_dedup[df_dedup['TBM Division'] == div_code]
        abms_temp = div_data_temp[['ABM Terr Code', 'ABM Name']].drop_duplicates()
        requests_temp = div_data_temp['Assigned Request Ids'].nunique()
        print(f"   {idx+1}. Division {div_code} ({affiliate} - {div_name}): {len(abms_temp)} ABMs, {requests_temp} requests")
    
    # Create output directory
    timestamp = datetime.now().strftime('%Y%m%d')
    output_dir = f"Division_Reports_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 Created output directory: {output_dir}")
    
    # Process each Division
    file_count = 0
    total_validation_errors = 0
    
    for _, div_row in divisions.iterrows():
        div_code = div_row['TBM Division']
        affiliate = div_row['AFFILIATE']
        div_name = div_row['DIV_NAME']
        
        print(f"\n📄 Processing Division: {div_code} - {affiliate} - {div_name}")
        
        # Filter data for this Division (using deduplicated data)
        div_data = df_dedup[df_dedup['TBM Division'] == div_code].copy()
        
        if len(div_data) == 0:
            print(f"No data found for Division: {div_code}")
            continue
        
        # Calculate totals for the entire Division (not individual ABMs)
        print(f"   📊 Calculating totals for entire Division")
        
        # Calculate all metrics using deduplicated data for the entire division
        # Use the dynamically found TBM created by column
        unique_tbms = div_data[tbm_created_by_col].nunique()
        unique_hcps = div_data['Doctor: Customer Code'].nunique()
        unique_requests = div_data['Assigned Request Ids'].nunique()
        
        # === SECTION A: Request Cancelled Out of Stock ===
        # Final Answer: Out of stock, On hold, Not permitted
        ho_statuses = ['Out of stock', 'On hold', 'Not permitted']
        request_cancelled_out_of_stock = div_data[div_data['Final Answer'].isin(ho_statuses)]['Assigned Request Ids'].nunique()
        
        # === SECTION B: Action Pending at HO ===
        # Final Answer: Request Raised, Action pending / In Process At HO
        pending_statuses = ['Request Raised', 'Action pending / In Process At HO']
        action_pending_at_ho = div_data[div_data['Final Answer'].isin(pending_statuses)]['Assigned Request Ids'].nunique()
        
        # === SECTION D: Pending for Invoicing ===
        # Final Answer: Action pending / In Process At Hub
        hub_pending_statuses = ['Action pending / In Process At Hub']
        pending_for_invoicing = div_data[div_data['Final Answer'].isin(hub_pending_statuses)]['Assigned Request Ids'].nunique()
        
        # === SECTION E: Pending for Dispatch ===
        # Final Answer: Dispatch Pending
        dispatch_pending_statuses = ['Dispatch  Pending', 'Dispatch Pending']
        pending_for_dispatch = div_data[div_data['Final Answer'].isin(dispatch_pending_statuses)]['Assigned Request Ids'].nunique()
        
        # === SECTION G: Delivered ===
        # Final Answer: Delivered
        delivered_statuses = ['Delivered']
        delivered = div_data[div_data['Final Answer'].isin(delivered_statuses)]['Assigned Request Ids'].nunique()
        
        # === SECTION H: Dispatched & In Transit ===
        # Final Answer: Dispatched & In Transit
        transit_statuses = ['Dispatched & In Transit']
        dispatched_in_transit = div_data[div_data['Final Answer'].isin(transit_statuses)]['Assigned Request Ids'].nunique()
        
        # === SECTION I: RTO (Return to Origin) ===
        # RTO Total: ONLY count requests with "Return" Final Answer
        rto_total = div_data[div_data['Final Answer'] == 'Return']['Assigned Request Ids'].nunique()
        
        # RTO Reasons: Count based on unique Request IDs that have RTO reasons
        # Get unique Request IDs for this Division that have Return status
        unique_request_ids = div_data[div_data['Final Answer'] == 'Return']['Assigned Request Ids'].unique()
        
        # For each unique Request ID, determine its RTO reason category based on priority
        incomplete_address = 0
        doctor_refused_to_accept = 0
        doctor_non_contactable = 0
        rto_due_to_hold_delivery = 0
    
        for req_id in unique_request_ids:
            # Get all rows for this Request ID under this Division
            req_rows = div_data[div_data['Assigned Request Ids'] == req_id]
            
            # Check RTO reasons in the Rto Reason column (check all rows for this request)
            rto_col = req_rows['Rto Reason'].astype(str).str.strip().str.lower()
            
            # Check which reasons are present for this Request ID
            has_incomplete = rto_col.str.contains('incomplete address', na=False, regex=False).any()
            has_refused = rto_col.str.contains('refused to accept', na=False, regex=False).any()
            has_non_contactable = rto_col.str.contains('non contactable', na=False, regex=False).any()
            has_rto_hold_delivery = rto_col.str.contains('hold delivery', na=False, regex=False).any()

            # Assign to EXACTLY ONE category based on priority
            # Priority: 1) Incomplete Address, 2) Doctor Refused, 3) Doctor Non Contactable
            if has_incomplete:
                incomplete_address += 1
            elif has_refused:
                doctor_refused_to_accept += 1
            elif has_non_contactable:
                doctor_non_contactable += 1
            elif has_rto_hold_delivery:
                rto_due_to_hold_delivery +=1
            # If no RTO reason found, don't count in any category
        
        # Validate RTO breakdown
        rto_reasons_sum = incomplete_address + doctor_non_contactable + doctor_refused_to_accept + rto_due_to_hold_delivery
        if rto_reasons_sum != rto_total:
            print(f"      RTO Breakdown mismatch for Division {div_code}:")
            print(f"         RTO Total: {rto_total}, Reasons Sum: {rto_reasons_sum}")
            print(f"         Incomplete: {incomplete_address}, Refused: {doctor_refused_to_accept}, Non-contactable: {doctor_non_contactable}")
            print(f"         RTO Hold due to hold delivery: {rto_due_to_hold_delivery}")
        
        # === CALCULATED FIELDS ===
        # F = Requests Dispatched = G + H + I
        requests_dispatched = delivered + dispatched_in_transit + rto_total
        
        # C = Sent to HUB = D + E + F
        sent_to_hub = pending_for_invoicing + pending_for_dispatch + requests_dispatched
        
        # Total = Requests Raised = A + B + C
        requests_raised_calc = request_cancelled_out_of_stock + action_pending_at_ho + sent_to_hub
        
        # Hold Delivery (not used in current logic)
        hold_delivery = 0
        
        # Check for unmapped requests
        all_mapped_statuses = ho_statuses + pending_statuses + hub_pending_statuses + dispatch_pending_statuses + delivered_statuses + transit_statuses + ['Return']
        mapped_requests = div_data[div_data['Final Answer'].isin(all_mapped_statuses)]['Assigned Request Ids'].nunique()
        unmapped_count = unique_requests - mapped_requests
        
        if unmapped_count > 0:
            print(f"      {unmapped_count} unmapped requests for Division {div_code}")
            unmapped_data = div_data[~div_data['Final Answer'].isin(all_mapped_statuses)]
            unmapped_statuses = unmapped_data['Final Answer'].value_counts().to_dict()
            print(f"         Unmapped statuses: {unmapped_statuses}")
        
        # Verify tally
        if requests_raised_calc != unique_requests:
            print(f"         TALLY MISMATCH for Division {div_code}:")
            print(f"         Calculated: {requests_raised_calc}, Actual: {unique_requests}, Diff: {unique_requests - requests_raised_calc}")
            print(f"         A={request_cancelled_out_of_stock}, B={action_pending_at_ho}, C={sent_to_hub}")
            print(f"         D={pending_for_invoicing}, E={pending_for_dispatch}, F={requests_dispatched}")
            print(f"         G={delivered}, H={dispatched_in_transit}, I={rto_total}")
            total_validation_errors += 1
        
        # Use actual unique request count to ensure accuracy
        requests_raised = unique_requests
        
        # Create single row summary data for the entire Division
        summary_data = [{
            'Affiliate': affiliate,
            'Division': div_code,
            'Division Name': div_name,
            'Area Name': f"Division {div_code} - {affiliate} - {div_name}",
            'ABM Name': 'Total',
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
            'RTO': rto_total,
            'Incomplete Address': incomplete_address,
            'Doctor Non Contactable': doctor_non_contactable,
            'Doctor Refused to Accept': doctor_refused_to_accept,
            'Hold Delivery': rto_due_to_hold_delivery,
        }]
        
        # Create DataFrame for this Division (single row)
        div_summary_df = pd.DataFrame(summary_data)
        
        # Validate Division total
        div_total_requests = div_data['Assigned Request Ids'].nunique()
        div_summary_total = div_summary_df['Requests Raised'].sum()
        
        if div_total_requests != div_summary_total:
            print(f"      WARNING: Division {div_code} total mismatch!")
            print(f"      Actual unique requests: {div_total_requests}")
            print(f"      Summary total: {div_summary_total}")
            print(f"      Difference: {div_summary_total - div_total_requests}")
        
        # Create Excel file for this Division
        create_division_excel_report(div_code, affiliate, div_name, div_summary_df, output_dir)
        file_count += 1
    
    print(f"\n🎉 Successfully created {file_count} Division reports in directory: {output_dir}")
    print(f"📊 Total Divisions processed: {file_count}")
    if total_validation_errors > 0:
        print(f"WARNING: {total_validation_errors} TBMs had validation errors")
    else:
        print(f"✅ All tallies match perfectly!")

def create_division_excel_report(div_code, affiliate, div_name, summary_df, output_dir):
    """Create Excel report for a specific Division with perfect formatting"""
    
    try:
        # Load template
        wb = load_workbook('division summary.xlsx')
        ws = wb.active  # Use the active sheet

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
        for row_idx in range(1, 15):
            for col_idx in range(1, min(30, ws.max_column + 1)):
                cell_value = get_cell_value_handling_merged(row_idx, col_idx)
                if cell_value and 'Affiliate' in str(cell_value):
                    header_row = row_idx
                    break
            if header_row:
                break
        
        if header_row is None:
            header_row = 5  # Default to row 5 based on template structure
        
        data_start_row = header_row + 1
        
        # Read actual column positions from template header row
        column_mapping = {}
        for col_idx in range(1, min(30, ws.max_column + 1)):
            header_val = get_cell_value_handling_merged(header_row, col_idx)
            if header_val:
                header_str = str(header_val).strip()
                
                if 'Affiliate' in header_str:
                    column_mapping['Affiliate'] = col_idx
                elif 'Division' in header_str and 'Name' not in header_str:
                    column_mapping['Division'] = col_idx
                elif 'Division Name' in header_str:
                    column_mapping['Division Name'] = col_idx
                elif 'TBMs' in header_str or '# TBMs' in header_str:
                    column_mapping['Unique TBMs'] = col_idx
                elif 'HCPs' in header_str or '# Unique HCPs' in header_str:
                    column_mapping['Unique HCPs'] = col_idx
                elif 'Requests raised' in header_str or '# Requests raised' in header_str:
                    column_mapping['Requests Raised'] = col_idx
                elif 'dispatched' in header_str and 'In Transit' not in header_str:
                    column_mapping['Requests Dispatched'] = col_idx
                elif 'Action pending' in header_str and 'HO' in header_str:
                    column_mapping['Action Pending at HO'] = col_idx
                elif 'Dispatched & In Transit' in header_str:
                    column_mapping['Dispatched In Transit'] = col_idx
                elif 'Out of stock' in header_str:
                    column_mapping['Request Cancelled Out of Stock'] = col_idx
                elif 'Delivered' in header_str:
                    column_mapping['Delivered'] = col_idx
                elif 'RTO' in header_str:
                    column_mapping['RTO'] = col_idx
                elif 'Incomplete Address' in header_str:
                    column_mapping['Incomplete Address'] = col_idx
                elif 'Non contactable' in header_str:
                    column_mapping['Doctor Non Contactable'] = col_idx
                elif 'refused to accept' in header_str:
                    column_mapping['Doctor Refused to Accept'] = col_idx
        
        # Clear existing data rows
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

        # Write single data row (total only)
        template_data_row = data_start_row
        target_row = data_start_row
        copy_row_style(template_data_row, target_row)
        
        for col_name, col_idx in column_mapping.items():
            if col_name in summary_df.columns:
                value = summary_df.iloc[0][col_name]
                
                try:
                    cell = ws.cell(row=target_row, column=col_idx)
                    cell.value = value
                    
                    if isinstance(value, (int, float)) and not pd.isna(value):
                        cell.number_format = '0'
                        cell.font = Font(bold=True, name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.font = Font(bold=True, name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                except:
                    pass

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

if __name__ == "__main__":
    create_division_hierarchical_reports()