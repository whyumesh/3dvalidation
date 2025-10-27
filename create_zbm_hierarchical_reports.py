import pandas as pd
import numpy as np
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from copy import copy as copy_style
import warnings

warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def create_zbm_hierarchical_reports():
    """
    Create hierarchical ZBM reports optimized for clean data files.
    Ensures exactly ONE report per unique ZBM Terr Code.
    """

    print("🔄 Starting ZBM Hierarchical Reports Creation...")

    master_file = "ZBM Automation Email 2410252.xlsx"
    logic_file = "logic.xlsx"
    template_file = "zbm_summary.xlsx"

    # Read master data
    print(f"📖 Reading {master_file}...")
    try:
        df = pd.read_excel(master_file, dtype=str)
        print(f"✅ Loaded {len(df)} rows")
    except Exception as e:
        print(f"❌ Error reading master file: {e}")
        return

    # --- Minimal cleaning for clean files ---
    df = df.fillna('')
    
    # Strip whitespace only (preserve case for now)
    for c in df.columns:
        if df[c].dtype == 'object':
            df[c] = df[c].str.strip()

    required_cols = [
        'ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID',
        'ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID',
        'TBM HQ', 'TBM EMAIL_ID', 'Doctor: Customer Code',
        'Assigned Request Ids', 'Request Status', 'Rto Reason'
    ]
    
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print(f"❌ Missing columns: {missing}")
        return

    # Only filter rows with missing critical identifiers
    original_len = len(df)
    df = df[
        (df['Assigned Request Ids'] != '') & 
        (df['ZBM Terr Code'] != '') & 
        (df['ABM Terr Code'] != '')
    ]
    print(f"📊 Filtered {original_len - len(df)} rows with missing critical fields")

    # Normalize only codes (keep names as-is)
    df['ZBM Terr Code'] = df['ZBM Terr Code'].str.upper()
    df['ABM Terr Code'] = df['ABM Terr Code'].str.upper()

    # --- Load logic mapping from logic.xlsx ---
    print("🧠 Loading status mapping from logic.xlsx...")
    status_mapping = {}
    
    try:
        logic_sheet = pd.read_excel(logic_file, sheet_name='Sheet2', dtype=str)
        logic_sheet = logic_sheet.fillna('')
        
        # Build mapping from each status to Final Answer
        for _, row in logic_sheet.iterrows():
            final_answer = row['Final Answer'] if 'Final Answer' in row else ''
            if not final_answer:
                continue
                
            # Get all status columns (all except 'Final Answer')
            for col in logic_sheet.columns:
                if col != 'Final Answer' and row[col]:
                    status_value = str(row[col]).strip()
                    if status_value:
                        status_mapping[status_value.lower()] = final_answer
        
        print(f"✅ Loaded {len(status_mapping)} status mappings")
    except Exception as e:
        print(f"⚠️ Could not read logic.xlsx ({e}). Using fallback mapping.")
        # Fallback mapping
        status_mapping = {
            'delivered': 'Delivered',
            'dispatched & in transit': 'Dispatched & In Transit',
            'dispatch pending': 'Dispatch Pending',
            'action pending / in process at hub': 'Action pending / In Process At Hub',
            'action pending / in process at ho': 'Action pending / In Process At HO',
            'out of stock': 'Out of stock',
            'on hold': 'On hold',
            'not permitted': 'Not permitted',
            'request raised': 'Request Raised'
        }

    # Apply mapping
    def map_status(status):
        status_lower = str(status).lower().strip()
        return status_mapping.get(status_lower, status)
    
    df['Final Answer'] = df['Request Status'].apply(map_status)
    
    # Show unique Final Answer values
    print(f"📊 Unique Final Answer values: {df['Final Answer'].unique()}")

    # --- Deduplicate by Request ID + ZBM + ABM ---
    print("🔄 Deduplicating requests...")
    
    df_dedup = df.groupby(['Assigned Request Ids', 'ZBM Terr Code', 'ABM Terr Code']).agg({
        'ZBM Name': 'first',
        'ZBM EMAIL_ID': 'first',
        'ABM Name': 'first',
        'ABM EMAIL_ID': 'first',
        'TBM HQ': 'first',
        'TBM EMAIL_ID': 'first',
        'Doctor: Customer Code': lambda x: ','.join(x.astype(str).unique()),
        'Final Answer': 'first',
        'Rto Reason': lambda x: ','.join([r for r in x.astype(str).unique() if r]),
        'Request Status': 'first'
    }).reset_index()

    print(f"📊 Deduplicated from {len(df)} → {len(df_dedup)} unique requests")

    # --- Build ZBM list (ONE entry per unique ZBM Terr Code) ---
    print("🔍 Identifying unique ZBM Terr Codes...")
    
    zbms = df_dedup.groupby('ZBM Terr Code').agg({
        'ZBM Name': lambda x: x.value_counts().index[0] if len(x) > 0 else x.iloc[0],  # Most common name
        'ZBM EMAIL_ID': lambda x: x.value_counts().index[0] if len(x) > 0 else x.iloc[0]  # Most common email
    }).reset_index().sort_values('ZBM Terr Code')

    print(f"📊 Found {len(zbms)} unique ZBM Terr Codes")

    timestamp = datetime.now().strftime("%Y%m%d")
    output_dir = f"ZBM_Reports_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)

    print(f"📁 Output directory: {output_dir}")

    # Track files created
    files_created = []

    # --- ZBM Processing Loop (ONE file per ZBM Terr Code) ---
    for idx, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        zbm_data = df_dedup[df_dedup['ZBM Terr Code'] == zbm_code].copy()

        print(f"\n🔄 [{idx+1}/{len(zbms)}] Processing ZBM: {zbm_code} - {zbm_name} ({len(zbm_data)} requests)")

        abms = zbm_data[['ABM Terr Code', 'ABM Name']].drop_duplicates().sort_values('ABM Terr Code')
        summary_data = []

        for _, abm_row in abms.iterrows():
            abm_code = abm_row['ABM Terr Code']
            abm_name = abm_row['ABM Name']
            abm_data = zbm_data[zbm_data['ABM Terr Code'] == abm_code].copy()

            # Count unique HCPs
            all_doctors = ','.join(abm_data['Doctor: Customer Code'].astype(str))
            unique_hcps = len([d for d in all_doctors.split(',') if d.strip()])
            
            # Total unique requests
            unique_requests = len(abm_data)

            # Count by Final Answer categories
            def count_status(statuses):
                if isinstance(statuses, str):
                    statuses = [statuses]
                return abm_data['Final Answer'].isin(statuses).sum()

            # A: Request Cancelled
            a = count_status(['Out of stock', 'On hold', 'Not permitted'])
            
            # B: Action Pending at HO
            b = count_status(['Request Raised', 'Action pending / In Process At HO'])
            
            # D: Pending for Invoicing (at Hub)
            d = count_status(['Action pending / In Process At Hub'])
            
            # E: Pending for Dispatch
            e = count_status(['Dispatch Pending'])
            
            # G: Delivered
            g = count_status(['Delivered'])
            
            # H: Dispatched In Transit
            h = count_status(['Dispatched & In Transit'])

            # I: RTO breakdown
            rto_reasons = abm_data['Rto Reason'].str.lower()
            i1 = rto_reasons.str.contains('incomplete', case=False, na=False).sum()
            i2 = rto_reasons.str.contains('non contact', case=False, na=False).sum()
            i3 = rto_reasons.str.contains('refus', case=False, na=False).sum()
            i = i1 + i2 + i3

            # Rollups
            f = g + h + i  # Requests Dispatched
            c = d + e + f  # Sent to HUB
            total = a + b + c

            # Validation
            if total != unique_requests:
                print(f"⚠️ Tally mismatch for {abm_code}: calculated={total}, actual={unique_requests}")
                print(f"   A={a}, B={b}, C={c}, D={d}, E={e}, F={f}, G={g}, H={h}, I={i}")

            summary_data.append({
                'Area Name': abm_code,
                'ABM Name': abm_name,
                'Unique HCPs': unique_hcps,
                'Requests Raised': total,
                'Request Cancelled Out of Stock': a,
                'Action Pending at HO': b,
                'Sent to HUB': c,
                'Pending for Invoicing': d,
                'Pending for Dispatch': e,
                'Requests Dispatched': f,
                'Delivered': g,
                'Dispatched In Transit': h,
                'RTO': i,
                'Incomplete Address': i1,
                'Doctor Non Contactable': i2,
                'Doctor Refused to Accept': i3
            })

        zbm_summary_df = pd.DataFrame(summary_data)
        filename = create_zbm_excel_report(zbm_code, zbm_name, zbm_email, zbm_summary_df, output_dir, template_file)
        if filename:
            files_created.append(filename)

    print(f"\n{'='*60}")
    print(f"🎉 Successfully created {len(files_created)} ZBM reports!")
    print(f"📊 Unique ZBM Terr Codes: {len(zbms)}")
    print(f"📁 Output directory: {output_dir}")
    print(f"{'='*60}")


def create_zbm_excel_report(zbm_code, zbm_name, zbm_email, summary_df, output_dir, template_file):
    """Writes Excel report using template formatting. Returns filename if successful."""
    try:
        wb = load_workbook(template_file)
        ws = wb['ZBM']

        # Find header row
        header_row = None
        for r in range(1, 15):
            for c in range(1, 30):
                val = ws.cell(row=r, column=c).value
                if val and 'Area Name' in str(val):
                    header_row = r
                    break
            if header_row:
                break
        
        if not header_row:
            header_row = 7

        data_row = header_row + 1
        max_clear = max(len(summary_df) + 10, 50)

        # Clear existing data
        for r in range(data_row, data_row + max_clear):
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).value = None

        # Write data
        for i, row in enumerate(summary_df.itertuples(index=False), start=data_row):
            for j, val in enumerate(row, start=1):
                cell = ws.cell(row=i, column=j)
                cell.value = val

        # Total row
        trow = data_row + len(summary_df)
        ws.cell(row=trow, column=1, value='').font = Font(bold=True)
        ws.cell(row=trow, column=2, value='Total').font = Font(bold=True)
        
        for j in range(3, summary_df.shape[1] + 1):
            cell = ws.cell(row=trow, column=j)
            cell.value = int(summary_df.iloc[:, j - 1].sum())
            cell.font = Font(bold=True)

        # Save file with sanitized ZBM name
        sanitized_name = zbm_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
        fname = f"ZBM_Summary_{zbm_code}_{sanitized_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        output_path = os.path.join(output_dir, fname)
        wb.save(output_path)
        print(f"✅ Created: {fname}")
        return fname

    except Exception as e:
        print(f"❌ Error creating Excel for {zbm_code}: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    create_zbm_hierarchical_reports()
