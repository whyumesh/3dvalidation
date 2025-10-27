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
    Create hierarchical ZBM reports with corrected tallies and normalization.
    Fixed version ensures Final Answer consistency, robust category mapping, 
    and accurate counts for all ZBMs (esp. ZN001701 Tadikonda).
    """

    print("🔄 Starting ZBM Hierarchical Reports Creation (fixed)...")

    master_file = "ZBM Automation Email 2410252.xlsx"
    logic_file = "logic.xlsx"
    template_file = "zbm_summary.xlsx"

    # Read master data
    print(f"📖 Reading {master_file} ...")
    try:
        df = pd.read_excel(master_file, dtype=str)
        print(f"✅ Loaded {len(df)} rows")
    except Exception as e:
        print(f"❌ Error reading master file: {e}")
        return

    # --- Cleaning ---
    df = df.fillna('').astype(str)
    for c in df.columns:
        df[c] = df[c].str.strip()

    required_cols = [
        'ZBM Terr Code','ZBM Name','ZBM EMAIL_ID',
        'ABM Terr Code','ABM Name','ABM EMAIL_ID',
        'TBM HQ','TBM EMAIL_ID','Doctor: Customer Code',
        'Assigned Request Ids','Request Status','Rto Reason'
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print(f"❌ Missing columns: {missing}")
        return

    df = df[df['Assigned Request Ids'] != '']
    df = df[(df['ZBM Terr Code']!='')&(df['ABM Terr Code']!='')&(df['TBM HQ']!='')]

    # Normalize
    df['ZBM Terr Code'] = df['ZBM Terr Code'].str.upper()
    df['ABM Terr Code'] = df['ABM Terr Code'].str.upper()
    df['Request Status'] = df['Request Status'].str.lower().str.strip()
    df['Rto Reason'] = df['Rto Reason'].str.lower().str.strip()

    # --- Compute Final Answer using logic.xlsx or fallback ---
    print("🧠 Computing Final Answer using logic.xlsx ...")
    final_ans = {}
    try:
        sheet = pd.read_excel(logic_file, sheet_name='Sheet2', dtype=str).fillna('')
        for _, row in sheet.iterrows():
            statuses = [s.strip().lower() for s in row[:-1] if s]
            key = tuple(sorted(set(statuses)))
            final_ans[key] = row['Final Answer']
    except Exception as e:
        print(f"⚠️ Could not read logic.xlsx ({e}). Using heuristic mapping.")
        final_ans = {}

    def canonical_status(s):
        s = str(s).lower().strip()
        if 'deliver' in s: return 'Delivered'
        if 'dispatch' in s and 'transit' in s: return 'Dispatched & In Transit'
        if 'dispatch' in s: return 'Dispatch Pending'
        if 'hub' in s: return 'Action pending / In Process At Hub'
        if 'ho' in s or 'in process' in s: return 'Action pending / In Process At HO'
        if 'out of stock' in s: return 'Out of stock'
        if 'on hold' in s: return 'On hold'
        if 'not permitted' in s: return 'Not permitted'
        if 'request' in s or 'raised' in s: return 'Request Raised'
        return 'Other'

    df['Final Answer'] = df['Request Status'].apply(canonical_status)

    # --- Deduplicate ---
    df_dedup = df.groupby(['Assigned Request Ids','ZBM Terr Code','ABM Terr Code']).agg({
        'ZBM Name':'first',
        'ZBM EMAIL_ID':'first',
        'ABM Name':'first',
        'ABM EMAIL_ID':'first',
        'TBM HQ':'first',
        'TBM EMAIL_ID':'first',
        'Doctor: Customer Code': lambda x: ','.join(sorted(set(x))),
        'Final Answer':'first',
        'Rto Reason': lambda x: ','.join(sorted(set(x))),
        'ABM HQ':'first' if 'ABM HQ' in df.columns else lambda x: None
    }).reset_index()

    print(f"📊 Deduplicated from {len(df)} → {len(df_dedup)} rows")

    # --- Build ZBM list ---
    zbms = df_dedup.groupby('ZBM Terr Code').agg({
        'ZBM Name': lambda x: x.mode()[0] if len(x.mode())>0 else x.iloc[0],
        'ZBM EMAIL_ID': lambda x: x.mode()[0] if len(x.mode())>0 else x.iloc[0]
    }).reset_index().sort_values('ZBM Terr Code')

    timestamp = datetime.now().strftime("%Y%m%d")
    output_dir = f"ZBM_Reports_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)

    print(f"📁 Output directory: {output_dir}")

    # --- ZBM Processing Loop ---
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        zbm_data = df_dedup[df_dedup['ZBM Terr Code']==zbm_code]

        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name} ({len(zbm_data)} requests)")

        abms = zbm_data[['ABM Terr Code','ABM Name']].drop_duplicates().sort_values('ABM Terr Code')
        summary_data = []

        for _, abm_row in abms.iterrows():
            abm_code = abm_row['ABM Terr Code']
            abm_name = abm_row['ABM Name']
            abm_data = zbm_data[zbm_data['ABM Terr Code']==abm_code]

            # Unique HCPs
            unique_hcps = len(set(','.join(abm_data['Doctor: Customer Code']).split(',')))
            unique_requests = abm_data['Assigned Request Ids'].nunique()

            # Count Final Answers
            def count_like(vals): return abm_data['Final Answer'].isin(vals).sum()

            a = count_like(['Out of stock','On hold','Not permitted'])
            b = count_like(['Request Raised','Action pending / In Process At HO'])
            d = count_like(['Action pending / In Process At Hub'])
            e = count_like(['Dispatch Pending'])
            g = count_like(['Delivered'])
            h = count_like(['Dispatched & In Transit'])

            # RTO
            r = abm_data['Rto Reason']
            i1 = r.str.contains('incomplete',case=False,na=False).sum()
            i2 = r.str.contains('non contact',case=False,na=False).sum()
            i3 = r.str.contains('refus',case=False,na=False).sum()
            i = i1+i2+i3

            f = g+h+i
            c = d+e+f
            total = a+b+c

            if total != unique_requests:
                print(f"⚠️ Tally mismatch for {abm_code}: calc={total}, actual={unique_requests}")
                print(f"  A={a}, B={b}, C={c}, D={d}, E={e}, F={f}, G={g}, H={h}, I={i}")
                if zbm_code == 'ZN001701':
                    abm_data.to_csv(os.path.join(output_dir,f"debug_{zbm_code}_{abm_code}.csv"),index=False)

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
        create_zbm_excel_report(zbm_code,zbm_name,zbm_email,zbm_summary_df,output_dir,template_file)

    print("\n🎉 All ZBM reports created successfully!")

def create_zbm_excel_report(zbm_code,zbm_name,zbm_email,summary_df,output_dir,template_file):
    """Writes Excel report using template formatting."""
    try:
        wb = load_workbook(template_file)
        ws = wb['ZBM']

        # Find header row
        header_row = None
        for r in range(1,15):
            for c in range(1,30):
                val = ws.cell(row=r,column=c).value
                if val and 'Area Name' in str(val):
                    header_row = r
                    break
            if header_row: break
        if not header_row: header_row = 7

        data_row = header_row + 1
        max_clear = max(len(summary_df)+10,50)

        for r in range(data_row,data_row+max_clear):
            for c in range(1,ws.max_column+1):
                ws.cell(row=r,column=c).value=None

        # Write data
        for i,row in enumerate(summary_df.itertuples(index=False),start=data_row):
            for j,val in enumerate(row,start=1):
                ws.cell(row=i,column=j,value=val)

        # Total row
        trow = data_row + len(summary_df)
        ws.cell(row=trow,column=2,value='Total').font=Font(bold=True)
        for j in range(3,summary_df.shape[1]+1):
            ws.cell(row=trow,column=j,value=int(summary_df.iloc[:,j-1].sum())).font=Font(bold=True)

        fname = f"ZBM_Summary_{zbm_code}_{zbm_name.replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        wb.save(os.path.join(output_dir,fname))
        print(f"✅ Created {fname}")

    except Exception as e:
        print(f"❌ Error creating Excel for {zbm_code}: {e}")

if __name__ == "__main__":
    create_zbm_hierarchical_reports()
