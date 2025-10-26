#!/usr/bin/env python3
"""
ZBM Email Display
Displays professional emails in Outlook for ZBMs with precise data matching
USES OUTLOOK TO DISPLAY EMAILS (NOT SEND AUTOMATICALLY)
"""

import pandas as pd
import os
from datetime import datetime
import warnings
import win32com.client
from openpyxl import load_workbook

# Suppress pandas warnings
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

def send_zbm_emails():
    """Display emails in Outlook for review without sending"""
    
    print("🚀 Starting ZBM Email Display...")
    print("📧 This will DISPLAY emails in Outlook for review - NOT SEND automatically")
    
    # Read Sample Master Tracker data
    print("📖 Reading Sample Master Tracker.xlsx...")
    try:
        df = pd.read_excel('Sample Master Tracker.xlsx')
        print(f"✅ Successfully loaded {len(df)} records from Sample Master Tracker.xlsx")
    except Exception as e:
        print(f"❌ Error reading Sample Master Tracker.xlsx: {e}")
        return
    
    # Required columns
    required_columns = [
        'ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID', 'ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID',
        'Assigned Request Ids', 'Doctor: Customer Code', 'Request Status', 'TBM EMAIL_ID', 'TBM HQ'
    ]
    
    # Check for missing columns
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing required columns: {missing}")
        return
    
    # Clean and filter data
    df = df.dropna(subset=['ZBM Terr Code', 'ZBM Name', 'ABM Terr Code', 'ABM Name'])
    df = df[df['ZBM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ABM Terr Code'].astype(str).str.strip() != '']
    df = df[df['ZBM Terr Code'].astype(str).str.startswith('ZN')]
    
    print(f"📊 After cleaning: {len(df)} records remaining")
    
    # Compute Final Status using logic.xlsx
    print("🧠 Computing final status...")
    try:
        xls_rules = pd.ExcelFile('logic.xlsx')
        
        # Check available sheet names
        sheet_names = xls_rules.sheet_names
        print(f"   📋 Available sheets in logic.xlsx: {sheet_names}")
        
        # Try to find the rules sheet (case-insensitive)
        rules_sheet = None
        for sheet in sheet_names:
            if 'rule' in sheet.lower():
                rules_sheet = sheet
                break
        
        if rules_sheet:
            print(f"   📖 Using sheet: {rules_sheet}")
            rules_df = pd.read_excel(xls_rules, rules_sheet)
        else:
            # Use the first sheet if no rules sheet found
            rules_sheet = sheet_names[0]
            print(f"   📖 Using first sheet: {rules_sheet}")
            rules_df = pd.read_excel(xls_rules, rules_sheet)
        
        # Check if required columns exist
        required_rule_columns = ['Request Status', 'Final Answer']
        missing_rule_columns = [col for col in required_rule_columns if col not in rules_df.columns]
        
        if missing_rule_columns:
            print(f"   ⚠️ Missing columns in rules sheet: {missing_rule_columns}")
            print(f"   📋 Available columns: {list(rules_df.columns)}")
            # Use alternative column names if available
            status_col = None
            answer_col = None
            
            for col in rules_df.columns:
                if 'request' in col.lower() and 'status' in col.lower():
                    status_col = col
                if 'final' in col.lower() and 'answer' in col.lower():
                    answer_col = col
            
            if status_col and answer_col:
                print(f"   🔄 Using alternative columns: {status_col} -> {answer_col}")
                status_mapping = {}
                for _, row in rules_df.iterrows():
                    if pd.notna(row[status_col]) and pd.notna(row[answer_col]):
                        status_mapping[row[status_col]] = row[answer_col]
            else:
                raise Exception("Cannot find suitable columns for status mapping")
        else:
            status_mapping = {}
            for _, row in rules_df.iterrows():
                if pd.notna(row['Request Status']) and pd.notna(row['Final Answer']):
                    status_mapping[row['Request Status']] = row['Final Answer']
        
        df['Final Status'] = df['Request Status'].map(status_mapping)
        df['Final Status'] = df['Final Status'].fillna(df['Request Status'])
        print("✅ Final status computed successfully")
        
    except Exception as e:
        print(f"❌ Error computing final status: {e}")
        print("   🔄 Using Request Status as Final Status")
        df['Final Status'] = df['Request Status']
    
    # Get unique ZBMs
    zbms = df[['ZBM Terr Code', 'ZBM Name', 'ZBM EMAIL_ID']].drop_duplicates().sort_values('ZBM Terr Code')
    print(f"📋 Found {len(zbms)} unique ZBMs")
    
    # Initialize Outlook with robust error handling
    print("📧 Initializing Outlook...")
    outlook = None
    
    # Try different Outlook initialization methods with more robust error handling
    outlook_methods = [
        "Outlook.Application",
        "Outlook.Application.16",  # Office 2016/2019/365
        "Outlook.Application.15",  # Office 2013
        "Outlook.Application.14",  # Office 2010
        "Outlook.Application.12",  # Office 2007
    ]
    
    for method in outlook_methods:
        try:
            print(f"   🔄 Trying: {method}")
            outlook = win32com.client.Dispatch(method)
            # Test if we can actually create a mail item
            test_mail = outlook.CreateItem(0)
            del test_mail
            print(f"✅ Outlook initialized successfully using: {method}")
            break
        except Exception as e:
            print(f"   ❌ Failed with {method}: {e}")
            continue
    
    if outlook is None:
        print("❌ Could not initialize Outlook with any method")
        print("🔧 Troubleshooting steps:")
        print("   1. Ensure Outlook is installed on this computer")
        print("   2. Try opening Outlook manually first")
        print("   3. Check if Outlook is running in the background")
        print("   4. Try running as administrator")
        print("   5. Install Microsoft Office/Outlook if not present")
        print("   6. Try installing pywin32: pip install pywin32")
        
        # Fallback: Create HTML email files
        print("\n🔄 Creating HTML email files as fallback...")
        create_html_email_files(df, zbms)
        return
    
    # Process each ZBM
    success_count = 0
    error_count = 0
    
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name}")
        
        try:
            # Read the actual ZBM summary report file
            zbm_summary_df = read_zbm_summary_report(zbm_code, zbm_name)
            
            if zbm_summary_df is None or zbm_summary_df.empty:
                print(f"⚠️ No ZBM summary report found for {zbm_code}")
                continue
            
            # Convert summary report data to email format
            summary_data = create_summary_data_from_report(zbm_summary_df)
            summary_df = pd.DataFrame(summary_data)
            
            # Get ABM emails for CC from the original data
            zbm_data = df[df['ZBM Terr Code'] == zbm_code]
            abms = zbm_data.groupby(['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']).agg({
                'TBM HQ': 'first'
            }).reset_index()
            
            # Generate email content
            email_content, cc_emails = generate_email_content(zbm_name, zbm_email, abms, summary_df)
            
            # Display email in Outlook (without sending)
            display_single_email(outlook, zbm_email, cc_emails, email_content, zbm_code, zbm_name)
            
            success_count += 1
            print(f"   ✅ Email displayed in Outlook for {zbm_name}")
            
        except Exception as e:
            error_count += 1
            print(f"   ❌ Error displaying email for {zbm_name}: {e}")
            continue
    
    print(f"\n🎉 Email display completed!")
    print(f"✅ Successfully displayed: {success_count} emails")
    print(f"❌ Failed to display: {error_count} emails")
    print(f"\n📧 All emails are now open in Outlook for your review and manual sending")

def read_zbm_summary_report(zbm_code, zbm_name):
    """Read the actual ZBM summary report file created by create_zbm_hierarchical_reports.py"""
    
    # Look for ZBM summary report files in current directory and subdirectories
    # Find all matching files and use the most recent one
    matching_files = []
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.startswith(f"ZBM_Summary_{zbm_code}_") and file.endswith('.xlsx'):
                filepath = os.path.join(root, file)
                matching_files.append(filepath)
    
    if not matching_files:
        print(f"   ⚠️ No ZBM summary report found for {zbm_code}")
        return None
    
    # Sort by modification time and use the most recent
    matching_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    latest_file = matching_files[0]
    
    print(f"   📊 Found ZBM summary report: {os.path.basename(latest_file)}")
    
    try:
        # Read the Excel file with proper header detection
        df = read_zbm_summary_report_properly(latest_file)
        if df is not None and not df.empty:
            print(f"   ✅ Successfully loaded ZBM summary report with {len(df)} ABMs")
            return df
        else:
            print(f"   ❌ Could not read ZBM summary report properly")
            return None
    except Exception as e:
        print(f"   ❌ Error reading ZBM summary report: {e}")
        return None
    
    print(f"   ⚠️ No ZBM summary report found for {zbm_code}")
    return None

def read_zbm_summary_report_properly(filepath):
    """Read ZBM summary report with proper column headers"""
    
    try:
        # Read all rows to find where the actual data starts
        raw_df = pd.read_excel(filepath, sheet_name='ZBM', header=None)
        
        # Look for the row that contains column headers
        header_row = None
        for i in range(min(10, len(raw_df))):
            row = raw_df.iloc[i]
            # Check if this row contains expected column names
            row_str = ' '.join([str(cell).strip() for cell in row if pd.notna(cell)])
            if 'Area Name' in row_str or 'ABM Name' in row_str:
                header_row = i
                break
        
        if header_row is not None:
            # Read with the found header row
            df = pd.read_excel(filepath, sheet_name='ZBM', header=header_row)
            
            # Clean up the data - remove rows where all key columns are NaN
            if 'Area Name' in df.columns:
                # Keep only rows where Area Name or ABM Name has data
                df = df.dropna(subset=['Area Name', 'ABM Name'], how='all')
                df = df[df['Area Name'].astype(str).str.strip() != '']
                df = df[df['ABM Name'].astype(str).str.strip() != '']
                
                # IMPORTANT: Filter out the Total row
                df = df[df['ABM Name'] != 'Total']
            
            # Convert numeric columns to proper types
            numeric_columns = ['# Unique TBMs', '# Unique HCPs', '# Requests Raised\n(A+B+C)',
                             'Request Cancelled / Out of Stock (A)', 'Action pending / In Process At HO (B)',
                             "Sent to HUB ('C)\n(D+E+F)", 'Pending for Invoicing (D)', 'Pending for Dispatch (E)',
                             '# Requests Dispatched (F)\n(G+H+I)', 'Delivered (G)', 'Dispatched & In Transit (H)',
                             'RTO (I)', 'Incomplete Address', 'Doctor Non Contactable', 'Doctor Refused to Accept', 'Hold Delivery']
            
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            
            return df
        else:
            print(f"   ❌ Could not find header row in {filepath}")
            return None
        
    except Exception as e:
        print(f"   ❌ Error reading Excel file {filepath}: {e}")
        return None

def create_summary_data_from_report(summary_df):
    """Convert ZBM summary report data to email format"""
    
    if summary_df is None or summary_df.empty:
        return []
    
    summary_data = []
    
    for _, row in summary_df.iterrows():
        # Map the data from the Excel report to email format using the actual column names
        summary_data.append({
            'Area Name': row.get('Area Name', ''),
            'ABM Name': row.get('ABM Name', ''),
            'Unique TBMs': row.get('# Unique TBMs', 0),
            'Unique HCPs': row.get('# Unique HCPs', 0),
            'Unique Requests': row.get('# Requests Raised\n(A+B+C)', 0),  # Use actual column name
            'Requests Raised': row.get('# Requests Raised\n(A+B+C)', 0),
            'Request Cancelled Out of Stock': row.get('Request Cancelled / Out of Stock (A)', 0),
            'Action Pending at HO': row.get('Action pending / In Process At HO (B)', 0),
            'Sent to HUB': row.get("Sent to HUB ('C)\n(D+E+F)", 0),
            'Pending for Invoicing': row.get('Pending for Invoicing (D)', 0),
            'Pending for Dispatch': row.get('Pending for Dispatch (E)', 0),
            'Requests Dispatched': row.get('# Requests Dispatched (F)\n(G+H+I)', 0),
            'Delivered': row.get('Delivered (G)', 0),
            'Dispatched In Transit': row.get('Dispatched & In Transit (H)', 0),
            'RTO': row.get('RTO (I)', 0),
            'Incomplete Address': row.get('Incomplete Address', 0),
            'Doctor Non Contactable': row.get('Doctor Non Contactable', 0),
            'Doctor Refused to Accept': row.get('Doctor Refused to Accept', 0),
            'Hold Delivery': row.get('Hold Delivery', 0)
        })
    
    return summary_data

def generate_email_content(zbm_name, zbm_email, abms, summary_df):
    """Generate professional email content"""
    
    current_date = datetime.now().strftime('%B %d, %Y')
    
    # Get ABM emails for CC
    abm_emails = abms['ABM EMAIL_ID'].dropna().unique().tolist()
    cc_emails = ', '.join(abm_emails)
    
    # Create summary table HTML
    table_html = create_summary_table_html(summary_df)
    
    email_content = f"""
Hi {zbm_name},

Please refer the status Sample requests raised in Abbworld for your area.

{table_html}

You can track your sample request at the following link with the Docket Number:

DTDC: <a href="https://www.dtdc.com/tracking">Click here</a>

Speed Post: <a href="https://www.indiapost.gov.in/vas/Pages/IndiaPostHome.aspx">Click Here</a>

In case of any query, please contact 1Point.

Regards,
Umesh Pawar.
"""
    
    return email_content, cc_emails

def create_summary_table_html(summary_df):
    """Create HTML table for summary data with all columns from summary report including section headers"""
    
    if summary_df.empty:
        return "<p>No data available</p>"
    
    # Create comprehensive HTML table with section headers and all data
    html = "<div style='font-family: Arial, sans-serif; margin: 20px 0;'>"
    
    # Add section headers with styling
    html += "<div style='background-color: #f8f9fa; padding: 15px; border-left: 4px solid #007bff; margin-bottom: 20px;'>"
    html += "<h3 style='margin: 0; color: #007bff;'>Sample Request Status Summary</h3>"
    html += "<p style='margin: 5px 0 0 0; color: #666;'>Complete breakdown of sample requests by ABM territory</p>"
    html += "</div>"
    
    # Create main summary table
    html += "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; font-size: 11px; margin-bottom: 20px;'>"
    
    # Header row with all columns from summary report
    html += "<tr style='background-color: #e9ecef; font-weight: bold; text-align: center;'>"
    html += "<th rowspan='2' style='vertical-align: middle;'>Area Name</th>"
    html += "<th rowspan='2' style='vertical-align: middle;'>ABM Name</th>"
    html += "<th rowspan='2' style='vertical-align: middle;'># Unique<br/>TBMs</th>"
    html += "<th rowspan='2' style='vertical-align: middle;'># Unique<br/>HCPs</th>"
    html += "<th rowspan='2' style='vertical-align: middle;'># Requests<br/>Raised<br/>(A+B+C)</th>"
    html += "<th colspan='2' style='background-color: #fff3cd;'>HO Section</th>"
    html += "<th colspan='3' style='background-color: #d1ecf1;'>HUB Section</th>"
    html += "<th colspan='3' style='background-color: #d4edda;'>Delivery Status</th>"
    html += "<th colspan='4' style='background-color: #f8d7da;'>RTO Reasons</th>"
    html += "</tr>"
    
    # Sub-header row
    html += "<tr style='background-color: #e9ecef; font-weight: bold; text-align: center;'>"
    html += "<th style='background-color: #fff3cd;'>Request Cancelled /<br/>Out of Stock (A)</th>"
    html += "<th style='background-color: #fff3cd;'>Action pending /<br/>In Process At HO (B)</th>"
    html += "<th style='background-color: #d1ecf1;'>Sent to HUB (C)<br/>(D+E+F)</th>"
    html += "<th style='background-color: #d1ecf1;'>Pending for<br/>Invoicing (D)</th>"
    html += "<th style='background-color: #d1ecf1;'>Pending for<br/>Dispatch (E)</th>"
    html += "<th style='background-color: #d4edda;'># Requests Dispatched (F)<br/>(G+H+I)</th>"
    html += "<th style='background-color: #d4edda;'>Delivered (G)</th>"
    html += "<th style='background-color: #d4edda;'>Dispatched &<br/>In Transit (H)</th>"
    html += "<th style='background-color: #f8d7da;'>RTO (I)</th>"
    html += "<th style='background-color: #f8d7da;'>Incomplete<br/>Address</th>"
    html += "<th style='background-color: #f8d7da;'>Doctor Non<br/>Contactable</th>"
    html += "<th style='background-color: #f8d7da;'>Doctor Refused<br/>to Accept</th>"
    html += "</tr>"
    
    # Data rows
    for _, row in summary_df.iterrows():
        html += "<tr style='border-bottom: 1px solid #dee2e6;'>"
        html += f"<td style='font-weight: bold;'>{row.get('Area Name', '')}</td>"
        html += f"<td>{row.get('ABM Name', '')}</td>"
        html += f"<td style='text-align: center;'>{row.get('Unique TBMs', 0)}</td>"
        html += f"<td style='text-align: center;'>{row.get('Unique HCPs', 0)}</td>"
        html += f"<td style='text-align: center; font-weight: bold; background-color: #f8f9fa;'>{row.get('Requests Raised', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #fff3cd;'>{row.get('Request Cancelled Out of Stock', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #fff3cd;'>{row.get('Action Pending at HO', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #d1ecf1;'>{row.get('Sent to HUB', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #d1ecf1;'>{row.get('Pending for Invoicing', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #d1ecf1;'>{row.get('Pending for Dispatch', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #d4edda;'>{row.get('Requests Dispatched', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #d4edda;'>{row.get('Delivered', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #d4edda;'>{row.get('Dispatched In Transit', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #f8d7da;'>{row.get('RTO', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #f8d7da;'>{row.get('Incomplete Address', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #f8d7da;'>{row.get('Doctor Non Contactable', 0)}</td>"
        html += f"<td style='text-align: center; background-color: #f8d7da;'>{row.get('Doctor Refused to Accept', 0)}</td>"
        html += "</tr>"
    
    # Total row with enhanced styling
    html += "<tr style='background-color: #343a40; color: white; font-weight: bold; text-align: center;'>"
    html += "<td style='font-size: 12px;'>TOTAL</td>"
    html += "<td></td>"
    html += f"<td>{summary_df.get('Unique TBMs', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Unique HCPs', pd.Series()).sum()}</td>"
    html += f"<td style='font-size: 12px;'>{summary_df.get('Requests Raised', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Request Cancelled Out of Stock', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Action Pending at HO', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Sent to HUB', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Pending for Invoicing', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Pending for Dispatch', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Requests Dispatched', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Delivered', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Dispatched In Transit', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('RTO', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Incomplete Address', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Doctor Non Contactable', pd.Series()).sum()}</td>"
    html += f"<td>{summary_df.get('Doctor Refused to Accept', pd.Series()).sum()}</td>"
    html += "</tr>"
    
    html += "</table>"
    
    # Add legend/explanation
    html += "<div style='background-color: #f8f9fa; padding: 10px; border-radius: 5px; font-size: 10px; color: #666;'>"
    html += "<strong>Legend:</strong> "
    html += "<span style='background-color: #fff3cd; padding: 2px 5px; margin: 0 5px;'>HO Section</span> "
    html += "<span style='background-color: #d1ecf1; padding: 2px 5px; margin: 0 5px;'>HUB Section</span> "
    html += "<span style='background-color: #d4edda; padding: 2px 5px; margin: 0 5px;'>Delivery Status</span> "
    html += "<span style='background-color: #f8d7da; padding: 2px 5px; margin: 0 5px;'>RTO Reasons</span>"
    html += "</div>"
    
    html += "</div>"
    
    return html

def display_single_email(outlook, zbm_email, cc_emails, email_content, zbm_code, zbm_name):
    """Display a single email in Outlook for review (without sending)"""
    
    # Create new mail item
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    
    # Set recipients
    mail.To = zbm_email
    
    # Set CC recipients
    if cc_emails:
        mail.CC = cc_emails
    
    # Set subject
    current_date = datetime.now().strftime('%B %d, %Y')
    mail.Subject = f"Sample Direct Dispatch to Doctors - Request Status as of {current_date}"
    
    # Set body
    mail.HTMLBody = email_content
    
    # Add attachment (consolidated file)
    consolidated_file = find_consolidated_file(zbm_code, zbm_name)
    if consolidated_file and os.path.exists(consolidated_file):
        mail.Attachments.Add(consolidated_file)
        print(f"   📎 Attached: {os.path.basename(consolidated_file)}")
    
    # Display email (don't send)
    mail.Display()
    
    print(f"   📧 Email displayed for: {zbm_email}")
    if cc_emails:
        print(f"   📧 CC'd to: {cc_emails}")
    print(f"   ⚠️  Review the email and send manually from Outlook")

def find_consolidated_file(zbm_code, zbm_name):
    """Find the consolidated file for this ZBM"""
    
    # Look for consolidated files in current directory and subdirectories
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.startswith(f"ZBM_Consolidated_{zbm_code}_") and file.endswith('.xlsx'):
                return os.path.join(root, file)
    
    return None

def create_html_email_files(df, zbms):
    """Create HTML email files as fallback when Outlook is not available"""
    
    print("📧 Creating HTML email files...")
    
    # Create output directory
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = f"ZBM_HTML_Emails_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 Created output directory: {output_dir}")
    
    success_count = 0
    
    for _, zbm_row in zbms.iterrows():
        zbm_code = zbm_row['ZBM Terr Code']
        zbm_name = zbm_row['ZBM Name']
        zbm_email = zbm_row['ZBM EMAIL_ID']
        
        print(f"\n🔄 Processing ZBM: {zbm_code} - {zbm_name}")
        
        try:
            # Read the actual ZBM summary report file
            zbm_summary_df = read_zbm_summary_report(zbm_code, zbm_name)
            
            if zbm_summary_df is None or zbm_summary_df.empty:
                print(f"⚠️ No ZBM summary report found for {zbm_code}")
                continue
            
            # Convert summary report data to email format
            summary_data = create_summary_data_from_report(zbm_summary_df)
            summary_df = pd.DataFrame(summary_data)
            
            # Get ABM emails for CC from the original data
            zbm_data = df[df['ZBM Terr Code'] == zbm_code]
            abms = zbm_data.groupby(['ABM Terr Code', 'ABM Name', 'ABM EMAIL_ID']).agg({
                'TBM HQ': 'first'
            }).reset_index()
            
            # Generate email content
            email_content, cc_emails = generate_email_content(zbm_name, zbm_email, abms, summary_df)
            
            # Create HTML email file
            create_single_html_email(zbm_code, zbm_name, zbm_email, cc_emails, email_content, output_dir)
            
            success_count += 1
            print(f"   ✅ HTML email created for {zbm_name}")
            
        except Exception as e:
            print(f"   ❌ Error creating HTML email for {zbm_name}: {e}")
            continue
    
    print(f"\n🎉 HTML email creation completed!")
    print(f"✅ Successfully created: {success_count} HTML email files")
    print(f"📁 Files saved in: {output_dir}")
    print(f"📧 You can open these HTML files in your browser and copy content to Outlook")

def create_single_html_email(zbm_code, zbm_name, zbm_email, cc_emails, email_content, output_dir):
    """Create a single HTML email file"""
    
    current_date = datetime.now().strftime('%B %d, %Y')
    
    # Create full HTML email
    html_email = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Sample Direct Dispatch to Doctors - Request Status as of {current_date}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .total-row {{ background-color: #e0e0e0; font-weight: bold; }}
        .header {{ background-color: #f0f0f0; padding: 10px; margin-bottom: 20px; }}
    </style>
</head>
<body>
    <div class="header">
        <h3>Email Details:</h3>
        <p><strong>To:</strong> {zbm_email}</p>
        <p><strong>CC:</strong> {cc_emails}</p>
        <p><strong>Subject:</strong> Sample Direct Dispatch to Doctors - Request Status as of {current_date}</p>
    </div>
    
    <div class="email-content">
        <p>Hi {zbm_name},</p>
        
        <p>Please refer the status Sample requests raised in Abbworld for your area.</p>
        
        {email_content}
        
        <p>You can track your sample request at the following link with the Docket Number:</p>
        <p>DTDC: <a href="https://www.dtdc.com/tracking">Click here</a></p>
        <p>Speed Post: <a href="https://www.indiapost.gov.in/vas/Pages/IndiaPostHome.aspx">Click Here</a></p>
        
        <p>In case of any query, please contact 1Point.</p>
        
        <p>Regards,<br>Umesh Pawar.</p>
    </div>
</body>
</html>
"""
    
    # Save HTML file
    safe_zbm_name = str(zbm_name).replace(' ', '_').replace('/', '_').replace('\\', '_')
    filename = f"Email_{zbm_code}_{safe_zbm_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    filepath = os.path.join(output_dir, filename)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(html_email)
    
    print(f"   📧 HTML email saved: {filename}")

if __name__ == "__main__":
    send_zbm_emails()
