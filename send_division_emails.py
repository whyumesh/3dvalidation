# ```python
import pandas as pd
import os
from jinja2 import Environment, FileSystemLoader
import win32com.client as win32
from datetime import datetime as dt
from openpyxl import load_workbook
import glob

# Load Jinja2 Template for email
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("email_template_Division.html")

# Get current date
z = dt.today()
current_date = z.date()

# Find the most recent Division folders
def find_latest_folder(pattern):
    """Find the most recently created folder matching the pattern"""
    folders = glob.glob(pattern)
    if not folders:
        return None
    return max(folders, key=os.path.getctime)

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))
if not current_dir:
    current_dir = os.getcwd()

print(f"📂 Working directory: {current_dir}")

# Locate the generated folders
consolidated_folder = find_latest_folder(os.path.join(current_dir, "Division_Consolidated_Files_*"))
reports_folder = find_latest_folder(os.path.join(current_dir, "Division_Reports_*"))

if not consolidated_folder:
    print("❌ Error: Could not find Division_Consolidated_Files folder. Please run create_division_consolidated_files.py first.")
    exit(1)

if not reports_folder:
    print("❌ Error: Could not find Division_Reports folder. Please run create_division_hierarchical_reports.py first.")
    exit(1)

print(f"✅ Found consolidated files folder: {consolidated_folder}")
print(f"✅ Found reports folder: {reports_folder}")

# Debug: List files in consolidated folder
print(f"\n🔍 Files in consolidated folder:")
if consolidated_folder and os.path.exists(consolidated_folder):
    files = os.listdir(consolidated_folder)
    print(f"   Total files: {len(files)}")
    for f in files[:5]:  # Show first 5 files
        print(f"   - {f}")
else:
    print("   Folder not accessible!")

# Debug: List files in reports folder
print(f"\n🔍 Files in reports folder:")
if reports_folder and os.path.exists(reports_folder):
    files = os.listdir(reports_folder)
    print(f"   Total files: {len(files)}")
    for f in files[:5]:  # Show first 5 files
        print(f"   - {f}")
else:
    print("   Folder not accessible!")

# Read Division Email Mapping file
print("📖 Reading Division Email Mapping file...")
try:
    # You need to provide the filename of your division email mapping file
    # Replace 'division_emails.xlsx' with your actual filename
    division_emails_df = pd.read_excel('division_emails.xlsx')
    print(f"✅ Successfully loaded division email mapping")
    print(f"📋 Columns in file: {list(division_emails_df.columns)}")
except Exception as e:
    print(f"❌ Error reading division email mapping file: {e}")
    print("Please ensure you have a file named 'division_emails.xlsx' with columns 'Division Name' and 'Email'")
    exit(1)

# Read ZBM Automation Email file to get division details
print("📖 Reading ZBM Automation Email 2410252.xlsx...")
df = pd.read_excel('ZBM Automation Email 2410252.xlsx')

# Get unique Divisions with their details
divisions = df.groupby('TBM Division').agg({
    'AFFILIATE': 'first',
    'DIV_NAME': 'first'
}).reset_index().sort_values('TBM Division')

print(f"📋 Found {len(divisions)} unique Divisions to process")

# Initialize Outlook
outlook = win32.Dispatch("Outlook.Application")

# Create output directory for sent email logs
output_dir = os.path.dirname(os.path.abspath(__file__))
email_log_folder = os.path.join(output_dir, f'Division_Email_Logs_{current_date}')
os.makedirs(email_log_folder, exist_ok=True)

def read_summary_report(div_code, div_name):
    """Read the summary report Excel file for a Division and extract data as HTML table with proper formatting"""
    try:
        # Find the summary report file
        pattern = os.path.join(reports_folder, f"Division_Summary_{div_code}_*.xlsx")
        files = glob.glob(pattern)
        
        if not files:
            print(f"   ⚠️ Warning: No summary report found for Division {div_code}")
            print(f"      Searched pattern: {pattern}")
            return None
        
        report_file = os.path.abspath(files[0])
        
        # Verify file exists
        if not os.path.exists(report_file):
            print(f"   ❌ Summary report file does not exist: {report_file}")
            return None
        
        print(f"   📊 Reading summary report: {os.path.basename(report_file)}")
        
        # Read the Excel file
        wb = load_workbook(report_file)
        ws = wb.active  # Use active sheet
        
        # Find header row and starting column (looking for "Affiliate")
        header_row = None
        start_col = None
        for row_idx in range(1, 15):
            for col_idx in range(1, 20):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value and 'Affiliate' in str(cell_value):
                    header_row = row_idx
                    start_col = col_idx
                    break
            if header_row:
                break

        if not header_row or not start_col:
            print(f"   ⚠️ Warning: Could not find header row in summary report")
            return None

        # Read headers starting from start_col
        headers = []
        header_colors = []
        for col_idx in range(start_col, ws.max_column + 1):
            cell = ws.cell(row=header_row, column=col_idx)
            header_val = cell.value
            if header_val is None or str(header_val).strip() == "":
                break
            headers.append(str(header_val).strip())
            # Get background color
            if cell.fill.start_color and cell.fill.start_color.rgb:
                rgb_value = cell.fill.start_color.rgb
                # Ensure it's a string
                if isinstance(rgb_value, str):
                    header_colors.append(rgb_value)
                else:
                    header_colors.append(None)
            else:
                header_colors.append(None)

        # Build merged cells map
        merged_cells_info = {}
        for merged_range in ws.merged_cells.ranges:
            min_row = merged_range.min_row
            max_row = merged_range.max_row
            min_col = merged_range.min_col
            max_col = merged_range.max_col
            
            # Store merge dimensions for top-left cell
            if min_row >= header_row and min_row <= ws.max_row:
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        if r != min_row or c != min_col:
                            # Mark as merged cell (not top-left)
                            merged_cells_info[(r, c)] = None
        
        # Read all data rows from Excel
        rows_data = []
        empty_row_count = 0
        
        for row_idx in range(header_row + 1, ws.max_row + 1):
            has_any_value = False
            is_total_row = False
            
            # Check if this is a "Total" row by looking at first column
            first_cell_value = ws.cell(row=row_idx, column=start_col).value
            if first_cell_value and str(first_cell_value).strip().lower() == 'total':
                is_total_row = True
            
            # Check if row has any value
            for col_idx in range(start_col, start_col + len(headers)):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and str(cell_value).strip() != "":
                    has_any_value = True
                    break
            
            if has_any_value:
                rows_data.append({
                    'row_idx': row_idx,
                    'is_total': is_total_row
                })
                empty_row_count = 0
            else:
                empty_row_count += 1
                if empty_row_count >= 2:
                    break
        
        # Build HTML table with matching Excel formatting
        html_table = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">\n'
        
        # Add header row
        html_table += '  <thead>\n    <tr style="background-color: #D3D3D3; font-weight: bold; text-align: center;">\n'
        for i, header in enumerate(headers):
            bg_color = ""
            if header_colors[i] and isinstance(header_colors[i], str):
                # Convert hex color if needed
                hex_color = header_colors[i]
                if hex_color.startswith('FF'):
                    hex_color = '#' + hex_color[2:]
                bg_color = f' background-color: {hex_color};'
            html_table += f'      <th style="{bg_color} padding: 8px; border: 1px solid #000;">{header}</th>\n'
        html_table += '    </tr>\n  </thead>\n'
        
        # Add data rows with merged cell handling
        html_table += '  <tbody>\n'
        
        for row_info in rows_data:
            row_idx = row_info['row_idx']
            is_total = row_info['is_total']
            
            row_style = 'font-weight: bold; background-color: #E6E6E6;' if is_total else ''
            html_table += f'    <tr style="{row_style}">\n'
            
            for col_idx in range(start_col, start_col + len(headers)):
                # Skip if this cell is part of a merge (not the top-left cell)
                if (row_idx, col_idx) in merged_cells_info:
                    continue
                
                # Check if this cell starts a merge
                colspan = 1
                rowspan = 1
                
                for merged_range in ws.merged_cells.ranges:
                    if row_idx == merged_range.min_row and col_idx == merged_range.min_col:
                        rowspan = merged_range.max_row - merged_range.min_row + 1
                        colspan = merged_range.max_col - merged_range.min_col + 1
                        break
                
                # Get cell value directly from Excel
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is None or pd.isna(cell_value):
                    value = '-'
                else:
                    value = str(cell_value).strip()
                
                # Add merge attributes if needed
                merge_attr = ''
                if rowspan > 1:
                    merge_attr += f' rowspan="{rowspan}"'
                if colspan > 1:
                    merge_attr += f' colspan="{colspan}"'
                
                html_table += f'      <td style="padding: 5px; border: 1px solid #000; text-align: center;"{merge_attr}>{value}</td>\n'
            
            html_table += '    </tr>\n'
        
        html_table += '  </tbody>\n</table>'
        
        return html_table
        
    except Exception as e:
        print(f"   ❌ Error reading summary report for Division {div_code}: {e}")
        import traceback
        traceback.print_exc()
        return None

def get_abm_emails_for_division(div_code):
    """Get all ABM email addresses under a specific Division for CC"""
    div_data = df[df['TBM Division'] == div_code]
    abm_emails = div_data['ABM EMAIL_ID'].dropna().unique()
    
    # Filter out invalid emails (0, '0', empty strings)
    abm_emails = [email for email in abm_emails if email and str(email) not in ['0', '0.0']]
    
    return '; '.join(abm_emails)

# Process each Division and send emails
email_count = 0
for _, div_row in divisions.iterrows():
    div_code = div_row['TBM Division']
    affiliate = div_row['AFFILIATE']
    div_name = div_row['DIV_NAME']
    
    print(f"\n🔄 Processing Division: {div_code} - {affiliate} - {div_name}")
    
    # Find email for this division
    div_email_row = division_emails_df[division_emails_df['Division Name'] == div_code]
    
    if div_email_row.empty:
        print(f"   ⚠️ No email found for Division {div_code}")
        continue
    
    div_email = div_email_row['Email'].iloc[0]
    
    # Skip if no valid email
    if not div_email or str(div_email) in ['0', '0.0']:
        print(f"   ⚠️ Skipping - No valid email address")
        continue
    
    print(f"   📧 Email: {div_email}")
    
    # Find consolidated file for this Division
    safe_div_name = str(div_name).replace(' ', '_').replace('/', '_').replace('\\', '_')
    consolidated_pattern = os.path.join(consolidated_folder, f"Division_Consolidated_{div_code}_*.xlsx")
    consolidated_files = glob.glob(consolidated_pattern)
    
    if not consolidated_files:
        print(f"   ⚠️ No consolidated file found for Division {div_code}")
        print(f"      Searched pattern: {consolidated_pattern}")
        continue
    
    consolidated_file = consolidated_files[0]
    
    # Convert to absolute path (Outlook requires absolute paths)
    consolidated_file = os.path.abspath(consolidated_file)
    
    # Verify file exists
    if not os.path.exists(consolidated_file):
        print(f"   ❌ File does not exist: {consolidated_file}")
        continue
    
    print(f"   📎 Attaching: {os.path.basename(consolidated_file)}")
    print(f"      Full path: {consolidated_file}")
    
    # Read summary report data
    summary_html = read_summary_report(div_code, div_name)
    
    if not summary_html:
        print(f"   ⚠️ No summary report data found for Division {div_code}")
        continue
    
    # Get ABM emails for CC
    abm_cc_emails = get_abm_emails_for_division(div_code)
    
    # Build CC list based on affiliate
    cc_list = []
    
    # Add affiliate-specific emails
    if affiliate == 'AIL':
        cc_list.extend(['ishan.mithbavkar@abbott.com', 'ashwini.suryavanshi@abbott.com'])
    elif affiliate == 'APC':
        cc_list.extend(['jenita.nadar@abbott.com', 'ashwini.suryavanshi@abbott.com'])
    elif affiliate == 'ASC':
        cc_list.extend(['sandesh.bhoir@abbott.com', 'ashwini.suryavanshi@abbott.com'])
    
    # Always add sandesh.bhoir@abbott.com if not already added
    if 'sandesh.bhoir@abbott.com' not in cc_list:
        cc_list.append('sandesh.bhoir@abbott.com')
    
    # Add ABM emails
    if abm_cc_emails:
        cc_list.extend(abm_cc_emails.split('; '))
    
    # Remove duplicates and join
    cc_list = list(set(cc_list))
    final_cc = '; '.join(cc_list)
    
    # Create email
    try:
        mail = outlook.CreateItem(0)
        mail.To = div_email
        
        # Add CC recipients
        if final_cc:
            mail.CC = final_cc
            print(f"   📧 CC: {final_cc}")
        
        # Set BCC
        mail.BCC = 'vaibhav.nalawade@abbott.com;kranti.vengurlekar@abbott.com'
        
        # Set subject
        mail.Subject = f"Sample Direct Dispatch - Division Summary Report as of {current_date}"
        
        # Render email body with summary table
        mail.HTMLBody = template.render(
            division_name=div_name,
            division_code=div_code,
            affiliate=affiliate,
            current_date=current_date,
            summary_table=summary_html
        )
        
        # Set sender
        mail.SentOnBehalfOfName = 'EPD_SFA@abbott.com'
        
        # Attach consolidated file - AFTER setting body
        try:
            mail.Attachments.Add(consolidated_file)
            print(f"   ✅ Attachment added successfully")
        except Exception as attach_error:
            print(f"   ❌ Error attaching file: {attach_error}")
            print(f"      File path: {consolidated_file}")
            continue
        
        # Display email
        mail.Display()
        
        email_count += 1
        print(f"   ✅ Email displayed successfully for {div_email}")
        
        # Log the sent email
        with open(os.path.join(email_log_folder, 'email_log.txt'), 'a') as log:
            log.write(f"{dt.now()} - Displayed email for Division {div_code} ({div_name}) - {div_email}\n")
            log.write(f"   CC: {final_cc}\n")
            log.write(f"   BCC: vaibhav.nalawade@abbott.com;kranti.vengurlekar@abbott.com\n\n")
        
    except Exception as e:
        print(f"   ❌ Error creating email for Division {div_code}: {e}")
        import traceback
        traceback.print_exc()
        continue

print(f"\n🎉 Email automation completed!")
print(f"📊 Total emails displayed: {email_count} out of {len(divisions)} Divisions")
print(f"📁 Email logs saved in: {email_log_folder}")
