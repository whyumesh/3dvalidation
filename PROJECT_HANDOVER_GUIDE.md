# ZBM Email Automation Project - Handover Guide

## 📋 Project Overview
This project automates the creation of ZBM (Zone Business Manager) summary reports and consolidated files for email distribution. It processes data from Sample Master Tracker.xlsx and creates hierarchical reports showing ABM (Area Business Manager) performance under each ZBM.

## 🗂️ File Structure & Purpose

### **Core Data Files**
| File Name | Purpose | Usage |
|-----------|---------|-------|
| `Sample Master Tracker.xlsx` | **Primary data source** | Contains all request data, ABM/ZBM mappings, and status information |
| `logic.xlsx` | **Business rules** | Contains status mapping rules to calculate Final Status from Request Status |
| `zbm_summary.xlsx` | **Template file** | Format template for summary reports (headers, styling, structure) |

### **Main Processing Scripts**
| File Name | Purpose | When to Use |
|-----------|---------|-------------|
| `create_zbm_hierarchical_reports.py` | **Creates summary reports** | Generate formatted summary reports for email body |
| `create_zbm_consolidated_files.py` | **Creates detailed files** | Generate detailed consolidated files for email attachments |
| `create_complete_zbm_reports.py` | **Master script** | Run both summary and consolidated files in one go |

### **Legacy/Reference Scripts**
| File Name | Purpose | Status |
|-----------|---------|--------|
| `zbm_summary_automation.py` | Original summary automation | **Legacy** - Use hierarchical version instead |
| `create_zbm_email_drafts.py` | Email draft creation | **Reference** - For email automation |
| `create_zbm_outlook_emails.py` | Outlook email creation | **Reference** - For email automation |
| `create_zbm_email_preview.py` | Email preview generation | **Reference** - For email automation |

### **Documentation Files**
| File Name | Purpose |
|-----------|---------|
| `PROJECT_HANDOVER_GUIDE.md` | This file - Complete project documentation |
| `ZBM_Automation_Documentation.md` | Technical documentation |
| `ZBM_Data_Flow_Diagram.md` | Data flow visualization |
| `Notes.txt` | Project notes and requirements |

## 🚀 How to Use the Project

### **Prerequisites**
1. **Python 3.7+** installed
2. **Required packages**: `pandas`, `openpyxl`
3. **Data files** in the same directory:
   - `Sample Master Tracker.xlsx`
   - `logic.xlsx`
   - `zbm_summary.xlsx`

### **Installation**
```bash
pip install pandas openpyxl
```

### **Quick Start (Recommended)**
```bash
# Generate both summary reports and consolidated files
python create_complete_zbm_reports.py
```

### **Individual Scripts**
```bash
# Generate only summary reports (for email body)
python create_zbm_hierarchical_reports.py

# Generate only consolidated files (for email attachments)
python create_zbm_consolidated_files.py
```

## 📊 Output Files

### **Summary Reports** (Email Body)
- **Location**: `ZBM_Reports_YYYYMMDD_HHMMSS/`
- **Format**: `ZBM_Summary_ZN000881_Manager_Name_YYYYMMDD_HHMMSS.xlsx`
- **Content**: Hierarchical summary with ABM performance metrics
- **Columns**: Area Name, ABM Name, Unique TBMs, Unique HCPs, Request counts, etc.

### **Consolidated Files** (Email Attachments)
- **Location**: `ZBM_Consolidated_Files_YYYYMMDD_HHMMSS/`
- **Format**: `ZBM_Consolidated_ZN000881_Manager_Name_YYYYMMDD_HHMMSS.xlsx`
- **Content**: Detailed request-level data for each ZBM
- **Columns**: Request IDs, Doctor details, Item codes, Status, Dates, etc.

## 📧 Email Workflow

### **For Each ZBM:**
1. **Email Body**: Use data from summary report
2. **Email Attachment**: Use corresponding consolidated file
3. **Recipient**: ZBM email address from the data

### **Email Content Structure:**
```
Subject: ZBM Summary Report - [ZBM Name] - [Date]

Body: [Summary report data in table format]

Attachment: ZBM_Consolidated_[ZBM_CODE]_[NAME].xlsx
```

## 🔧 Configuration & Customization

### **Data Source Changes**
- Update `Sample Master Tracker.xlsx` with new data
- Ensure column names match the required columns in scripts
- Run the processing scripts to generate new reports

### **Column Mapping**
- **Summary Reports**: Edit `create_zbm_hierarchical_reports.py`
- **Consolidated Files**: Edit `create_zbm_consolidated_files.py`
- **Required columns** are defined in the `required_columns` list

### **Business Rules**
- Update `logic.xlsx` to modify status mapping rules
- The `Rules` sheet contains Request Status → Final Status mappings

## 🐛 Troubleshooting

### **Common Issues**

#### **1. Missing Columns Error**
```
❌ Missing required columns in Sample Master Tracker.xlsx: ['Column Name']
```
**Solution**: Check column names in the data file and update the script

#### **2. Empty RTO Reason Column**
**Debug**: The script shows RTO Reason analysis in the output
**Solution**: Check if RTO Reason data exists in the source file

#### **3. Wrong Counts in Summary**
**Debug**: Check the debugging output for actual vs expected counts
**Solution**: Verify data filtering and aggregation logic

#### **4. File Not Found Errors**
**Solution**: Ensure all required files are in the same directory as the scripts

### **Debug Mode**
All scripts include detailed debugging output:
- Column analysis
- Data counts
- Sample data preview
- Error messages with context

## 📈 Data Flow

```
Sample Master Tracker.xlsx
    ↓
[Data Cleaning & Filtering]
    ↓
[Apply Business Rules from logic.xlsx]
    ↓
[Group by ZBM]
    ↓
┌─────────────────┬─────────────────┐
│ Summary Reports │ Consolidated    │
│ (Email Body)    │ Files           │
│                 │ (Attachments)   │
└─────────────────┴─────────────────┘
```

## 🎯 Key Features

### **Summary Reports**
- ✅ Hierarchical structure (ZBM → ABM)
- ✅ Perfect formatting matching template
- ✅ Calculated metrics and tallies
- ✅ ZBM-specific data only

### **Consolidated Files**
- ✅ Detailed request-level data
- ✅ All required columns mapped
- ✅ Final Status calculated from business rules
- ✅ Sorted by ABM and Request ID

### **Automation Features**
- ✅ Batch processing for all ZBMs
- ✅ Automatic file naming with timestamps
- ✅ Error handling and validation
- ✅ Debugging and logging

## 📞 Support & Maintenance

### **Regular Tasks**
1. **Update data**: Replace `Sample Master Tracker.xlsx` with new data
2. **Run scripts**: Execute `create_complete_zbm_reports.py`
3. **Distribute reports**: Send emails with generated files

### **Monitoring**
- Check console output for errors
- Verify file counts match ZBM count
- Validate data in generated files

### **Updates**
- Modify column mappings in scripts as needed
- Update business rules in `logic.xlsx`
- Adjust formatting in template file

## 🔒 Security Notes

- **Data Privacy**: Ensure sensitive data is handled according to company policies
- **File Access**: Limit access to authorized personnel only
- **Backup**: Keep backups of data files and generated reports

## 📝 Version History

- **v1.0**: Initial hierarchical reporting system
- **v1.1**: Added consolidated files generation
- **v1.2**: Enhanced debugging and error handling
- **v1.3**: Fixed column mapping and data validation

---

**Last Updated**: [Current Date]
**Maintained By**: [Your Name]
**Contact**: [Your Email]
