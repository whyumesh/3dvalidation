# ZBM Automation - Quick Reference Card

## 🚀 Daily Usage

### **Step 1: Prepare Data**
- Ensure `Sample Master Tracker.xlsx` is updated with latest data
- Verify `logic.xlsx` and `zbm_summary.xlsx` are present

### **Step 2: Generate Reports**
```bash
python create_complete_zbm_reports.py
```

### **Step 3: Check Output**
- Summary reports: `ZBM_Reports_YYYYMMDD_HHMMSS/`
- Consolidated files: `ZBM_Consolidated_Files_YYYYMMDD_HHMMSS/`

## 📁 File Locations

| What You Need | Where to Find |
|---------------|---------------|
| **Summary Reports** | `ZBM_Reports_YYYYMMDD_HHMMSS/` |
| **Consolidated Files** | `ZBM_Consolidated_Files_YYYYMMDD_HHMMSS/` |
| **Data Source** | `Sample Master Tracker.xlsx` |
| **Business Rules** | `logic.xlsx` |
| **Template** | `zbm_summary.xlsx` |

## 🔧 Common Commands

```bash
# Generate everything
python create_complete_zbm_reports.py

# Generate only summaries
python create_zbm_hierarchical_reports.py

# Generate only consolidated files
python create_zbm_consolidated_files.py
```

## ⚠️ Troubleshooting

| Error | Solution |
|-------|----------|
| Missing columns | Check column names in Sample Master Tracker.xlsx |
| Empty RTO Reason | Check if data exists in source file |
| Wrong counts | Check debugging output in console |
| File not found | Ensure all files are in same directory |

## 📧 Email Workflow

1. **For each ZBM**:
   - Use summary report data in email body
   - Attach corresponding consolidated file
   - Send to ZBM email address

2. **File naming**:
   - Summary: `ZBM_Summary_ZN000881_Manager_Name_YYYYMMDD_HHMMSS.xlsx`
   - Consolidated: `ZBM_Consolidated_ZN000881_Manager_Name_YYYYMMDD_HHMMSS.xlsx`

## 📊 What Each Report Contains

### **Summary Report** (Email Body)
- Area Name, ABM Name
- Unique TBMs, Unique HCPs
- Request counts by status
- Performance metrics

### **Consolidated File** (Attachment)
- Detailed request data
- Doctor information
- Item codes and quantities
- Status and dates
- RTO reasons

## 🎯 Key Points

- ✅ Each ZBM gets their own summary + consolidated file
- ✅ Only ZBM codes starting with "ZN" are processed
- ✅ All data is ZBM-specific (no cross-contamination)
- ✅ Perfect tallies and formatting maintained
- ✅ Ready for email distribution

---
**Need Help?** Check `PROJECT_HANDOVER_GUIDE.md` for detailed documentation
