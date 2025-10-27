# Safe First Run Guide - ZBM Email System

## ⚠️ IMPORTANT: Before Sending to All ZBMs

### STEP 1: Generate Fresh Reports First
```bash
python create_complete_zbm_reports.py
```

Wait for completion. Verify:
- ZBM_Reports_[date] folder created
- ZBM_Consolidated_Files_[date] folder created
- Same number of files in both folders

### STEP 2: Verify Data with ONE ZBM (CRITICAL TEST)

1. **Modify send_zbm_emails.py** - Line 272, change to:
```python
# Change this line:
for _, zbm_row in zbms.iterrows():

# To this (test with only first ZBM):
for _, zbm_row in zbms.head(1).iterrows():
```

2. **Run the script**:
```bash
python send_zbm_emails.py
```

3. **REVIEW THE EMAIL CAREFULLY**:
   - ✅ Does the name match the ZBM?
   - ✅ Is the summary table readable?
   - ✅ Are merged cells displaying correctly?
   - ✅ Is the correct consolidated file attached?
   - ✅ Does the attachment open correctly?

4. **Manually verify data accuracy**:
   - Open the attached consolidated file
   - Pick 5 random rows
   - Find those same records in Sample Master Tracker.xlsx
   - Verify they match exactly
   - Verify they belong to the same ZBM code

5. **Verify summary totals**:
   - Open the attached consolidated file
   - Count records manually (or use Excel filter)
   - Compare with summary table totals in email body
   - They should be reasonably close (may differ slightly due to Final Status logic)

### STEP 3: Test with 3-5 ZBMs

1. **Modify send_zbm_emails.py** - Line 272:
```python
for _, zbm_row in zbms.head(5).iterrows():  # Test with 5 ZBMs
```

2. **Run and review all 5 emails**

3. **Manual spot check each one**

### STEP 4: Full Production Run

Only after Steps 1-3 pass perfectly:
```python
# Change back to original:
for _, zbm_row in zbms.iterrows():
```

## 🔍 What to Check During Manual Verification

### For Each Email:

1. **ZBM Code/Name Match**:
   - Email subject should have correct ZBM code
   - Email body should address correct ZBM name
   - From: "Hi [ZBM Name]"

2. **Attachment Verification**:
   - Filename should contain ZBM code
   - Open attachment
   - Check ABM Terr Code column - should only have ABMs under this ZBM
   - Check ZBM Terr Code - should all be the same code

3. **Summary Table Verification**:
   - Area Names should match ABM codes under this ZBM
   - Total row should be bold
   - Table should have proper borders

4. **Data Count Check**:
   - Count rows in consolidated file
   - Check "Requests Raised" total in summary table
   - Should be in same ballpark (may differ due to Final Status logic)

## 🚨 Stop Conditions (DO NOT PROCEED IF):

- ❌ Email shows wrong ZBM name or code
- ❌ Attachment contains data from other ZBMs
- ❌ Summary table numbers don't make sense
- ❌ Any ZBM receives someone else's data
- ❌ More than 5% discrepancy between attachment and summary

## 📋 Final Checklist Before Full Run

- [ ] All tests with 1-5 ZBMs passed
- [ ] Manual verification completed for at least 3 emails
- [ ] Data accuracy confirmed
- [ ] Email formatting looks correct
- [ ] You have backup of Sample Master Tracker.xlsx
- [ ] You have generated fresh reports folder timestamp
- [ ] Ready to proceed with full run

## 🎯 Expected Outcome

After full run:
- Each ZBM receives exactly ONE email
- Email contains their own summary table
- Email has their own consolidated file attached
- No duplicate emails to same ZBM
- No emails with wrong recipient data

## 💡 Pro Tips

1. **Log Files**: Check `ZBM_Email_Logs_[date]` folder after run
2. **Monitor First Batch**: Watch first 10 emails before letting it continue
3. **Export Log**: Save the email_log.txt for record keeping
4. **Backup Reports**: Keep the generated folders as audit trail

## 🆘 If Something Goes Wrong

1. **STOP THE SCRIPT IMMEDIATELY** (Ctrl+C)
2. Check the log file to see how many emails were sent
3. Review the consolidated files and reports folders
4. If data mismatch found, DO NOT continue
5. Contact support or review the source data


