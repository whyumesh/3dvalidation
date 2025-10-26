#!/usr/bin/env python3
"""
Complete ZBM Reports Generator
Creates both summary reports and consolidated files for each ZBM
"""

import subprocess
import sys
import os
from datetime import datetime

def run_script(script_name, description):
    """Run a Python script and handle errors"""
    print(f"\n{'='*60}")
    print(f"🚀 {description}")
    print(f"{'='*60}")
    
    try:
        result = subprocess.run([sys.executable, script_name], 
                              capture_output=True, text=True, check=True)
        print(result.stdout)
        if result.stderr:
            print("Warnings/Errors:")
            print(result.stderr)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error running {script_name}:")
        print(f"Return code: {e.returncode}")
        print(f"STDOUT: {e.stdout}")
        print(f"STDERR: {e.stderr}")
        return False

def create_complete_zbm_reports():
    """Create both summary reports and consolidated files for all ZBMs"""
    
    print("🎯 COMPLETE ZBM REPORTS GENERATOR")
    print("="*60)
    print("This will create:")
    print("1. 📊 Summary reports for each ZBM (for email body)")
    print("2. 📁 Consolidated files for each ZBM (for email attachment)")
    print("="*60)
    
    # Check if required files exist
    required_files = [
        'Sample Master Tracker.xlsx',
        'logic.xlsx',
        'zbm_summary.xlsx'
    ]
    
    missing_files = [f for f in required_files if not os.path.exists(f)]
    if missing_files:
        print(f"❌ Missing required files: {missing_files}")
        return
    
    print("✅ All required files found")
    
    # Step 1: Create summary reports
    success1 = run_script('create_zbm_hierarchical_reports.py', 
                         'Creating ZBM Summary Reports (for email body)')
    
    if not success1:
        print("❌ Failed to create summary reports. Stopping.")
        return
    
    # Step 2: Create consolidated files
    success2 = run_script('create_zbm_consolidated_files.py', 
                         'Creating ZBM Consolidated Files (for email attachments)')
    
    if not success2:
        print("❌ Failed to create consolidated files. Stopping.")
        return
    
    # Summary
    print(f"\n{'='*60}")
    print("🎉 COMPLETE ZBM REPORTS GENERATION FINISHED!")
    print(f"{'='*60}")
    print("✅ Summary reports created (ready for email body)")
    print("✅ Consolidated files created (ready for email attachments)")
    print("\n📧 Next steps:")
    print("1. Use summary reports in email body")
    print("2. Attach corresponding consolidated file to each ZBM email")
    print("3. Each ZBM gets their own summary + consolidated file")

if __name__ == "__main__":
    create_complete_zbm_reports()
