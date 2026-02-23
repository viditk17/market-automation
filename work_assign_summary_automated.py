"""
OM INSIGHTS - MULTI-REPORT BULK DOWNLOADER (DIRECT API)
pip install requests
"""

import requests
import calendar
import os
import time

print("=" * 70)
print("OM INSIGHTS - MULTI-REPORT BULK DOWNLOADER")
print("=" * 70)

EMAIL = "vidit.kalra@olsc.in"
PASSWORD = "v#1234K%"
DOWNLOAD_DIR = os.path.dirname(os.path.abspath(__file__))

# Report configurations
REPORTS = {
    "1": {
        "name": "0016 - WORK ASSIGN SUMMARY",
        "card_id": "1293",
        "params": [
            {"name": "P_EMP_CODE", "type": "category", "value": "*"},
            {"name": "P_FROM_DT", "type": "date/single"},
            {"name": "P_TO_DT", "type": "date/single"}
        ],
        "filename": "0016_WORK_ASSIGN_SUMMARY",
        "format": "xlsx"
    },
    "2": {
        "name": "0003 - VEHICLE HIRING INCENTIVE",
        "card_id": "1272",
        "params": [
            {"name": "FROM_DT", "type": "date/single"},
            {"name": "TO_DT", "type": "date/single"}
        ],
        "filename": "0003_VEHICLE_HIRING_INCENTIVE",
        "format": "csv"
    },
    "3": {
        "name": "0005 - BROKER MASTER REPORT",
        "card_id": "1274",
        "params": [],
        "filename": "0005_BROKER_MASTER_REPORT",
        "format": "csv"
    }
}

# Input - Select mode
print("\nSelect Mode:")
print("  1. Download Single Report")
print("  2. Download All Reports")
print("  3. Download Multiple Selected Reports")

mode = input("\nEnter choice (1/2/3): ").strip()

selected_reports = []

if mode == "1":
    # Single report
    print("\nAvailable Reports:")
    for key, report in REPORTS.items():
        print(f"  {key}. {report['name']}")
    
    report_choice = input("\nEnter choice (1/2/3): ").strip()
    
    if report_choice not in REPORTS:
        print("❌ Invalid choice!")
        exit()
    
    selected_reports = [report_choice]

elif mode == "2":
    # All reports
    selected_reports = list(REPORTS.keys())
    print("\n✅ All reports selected")

elif mode == "3":
    # Multiple selected
    print("\nAvailable Reports:")
    for key, report in REPORTS.items():
        print(f"  {key}. {report['name']}")
    
    choices = input("\nEnter choices separated by comma (e.g., 1,2): ").strip()
    selected_reports = [c.strip() for c in choices.split(",")]
    
    # Validate
    for choice in selected_reports:
        if choice not in REPORTS:
            print(f"❌ Invalid choice: {choice}")
            exit()
else:
    print("❌ Invalid mode!")
    exit()

# Input - Year and Month (for reports that need dates)
needs_dates = any(REPORTS[r]['params'] for r in selected_reports)

if needs_dates:
    selected_year = int(input("\nYear (YYYY): "))
    selected_month_num = int(input("Month (1-12): "))
    
    last_day = calendar.monthrange(selected_year, selected_month_num)[1]
    from_date = f"{selected_year}-{selected_month_num:02d}-01"
    to_date = f"{selected_year}-{selected_month_num:02d}-{last_day:02d}"
    
    print(f"\n📆 Date Range: {from_date} to {to_date}")

print(f"\n📊 Total reports to download: {len(selected_reports)}")

session = requests.Session()

try:
    # Login
    print("\n🔗 Logging in...")
    login_response = session.post(
        "https://ominsights.omlogistics.co.in/api/session",
        json={"username": EMAIL, "password": PASSWORD},
        headers={"Content-Type": "application/json"}
    )
    
    if login_response.status_code != 200:
        print(f"❌ Login failed: {login_response.status_code}")
        exit()
    
    print("✅ Logged in")
    
    # Download each report
    downloaded_files = []
    
    for idx, report_key in enumerate(selected_reports, 1):
        report = REPORTS[report_key]
        
        print(f"\n[{idx}/{len(selected_reports)}] 💾 Downloading: {report['name']}")
        
        download_url = f"https://ominsights.omlogistics.co.in/api/card/{report['card_id']}/query/{report['format']}"
        
        # Build parameters
        parameters = []
        
        for param in report['params']:
            if param['type'] == 'category':
                parameters.append({
                    "type": param['type'],
                    "target": ["variable", ["template-tag", param['name']]],
                    "value": param['value']
                })
            elif param['type'] == 'date/single':
                # Add FROM_DT
                if 'FROM' in param['name'] or 'P_FROM' in param['name']:
                    parameters.append({
                        "type": param['type'],
                        "target": ["variable", ["template-tag", param['name']]],
                        "value": from_date
                    })
                # Add TO_DT
                elif 'TO' in param['name'] or 'P_TO' in param['name']:
                    parameters.append({
                        "type": param['type'],
                        "target": ["variable", ["template-tag", param['name']]],
                        "value": to_date
                    })
        
        payload = {"parameters": parameters} if parameters else {}
        
        try:
            file_response = session.post(
                download_url, 
                json=payload,
                headers={"Content-Type": "application/json"},
                timeout=120
            )
            
            if file_response.status_code == 200:
                if needs_dates and report['params']:
                    filename = f"{report['filename']}_{from_date}_to_{to_date}.{report['format']}"
                else:
                    filename = f"{report['filename']}.{report['format']}"
                
                file_path = os.path.join(DOWNLOAD_DIR, filename)
                
                with open(file_path, 'wb') as f:
                    f.write(file_response.content)
                
                file_size = os.path.getsize(file_path) / 1024
                
                print(f"   ✅ Downloaded: {filename} ({file_size:.1f} KB)")
                downloaded_files.append(filename)
                
            else:
                print(f"   ❌ Failed: {file_response.status_code}")
                print(f"   Response: {file_response.text[:200]}")
        
        except Exception as e:
            print(f"   ❌ Error: {e}")
        
        # Small delay between downloads
        if idx < len(selected_reports):
            time.sleep(1)
    
    # Summary
    print("\n" + "=" * 70)
    print(f"✅ DOWNLOAD COMPLETED!")
    print(f"📁 Total files downloaded: {len(downloaded_files)}")
    print(f"📂 Location: {DOWNLOAD_DIR}")
    print("\nDownloaded files:")
    for f in downloaded_files:
        print(f"   • {f}")
    print("=" * 70)

except Exception as e:
    print(f"\n❌ ERROR: {e}")

print("\n✅ DONE!")