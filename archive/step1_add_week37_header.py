#!/usr/bin/env python3
"""
Step 1: Add Week 37 column header only
Brick by brick approach
"""

import requests
import json

# Your Feishu credentials
FEISHU_APP_ID = "cli_a86397ecf2f89013"
FEISHU_APP_SECRET = "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt"
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"

SPREADSHEET_TOKEN = "SBYJsQ4KkhQ1Svti88ouShLtsK0"
SHEET_ID = "IqubAk"  # "Lululemon Adobe Data 2025"

def get_feishu_access_token():
    """Get Feishu access token"""
    url = f"{FEISHU_BASE_URL}/auth/v3/tenant_access_token/internal"
    
    payload = {
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    }
    
    response = requests.post(url, json=payload)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            return data["tenant_access_token"]
        else:
            raise Exception(f"Failed to get token: {data.get('msg')}")
    else:
        raise Exception(f"HTTP {response.status_code}: {response.text}")

def add_week37_column():
    """Add Week 37 column header"""
    
    token = get_feishu_access_token()
    print(f"✅ Got access token")
    
    # First, let's see where to add the column
    # Read current headers to find the right position
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:Z1"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            values = data.get("data", {}).get("values", [])
            
            if values and len(values) > 0:
                header_row = values[0]
                print(f"Current headers: {header_row}")
                
                # Find where to insert Week 37 (after the last week)
                last_week_col = len(header_row)
                print(f"Adding Week 37 at column {last_week_col + 1}")
                
                # Write "Week 37" to the appropriate header cell
                target_cell = f"{SHEET_ID}!{column_to_letter(last_week_col + 1)}1"
                
                # Use the values API to write the header
                write_url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
                
                write_payload = {
                    "valueRange": {
                        "range": target_cell,
                        "values": [["Week 37"]]
                    }
                }
                
                write_response = requests.put(write_url, json=write_payload, headers=headers)
                
                if write_response.status_code == 200:
                    write_data = write_response.json()
                    if write_data.get("code") == 0:
                        print("✅ Successfully added Week 37 column header!")
                        return True
                    else:
                        print(f"❌ Write error: {write_data.get('msg')}")
                else:
                    print(f"❌ Write HTTP {write_response.status_code}: {write_response.text}")
            else:
                print("❌ No header row found")
        else:
            print(f"❌ Read error: {data.get('msg')}")
    else:
        print(f"❌ HTTP {response.status_code}: {response.text}")
    
    return False

def column_to_letter(col_num):
    """Convert column number to Excel letter"""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + 65) + letter
        col_num //= 26
    return letter

def main():
    print("🧱 Step 1: Add Week 37 Column Header")
    print("="*50)
    
    try:
        success = add_week37_column()
        
        if success:
            print("\n🎉 Step 1 complete!")
            print("Next: Step 2 - Add campaign data rows")
        else:
            print("\n❌ Step 1 failed")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    main()