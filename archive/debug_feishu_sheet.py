#!/usr/bin/env python3
"""
Debug Feishu Sheet Structure
Read the actual sheet to understand the layout
"""

import requests
import json

# Your Feishu credentials
FEISHU_APP_ID = "cli_a86397ecf2f89013"
FEISHU_APP_SECRET = "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt"
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"

SPREADSHEET_TOKEN = "SBYJsQ4KkhQ1Svti88ouShLtsK0"
SHEET_ID = "IqubAk"

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

def read_sheet_data():
    """Read the actual sheet data to understand structure"""
    
    token = get_feishu_access_token()
    print(f"✅ Got access token: {token[:20]}...")
    
    # Read a large range to see the full structure
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AZ200"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print("📖 Reading sheet data...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            values = data.get("data", {}).get("values", [])
            
            print(f"📊 Found {len(values)} rows of data")
            print("\n=== SHEET STRUCTURE ANALYSIS ===")
            
            # Show first 20 rows to understand layout
            for i, row in enumerate(values[:20]):
                if row:
                    # Show first 10 columns to keep it readable
                    display_row = row[:10]
                    print(f"Row {i+1:2d}: {display_row}")
            
            if len(values) > 20:
                print(f"... and {len(values) - 20} more rows")
            
            # Look for patterns
            print("\n=== PATTERN ANALYSIS ===")
            
            # Find week headers
            week_cols = []
            if values and len(values) > 0:
                header_row = values[0]
                for j, cell in enumerate(header_row):
                    if cell and "week" in str(cell).lower():
                        col_letter = column_to_letter(j + 1)
                        week_cols.append((j, col_letter, str(cell)))
            
            print(f"Week columns found: {week_cols}")
            
            # Find campaign rows
            campaign_rows = []
            for i, row in enumerate(values):
                if row and len(row) > 0:
                    first_cell = str(row[0]).strip()
                    if any(keyword in first_cell.lower() for keyword in ['evergreen', 'prospecting', 'retargeting', 'cns', 's+']):
                        campaign_rows.append((i + 1, first_cell[:50]))  # First 50 chars
            
            print(f"Campaign rows found: {len(campaign_rows)}")
            for row_num, campaign in campaign_rows[:10]:  # Show first 10
                print(f"  Row {row_num}: {campaign}")
            
            return values
        else:
            print(f"❌ Error: {data.get('msg')}")
    else:
        print(f"❌ HTTP {response.status_code}: {response.text}")
    
    return None

def column_to_letter(col_num):
    """Convert column number to Excel letter"""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + 65) + letter
        col_num //= 26
    return letter

def main():
    print("🔍 Debugging Feishu Sheet Structure")
    print("="*50)
    
    try:
        values = read_sheet_data()
        
        if values:
            print(f"\n✅ Successfully read sheet structure")
            
            # Save to file for analysis
            with open('feishu_sheet_debug.json', 'w') as f:
                json.dump(values, f, indent=2)
            
            print(f"💾 Full data saved to: feishu_sheet_debug.json")
            
            # Show copy-paste ready format
            print(f"\n📋 If manual copy-paste is needed:")
            print("Week 37 ROAS values in order:")
            print("3.24")  # Overall
            for campaign in [
                "US Evergreen Prospecting (Lowest Cost)",
                "S+ 2.0 US Evergreen Prospecting (Lowest Cost)", 
                "US Evergreen Prospecting (Cost Cap)",
                "US Evergreen Retargeting (Lowest Cost)",
                "Canada Evergreen Prospecting (Lowest Cost)",
                "S+ 2.0 Canada Evergreen Prospecting (Lowest Cost)",
                "Canada Evergreen Retargeting (Lowest Cost)",
                "US CNS S+ 2.0 (carousel ads only - Lowest Cost)",
                "Canada CNS S+ 2.0 (carousel ads only - Lowest Cost)"
            ]:
                print(f"[Find row for: {campaign}]")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()