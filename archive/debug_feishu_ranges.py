#!/usr/bin/env python3
"""
Debug Feishu sheet - try different ranges
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

def try_different_ranges():
    """Try reading different ranges to find the data"""
    
    token = get_feishu_access_token()
    print(f"✅ Got access token")
    
    ranges_to_try = [
        f"{SHEET_ID}!A1:Z10",
        f"{SHEET_ID}!A1:D20", 
        f"{SHEET_ID}!B2:AA50",
        f"{SHEET_ID}!A:A",
        f"{SHEET_ID}!1:10"
    ]
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    for range_str in ranges_to_try:
        print(f"\n📖 Trying range: {range_str}")
        
        url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{range_str}"
        
        response = requests.get(url, headers=headers)
        
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            if data.get("code") == 0:
                values = data.get("data", {}).get("values", [])
                print(f"Found {len(values)} rows")
                
                if values:
                    print("Sample data:")
                    for i, row in enumerate(values[:5]):  # First 5 rows
                        print(f"  Row {i+1}: {row}")
                    
                    # Save this working range
                    return range_str, values
            else:
                print(f"Error: {data.get('msg')}")
        else:
            print(f"HTTP Error: {response.text}")
    
    return None, []

def main():
    print("🔍 Debugging Feishu Sheet Ranges")
    print("="*50)
    
    try:
        working_range, data = try_different_ranges()
        
        if working_range and data:
            print(f"\n✅ Found data using range: {working_range}")
            print(f"📊 Total rows: {len(data)}")
            
            # Analyze the data structure
            print(f"\n📋 Data structure analysis:")
            if data:
                print(f"First row: {data[0]}")
                print(f"Row count: {len(data)}")
                if len(data) > 1:
                    print(f"Second row: {data[1]}")
            
            # Save for reference
            with open('feishu_sheet_structure.json', 'w') as f:
                json.dump({
                    'working_range': working_range,
                    'total_rows': len(data),
                    'sample_data': data[:20]  # First 20 rows
                }, f, indent=2)
            
            print(f"\n💾 Saved structure to feishu_sheet_structure.json")
            
        else:
            print("❌ No data found in any range")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()