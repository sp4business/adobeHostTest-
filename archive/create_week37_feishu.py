#!/usr/bin/env python3
"""
Create Week 37 ROAS data in Feishu sheet
Since the sheet appears empty, let's create the structure and populate it
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

def create_week37_structure():
    """Create the Week 37 ROAS structure in the sheet"""
    
    token = get_feishu_access_token()
    print(f"✅ Got access token")
    
    # Week 37 ROAS data structure
    week37_data = [
        ["Campaign Name", "Week 37"],  # Headers
        ["Week 37 Overall ROAS", 3.24],
        ["", ""],  # Empty row
        ["US Evergreen Prospecting (Lowest Cost)", 3.56],
        ["S+ 2.0 US Evergreen Prospecting (Lowest Cost)", 2.76],
        ["US Evergreen Prospecting (Cost Cap)", 3.03],
        ["US Evergreen Retargeting (Lowest Cost)", 3.01],
        ["Canada Evergreen Prospecting (Lowest Cost)", 2.86],
        ["S+ 2.0 Canada Evergreen Prospecting (Lowest Cost)", 0.45],
        ["Canada Evergreen Retargeting (Lowest Cost)", 2.4],
        ["US CNS S+ 2.0 (carousel ads only - Lowest Cost)", 3.25],
        ["Canada CNS S+ 2.0 (carousel ads only - Lowest Cost)", 3.59]
    ]
    
    print(f"📝 Creating Week 37 structure with {len(week37_data)} rows...")
    
    # Write data to sheet
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    
    payload = {
        "valueRange": {
            "range": f"{SHEET_ID}!A1:B{len(week37_data)}",
            "values": week37_data
        }
    }
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.put(url, json=payload, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            print("✅ Successfully created Week 37 ROAS structure!")
            print(f"📊 Written {len(week37_data)} rows to sheet")
            print(f"📋 Check your sheet: https://bytedance.us.larkoffice.com/sheets/{SPREADSHEET_TOKEN}")
            return True
        else:
            print(f"❌ Write error: {data.get('msg')}")
    else:
        print(f"❌ HTTP {response.status_code}: {response.text}")
    
    return False

def test_simple_write():
    """Test with a simple write first"""
    
    token = get_feishu_access_token()
    print(f"✅ Got access token")
    
    # Simple test data
    test_data = [
        ["Test", "Week 37"],
        ["Overall ROAS", 3.24]
    ]
    
    print(f"📝 Testing simple write...")
    
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    
    payload = {
        "valueRange": {
            "range": f"{SHEET_ID}!A1:B2",
            "values": test_data
        }
    }
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.put(url, json=payload, headers=headers)
    
    print(f"Status: {response.status_code}")
    print(f"Response: {response.text}")
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            print("✅ Simple write successful!")
            return True
        else:
            print(f"❌ Write error: {data.get('msg')}")
    
    return False

def main():
    print("🚀 Creating Week 37 ROAS Data in Feishu")
    print("="*50)
    
    try:
        # First try simple test
        if test_simple_write():
            # Then create full structure
            create_week37_structure()
        else:
            print("❌ Simple write failed - checking permissions...")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()