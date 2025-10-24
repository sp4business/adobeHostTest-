#!/usr/bin/env python3
"""
Simple Feishu Write Test
Test basic write operations to understand what's possible
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

def test_simple_write():
    """Test simple write operation"""
    
    token = get_feishu_access_token()
    print(f"✅ Got token: {token[:20]}...")
    
    # Test 1: Write to a simple cell
    print("\n📝 Testing simple cell write...")
    
    # Try writing "3.24" to cell A1 as a test
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    
    payload = {
        "valueRange": {
            "range": f"{SHEET_ID}!A1:A1",
            "values": [["3.24"]]
        }
    }
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.put(url, json=payload, headers=headers)
    
    print(f"Status: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        print(f"Response: {json.dumps(data, indent=2)}")
        if data.get("code") == 0:
            print("✅ Simple write successful!")
        else:
            print(f"❌ Write error: {data.get('msg')}")
    else:
        print(f"❌ HTTP error: {response.text}")
    
    # Test 2: Try to get sheet metadata
    print("\n📋 Testing sheet metadata...")
    
    meta_url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/metainfo"
    meta_response = requests.get(meta_url, headers=headers)
    
    print(f"Metadata status: {meta_response.status_code}")
    if meta_response.status_code == 200:
        meta_data = meta_response.json()
        print(f"Metadata: {json.dumps(meta_data, indent=2)}")
    else:
        print(f"Metadata error: {meta_response.text}")

def main():
    print("🧪 Feishu API Write Test")
    print("="*50)
    
    try:
        test_simple_write()
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()