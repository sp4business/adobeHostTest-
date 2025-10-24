#!/usr/bin/env python3
"""
Feishu Sheet Analyzer - CLEANED VERSION
Analyzes the Feishu sheet structure to understand campaign layout and week columns
NOTE: This is a cleaned version - API credentials should be loaded from config file
"""

import requests
import json
import os
from typing import Optional, Dict, List

# Configuration placeholder - load from config file
FEISHU_CONFIG_FILE = "feishu_config.json"
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"

def load_feishu_config() -> Optional[Dict]:
    """Load Feishu configuration from file"""
    if os.path.exists(FEISHU_CONFIG_FILE):
        try:
            with open(FEISHU_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️  Error loading config: {e}")
    return None

def get_feishu_access_token():
    """Get Feishu access token - requires config file"""
    config = load_feishu_config()
    if not config or 'app_id' not in config or 'app_secret' not in config:
        print("❌ Feishu config not found. Please create feishu_config.json with app_id and app_secret")
        return None
    
    url = f"{FEISHU_BASE_URL}/auth/v3/tenant_access_token/internal"
    
    payload = {
        "app_id": config['app_id'],
        "app_secret": config['app_secret']
    }
    
    print("🔑 Getting Feishu access token...")
    response = requests.post(url, json=payload)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            print("✅ Successfully obtained Feishu access token")
            return data["tenant_access_token"]
        else:
            raise Exception(f"Failed to get token: {data.get('msg')}")
    else:
        raise Exception(f"HTTP {response.status_code}: {response.text}")

def get_sheet_data(token, spreadsheet_token, sheet_id, range_start="A1", range_end="Z100"):
    """Get data from specific range in sheet"""
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!{range_start}:{range_end}"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print(f"📊 Fetching sheet data from range {range_start}:{range_end}...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            return data.get("data", {}).get("values", [])
        else:
            print(f"⚠️  API returned code {data.get('code')}: {data.get('msg')}")
            return []
    else:
        print(f"❌ HTTP {response.status_code}: {response.text}")
        return []

def analyze_sheet_structure(spreadsheet_token: str = None, sheet_id: str = None):
    """Analyze the Feishu sheet structure"""
    
    # Load config for sheet details
    config = load_feishu_config()
    if not config:
        print("❌ No Feishu config found")
        return None
    
    # Use provided params or config
    spreadsheet_token = spreadsheet_token or config.get('spreadsheet_token')
    sheet_id = sheet_id or config.get('sheet_id')
    
    if not spreadsheet_token or not sheet_id:
        print("❌ Missing spreadsheet_token or sheet_id in config")
        return None
    
    # Get access token
    token = get_feishu_access_token()
    if not token:
        return None
    
    print(f"\n🔍 Analyzing Feishu sheet structure...")
    print(f"📋 Spreadsheet Token: {spreadsheet_token}")
    print(f"📄 Sheet ID: {sheet_id}")
    
    # Get a large range to understand the structure
    values = get_sheet_data(token, spreadsheet_token, sheet_id, "A1", "AZ200")
    
    if not values:
        print("❌ No data found in sheet")
        return
    
    print(f"\n📈 Found {len(values)} rows of data")
    
    # Analyze the structure
    print("\n=== SHEET STRUCTURE ANALYSIS ===")
    
    # Find campaign names (likely in first column)
    campaigns = []
    week_headers = []
    
    for row_idx, row in enumerate(values):
        if row and len(row) > 0:
            first_cell = str(row[0]).strip()
            
            # Look for campaign names (contain keywords like "Evergreen", "Prospecting", etc.)
            if any(keyword in first_cell.lower() for keyword in ["evergreen", "prospecting", "retargeting", "cns", "s+"]):
                campaigns.append({
                    "row": row_idx + 1,
                    "name": first_cell,
                    "data": row[1:] if len(row) > 1 else []
                })
            
            # Look for week headers (contain "Week" or numbers)
            elif "week" in first_cell.lower() or any(str(i) in first_cell for i in range(1, 53)):
                if len(row) > 1:
                    week_headers = row
                    print(f"🗓️  Week headers found in row {row_idx + 1}: {row[:10]}...")  # Show first 10
    
    print(f"\n🎯 Found {len(campaigns)} campaigns:")
    for i, campaign in enumerate(campaigns[:10]):  # Show first 10
        print(f"  Row {campaign['row']}: {campaign['name']}")
    
    if len(campaigns) > 10:
        print(f"  ... and {len(campaigns) - 10} more campaigns")
    
    # Find the current/latest week column
    current_week_col = None
    if week_headers:
        for col_idx, header in enumerate(week_headers):
            if header and "week" in str(header).lower():
                # Look for the latest week (highest number)
                week_num_match = re.search(r'week\s*(\d+)', str(header).lower())
                if week_num_match:
                    week_num = int(week_num_match.group(1))
                    if week_num > 30:  # Assuming we're in the current year (week 30+)
                        current_week_col = col_idx + 1  # 1-based indexing for Excel
    
    print(f"\n📅 Current week column: {current_week_col}")
    
    # Save analysis results
    analysis_results = {
        "spreadsheet_token": spreadsheet_token,
        "sheet_id": sheet_id,
        "total_rows": len(values),
        "campaigns": campaigns,
        "week_headers": week_headers,
        "current_week_column": current_week_col,
        "sample_data": values[:20]  # First 20 rows for reference
    }
    
    with open('/Users/bytedance/Desktop/studio/lulu-adobe-automation-studio/feishu_sheet_analysis.json', 'w') as f:
        json.dump(analysis_results, f, indent=2)
    
    print(f"\n✅ Analysis complete! Results saved to feishu_sheet_analysis.json")
    
    # Return key findings
    return {
        "campaigns": campaigns,
        "current_week_column": current_week_col,
        "week_headers": week_headers
    }

if __name__ == "__main__":
    print("🧹 Feishu Sheet Analyzer - Cleaned Version")
    print("==========================================")
    print("This script now requires a config file: feishu_config.json")
    print("Create it with: app_id, app_secret, spreadsheet_token, sheet_id")
    print()
    
    try:
        results = analyze_sheet_structure()
        
        if results and results["campaigns"]:
            print(f"\n🎯 Found {len(results['campaigns'])} campaigns in sheet!")
            print("📋 Analysis saved to feishu_sheet_analysis.json")
        else:
            print("\n⚠️  No campaigns found. Please check the sheet structure or config.")
            
    except Exception as e:
        print(f"❌ Error: {e}")