#!/usr/bin/env python3
"""
Direct Feishu API Transfer - Week 37 ROAS Data
Uses your working authentication code to push data directly
"""

import requests
import json

# Your working Feishu configuration
FEISHU_APP_ID = "cli_a86397ecf2f89013"
FEISHU_APP_SECRET = "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt"
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"

# Target sheet details from your URL
SPREADSHEET_TOKEN = "SBYJsQ4KkhQ1Svti88ouShLtsK0"
SHEET_ID = "IqubAk"

def get_feishu_access_token():
    """Get Feishu access token using your working code"""
    url = f"{FEISHU_BASE_URL}/auth/v3/tenant_access_token/internal"
    
    payload = {
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
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

def analyze_sheet_structure(token):
    """Analyze the sheet structure to find where to put Week 37 data"""
    
    print("🔍 Analyzing sheet structure...")
    
    # Read a large range to understand the layout
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:AZ100"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            values = data.get("data", {}).get("values", [])
            
            structure = {
                'week_headers': {},
                'campaign_rows': {},
                'data': values
            }
            
            # Find week headers (row 1)
            if values and len(values) > 0:
                header_row = values[0]
                for col_idx, cell in enumerate(header_row):
                    if cell and "week" in str(cell).lower():
                        week_match = extract_week_number(str(cell))
                        if week_match:
                            col_letter = column_to_letter(col_idx + 1)
                            structure['week_headers'][week_match] = col_letter
                            print(f"   Found Week {week_match} at column {col_letter}")
            
            # Find campaign rows
            for row_idx, row in enumerate(values):
                if row and len(row) > 0:
                    first_cell = str(row[0]).strip()
                    if any(keyword in first_cell.lower() for keyword in ['evergreen', 'prospecting', 'retargeting', 'cns', 's+']):
                        structure['campaign_rows'][first_cell] = row_idx + 1  # 1-based
                        print(f"   Found campaign: {first_cell[:50]}... at row {row_idx + 1}")
            
            print(f"📊 Found {len(structure['week_headers'])} week columns")
            print(f"📊 Found {len(structure['campaign_rows'])} campaign rows")
            
            return structure
        else:
            print(f"❌ Sheet analysis error: {data.get('msg')}")
    else:
        print(f"❌ HTTP {response.status_code}: {response.text}")
    
    return None

def extract_week_number(text):
    """Extract week number from text"""
    import re
    week_match = re.search(r'Week\s*(\d+)', text, re.IGNORECASE)
    return week_match.group(1) if week_match else None

def column_to_letter(col_num):
    """Convert column number to Excel letter"""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + 65) + letter
        col_num //= 26
    return letter

def prepare_week37_updates():
    """Prepare Week 37 ROAS data for upload"""
    
    # Week 37 data from our Excel analysis
    week37_data = {
        "week": "37",
        "overall_roas": 3.24,
        "campaigns": [
            {"name": "US Evergreen Prospecting (Lowest Cost)", "roas": 3.56},
            {"name": "S+ 2.0 US Evergreen Prospecting (Lowest Cost)", "roas": 2.76},
            {"name": "US Evergreen Prospecting (Cost Cap)", "roas": 3.03},
            {"name": "US Evergreen Retargeting (Lowest Cost)", "roas": 3.01},
            {"name": "Canada Evergreen Prospecting (Lowest Cost)", "roas": 2.86},
            {"name": "S+ 2.0 Canada Evergreen Prospecting (Lowest Cost)", "roas": 0.45},
            {"name": "Canada Evergreen Retargeting (Lowest Cost)", "roas": 2.4},
            {"name": "US CNS S+ 2.0 (carousel ads only - Lowest Cost)", "roas": 3.25},
            {"name": "Canada CNS S+ 2.0 (carousel ads only - Lowest Cost)", "roas": 3.59}
        ]
    }
    
    print(f"📊 Week 37 data prepared:")
    print(f"   Overall ROAS: {week37_data['overall_roas']}x")
    print(f"   Campaigns: {len(week37_data['campaigns'])}")
    
    return week37_data

def campaigns_match(excel_campaign, sheet_campaign):
    """Check if campaign names match"""
    excel_lower = excel_campaign.lower()
    sheet_lower = sheet_campaign.lower()
    
    # Direct match
    if excel_lower == sheet_lower:
        return True
    
    # Key component matching
    key_terms = ['us', 'canada', 'evergreen', 'prospecting', 'retargeting', 'cns', 's+', 'cost cap', 'lowest cost']
    
    excel_terms = [term for term in key_terms if term in excel_lower]
    sheet_terms = [term for term in key_terms if term in sheet_lower]
    
    # Match if most key terms align
    if len(excel_terms) > 0 and len(sheet_terms) > 0:
        common_terms = set(excel_terms) & set(sheet_terms)
        if len(common_terms) >= min(len(excel_terms), len(sheet_terms)) * 0.7:
            return True
    
    return False

def upload_week37_data(token, sheet_structure):
    """Upload Week 37 data to the sheet"""
    
    week37_data = prepare_week37_updates()
    
    # Find Week 37 column
    week37_col = sheet_structure['week_headers'].get('37')
    if not week37_col:
        print(f"❌ Week 37 column not found. Available weeks: {list(sheet_structure['week_headers'].keys())}")
        return False
    
    print(f"🎯 Targeting Week 37 column: {week37_col}")
    
    # Prepare batch updates
    updates = []
    
    # Update campaign ROAS values
    for campaign in week37_data['campaigns']:
        campaign_name = campaign['name']
        roas_value = campaign['roas']
        
        # Find matching row in sheet
        found_row = None
        for sheet_campaign, row_num in sheet_structure['campaign_rows'].items():
            if campaigns_match(campaign_name, sheet_campaign):
                found_row = row_num
                break
        
        if found_row:
            cell_range = f"{SHEET_ID}!{week37_col}{found_row}"
            updates.append({
                "range": cell_range,
                "values": [[roas_value]]
            })
            print(f"   ✓ {campaign_name} → Row {found_row}: {roas_value}x")
        else:
            print(f"   ⚠️  No match found for: {campaign_name}")
    
    if not updates:
        print("❌ No valid updates to make")
        return False
    
    print(f"📝 Prepared {len(updates)} cell updates")
    
    # Execute batch update
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    
    payload = {
        "valueRanges": updates
    }
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    print("🚀 Executing batch update...")
    response = requests.put(url, json=payload, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            print("✅ Successfully updated Feishu sheet!")
            print(f"📊 Updated {len(updates)} cells with Week 37 ROAS data")
            return True
        else:
            print(f"❌ Update error: {data.get('msg')}")
    else:
        print(f"❌ HTTP {response.status_code}: {response.text}")
    
    return False

def main():
    """Main function to transfer Week 37 data"""
    
    print("🚀 Week 37 ROAS Data Transfer to Feishu")
    print("="*50)
    
    try:
        # Step 1: Get access token
        token = get_feishu_access_token()
        if not token:
            return
        
        # Step 2: Analyze sheet structure
        sheet_structure = analyze_sheet_structure(token)
        if not sheet_structure:
            return
        
        # Step 3: Upload Week 37 data
        success = upload_week37_data(token, sheet_structure)
        
        if success:
            print(f"\n🎉 SUCCESS! Week 37 ROAS data transferred to Feishu!")
            print(f"📊 Check your sheet: https://bytedance.us.larkoffice.com/sheets/{SPREADSHEET_TOKEN}")
        else:
            print(f"\n❌ Transfer failed")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    main()