#!/usr/bin/env python3
"""
Feishu API Data Transfer - Week 37 ROAS Data
Pushes the extracted Excel data directly to your Feishu sheet
"""

import requests
import json
import os
from typing import Dict, List, Optional, Any
from excel_data_reader import ExcelDataReader

class FeishuDataTransfer:
    """Handles complete data transfer from Excel to Feishu"""
    
    def __init__(self, config_file: str = "feishu_config.json"):
        self.config = self.load_config(config_file)
        self.base_url = "https://open.feishu.cn/open-apis"
        self.access_token = None
        self.excel_reader = ExcelDataReader()
        
    def load_config(self, config_file: str) -> Dict:
        """Load Feishu configuration"""
        if not os.path.exists(config_file):
            raise FileNotFoundError(f"Config file {config_file} not found")
        
        try:
            with open(config_file, 'r') as f:
                config = json.load(f)
                
            required_keys = ['app_id', 'app_secret', 'spreadsheet_token', 'sheet_id']
            for key in required_keys:
                if key not in config:
                    raise ValueError(f"Missing required config key: {key}")
                    
            return config
        except Exception as e:
            raise Exception(f"Error loading config: {e}")
    
    def get_access_token(self) -> Optional[str]:
        """Get Feishu access token"""
        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        
        payload = {
            "app_id": self.config['app_id'],
            "app_secret": self.config['app_secret']
        }
        
        try:
            response = requests.post(url, json=payload)
            if response.status_code == 200:
                data = response.json()
                if data.get("code") == 0:
                    self.access_token = data["tenant_access_token"]
                    print(f"✅ Access token obtained successfully")
                    return self.access_token
                else:
                    print(f"❌ Token error: {data.get('msg')}")
            else:
                print(f"❌ HTTP {response.status_code}: {response.text}")
        except Exception as e:
            print(f"❌ Token request failed: {e}")
        
        return None
    
    def get_sheet_structure(self) -> Optional[Dict]:
        """Analyze the target Feishu sheet structure"""
        if not self.access_token:
            return None
            
        # Get sheet data to understand structure
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.config['spreadsheet_token']}/values/{self.config['sheet_id']}!A1:Z100"
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                data = response.json()
                if data.get("code") == 0:
                    values = data.get("data", {}).get("values", [])
                    
                    # Analyze structure
                    structure = {
                        'campaign_rows': {},
                        'week_columns': {},
                        'data': values
                    }
                    
                    # Find campaign rows and week columns
                    for row_idx, row in enumerate(values):
                        if row and len(row) > 0:
                            first_cell = str(row[0]).strip()
                            
                            # Look for campaign names
                            if any(keyword in first_cell.lower() for keyword in ['evergreen', 'prospecting', 'retargeting', 'cns', 's+']):
                                structure['campaign_rows'][first_cell] = row_idx + 1  # 1-based
                            
                            # Look for week headers in first row
                            if row_idx == 0:
                                for col_idx, cell in enumerate(row):
                                    if cell and "week" in str(cell).lower():
                                        week_match = self.extract_week_number(str(cell))
                                        if week_match:
                                            structure['week_columns'][week_match] = self.column_to_letter(col_idx + 1)
                    
                    print(f"📊 Sheet structure analyzed:")
                    print(f"   Found {len(structure['campaign_rows'])} campaigns")
                    print(f"   Found {len(structure['week_columns'])} week columns")
                    
                    return structure
                else:
                    print(f"⚠️  Sheet data error: {data.get('msg')}")
            else:
                print(f"❌ HTTP {response.status_code}: {response.text}")
        except Exception as e:
            print(f"❌ Sheet data request failed: {e}")
        
        return None
    
    def extract_week_number(self, text: str) -> Optional[str]:
        """Extract week number from text"""
        import re
        week_match = re.search(r'Week\s*(\d+)', text, re.IGNORECASE)
        return week_match.group(1) if week_match else None
    
    def column_to_letter(self, col_num: int) -> str:
        """Convert column number to Excel letter (1 = A, 2 = B, etc.)"""
        letter = ""
        while col_num > 0:
            col_num -= 1
            letter = chr(col_num % 26 + 65) + letter
            col_num //= 26
        return letter
    
    def prepare_update_data(self, excel_data: Dict, sheet_structure: Dict) -> Optional[Dict]:
        """Prepare the data update for Feishu API"""
        week_num = excel_data.get('week')
        if not week_num:
            print("❌ No week number in Excel data")
            return None
        
        # Find the target week column
        target_week_col = sheet_structure['week_columns'].get(week_num)
        if not target_week_col:
            print(f"❌ Week {week_num} column not found in sheet")
            print(f"Available weeks: {list(sheet_structure['week_columns'].keys())}")
            return None
        
        # Build update data
        updates = []
        
        # Update overall ROAS (if we can find where it goes)
        if excel_data['overall_roas']:
            # Look for overall ROAS row (usually near the top)
            for row_name, row_num in sheet_structure['campaign_rows'].items():
                if "overall" in row_name.lower() and "roas" in row_name.lower():
                    cell_range = f"{self.config['sheet_id']}!{target_week_col}{row_num}"
                    updates.append({
                        'range': cell_range,
                        'values': [[excel_data['overall_roas']]]
                    })
                    break
        
        # Update campaign ROAS values
        for campaign in excel_data['campaigns']:
            campaign_name = campaign['name']
            roas_value = campaign['roas']
            
            # Find matching row in sheet
            found_row = None
            for sheet_campaign, row_num in sheet_structure['campaign_rows'].items():
                if self.campaigns_match(campaign_name, sheet_campaign):
                    found_row = row_num
                    break
            
            if found_row:
                cell_range = f"{self.config['sheet_id']}!{target_week_col}{found_row}"
                updates.append({
                    'range': cell_range,
                    'values': [[roas_value]]
                })
                print(f"📝 Mapping: {campaign_name} → Row {found_row}")
            else:
                print(f"⚠️  No matching row found for: {campaign_name}")
        
        if not updates:
            print("❌ No valid updates to make")
            return None
        
        print(f"📊 Prepared {len(updates)} cell updates for week {week_num}")
        return {
            'week': week_num,
            'week_column': target_week_col,
            'updates': updates
        }
    
    def campaigns_match(self, excel_campaign: str, sheet_campaign: str) -> bool:
        """Check if campaign names match (fuzzy matching)"""
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
    
    def execute_batch_update(self, update_data: Dict) -> bool:
        """Execute batch update to Feishu sheet"""
        if not self.access_token:
            print("❌ No access token")
            return False
        
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.config['spreadsheet_token']}/values"
        
        # Prepare batch update payload
        data = []
        for update in update_data['updates']:
            data.append({
                "range": update['range'],
                "values": update['values']
            })
        
        payload = {
            "valueRanges": data
        }
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            print(f"🚀 Executing batch update for week {update_data['week']}...")
            response = requests.put(url, json=payload, headers=headers)
            
            if response.status_code == 200:
                result = response.json()
                if result.get("code") == 0:
                    print(f"✅ Successfully updated {len(data)} cells!")
                    return True
                else:
                    print(f"❌ Update error: {result.get('msg')}")
            else:
                print(f"❌ HTTP {response.status_code}: {response.text}")
                
        except Exception as e:
            print(f"❌ Batch update failed: {e}")
        
        return False
    
    def transfer_excel_to_feishu(self, excel_file: str) -> bool:
        """Complete transfer workflow"""
        print(f"🚀 Starting Excel to Feishu transfer...")
        print(f"📁 Excel file: {excel_file}")
        
        # Step 1: Get access token
        print(f"🔑 Getting Feishu access token...")
        if not self.get_access_token():
            return False
        
        # Step 2: Read Excel data
        print(f"📖 Reading Excel data...")
        if "AB_format" in excel_file:
            excel_data = self.excel_reader.read_ab_format_excel(excel_file)
        else:
            excel_data = self.excel_reader.read_feishu_format_excel(excel_file)
        
        if not excel_data or not excel_data['campaigns']:
            print("❌ No campaign data found in Excel")
            return False
        
        # Step 3: Analyze sheet structure
        print(f"🔍 Analyzing Feishu sheet structure...")
        sheet_structure = self.get_sheet_structure()
        if not sheet_structure:
            print("❌ Could not analyze sheet structure")
            return False
        
        # Step 4: Prepare update data
        print(f"📝 Preparing update data...")
        update_data = self.prepare_update_data(excel_data, sheet_structure)
        if not update_data:
            print("❌ Could not prepare update data")
            return False
        
        # Step 5: Execute update
        return self.execute_batch_update(update_data)

def main():
    """Main function"""
    
    # Find Excel file
    excel_files = [
        "week_37_AB_format.xlsx",
        "week_37_feishu_ready.xlsx"
    ]
    
    excel_file = None
    for file in excel_files:
        if os.path.exists(file):
            excel_file = file
            break
    
    if not excel_file:
        print("❌ No Excel file found. Please generate one first:")
        print("💡 python ab_format_excel_generator.py")
        return
    
    print(f"🎯 Using: {excel_file}")
    
    # Check config
    if not os.path.exists("feishu_config.json"):
        print("❌ feishu_config.json not found. Please set up your API credentials.")
        return
    
    # Execute transfer
    try:
        transfer = FeishuDataTransfer()
        success = transfer.transfer_excel_to_feishu(excel_file)
        
        if success:
            print(f"\n🎉 SUCCESS! Week 37 ROAS data transferred to Feishu!")
            print(f"📊 Check your sheet: https://bytedance.us.larkoffice.com/sheets/{transfer.config['spreadsheet_token']}")
        else:
            print(f"\n❌ Transfer failed. Check error messages above.")
            
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    import os
    main()