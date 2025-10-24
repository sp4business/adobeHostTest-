#!/usr/bin/env python3
"""
Feishu Automation - Adobe ROAS Data to Feishu Sheets
Complete workflow for pushing parsed Adobe campaign data to Lululemon tracking sheet
"""

import requests
import json
import os
import re
from typing import Dict, List, Optional, Tuple
from datetime import datetime

class FeishuAutomation:
    """Complete automation for Adobe data → Feishu sheet transfer"""

    def __init__(self, config_file: str = "feishu_config.json"):
        self.config = self.load_config(config_file)
        # Use config base_url if provided, otherwise default to Lark international
        self.base_url = self.config.get('base_url', 'https://open.larksuite.com/open-apis')
        self.access_token = None
        self.sheet_structure = None
        print(f"🚀 FeishuAutomation initialized (API: {self.base_url})")

    def load_config(self, config_file: str) -> Dict:
        """Load Feishu API configuration"""
        if not os.path.exists(config_file):
            raise FileNotFoundError(f"❌ Config file not found: {config_file}")

        with open(config_file, 'r') as f:
            config = json.load(f)

        required = ['app_id', 'app_secret', 'spreadsheet_token', 'sheet_id']
        for key in required:
            if key not in config:
                raise ValueError(f"❌ Missing config key: {key}")

        return config

    def get_access_token(self) -> Optional[str]:
        """Get Feishu tenant access token"""
        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        payload = {
            "app_id": self.config['app_id'],
            "app_secret": self.config['app_secret']
        }

        try:
            print("🔑 Requesting access token...")
            print(f"   API endpoint: {url}")
            print(f"   App ID: {payload['app_id']}")
            response = requests.post(url, json=payload)

            print(f"   Response status: {response.status_code}")
            if response.status_code == 200:
                data = response.json()
                print(f"   Response data: {data}")
                if data.get("code") == 0:
                    self.access_token = data["tenant_access_token"]
                    print("✅ Access token obtained")
                    return self.access_token
                else:
                    print(f"❌ Token error: {data.get('msg')}")
                    print(f"   Full response: {data}")
            else:
                print(f"❌ HTTP {response.status_code}: {response.text}")
                print(f"\n💡 Troubleshooting:")
                print(f"   - Check if app credentials are correct")
                print(f"   - Verify app is created on the correct platform:")
                print(f"     * Lark (International): https://open.larksuite.com")
                print(f"     * Feishu (China): https://open.feishu.cn")
                print(f"   - Ensure app has required permissions")
        except Exception as e:
            print(f"❌ Token request failed: {e}")

        return None

    def read_sheet_structure(self) -> Optional[Dict]:
        """Read Feishu sheet to map campaigns to rows and find week columns"""
        if not self.access_token:
            print("❌ No access token")
            return None

        # Read first 100 rows and columns A-Z
        range_str = f"{self.config['sheet_id']}!A1:Z100"
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.config['spreadsheet_token']}/values/{range_str}"

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        try:
            print("📊 Reading sheet structure...")
            response = requests.get(url, headers=headers)

            if response.status_code == 200:
                data = response.json()
                if data.get("code") == 0:
                    values = data.get("data", {}).get("valueRange", {}).get("values", [])

                    structure = {
                        'campaign_rows': {},  # {"campaign_name": row_number}
                        'week_columns': {},   # {week_number: column_letter}
                        'overall_roas_row': None,
                        'raw_data': values
                    }

                    # Analyze structure
                    for row_idx, row in enumerate(values):
                        if not row or len(row) == 0:
                            continue

                        # Row 1: Look for week headers
                        if row_idx == 0:
                            for col_idx, cell in enumerate(row):
                                if cell and "week" in str(cell).lower():
                                    week_num = self.extract_week_number(str(cell))
                                    if week_num:
                                        col_letter = self.column_to_letter(col_idx + 1)
                                        structure['week_columns'][week_num] = col_letter

                        # Other rows: Look for campaign names in column A
                        first_cell = str(row[0]).strip() if row else ""

                        # Check for Overall ROAS row
                        if "overall" in first_cell.lower() and "roas" in first_cell.lower():
                            structure['overall_roas_row'] = row_idx + 1  # 1-based

                        # Check for campaign rows
                        campaign_keywords = ['evergreen', 'prospecting', 'retargeting', 'cns', 's+']
                        if any(keyword in first_cell.lower() for keyword in campaign_keywords):
                            structure['campaign_rows'][first_cell] = row_idx + 1  # 1-based

                    self.sheet_structure = structure

                    print(f"✅ Sheet structure analyzed:")
                    print(f"   📍 Found {len(structure['campaign_rows'])} campaign rows")
                    print(f"   📅 Found {len(structure['week_columns'])} week columns")
                    print(f"   📈 Overall ROAS row: {structure['overall_roas_row']}")

                    return structure
                else:
                    print(f"❌ Sheet read error: {data.get('msg')}")
            else:
                print(f"❌ HTTP {response.status_code}: {response.text}")
        except Exception as e:
            print(f"❌ Sheet read failed: {e}")

        return None

    def extract_week_number(self, text: str) -> Optional[str]:
        """Extract week number from text like 'Week 37'"""
        match = re.search(r'[Ww]eek\s*(\d+)', text)
        return match.group(1) if match else None

    def column_to_letter(self, col_num: int) -> str:
        """Convert column number to Excel letter (1=A, 27=AA)"""
        letter = ""
        while col_num > 0:
            col_num -= 1
            letter = chr(col_num % 26 + 65) + letter
            col_num //= 26
        return letter

    def find_or_create_week_column(self, week_number: str) -> Optional[str]:
        """Find existing week column or create new one"""
        if not self.sheet_structure:
            print("❌ No sheet structure loaded")
            return None

        # Check if week column already exists
        if week_number in self.sheet_structure['week_columns']:
            col = self.sheet_structure['week_columns'][week_number]
            print(f"✅ Found existing Week {week_number} column: {col}")
            return col

        # Need to create new column
        print(f"📝 Week {week_number} column not found, creating...")

        # Get current column count
        current_columns = len(self.sheet_structure['raw_data'][0]) if self.sheet_structure['raw_data'] else 0
        new_col_index = current_columns  # 0-based

        # Insert column
        if self.insert_column(new_col_index):
            new_col_letter = self.column_to_letter(new_col_index + 1)

            # Write week header
            header_range = f"{new_col_letter}1"
            if self.write_to_sheet(header_range, [[f"Week {week_number}"]]):
                print(f"✅ Created Week {week_number} column: {new_col_letter}")
                self.sheet_structure['week_columns'][week_number] = new_col_letter
                return new_col_letter

        print(f"❌ Failed to create Week {week_number} column")
        return None

    def insert_column(self, start_index: int) -> bool:
        """Insert new column at specified index"""
        if not self.access_token:
            return False

        url = f"{self.base_url}/sheets/v3/spreadsheets/{self.config['spreadsheet_token']}/sheets/{self.config['sheet_id']}/insert_dimension"

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "dimension_range": {
                "major_dimension": "COLUMNS",
                "start_index": start_index,
                "end_index": start_index + 1
            }
        }

        try:
            response = requests.post(url, json=payload, headers=headers)
            if response.status_code == 200:
                data = response.json()
                return data.get("code") == 0
        except Exception as e:
            print(f"❌ Insert column error: {e}")

        return False

    def campaigns_match(self, campaign1: str, campaign2: str) -> bool:
        """Fuzzy match campaign names"""
        c1 = campaign1.lower().strip()
        c2 = campaign2.lower().strip()

        # Direct match
        if c1 == c2:
            return True

        # Key terms matching
        key_terms = ['us', 'canada', 'evergreen', 'prospecting',
                     'retargeting', 'cns', 's+', 'cost cap', 'lowest cost',
                     'manual', 'carousel']

        c1_terms = [term for term in key_terms if term in c1]
        c2_terms = [term for term in key_terms if term in c2]

        if len(c1_terms) > 0 and len(c2_terms) > 0:
            common_terms = set(c1_terms) & set(c2_terms)
            threshold = min(len(c1_terms), len(c2_terms)) * 0.7
            return len(common_terms) >= threshold

        return False

    def build_update_payload(self, adobe_data: Dict, week_column: str) -> Optional[Dict]:
        """Build payload for batch update"""
        if not self.sheet_structure:
            print("❌ No sheet structure")
            return None

        print(f"\n📝 Building update payload for Week {adobe_data['week']}...")

        updates = []
        matched_campaigns = 0
        unmatched_campaigns = []

        # Update Overall ROAS if we know the row
        if self.sheet_structure['overall_roas_row'] and adobe_data.get('overall_roas'):
            cell_range = f"{week_column}{self.sheet_structure['overall_roas_row']}"
            updates.append({
                'range': cell_range,
                'value': adobe_data['overall_roas'],
                'description': 'Overall ROAS'
            })
            print(f"   ✓ Overall ROAS: {adobe_data['overall_roas']} → {cell_range}")

        # Update campaign ROAS values
        for campaign in adobe_data['campaigns']:
            campaign_name = campaign['name']
            roas_value = campaign['roas']

            # Find matching row in sheet
            matched_row = None
            for sheet_campaign, row_num in self.sheet_structure['campaign_rows'].items():
                if self.campaigns_match(campaign_name, sheet_campaign):
                    matched_row = row_num
                    break

            if matched_row:
                cell_range = f"{week_column}{matched_row}"
                updates.append({
                    'range': cell_range,
                    'value': roas_value,
                    'description': campaign_name
                })
                print(f"   ✓ {campaign_name[:50]}: {roas_value} → {cell_range}")
                matched_campaigns += 1
            else:
                unmatched_campaigns.append(campaign_name)
                print(f"   ⚠️  No match found for: {campaign_name}")

        print(f"\n📊 Summary: {matched_campaigns}/{len(adobe_data['campaigns'])} campaigns matched")

        if unmatched_campaigns:
            print(f"⚠️  Unmatched campaigns ({len(unmatched_campaigns)}):")
            for uc in unmatched_campaigns:
                print(f"   - {uc}")

        if not updates:
            print("❌ No updates to make")
            return None

        return {
            'week': adobe_data['week'],
            'week_column': week_column,
            'updates': updates,
            'matched_count': matched_campaigns,
            'total_count': len(adobe_data['campaigns'])
        }

    def execute_batch_update(self, update_data: Dict) -> bool:
        """Execute batch update to Feishu sheet"""
        if not self.access_token:
            print("❌ No access token")
            return False

        if not update_data or not update_data.get('updates'):
            print("❌ No updates to execute")
            return False

        print(f"\n🚀 Executing batch update for Week {update_data['week']}...")

        # Build valueRanges array for batch update
        value_ranges = []
        for update in update_data['updates']:
            full_range = f"{self.config['sheet_id']}!{update['range']}"
            value_ranges.append({
                "range": full_range,
                "values": [[update['value']]]  # Always 2D array
            })

        # Use batch update API
        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.config['spreadsheet_token']}/values_batch_update"

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "valueRanges": value_ranges
        }

        try:
            print(f"📤 Sending {len(value_ranges)} cell updates...")
            response = requests.put(url, json=payload, headers=headers)

            if response.status_code == 200:
                data = response.json()
                if data.get("code") == 0:
                    print(f"✅ Successfully updated {len(value_ranges)} cells!")
                    print(f"   📊 Week {update_data['week']} ROAS data is now in Feishu")
                    print(f"   🔗 View: https://bytedance.us.larkoffice.com/sheets/{self.config['spreadsheet_token']}")
                    return True
                else:
                    print(f"❌ Batch update error: {data.get('msg')}")
                    print(f"   Code: {data.get('code')}")
            else:
                print(f"❌ HTTP {response.status_code}: {response.text}")
        except Exception as e:
            print(f"❌ Batch update failed: {e}")

        return False

    def write_to_sheet(self, range_str: str, values: List[List]) -> bool:
        """Write to single range (helper method)"""
        if not self.access_token:
            return False

        url = f"{self.base_url}/sheets/v2/spreadsheets/{self.config['spreadsheet_token']}/values"

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "valueRange": {
                "range": f"{self.config['sheet_id']}!{range_str}",
                "values": values
            }
        }

        try:
            response = requests.put(url, json=payload, headers=headers)
            if response.status_code == 200:
                data = response.json()
                return data.get("code") == 0
        except Exception as e:
            print(f"❌ Write error: {e}")

        return False

    def load_adobe_data(self, data_file: str) -> Optional[Dict]:
        """Load parsed Adobe campaign data"""
        if not os.path.exists(data_file):
            print(f"❌ Data file not found: {data_file}")
            return None

        try:
            with open(data_file, 'r') as f:
                data = json.load(f)

            print(f"📖 Loaded Adobe data:")
            print(f"   Week: {data.get('week')}")
            print(f"   Overall ROAS: {data.get('overall_roas')}")
            print(f"   Campaigns: {len(data.get('campaigns', []))}")

            return data
        except Exception as e:
            print(f"❌ Error loading data: {e}")
            return None

    def run_automation(self, data_file: str = "output/week_37_transfer_data.json") -> bool:
        """Complete automation workflow"""
        print("=" * 70)
        print("🎯 FEISHU AUTOMATION - Adobe ROAS Data Transfer")
        print("=" * 70)

        # Step 1: Authenticate
        if not self.get_access_token():
            print("❌ Authentication failed")
            return False

        # Step 2: Load Adobe data
        adobe_data = self.load_adobe_data(data_file)
        if not adobe_data:
            print("❌ Failed to load Adobe data")
            return False

        # Step 3: Read sheet structure
        if not self.read_sheet_structure():
            print("❌ Failed to read sheet structure")
            return False

        # Step 4: Find/create week column
        week_column = self.find_or_create_week_column(adobe_data['week'])
        if not week_column:
            print("❌ Failed to find/create week column")
            return False

        # Step 5: Build update payload
        update_payload = self.build_update_payload(adobe_data, week_column)
        if not update_payload:
            print("❌ Failed to build update payload")
            return False

        # Step 6: Execute batch update
        success = self.execute_batch_update(update_payload)

        if success:
            print("\n" + "=" * 70)
            print("🎉 AUTOMATION COMPLETE!")
            print("=" * 70)
            print(f"✅ Week {adobe_data['week']} ROAS data successfully pushed to Feishu")
            print(f"📊 Updated {update_payload['matched_count']}/{update_payload['total_count']} campaigns")
        else:
            print("\n❌ Automation failed - check errors above")

        return success

def main():
    """Main entry point"""
    import sys

    # Default data file
    data_file = "output/week_37_transfer_data.json"

    # Allow override from command line
    if len(sys.argv) > 1:
        data_file = sys.argv[1]

    # Check if data file exists
    if not os.path.exists(data_file):
        print(f"❌ Data file not found: {data_file}")
        print("\n💡 Available data files:")
        if os.path.exists("output"):
            for f in os.listdir("output"):
                if f.endswith("_data.json") or f.endswith("_transfer_data.json"):
                    print(f"   - output/{f}")
        print("\n📝 Usage: python 5_run_feishu_automation.py [data_file]")
        return

    # Run automation
    try:
        automation = FeishuAutomation()
        success = automation.run_automation(data_file)
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
