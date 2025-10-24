import requests
import json
import openpyxl
from typing import Dict, List, Optional

class FeishuAutomation:
    def __init__(self, config_file: str = "feishu_config.json"):
        self.config = self.load_config(config_file)
        self.base_url = "https://open.feishu.cn/open-apis"
        self.access_token = self.get_feishu_access_token()

    def load_config(self, config_file: str) -> Dict:
        with open(config_file, 'r') as f:
            return json.load(f)

    def get_feishu_access_token(self) -> Optional[str]:
        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        payload = {
            "app_id": self.config['app_id'],
            "app_secret": self.config['app_secret']
        }
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            data = response.json()
            if data.get("code") == 0:
                return data["tenant_access_token"]
        return None

    def get_sheet_dimensions(self, spreadsheet_token: str, sheet_id: str) -> Optional[Dict]:
        if not self.access_token:
            return None
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/metainfo"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            if data.get("code") == 0:
                for sheet in data.get("data", {}).get("sheets", []):
                    if sheet.get("sheetId") == sheet_id:
                        return {
                            'row_count': sheet.get("rowCount", 0),
                            'column_count': sheet.get("columnCount", 0)
                        }
        return None

    def insert_column(self, spreadsheet_token: str, sheet_id: str, start_index: int):
        if not self.access_token:
            return
        url = f"{self.base_url}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}/insert_dimension"
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
        response = requests.post(url, json=payload, headers=headers)
        return response.status_code == 200 and response.json().get("code") == 0

    def write_to_sheet(self, spreadsheet_token: str, sheet_id: str, range_str: str, values: List[List[str]]):
        if not self.access_token:
            return
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/values"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        payload = {
            "valueRange": {
                "range": f"{sheet_id}!{range_str}",
                "values": values
            }
        }
        response = requests.put(url, json=payload, headers=headers)
        return response.status_code == 200 and response.json().get("code") == 0

def main():
    automation = FeishuAutomation()
    spreadsheet_token = "SBYJsQ4KkhQ1Svti88ouShLtsK0"
    sheet_id = "IqubAk"
    
    dimensions = automation.get_sheet_dimensions(spreadsheet_token, sheet_id)
    if dimensions:
        new_column_index = dimensions['column_count'] -1
        header_range = f"{chr(ord('A') + new_column_index)}1"
        if automation.write_to_sheet(spreadsheet_token, sheet_id, header_range, [["Week 37"]]):
            print("Successfully wrote 'Week 37' to the new column header.")

if __name__ == "__main__":
    main()
