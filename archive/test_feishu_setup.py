#!/usr/bin/env python3
"""
Feishu API Setup and Test Script
Helps you set up and test the Feishu API connection
"""

import requests
import json
import os

def test_feishu_connection():
    """Test basic Feishu API connection"""
    
    print("🧪 Testing Feishu API Connection")
    print("="*50)
    
    # Load config
    try:
        with open('feishu_config.json', 'r') as f:
            config = json.load(f)
    except FileNotFoundError:
        print("❌ feishu_config.json not found")
        return False
    except Exception as e:
        print(f"❌ Error reading config: {e}")
        return False
    
    # Check if credentials are still placeholders
    if config['app_id'] == "YOUR_APP_ID_HERE" or config['app_secret'] == "YOUR_APP_SECRET_HERE":
        print("❌ API credentials are still set to placeholder values")
        print("\n📝 To get your Feishu API credentials:")
        print("1. Go to https://open.feishu.cn/")
        print("2. Create a developer application")
        print("3. Get your App ID and App Secret")
        print("4. Update feishu_config.json with real credentials")
        return False
    
    print(f"📋 Using App ID: {config['app_id'][:10]}...")
    print(f"📊 Spreadsheet: {config['spreadsheet_token']}")
    print(f"📄 Sheet: {config['sheet_id']}")
    
    # Test authentication
    print("\n🔑 Testing authentication...")
    
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    payload = {
        "app_id": config['app_id'],
        "app_secret": config['app_secret']
    }
    
    try:
        response = requests.post(url, json=payload)
        
        if response.status_code == 200:
            data = response.json()
            if data.get("code") == 0:
                print("✅ Authentication successful!")
                access_token = data["tenant_access_token"]
                
                # Test sheet access
                print("\n📋 Testing sheet access...")
                sheet_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{config['spreadsheet_token']}/metainfo"
                headers = {
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/json"
                }
                
                sheet_response = requests.get(sheet_url, headers=headers)
                
                if sheet_response.status_code == 200:
                    sheet_data = sheet_response.json()
                    if sheet_data.get("code") == 0:
                        print("✅ Sheet access successful!")
                        
                        # Get sheet info
                        sheet_info = sheet_data.get("data", {})
                        print(f"📊 Sheet title: {sheet_info.get('title', 'Unknown')}")
                        print(f"📄 Total sheets: {len(sheet_info.get('sheets', []))}")
                        
                        # Test reading a small range
                        print("\n📖 Testing data read...")
                        read_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{config['spreadsheet_token']}/values/{config['sheet_id']}!A1:C10"
                        read_response = requests.get(read_url, headers=headers)
                        
                        if read_response.status_code == 200:
                            read_data = read_response.json()
                            if read_data.get("code") == 0:
                                values = read_data.get("data", {}).get("values", [])
                                print(f"✅ Data read successful!")
                                print(f"📊 Found {len(values)} rows of data")
                                
                                if values:
                                    print(f"📝 Sample data (first 3 rows):")
                                    for i, row in enumerate(values[:3]):
                                        print(f"   Row {i+1}: {row[:3]}")  # First 3 columns
                                
                                return True
                            else:
                                print(f"❌ Data read error: {read_data.get('msg')}")
                        else:
                            print(f"❌ Data read HTTP {read_response.status_code}")
                            
                    else:
                        print(f"❌ Sheet access error: {sheet_data.get('msg')}")
                else:
                    print(f"❌ Sheet access HTTP {sheet_response.status_code}")
                    
            else:
                print(f"❌ Authentication error: {data.get('msg')}")
        else:
            print(f"❌ Authentication HTTP {response.status_code}")
            
    except Exception as e:
        print(f"❌ Request failed: {e}")
    
    return False

def show_setup_instructions():
    """Show detailed setup instructions"""
    print("\n📖 Feishu API Setup Instructions")
    print("="*50)
    
    print("\n1. 🔐 Get API Credentials:")
    print("   • Go to https://open.feishu.cn/")
    print("   • Log in with your ByteDance account")
    print("   • Navigate to 'Developer' section")
    print("   • Create a new app or use existing one")
    print("   • Get App ID and App Secret")
    
    print("\n2. 📋 Get Spreadsheet Info:")
    print("   • Open your target spreadsheet: https://bytedance.us.larkoffice.com/sheets/SBYJsQ4KkhQ1Svti88ouShLtsK0")
    print("   • The spreadsheet token is: SBYJsQ4KkhQ1Svti88ouShLtsK0")
    print("   • The sheet ID is: IqubAk (from the URL)")
    
    print("\n3. 🔑 Update Configuration:")
    print("   • Edit feishu_config.json")
    print("   • Replace YOUR_APP_ID_HERE with real App ID")
    print("   • Replace YOUR_APP_SECRET_HERE with real App Secret")
    print("   • Keep the existing spreadsheet_token and sheet_id")
    
    print("\n4. 🧪 Test Connection:")
    print("   • Run: python3 test_feishu_setup.py")
    print("   • Check for successful authentication and sheet access")
    
    print("\n5. 📤 Transfer Data:")
    print("   • Run: python3 transfer_to_feishu.py")
    print("   • This will upload Week 37 ROAS data to your sheet")

def main():
    """Main function"""
    print("🚀 Feishu API Setup and Test")
    print("="*50)
    
    # Test connection
    success = test_feishu_connection()
    
    if not success:
        show_setup_instructions()
        print("\n❌ Setup incomplete. Please follow the instructions above.")
    else:
        print("\n✅ Setup complete! You can now transfer data to Feishu.")
        print("💡 Next step: python3 transfer_to_feishu.py")

if __name__ == "__main__":
    main()