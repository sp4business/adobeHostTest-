#!/usr/bin/env python3
"""
Week 37 ROAS Data Transfer to Feishu
Ready-to-run script once you have API credentials
"""

import json
import os
from typing import Dict, List

def show_current_data():
    """Show the Week 37 data we have ready to transfer"""
    
    # Week 37 data from our Excel analysis
    week_37_data = {
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
    
    print("📊 Week 37 ROAS Data Ready for Transfer")
    print("="*50)
    print(f"📅 Week: {week_37_data['week']}")
    print(f"📈 Overall ROAS: {week_37_data['overall_roas']}x")
    print(f"🎯 Campaigns: {len(week_37_data['campaigns'])}")
    print()
    
    print("Campaign Breakdown:")
    for campaign in week_37_data['campaigns']:
        print(f"  • {campaign['name']}: {campaign['roas']}x")
    
    return week_37_data

def create_transfer_instructions():
    """Create step-by-step instructions for manual transfer"""
    
    print("\n\n📋 MANUAL TRANSFER INSTRUCTIONS")
    print("="*50)
    
    print("\n1. 📊 Open your Feishu sheet:")
    print("   https://bytedance.us.larkoffice.com/sheets/SBYJsQ4KkhQ1Svti88ouShLtsK0")
    
    print("\n2. 🎯 Find Week 37 column:")
    print("   Look for 'Week 37' header in your sheet")
    
    print("\n3. 📤 Copy data from Excel:")
    print("   • Open: week_37_AB_format.xlsx")
    print("   • Copy Column B (ROAS values)")
    
    print("\n4. 📥 Paste into Feishu:")
    print("   • Paste into Week 37 column")
    print("   • Match rows by campaign name")
    
    print("\n📊 Data to copy (Column B from Excel):")
    data = show_current_data()
    
    print(f"\n📋 Copy-paste ready format:")
    print("="*30)
    print(f"Week 37 Overall: {data['overall_roas']}")
    for campaign in data['campaigns']:
        print(f"{campaign['roas']}")

def create_api_setup_guide():
    """Create API setup guide for future automation"""
    
    print("\n\n🔧 API AUTOMATION SETUP (Optional)")
    print("="*50)
    
    print("\nTo enable direct API transfer, you need:")
    
    print("\n1. 🔐 Get Feishu API Credentials:")
    print("   • Go to https://open.feishu.cn/")
    print("   • Create developer application")
    print("   • Get App ID and App Secret")
    
    print("\n2. 📝 Update feishu_config.json:")
    config_template = {
        "app_id": "YOUR_APP_ID_HERE",
        "app_secret": "YOUR_APP_SECRET_HERE", 
        "spreadsheet_token": "SBYJsQ4KkhQ1Svti88ouShLtsK0",
        "sheet_id": "IqubAk"
    }
    
    print("   Replace with your actual credentials:")
    print(f"   {json.dumps(config_template, indent=2)}")
    
    print("\n3. 🚀 Run automation:")
    print("   python3 transfer_to_feishu.py")
    
    print("\n4. ✅ Verify transfer:")
    print("   Check your Feishu sheet for updated Week 37 data")

def main():
    """Main function"""
    
    print("🚀 Week 37 ROAS Data Transfer")
    print("="*50)
    
    # Show current data
    data = show_current_data()
    
    # Show manual transfer instructions
    create_transfer_instructions()
    
    # Show API setup guide
    create_api_setup_guide()
    
    print(f"\n\n🎉 READY TO TRANSFER WEEK 37 DATA!")
    print("="*50)
    print("Choose your method:")
    print("1. 📋 Manual copy-paste (immediate)")
    print("2. 🔧 Set up API automation (future)")
    
    # Save data to file for reference
    with open('week_37_transfer_data.json', 'w') as f:
        json.dump(data, f, indent=2)
    
    print(f"\n💾 Data saved to: week_37_transfer_data.json")

if __name__ == "__main__":
    main()