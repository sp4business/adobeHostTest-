#!/usr/bin/env python3
"""
Quick Platform Detector - Find out if your app is on Feishu (China) or Lark (International)
"""

import requests

# Your credentials
APP_ID = "cli_a86397ecf2f89013"
APP_SECRET = "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt"

print("=" * 70)
print("  FEISHU/LARK PLATFORM DETECTOR")
print("=" * 70)
print(f"\nApp ID: {APP_ID}")
print(f"\nTesting which platform your app is registered on...\n")

# Test 1: Feishu (China)
print("🇨🇳 Test 1: Feishu (China)")
print("   Endpoint: https://open.feishu.cn")

try:
    response = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
        timeout=10
    )

    print(f"   Status Code: {response.status_code}")

    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            print("   ✅ SUCCESS - Your app is on Feishu (China)!")
            print(f"   Token received: {data['tenant_access_token'][:20]}...")
            feishu_works = True
        else:
            print(f"   ❌ Error: {data.get('msg')}")
            feishu_works = False
    else:
        print(f"   ❌ Failed: {response.text[:100]}")
        feishu_works = False
except Exception as e:
    print(f"   ❌ Exception: {e}")
    feishu_works = False

print()

# Test 2: Lark (International)
print("🌎 Test 2: Lark (International/Global)")
print("   Endpoint: https://open.larksuite.com")

try:
    response = requests.post(
        "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
        timeout=10
    )

    print(f"   Status Code: {response.status_code}")

    if response.status_code == 200:
        data = response.json()
        if data.get("code") == 0:
            print("   ✅ SUCCESS - Your app is on Lark (International)!")
            print(f"   Token received: {data['tenant_access_token'][:20]}...")
            lark_works = True
        else:
            print(f"   ❌ Error: {data.get('msg')}")
            lark_works = False
    else:
        print(f"   ❌ Failed: {response.text[:100]}")
        lark_works = False
except Exception as e:
    print(f"   ❌ Exception: {e}")
    lark_works = False

# Summary
print("\n" + "=" * 70)
print("  RESULT")
print("=" * 70)

if feishu_works and not lark_works:
    print("\n✅ Your app is registered on: Feishu (China)")
    print("\n📋 Configuration:")
    print('   Update feishu_config.json:')
    print('   "base_url": "https://open.feishu.cn/open-apis"')
    print("\n🔗 Developer Console:")
    print("   https://open.feishu.cn/app")

elif lark_works and not feishu_works:
    print("\n✅ Your app is registered on: Lark (International)")
    print("\n📋 Configuration:")
    print('   Update feishu_config.json:')
    print('   "base_url": "https://open.larksuite.com/open-apis"')
    print("\n🔗 Developer Console:")
    print("   https://open.larksuite.com/app")
    print("\n⚠️  IMPORTANT: Your spreadsheet URL shows:")
    print("   https://bytedance.us.larkoffice.com/sheets/...")
    print("   This confirms you're on Lark (International) ✓")

elif feishu_works and lark_works:
    print("\n⚠️  UNUSUAL: App works on BOTH platforms")
    print("   This is rare - you may have duplicate apps")
    print("   Use the one matching your spreadsheet platform")

else:
    print("\n❌ App doesn't work on either platform!")
    print("\n🔧 Possible issues:")
    print("   1. App ID or App Secret is incorrect")
    print("   2. App has been disabled/deleted")
    print("   3. Network/firewall blocking requests")
    print("\n💡 Solutions:")
    print("   - Double-check credentials in developer console")
    print("   - Try creating a new app on the correct platform")
    print("   - Check your spreadsheet URL to determine platform:")
    print("     • *.feishu.cn → Use Feishu (China)")
    print("     • *.larkoffice.com → Use Lark (International)")

print("\n" + "=" * 70)
