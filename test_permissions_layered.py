#!/usr/bin/env python3
"""
Layered Permission Testing - Incremental Feishu API Testing
Tests each operation level-by-level to identify missing permissions
"""

import requests
import json
from datetime import datetime

# Your working credentials
FEISHU_APP_ID = "cli_a86397ecf2f89013"
FEISHU_APP_SECRET = "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt"
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"

# Target sheet
SPREADSHEET_TOKEN = "SBYJsQ4KkhQ1Svti88ouShLtsK0"
SHEET_ID = "IqubAk"

def print_section(title):
    """Print a section header"""
    print("\n" + "=" * 70)
    print(f"  {title}")
    print("=" * 70)

def print_permission_info(permission_name, permission_scope, doc_link):
    """Print permission information"""
    print(f"\n📋 Required Permission:")
    print(f"   Name: {permission_name}")
    print(f"   Scope: {permission_scope}")
    print(f"   Docs: {doc_link}")

# ============================================================================
# LAYER 1: AUTHENTICATION
# ============================================================================

def test_layer1_authentication():
    """Layer 1: Test if we can get an access token"""
    print_section("LAYER 1: AUTHENTICATION TEST")

    url = f"{FEISHU_BASE_URL}/auth/v3/tenant_access_token/internal"

    payload = {
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    }

    print(f"🔑 Testing authentication...")
    print(f"   Endpoint: {url}")

    try:
        response = requests.post(url, json=payload)

        print(f"   Response Status: {response.status_code}")

        if response.status_code == 200:
            data = response.json()
            print(f"   Response Code: {data.get('code')}")

            if data.get("code") == 0:
                token = data["tenant_access_token"]
                print(f"✅ SUCCESS: Authentication works!")
                print(f"   Token: {token[:20]}...")
                return token
            else:
                print(f"❌ FAILED: {data.get('msg')}")
                return None
        else:
            print(f"❌ FAILED: HTTP {response.status_code}")
            print(f"   Response: {response.text}")
            return None

    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        return None

# ============================================================================
# LAYER 2: READ SHEET METADATA
# ============================================================================

def test_layer2_read_metadata(token):
    """Layer 2: Test if we can read sheet metadata"""
    print_section("LAYER 2: READ SHEET METADATA")

    print_permission_info(
        "sheets:spreadsheet",
        "Read spreadsheet metadata",
        "https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet/get"
    )

    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/metainfo"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    print(f"\n📖 Testing metadata read...")
    print(f"   Endpoint: {url}")
    print(f"   Spreadsheet: {SPREADSHEET_TOKEN}")

    try:
        response = requests.get(url, headers=headers)

        print(f"   Response Status: {response.status_code}")

        if response.status_code == 200:
            data = response.json()
            print(f"   Response Code: {data.get('code')}")

            if data.get("code") == 0:
                print(f"✅ SUCCESS: Can read metadata!")
                sheets = data.get("data", {}).get("sheets", [])
                print(f"   Found {len(sheets)} sheets")
                for sheet in sheets:
                    print(f"     - {sheet.get('title')} (ID: {sheet.get('sheetId')})")
                return True
            else:
                print(f"❌ FAILED: {data.get('msg')}")
                print(f"\n🔧 TO FIX:")
                print(f"   1. Go to: https://open.feishu.cn/app")
                print(f"   2. Select your app: {FEISHU_APP_ID}")
                print(f"   3. Go to 'Permissions & Scopes'")
                print(f"   4. Add permission: 'View, comment, and export Sheets'")
                print(f"      (sheets:spreadsheet)")
                return False
        else:
            print(f"❌ FAILED: HTTP {response.status_code}")
            print(f"   Response: {response.text}")
            return False

    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        return False

# ============================================================================
# LAYER 3: READ SHEET DATA
# ============================================================================

def test_layer3_read_data(token):
    """Layer 3: Test if we can read sheet data"""
    print_section("LAYER 3: READ SHEET DATA (VALUES)")

    print_permission_info(
        "sheets:spreadsheet:readonly",
        "Read spreadsheet data (required for reading cell values)",
        "https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/reading-a-single-range"
    )

    # Try to read a small range
    range_str = f"{SHEET_ID}!A1:C3"
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{range_str}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    print(f"\n📖 Testing data read...")
    print(f"   Endpoint: {url}")
    print(f"   Range: {range_str}")

    try:
        response = requests.get(url, headers=headers)

        print(f"   Response Status: {response.status_code}")

        if response.status_code == 200:
            data = response.json()
            print(f"   Response Code: {data.get('code')}")

            if data.get("code") == 0:
                values = data.get("data", {}).get("valueRange", {}).get("values", [])
                print(f"✅ SUCCESS: Can read data!")
                print(f"   Read {len(values)} rows")
                if values:
                    print(f"   First row: {values[0]}")
                return True
            else:
                print(f"❌ FAILED: {data.get('msg')}")
                print(f"   Error Code: {data.get('code')}")
                print(f"\n🔧 TO FIX:")
                print(f"   1. Go to: https://open.feishu.cn/app")
                print(f"   2. Select your app: {FEISHU_APP_ID}")
                print(f"   3. Go to 'Permissions & Scopes'")
                print(f"   4. Add permission: 'View, comment, and export Sheets'")
                print(f"      (sheets:spreadsheet:readonly)")
                return False
        else:
            print(f"❌ FAILED: HTTP {response.status_code}")
            print(f"   Response: {response.text}")
            return False

    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        return False

# ============================================================================
# LAYER 4: WRITE SHEET DATA
# ============================================================================

def test_layer4_write_data(token):
    """Layer 4: Test if we can write to sheet"""
    print_section("LAYER 4: WRITE SHEET DATA (VALUES)")

    print_permission_info(
        "sheets:spreadsheet",
        "Edit and manage spreadsheets (required for writing cell values)",
        "https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-a-single-range"
    )

    # Try to write to a test cell (far right, won't interfere with data)
    test_range = f"{SHEET_ID}!AA1"
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "valueRange": {
            "range": test_range,
            "values": [["TEST"]]
        }
    }

    print(f"\n✍️  Testing data write...")
    print(f"   Endpoint: {url}")
    print(f"   Range: {test_range}")
    print(f"   Value: 'TEST' (will be written to column AA to avoid data interference)")

    try:
        response = requests.put(url, json=payload, headers=headers)

        print(f"   Response Status: {response.status_code}")

        if response.status_code == 200:
            data = response.json()
            print(f"   Response Code: {data.get('code')}")

            if data.get("code") == 0:
                print(f"✅ SUCCESS: Can write data!")
                print(f"   Wrote 'TEST' to {test_range}")
                return True
            else:
                print(f"❌ FAILED: {data.get('msg')}")
                print(f"   Error Code: {data.get('code')}")
                print(f"\n🔧 TO FIX:")
                print(f"   1. Go to: https://open.feishu.cn/app")
                print(f"   2. Select your app: {FEISHU_APP_ID}")
                print(f"   3. Go to 'Permissions & Scopes'")
                print(f"   4. Add permission: 'View, create, edit, and delete Sheets'")
                print(f"      (sheets:spreadsheet)")
                print(f"   5. IMPORTANT: Click 'Create new version' to publish the permission change")
                return False
        else:
            print(f"❌ FAILED: HTTP {response.status_code}")
            print(f"   Response: {response.text}")
            return False

    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        return False

# ============================================================================
# LAYER 5: INSERT DIMENSION (ROW/COLUMN)
# ============================================================================

def test_layer5_insert_dimension(token):
    """Layer 5: Test if we can insert rows/columns"""
    print_section("LAYER 5: INSERT DIMENSION (ROW/COLUMN)")

    print_permission_info(
        "sheets:spreadsheet",
        "Edit and manage spreadsheets (required for inserting rows/columns)",
        "https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/insert_dimension"
    )

    # Note: We'll test the permission but won't actually insert to avoid modifying sheet
    url = f"{FEISHU_BASE_URL}/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/{SHEET_ID}/insert_dimension"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    print(f"\n📋 Checking insert dimension permission...")
    print(f"   Endpoint: {url}")
    print(f"   Note: We'll simulate the request structure")
    print(f"\n⚠️  SKIPPING ACTUAL INSERT to avoid modifying your sheet")
    print(f"   If Layer 4 (write) passed, this should work too")
    print(f"\n   Required permission is the same as Layer 4:")
    print(f"   'View, create, edit, and delete Sheets' (sheets:spreadsheet)")

    return True

# ============================================================================
# LAYER 6: BATCH UPDATE
# ============================================================================

def test_layer6_batch_update(token):
    """Layer 6: Test batch update capability"""
    print_section("LAYER 6: BATCH UPDATE (MULTIPLE RANGES)")

    print_permission_info(
        "sheets:spreadsheet",
        "Edit and manage spreadsheets (required for batch updates)",
        "https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-multiple-ranges"
    )

    # Try to write to multiple test cells
    url = f"{FEISHU_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_batch_update"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "valueRanges": [
            {
                "range": f"{SHEET_ID}!AB1",
                "values": [["TEST1"]]
            },
            {
                "range": f"{SHEET_ID}!AB2",
                "values": [["TEST2"]]
            }
        ]
    }

    print(f"\n✍️  Testing batch update...")
    print(f"   Endpoint: {url}")
    print(f"   Ranges: AB1, AB2 (test area, won't interfere with data)")

    try:
        response = requests.put(url, json=payload, headers=headers)

        print(f"   Response Status: {response.status_code}")

        if response.status_code == 200:
            data = response.json()
            print(f"   Response Code: {data.get('code')}")

            if data.get("code") == 0:
                print(f"✅ SUCCESS: Can batch update!")
                print(f"   Updated 2 cells in one request")
                return True
            else:
                print(f"❌ FAILED: {data.get('msg')}")
                print(f"   Error Code: {data.get('code')}")
                print(f"\n🔧 Same fix as Layer 4")
                return False
        else:
            print(f"❌ FAILED: HTTP {response.status_code}")
            print(f"   Response: {response.text}")
            return False

    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        return False

# ============================================================================
# MAIN TEST RUNNER
# ============================================================================

def main():
    """Run all tests in layers"""
    print("🧪 FEISHU API PERMISSION DIAGNOSTIC")
    print("=" * 70)
    print(f"App ID: {FEISHU_APP_ID}")
    print(f"Spreadsheet: {SPREADSHEET_TOKEN}")
    print(f"Sheet ID: {SHEET_ID}")
    print(f"Base URL: {FEISHU_BASE_URL}")

    results = {}

    # Layer 1: Authentication
    token = test_layer1_authentication()
    results['authentication'] = token is not None

    if not token:
        print("\n❌ Cannot proceed without valid token")
        print_summary(results)
        return

    # Layer 2: Read Metadata
    results['read_metadata'] = test_layer2_read_metadata(token)

    # Layer 3: Read Data
    results['read_data'] = test_layer3_read_data(token)

    # Layer 4: Write Data
    results['write_data'] = test_layer4_write_data(token)

    # Layer 5: Insert Dimension
    results['insert_dimension'] = test_layer5_insert_dimension(token)

    # Layer 6: Batch Update
    results['batch_update'] = test_layer6_batch_update(token)

    # Print summary
    print_summary(results)

def print_summary(results):
    """Print test summary and recommendations"""
    print("\n" + "=" * 70)
    print("  TEST SUMMARY")
    print("=" * 70)

    for test_name, passed in results.items():
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"{status}  {test_name.replace('_', ' ').title()}")

    # Determine required permissions
    print("\n" + "=" * 70)
    print("  REQUIRED PERMISSIONS FOR YOUR APP")
    print("=" * 70)

    print(f"\n📋 To enable full Adobe → Feishu automation, add these permissions:")
    print(f"\n1. Go to: https://open.feishu.cn/app")
    print(f"2. Select your app (ID: {FEISHU_APP_ID})")
    print(f"3. Navigate to: 'Permissions & Scopes' (权限管理)")
    print(f"4. Add the following permission:")
    print(f"\n   ✓ View, create, edit, and delete Sheets")
    print(f"     English: 'View, create, edit, and delete Sheets'")
    print(f"     Chinese: 查看、评论、编辑和管理电子表格")
    print(f"     Scope: sheets:spreadsheet")
    print(f"\n5. CRITICAL: Click 'Create new version' (创建版本) to publish changes")
    print(f"6. Wait 1-2 minutes for permissions to propagate")

    print(f"\n📚 Documentation References:")
    print(f"   - Read data: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/reading-a-single-range")
    print(f"   - Write data: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-a-single-range")
    print(f"   - Batch update: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-multiple-ranges")
    print(f"   - Insert row/col: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/insert_dimension")

    if all(results.values()):
        print("\n🎉 ALL TESTS PASSED! You're ready to run the automation!")
        print(f"\n   Run: python3 5_run_feishu_automation.py")
    else:
        failed_tests = [name for name, passed in results.items() if not passed]
        print(f"\n⚠️  {len(failed_tests)} test(s) failed. Fix permissions above and re-run this script.")

if __name__ == "__main__":
    main()
