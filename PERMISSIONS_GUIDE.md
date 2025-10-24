# Feishu/Lark App Permissions Guide

## Issue: Authentication 403 Error

Your app credentials are correct (they work for creating spreadsheets), but getting 403 on authentication suggests a **platform mismatch** issue.

---

## CRITICAL: Which Platform is Your App On?

There are TWO separate Feishu/Lark platforms:

### 1. Feishu (China) 🇨🇳
- **Developer Console**: https://open.feishu.cn
- **API Endpoint**: `https://open.feishu.cn/open-apis`
- **Sheet URL Pattern**: `https://xxx.feishu.cn/sheets/...`
- **Data Location**: Beijing, China

### 2. Lark (International) 🌎
- **Developer Console**: https://open.larksuite.com
- **API Endpoint**: `https://open.larksuite.com/open-apis`
- **Sheet URL Pattern**: `https://bytedance.us.larkoffice.com/sheets/...` (US)
- **Data Location**: Singapore (international)

### Your Current Setup:
- App ID: `cli_a86397ecf2f89013`
- Spreadsheet URL: `https://bytedance.us.larkoffice.com/sheets/SBYJsQ4KkhQ1Svti88ouShLtsK0`
- This URL shows you're using **Lark (International/US)**

### The Problem:
If your app was created on Feishu (China) but your spreadsheet is on Lark (US), they won't work together!

---

## Step 1: Verify Your App Platform

**Check where your app was created:**

1. Try logging in to Feishu Developer Console:
   - https://open.feishu.cn/app
   - If you can see app `cli_a86397ecf2f89013` here → App is on Feishu (China)

2. Try logging in to Lark Developer Console:
   - https://open.larksuite.com/app
   - If you can see app `cli_a86397ecf2f89013` here → App is on Lark (International)

**Important**: Your spreadsheet URL (`bytedance.us.larkoffice.com`) indicates you need a **Lark (International)** app.

---

## Step 2: Fix Platform Mismatch

### Option A: If App is on Feishu (China) [Most Likely]

You need to create a NEW app on Lark (International):

1. Go to: https://open.larksuite.com
2. Click "Create Custom App"
3. Name it "Lululemon Adobe Automation" (or similar)
4. Save the new App ID and App Secret
5. Update `feishu_config.json` with new credentials

### Option B: If App is on Lark (International)

Update your config to use Lark API endpoint:

```json
{
  "app_id": "cli_a86397ecf2f89013",
  "app_secret": "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt",
  "spreadsheet_token": "SBYJsQ4KkhQ1Svti88ouShLtsK0",
  "sheet_id": "IqubAk",
  "base_url": "https://open.larksuite.com/open-apis"
}
```

---

## Step 3: Add Required Permissions

Once you're on the correct platform, add these permissions:

### For Lark (International):
1. Go to: https://open.larksuite.com/app
2. Select your app
3. Navigate to: **Permissions & Scopes**
4. Add permission:

   **"View, create, edit, and delete Sheets"**
   - Scope: `sheets:spreadsheet`
   - This is the ALL-IN-ONE permission for sheets

5. **CRITICAL**: Click "Create Version" to publish
6. Wait 1-2 minutes for changes to take effect

### For Feishu (China):
1. Go to: https://open.feishu.cn/app
2. Select your app
3. Navigate to: **权限管理** (Permissions & Scopes)
4. Add permission:

   **查看、评论、编辑和管理电子表格**
   - Scope: `sheets:spreadsheet`

5. **CRITICAL**: Click **创建版本** (Create Version)
6. Wait 1-2 minutes

---

## Step 4: Re-run Diagnostic

After fixing the platform and adding permissions:

```bash
python3 test_permissions_layered.py
```

Expected output when working:
```
✅ PASS  Authentication
✅ PASS  Read Metadata
✅ PASS  Read Data
✅ PASS  Write Data
✅ PASS  Insert Dimension
✅ PASS  Batch Update
```

---

## Required Permissions Summary

For full Adobe → Feishu automation, you need **ONE** permission:

| Permission Name | Scope | What it Enables |
|----------------|-------|-----------------|
| View, create, edit, and delete Sheets | `sheets:spreadsheet` | Read metadata, read data, write data, insert rows/columns, batch updates |

### Alternative (Read-Only):
If you only need to read:

| Permission Name | Scope | What it Enables |
|----------------|-------|-----------------|
| View, comment, and export Sheets | `sheets:spreadsheet:readonly` | Read metadata, read data only |

---

## Documentation References

### Official API Docs:
- **Lark Scope List**: https://open.larksuite.com/document/server-docs/getting-started/scope-list
- **Feishu Scope List**: https://open.feishu.cn/document/server-docs/application-scope/scope-list
- **Sheets API Overview**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/overview

### Specific Operations:
1. **Read Data**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/reading-a-single-range
   - Permission: `sheets:spreadsheet:readonly` or `sheets:spreadsheet`

2. **Write Data**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-a-single-range
   - Permission: `sheets:spreadsheet`

3. **Batch Update**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-multiple-ranges
   - Permission: `sheets:spreadsheet`

4. **Insert Row/Column**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/insert_dimension
   - Permission: `sheets:spreadsheet`

---

## Troubleshooting Common Errors

### Error: 403 on Authentication
- **Cause**: Platform mismatch (app on Feishu, sheet on Lark)
- **Fix**: Create app on matching platform

### Error: 99991663 (Invalid Token)
- **Cause**: Token expired or invalid credentials
- **Fix**: Regenerate token, check app_id/app_secret

### Error: 1254301 (Invalid Range)
- **Cause**: Incorrect sheet_id or range format
- **Fix**: Verify sheet_id, use format: `sheetId!A1:B2`

### Error: Permission Denied on Read/Write
- **Cause**: Missing `sheets:spreadsheet` permission
- **Fix**: Add permission and publish new version

---

## Next Steps

1. **Identify platform** (Feishu CN vs Lark International)
2. **Create/verify app** on correct platform
3. **Add permission**: `sheets:spreadsheet`
4. **Publish version** (critical!)
5. **Update config** with correct base_url
6. **Run test**: `python3 test_permissions_layered.py`
7. **Run automation**: `python3 5_run_feishu_automation.py`

---

## Quick Test: Which Platform Am I On?

Run this to test both endpoints:

```python
import requests

app_id = "cli_a86397ecf2f89013"
app_secret = "nhZh2vVIrdngJ42LMGwA7fuB5AHp5RXt"

# Test Feishu (China)
print("Testing Feishu (China)...")
r1 = requests.post(
    "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": app_id, "app_secret": app_secret}
)
print(f"  Status: {r1.status_code}")
if r1.status_code == 200:
    print(f"  ✅ App exists on Feishu (China)")

# Test Lark (International)
print("\nTesting Lark (International)...")
r2 = requests.post(
    "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": app_id, "app_secret": app_secret}
)
print(f"  Status: {r2.status_code}")
if r2.status_code == 200:
    print(f"  ✅ App exists on Lark (International)")
```

The one that returns 200 is your platform!

---

**Last Updated**: 2025-10-24
**Project**: Lululemon Adobe-to-Feishu Automation
