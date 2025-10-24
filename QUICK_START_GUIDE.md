# Quick Start Guide - Adobe to Feishu Automation

## Current Situation

Your spreadsheet is on **Lark (US)**: `https://bytedance.us.larkoffice.com/sheets/SBYJsQ4KkhQ1Svti88ouShLtsK0`

But authentication with app `cli_a86397ecf2f89013` is failing on BOTH platforms.

Since you mentioned your `create_spreadsheet` code works, there might be:
1. **Updated credentials** you're using elsewhere
2. **Different app** for spreadsheet creation
3. **Expired credentials** (app was disabled/recreated)

---

## Recommended Action Plan

### STEP 1: Verify/Create Lark App (5 minutes)

Since your spreadsheet is on `bytedance.us.larkoffice.com`, you need a **Lark (International)** app.

**Go to**: https://open.larksuite.com

#### Option A: If App Exists
1. Login and find app `cli_a86397ecf2f89013`
2. Go to "Credentials & Basic Info"
3. Copy the **current** App ID and App Secret
4. They may have changed!

#### Option B: Create New App (Recommended)
1. Click "Create Custom App"
2. App Name: "Lululemon Adobe ROAS Automation"
3. Description: "Automated Adobe campaign data to Lark Sheets"
4. Save App ID and App Secret from "Credentials & Basic Info"

---

### STEP 2: Add Required Permissions (2 minutes)

1. In your app, go to: **Permissions & Scopes**

2. Search for and add:
   - **"View, create, edit, and delete Sheets"**
   - Scope identifier: `sheets:spreadsheet`

3. **CRITICAL**: Click "Create Version" button
   - Without this, permissions won't activate!

4. Wait 1-2 minutes for propagation

---

### STEP 3: Update Configuration

Update `feishu_config.json` with your app credentials:

```json
{
  "app_id": "YOUR_NEW_APP_ID_HERE",
  "app_secret": "YOUR_NEW_APP_SECRET_HERE",
  "spreadsheet_token": "SBYJsQ4KkhQ1Svti88ouShLtsK0",
  "sheet_id": "IqubAk",
  "base_url": "https://open.larksuite.com/open-apis"
}
```

---

### STEP 4: Test Connection

Run the platform detector:

```bash
python3 test_which_platform.py
```

Expected output when working:
```
✅ Your app is registered on: Lark (International)
```

---

### STEP 5: Test Permissions Layer-by-Layer

```bash
python3 test_permissions_layered.py
```

This will test each permission incrementally:
- ✅ Layer 1: Authentication
- ✅ Layer 2: Read Metadata
- ✅ Layer 3: Read Data
- ✅ Layer 4: Write Data
- ✅ Layer 5: Insert Dimension
- ✅ Layer 6: Batch Update

If any layer fails, the script tells you exactly which permission to add.

---

### STEP 6: Run Full Automation

Once all tests pass:

```bash
python3 5_run_feishu_automation.py
```

This will:
1. Authenticate with Lark
2. Read your Lululemon tracking sheet structure
3. Find/create Week 37 column
4. Map Adobe campaigns to sheet rows
5. Batch update all ROAS values
6. Validate success

---

## Troubleshooting

### Issue: "Access denied" (403) during authentication

**Causes:**
- Wrong platform (app on Feishu, sheet on Lark)
- Incorrect credentials
- App disabled/deleted

**Solutions:**
1. Verify app exists at https://open.larksuite.com/app
2. Copy fresh credentials from "Credentials & Basic Info"
3. If needed, create a new app

---

### Issue: Authentication works, but "Permission denied" on read/write

**Cause:** Missing `sheets:spreadsheet` permission

**Solution:**
1. Go to app → "Permissions & Scopes"
2. Add "View, create, edit, and delete Sheets"
3. **Click "Create Version"** (don't skip this!)
4. Wait 2 minutes
5. Re-run tests

---

### Issue: Can't find my app in developer console

**Solution:**
- Check you're on the right platform:
  - **Lark (International)**: https://open.larksuite.com
  - **Feishu (China)**: https://open.feishu.cn
- Your sheet URL (`bytedance.us.larkoffice.com`) means you need Lark International

---

### Issue: Tests pass but automation fails on campaign matching

**Cause:** Campaign names in sheet don't match Adobe campaign names

**Solution:**
- The script uses fuzzy matching (70% keyword match)
- Check sheet row 1 for campaign names
- Verify they contain keywords like: "US", "Canada", "Evergreen", "Prospecting", "S+ 2.0", "CNS"
- See `campaign_mapping_framework.md` for mapping rules

---

## Files Created for You

### Testing & Diagnostics:
- `test_which_platform.py` - Detect if app is on Feishu or Lark
- `test_permissions_layered.py` - Test each permission level incrementally

### Documentation:
- `IMPLEMENTATION_PLAN.md` - Full technical implementation details
- `PERMISSIONS_GUIDE.md` - Complete permissions reference
- `QUICK_START_GUIDE.md` - This file

### Production Scripts:
- `5_run_feishu_automation.py` - Main automation script (READY TO USE)
- `feishu_config.json` - Configuration file (needs your credentials)

---

## Expected Workflow (After Setup)

### Weekly Process:
1. Receive Adobe ROAS email (e.g., Week 38)
2. Copy email content to `input/week_38_email_input.txt`
3. Run: `python3 parsers.py` (extracts campaign data)
4. Run: `python3 5_run_feishu_automation.py` (pushes to Lark)
5. Done! Data is now in your Lululemon tracking sheet

### Automation Output:
```
🎯 FEISHU AUTOMATION - Adobe ROAS Data Transfer
✅ Access token obtained
📊 Sheet structure analyzed
✅ Found existing Week 37 column
📝 Building update payload for Week 37
   ✓ Overall ROAS: 3.24 → D5
   ✓ US Evergreen Prospecting: 3.56 → D10
   ✓ S+ 2.0 US Evergreen: 2.76 → D11
   ... (9 campaigns total)
🚀 Executing batch update
✅ Successfully updated 10 cells!
🎉 AUTOMATION COMPLETE!
```

---

## Permissions Needed Summary

| Operation | Permission Scope | How to Add |
|-----------|-----------------|------------|
| Read sheet structure | `sheets:spreadsheet` | App → Permissions & Scopes |
| Read cell values | `sheets:spreadsheet` | → Add "View, create, edit, and delete Sheets" |
| Write cell values | `sheets:spreadsheet` | → Create Version |
| Insert rows/columns | `sheets:spreadsheet` | → Wait 2 minutes |
| Batch updates | `sheets:spreadsheet` | Same as above |

**Note**: One permission (`sheets:spreadsheet`) gives you ALL capabilities!

---

## Next Steps - Do This Now

1. ✅ Go to https://open.larksuite.com and verify/create your app
2. ✅ Add permission: "View, create, edit, and delete Sheets"
3. ✅ Click "Create Version" to publish
4. ✅ Copy App ID and App Secret to `feishu_config.json`
5. ✅ Update `base_url` to `https://open.larksuite.com/open-apis`
6. ✅ Run: `python3 test_which_platform.py`
7. ✅ Run: `python3 test_permissions_layered.py`
8. ✅ Run: `python3 5_run_feishu_automation.py`

---

## Support & Documentation

### Lark Developer Docs:
- **Console**: https://open.larksuite.com/app
- **API Docs**: https://open.larksuite.com/document
- **Scope List**: https://open.larksuite.com/document/server-docs/getting-started/scope-list

### Specific Endpoints:
- **Sheets API**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/overview
- **Write Data**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-a-single-range
- **Batch Update**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/data-operation/write-data-to-multiple-ranges
- **Insert Row/Col**: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/insert_dimension

---

**Last Updated**: 2025-10-24
**Status**: Ready for testing once credentials are updated
