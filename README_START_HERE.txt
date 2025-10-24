================================================================================
  ADOBE TO FEISHU AUTOMATION - START HERE
================================================================================

Your automation is READY! Just need to fix app credentials.

PROBLEM:
--------
App credentials (cli_a86397ecf2f89013) are not authenticating.
This could mean:
  - Credentials have changed/expired
  - App was recreated with new ID
  - App is on wrong platform (Feishu vs Lark)

YOUR SPREADSHEET:
----------------
URL: https://bytedance.us.larkoffice.com/sheets/SBYJsQ4KkhQ1Svti88ouShLtsK0
Platform: Lark (International/US)
→ You MUST create app on: https://open.larksuite.com

QUICK FIX (10 minutes):
-----------------------
1. Go to: https://open.larksuite.com
2. Create new app: "Lululemon Adobe Automation"
3. Add permission: "View, create, edit, and delete Sheets"
4. Click "Create Version" to publish
5. Copy App ID + App Secret
6. Update feishu_config.json with new credentials
7. Run: python3 test_permissions_layered.py
8. Run: python3 5_run_feishu_automation.py

DOCUMENTATION:
--------------
📘 QUICK_START_GUIDE.md     ← Start here!
📘 PERMISSIONS_GUIDE.md     ← Permission troubleshooting
📘 IMPLEMENTATION_PLAN.md   ← Full technical details

TESTING SCRIPTS:
----------------
🧪 test_which_platform.py        ← Test if app is on Feishu or Lark
🧪 test_permissions_layered.py   ← Test each permission incrementally

AUTOMATION SCRIPT:
------------------
🚀 5_run_feishu_automation.py    ← Main script (ready to use)

NEED HELP?
----------
1. Read QUICK_START_GUIDE.md
2. Run test_which_platform.py to diagnose platform
3. Run test_permissions_layered.py to verify permissions
4. Check PERMISSIONS_GUIDE.md for specific error codes

EXPECTED WORKFLOW (after setup):
---------------------------------
Week 38 arrives:
1. Copy email → input/week_38_email_input.txt
2. Run: python3 parsers.py
3. Run: python3 5_run_feishu_automation.py
4. Done! Data in Lark sheet

================================================================================
