# Feishu Configuration Template

This file contains the configuration for Feishu API integration.

## Setup Instructions

1. **Get Feishu API Credentials:**
   - Go to Feishu Developer Console
   - Create a new app or use existing one
   - Get App ID and App Secret

2. **Get Spreadsheet Info:**
   - Open your target Feishu spreadsheet
   - Extract spreadsheet token from URL
   - Get sheet ID from sheet properties

3. **Update Configuration:**
   - Replace placeholder values in `feishu_config.json`
   - Keep this file secure (add to .gitignore)

## Security Notes

- Never commit actual credentials to version control
- Use environment variables for production
- Rotate API keys regularly
- Limit app permissions to minimum required

## File Structure

```json
{
  "app_id": "cli_xxxxxxxxxxxxxxxx",
  "app_secret": "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  "spreadsheet_token": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
  "sheet_id": "XXXXXXXX"
}
```