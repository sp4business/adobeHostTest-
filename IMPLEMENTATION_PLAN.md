# Adobe Data to Feishu - Implementation Plan

## Overview
This document outlines the complete implementation strategy for automatically pushing Adobe ROAS campaign data to Feishu spreadsheets using the Lark API.

---

## Current Status

### ✅ What's Working
- ✓ Feishu API authentication (tenant access token)
- ✓ Adobe email parsing and campaign data extraction
- ✓ Campaign name mapping (Adobe → Target format)
- ✓ Excel file generation with proper data structure
- ✓ Column insertion (new week columns can be added)
- ✓ Sheet metadata retrieval

### ❌ What's Missing
- ✗ Writing ROAS values to correct cells
- ✗ Batch update implementation
- ✗ Row-to-campaign mapping logic
- ✗ Complete end-to-end automation

---

## Feishu API Endpoints

### 1. Authentication
```
POST /auth/v3/tenant_access_token/internal
```
**Payload:**
```json
{
  "app_id": "your_app_id",
  "app_secret": "your_app_secret"
}
```
**Response:** `{"tenant_access_token": "..."}`

---

### 2. Get Sheet Metadata (v2)
```
GET /sheets/v2/spreadsheets/{spreadsheetToken}/metainfo
```
**Purpose:** Get sheet structure, row count, column count

---

### 3. Read Sheet Data (v2)
```
GET /sheets/v2/spreadsheets/{spreadsheetToken}/values/{range}
```
**Range Format:** `{sheet_id}!A1:Z100`
**Purpose:** Read existing data to map campaigns to rows

---

### 4. Write Single Range (v2) ⭐ RECOMMENDED
```
PUT /sheets/v2/spreadsheets/{spreadsheetToken}/values
```
**Payload:**
```json
{
  "valueRange": {
    "range": "IqubAk!D5:D15",
    "values": [
      [3.24],
      [3.56],
      [2.76],
      [3.03]
    ]
  }
}
```
**Best For:** Writing multiple values in a contiguous block (single column)

---

### 5. Batch Update Multiple Ranges (v2) ⭐ ALTERNATIVE
```
PUT /sheets/v2/spreadsheets/{spreadsheetToken}/values_batch_update
```
**Payload:**
```json
{
  "valueRanges": [
    {
      "range": "IqubAk!D5",
      "values": [[3.24]]
    },
    {
      "range": "IqubAk!D10",
      "values": [[3.56]]
    }
  ]
}
```
**Best For:** Writing to scattered cells across the sheet

---

### 6. Insert Dimension (v3)
```
POST /sheets/v3/spreadsheets/{token}/sheets/{sheet_id}/insert_dimension
```
**Purpose:** Insert new rows/columns (you've already used this!)

---

## Data Flow Architecture

```
┌─────────────────────────────────────────────────────────────┐
│ STEP 1: Parse Adobe Email                                   │
│ Input: input/week_37_email_input.txt                        │
│ Output: output/parsed_campaign_data.json                    │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ STEP 2: Get Feishu Access Token                             │
│ API: POST /auth/v3/tenant_access_token/internal             │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ STEP 3: Read Feishu Sheet Structure                         │
│ API: GET /sheets/v2/spreadsheets/{token}/values/{range}     │
│ Purpose: Map campaign names to row numbers                  │
│ Example: "US Evergreen Prospecting" → Row 10                │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ STEP 4: Find Week Column                                    │
│ - Look for "Week 37" header in row 1                        │
│ - Get column letter (A, B, C... or AA, AB...)               │
│ - If not found, insert new column first                     │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ STEP 5: Build Update Payload                                │
│ For each campaign in parsed_data:                           │
│   - Find matching row in sheet                              │
│   - Build range: "{sheet_id}!{col}{row}"                    │
│   - Add ROAS value                                          │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ STEP 6: Execute Batch Update                                │
│ API: PUT /sheets/v2/spreadsheets/{token}/values             │
│ Payload: Single range with all values in column             │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ STEP 7: Validate Write Success                              │
│ - Check response code == 0                                  │
│ - Log successful updates                                    │
│ - Retry failed updates if any                               │
└─────────────────────────────────────────────────────────────┘
```

---

## Implementation Strategy

### Option 1: Single Contiguous Range (RECOMMENDED) ✅

**Pros:**
- Most efficient (single API call)
- Simpler error handling
- Faster execution

**Cons:**
- Requires campaign rows to be in order
- Need to handle missing campaigns (insert N/A or empty)

**Implementation:**
```python
# Build values array in row order
values = []
for row in range(start_row, end_row + 1):
    campaign_name = sheet_data[row][0]  # Column A has campaign names

    # Find matching ROAS value
    roas_value = find_roas_for_campaign(campaign_name, parsed_data)
    values.append([roas_value if roas_value else "N/A"])

# Single API call
payload = {
    "valueRange": {
        "range": f"{sheet_id}!{week_column}{start_row}:{week_column}{end_row}",
        "values": values
    }
}
```

---

### Option 2: Batch Update Multiple Ranges

**Pros:**
- Can update any cells independently
- Flexible for non-contiguous data

**Cons:**
- More complex payload
- Slightly slower

**Implementation:**
```python
# Build valueRanges array
value_ranges = []
for campaign in parsed_data['campaigns']:
    row_num = find_row_for_campaign(campaign['name'])
    if row_num:
        value_ranges.append({
            "range": f"{sheet_id}!{week_column}{row_num}",
            "values": [[campaign['roas']]]
        })

# Batch update API call
payload = {"valueRanges": value_ranges}
```

---

## Critical Implementation Details

### 1. Range Format
- Format: `{sheet_id}!{column}{row}`
- Examples:
  - Single cell: `IqubAk!D15`
  - Range: `IqubAk!D5:D20`
  - Entire column: `IqubAk!D:D`

### 2. Values Format
- **Always 2D array**, even for single cell
- Single cell: `[[3.24]]` ✅ NOT `[3.24]` ❌
- Multiple cells: `[[3.24], [2.76], [3.56]]`

### 3. Column Letter Conversion
```python
def column_to_letter(col_num: int) -> str:
    """Convert 1-based column number to letter (1=A, 27=AA)"""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + 65) + letter
        col_num //= 26
    return letter
```

### 4. Campaign Name Matching
```python
def campaigns_match(excel_campaign: str, sheet_campaign: str) -> bool:
    """Fuzzy match campaign names"""
    # Normalize both strings
    excel_lower = excel_campaign.lower().strip()
    sheet_lower = sheet_campaign.lower().strip()

    # Direct match
    if excel_lower == sheet_lower:
        return True

    # Key terms matching
    key_terms = ['us', 'canada', 'evergreen', 'prospecting',
                 'retargeting', 'cns', 's+', 'cost cap']

    excel_terms = [term for term in key_terms if term in excel_lower]
    sheet_terms = [term for term in key_terms if term in sheet_lower]

    # Match if 70%+ key terms align
    if len(excel_terms) > 0 and len(sheet_terms) > 0:
        common = set(excel_terms) & set(sheet_terms)
        threshold = min(len(excel_terms), len(sheet_terms)) * 0.7
        return len(common) >= threshold

    return False
```

### 5. Error Handling
```python
response = requests.put(url, json=payload, headers=headers)

if response.status_code == 200:
    data = response.json()
    if data.get("code") == 0:
        print("✅ Success!")
        return True
    else:
        print(f"❌ API Error: {data.get('msg')}")
        return False
else:
    print(f"❌ HTTP {response.status_code}: {response.text}")
    return False
```

---

## Expected Sheet Structure

### Lululemon Final Sheet Format:
```
Row 1:  [Campaign Name] [Week 1] [Week 2] ... [Week 37] [Week 38]
Row 5:  [Overall ROAS]  [3.45]   [3.12]  ... [3.24]    [...]
Row 10: [US Evergreen Prospecting] [2.89] ... [3.56]
Row 11: [S+ 2.0 US Evergreen] [0.00] ... [2.76]
Row 12: [US Evergreen (Cost Cap)] [3.15] ... [3.03]
...
```

### What We Need to Know:
1. **Campaign row mapping** (read column A to find each campaign)
2. **Week column** (find "Week 37" in row 1)
3. **Overall ROAS row** (usually row 5, but verify)

---

## Complete Workflow Code Structure

```python
class FeishuAutomation:
    def __init__(self):
        self.config = load_config()
        self.access_token = None
        self.sheet_structure = None

    def run_automation(self, week_data_file: str):
        """Complete automation workflow"""

        # 1. Authenticate
        self.access_token = self.get_access_token()

        # 2. Load parsed Adobe data
        adobe_data = self.load_adobe_data(week_data_file)

        # 3. Read sheet structure
        self.sheet_structure = self.read_sheet_structure()

        # 4. Find/create week column
        week_col = self.find_or_create_week_column(adobe_data['week'])

        # 5. Build update payload
        update_payload = self.build_update_payload(
            adobe_data,
            week_col
        )

        # 6. Execute update
        success = self.execute_batch_update(update_payload)

        # 7. Validate
        if success:
            self.validate_written_data(adobe_data, week_col)

        return success
```

---

## Testing Strategy

### Test 1: Read Sheet Structure
```python
python -c "from 5_run_feishu_automation import *;
           auto = FeishuAutomation();
           auto.access_token = auto.get_access_token();
           structure = auto.read_sheet_structure();
           print(structure)"
```

### Test 2: Find Campaign Rows
```python
# Should output: {"US Evergreen Prospecting": 10, "S+ 2.0 US Evergreen": 11, ...}
```

### Test 3: Write Single Cell
```python
# Test writing ROAS value to one cell before batch update
auto.write_single_cell("IqubAk!D10", 3.56)
```

### Test 4: Full Batch Update
```python
python 5_run_feishu_automation.py
```

---

## Error Scenarios & Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| `code: 99991663` | Invalid token | Refresh access token |
| `code: 1254301` | Invalid range | Check sheet_id and range format |
| `HTTP 403` | Permission denied | Grant app sheet write permissions |
| `HTTP 429` | Rate limit | Add retry with exponential backoff |
| Campaign not found | Name mismatch | Improve fuzzy matching logic |

---

## Performance Considerations

- **API Rate Limits:** Feishu allows ~100 requests/minute
- **Batch Size:** Can update 5000 rows × 100 columns in one call
- **Recommended Approach:** Single batch update for all Week 37 data (~10 campaigns)
- **Execution Time:** Expected < 3 seconds for full automation

---

## Next Steps

1. ✅ Complete `5_run_feishu_automation.py` with all functions
2. ✅ Test with Week 37 data
3. ✅ Add comprehensive error handling
4. ✅ Create validation checks
5. ⏭️ Schedule weekly automation (cron job or manual trigger)

---

## References

- Feishu API Docs: https://open.feishu.cn/document/server-docs/docs/sheets-v3/overview
- Insert Dimension: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/insert_dimension
- Insert Values: https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet-value/insert
- Batch Update: https://open.feishu.cn/document/ukTMukTMukTM/uEjMzUjLxIzM14SMyMTN

---

**Document Version:** 1.0
**Last Updated:** 2025-10-24
**Project:** Lululemon Adobe-to-Feishu Automation
