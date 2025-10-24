# Lululemon TikTok-Adobe Campaign Mapping & Correlation Framework

## Campaign Name Decoding Key

### Adobe to TikTok Name Mapping
| Adobe Campaign Type | TikTok Abbreviation | Full Pattern |
|-------------------|-------------------|--------------|
| Evergreen | EVRGN | Contains "EVRGN" |
| Prospecting | TOF or PROS | Contains "TOF" or "PROS" |
| Cost Cap | CC | Contains "CC" |
| Smart+ 2.0 | SMART | Contains "SMART" → Maps to S+ 2.0 |
| Conversions/Performance | CNS | Contains "CNS" (Best Sellers/Conversions) |
| Re-engagement | REENG | Contains "REENG" |
| Retargeting | RTG | Contains "RTG" |

### Geographic Identifiers
- **CA**: Canada campaigns
- **US**: United States campaigns

## Smart+ 2.0 Product Suite Logic
**Rule**: Any campaign containing "SMART" automatically maps to S+ 2.0 version in target CSV
**Examples**:
- `EVRGN-SMART_PROS_DM (US)` → `S+ 2.0 US Evergreen Prospecting (Lowest Cost)`
- `EVRGN-CNS-SMART_PROS_DM (CA)` → `Canada CNS S+ 2.0 (carousel ads only - Lowest Cost)`

## CNS Campaign Category
**CNS = Best Sellers/Conversions Performance Campaigns**
- Separate from funnel campaigns (Prospecting/Retargeting/Re-engagement)
- Has both Manual and S+ 2.0 versions

## Campaign Matching Methodology

### Step 1: Extract Campaign Names from Adobe
1. Pull all campaign names from Adobe ROAS tracker
2. Categorize by type using patterns above
3. Note geography (CA/US) from campaign names

### Step 2: Decode TikTok Campaign Names
1. Export all TikTok campaign names from TTAM
2. Apply decoding key to identify campaign types
3. Match geography based on CA/US presence

### Step 3: 1:1 Campaign Mapping
**Perfect Mapping Logic (Updated for Smart+ 2.0 & CNS):**

| Email Campaign Name | Target CSV Campaign Name | Smart+ Version |
|-------------------|------------------------|---------------|
| `EVRGN-SMART_PROS_DM (US)` | `S+ 2.0 US Evergreen Prospecting (Lowest Cost)` | **S+ 2.0** |
| `EVRGN_TOF_DM (US)` | `US Evergreen Prospecting (Lowest Cost)` | Manual |
| `EVRGN_TOF_DM (US) CC` | `US Evergreen Prospecting (Cost Cap)` | Manual |
| `EVRGN_RTG_DM (US)` | `US Evergreen Retargeting (Lowest Cost)` | Manual |
| `EVRGN-SMART_PROS_DM (CA)` | `S+ 2.0 Canada Evergreen Prospecting (Lowest Cost)` | **S+ 2.0** |
| `EVRGN_TOF_DM (CA)` | `Canada Evergreen Prospecting (Lowest Cost)` | Manual |
| `EVRGN_RTG_DM (CA)` | `Canada Evergreen Retargeting (Lowest Cost)` | Manual |
| `EVRGN-CNS-SMART_PROS_DM (US)` | `US CNS S+ 2.0 (carousel ads only - Lowest Cost)` | **S+ 2.0** |
| `EVRGN-CNS-SMART_PROS_DM (CA)` | `Canada CNS S+ 2.0 (carousel ads only - Lowest Cost)` | **S+ 2.0** |

**Priority Matching Order:**
1. Geography match (CA/US)
2. Campaign type match (CNS vs Prospecting/Retargeting)
3. Smart+ detection (SMART → S+ 2.0)
4. Cost cap detection (CC)
5. Time period alignment

**Priority Matching Order:**
1. Geography match (CA/US)
2. Campaign type match (CNS vs Prospecting/Retargeting)
3. Smart+ detection (SMART → S+ 2.0)
4. Cost cap detection (CC)
5. Time period alignment

### Step 4: Validation Checklist
For each matched pair:
- [ ] Same geography (CA/US)
- [ ] Same campaign type
- [ ] Similar spend levels (±30%)
- [ ] Active in same time period
- [ ] Minimum $500 weekly spend

## Correlation Analysis Process

### Data Preparation
1. **Adobe Data**: Weekly ROAS from tracker (use as baseline)
2. **TikTok Data**: Daily Purchase ROAS + CPA from TTAM
3. **Time Alignment**: Aggregate TikTok to weekly periods matching Adobe

### Analysis Table Template
| Matched Campaign | Geography | Type | Adobe ROAS | TikTok ROAS | TikTok CPA | ROAS Gap | Notes |
|------------------|-----------|------|------------|-------------|------------|----------|--------|
| Campaign_1 | US | EVRGN | $X.XX | $Y.YY | $Z.ZZ | X.Yx | |
| Campaign_2 | CA | PROS | $X.XX | $Y.YY | $Z.ZZ | X.Yx | |

### Correlation Calculations
- **Primary**: TikTok Purchase ROAS vs Adobe ROAS
- **Secondary**: TikTok CPA vs Adobe ROAS (expect inverse)
- **Segmented by**: Campaign type, geography

## Expected Challenges & Solutions

**Challenge 1**: Naming inconsistencies
**Solution**: Fuzzy matching on key terms (EVRGN, PROS, etc.)

**Challenge 2**: Multiple TikTok campaigns map to one Adobe campaign
**Solution**: Aggregate TikTok campaigns before correlation

**Challenge 3**: Time period mismatches
**Solution**: Use overlapping weeks only (Sept 1 - Oct 5)

## Quick Start Process
1. **Export**: Get Adobe campaign list from ROAS tracker
2. **Decode**: Apply naming key to identify campaign types
3. **Match**: Find corresponding TikTok campaigns in TTAM export
4. **Validate**: Check spend levels and time periods align
5. **Analyze**: Run correlation on matched pairs only