#!/usr/bin/env python3
"""
Lululemon Adobe ROAS Email Parser
Parses weekly performance emails and extracts campaign data with perfect mapping
"""

import re
import json
from typing import Dict, List, Optional

class CampaignMapper:
    """Handles perfect campaign name mapping from email to target CSV format"""
    
    @staticmethod
    def map_email_campaign_to_csv(email_campaign_name: str) -> Optional[str]:
        """Perfect mapping with Smart+ 2.0 and CNS logic"""
        
        if not email_campaign_name:
            return None
            
        # Geography detection
        geography = "US" if "(US)" in email_campaign_name else "CA"
        
        # Smart+ detection
        is_smart = "SMART" in email_campaign_name
        
        # Campaign type detection
        if "CNS" in email_campaign_name:
            # CNS campaigns
            if is_smart:
                if geography == "US":
                    return "US CNS S+ 2.0 (carousel ads only - Lowest Cost)"
                else:
                    return "Canada CNS S+ 2.0 (carousel ads only - Lowest Cost)"
            else:
                # Manual CNS (if exists in future)
                if geography == "US":
                    return "US CNS Manual (carousel ads only - Lowest Cost)"
                else:
                    return "Canada CNS Manual (carousel ads only - Lowest Cost)"
        
        elif "RTG" in email_campaign_name:
            # Retargeting campaigns
            if geography == "US":
                return "US Evergreen Retargeting (Lowest Cost)"
            else:
                return "Canada Evergreen Retargeting (Lowest Cost)"
        
        elif "TOF" in email_campaign_name or "PROS" in email_campaign_name:
            # Prospecting campaigns
            if "CC" in email_campaign_name:
                return "US Evergreen Prospecting (Cost Cap)"  # Only US has cost cap
            else:
                if is_smart:
                    if geography == "US":
                        return "S+ 2.0 US Evergreen Prospecting (Lowest Cost)"
                    else:
                        return "S+ 2.0 Canada Evergreen Prospecting (Lowest Cost)"
                else:
                    if geography == "US":
                        return "US Evergreen Prospecting (Lowest Cost)"
                    else:
                        return "Canada Evergreen Prospecting (Lowest Cost)"
        
        return None  # Unknown campaign type

class EmailParser:
    """Parses Adobe ROAS performance emails"""
    
    def __init__(self):
        self.campaign_mapper = CampaignMapper()
        
    def extract_week_info(self, email_text: str) -> Optional[str]:
        """Extract week number from email text"""
        week_match = re.search(r'WK\s+(\d+)', email_text, re.IGNORECASE)
        return week_match.group(1) if week_match else None
    
    def extract_overall_roas(self, email_text: str) -> Optional[float]:
        """Extract overall ROAS from summary section"""
        # Look for pattern like "3.24x ROAS"
        roas_match = re.search(r'(\d+\.?\d*)x\s+ROAS', email_text, re.IGNORECASE)
        return float(roas_match.group(1)) if roas_match else None
    
    def extract_campaign_data(self, email_text: str) -> List[Dict]:
        """Extract individual campaign data from email"""
        campaigns = []
        
        # Split email into lines for processing
        lines = email_text.strip().split('\n')
        
        # Find campaign data section (starts after "Campaign Name" header)
        in_campaign_section = False
        current_campaign = {}
        
        for line in lines:
            line = line.strip()
            
            # Skip empty lines
            if not line:
                continue
                
            # Look for campaign section start
            if "Campaign Name" in line:
                in_campaign_section = True
                continue
                
            if not in_campaign_section:
                continue
                
            # Look for TOTAL line (end of campaign data)
            if "TOTAL" in line:
                break
                
            # Check if this is a campaign name line (contains EVRGN)
            if "EVRGN" in line and not current_campaign:
                current_campaign['raw_campaign_name'] = line
                continue
                
            # Check for ad group
            if current_campaign and 'ad_group' not in current_campaign:
                if "US_NAT" in line or "CA_NAT" in line:
                    current_campaign['ad_group'] = line
                    continue
                
            # Check for spend (starts with $)
            if current_campaign and 'spend' not in current_campaign:
                if line.startswith('$') and ',' in line:
                    current_campaign['spend'] = line
                    continue
                elif line.startswith('$') and line.replace('$', '').replace(',', '').isdigit():
                    current_campaign['spend'] = line
                    continue
            
            # Check for revenue (starts with $)
            if current_campaign and 'revenue' not in current_campaign:
                if line.startswith('$') and ',' in line and 'spend' in current_campaign:
                    current_campaign['revenue'] = line
                    continue
                elif line.startswith('$') and line.replace('$', '').replace(',', '').isdigit() and 'spend' in current_campaign:
                    current_campaign['revenue'] = line
                    continue
            
            # Check for ROAS (decimal number)
            if current_campaign and 'roas' not in current_campaign:
                if re.match(r'^\d+\.\d+$', line):
                    current_campaign['roas'] = line
                    
                    # Campaign data is complete, process it
                    if self.validate_campaign_data(current_campaign):
                        processed_campaign = self.process_campaign_data(current_campaign)
                        if processed_campaign:
                            campaigns.append(processed_campaign)
                    
                    # Reset for next campaign
                    current_campaign = {}
                    continue
        
        return campaigns
    
    def validate_campaign_data(self, campaign: Dict) -> bool:
        """Validate that campaign has all required fields"""
        required_fields = ['raw_campaign_name', 'spend', 'revenue', 'roas']
        return all(field in campaign for field in required_fields)
    
    def process_campaign_data(self, campaign: Dict) -> Optional[Dict]:
        """Process raw campaign data and apply mapping"""
        try:
            # Clean and convert values
            spend = float(campaign['spend'].replace('$', '').replace(',', ''))
            revenue = float(campaign['revenue'].replace('$', '').replace(',', ''))
            roas = float(campaign['roas'])
            
            # Apply campaign mapping
            mapped_campaign_name = self.campaign_mapper.map_email_campaign_to_csv(
                campaign['raw_campaign_name']
            )
            
            if not mapped_campaign_name:
                return None
            
            return {
                'target_campaign_name': mapped_campaign_name,
                'raw_campaign_name': campaign['raw_campaign_name'],
                'ad_group': campaign.get('ad_group', ''),
                'spend': spend,
                'revenue': revenue,
                'roas': roas
            }
            
        except (ValueError, AttributeError):
            return None
    
    def parse_email(self, email_text: str) -> Dict:
        """Main parsing function"""
        week = self.extract_week_info(email_text)
        overall_roas = self.extract_overall_roas(email_text)
        campaigns = self.extract_campaign_data(email_text)
        
        return {
            'week': week,
            'overall_roas': overall_roas,
            'campaigns': campaigns,
            'total_campaigns': len(campaigns)
        }

def main():
    """Test the parser with WK 37 data"""
    
    # Read the WK 37 email file
    with open('input/week_37_email_input.txt', 'r') as f:
        email_text = f.read()
    
    # Parse the email
    parser = EmailParser()
    result = parser.parse_email(email_text)
    
    # Output results
    print("=== LULULEMON ADOBE ROAS PARSER RESULTS ===")
    print(f"Week: {result['week']}")
    print(f"Overall ROAS: {result['overall_roas']}x")
    print(f"Total Campaigns Parsed: {result['total_campaigns']}")
    print("\n=== CAMPAIGN DATA ===")
    
    for campaign in result['campaigns']:
        print(f"\nTarget: {campaign['target_campaign_name']}")
        print(f"Raw: {campaign['raw_campaign_name']}")
        print(f"Spend: ${campaign['spend']:,}")
        print(f"Revenue: ${campaign['revenue']:,}")
        print(f"ROAS: {campaign['roas']}x")
    
    # Save to JSON for further processing
    with open('output/parsed_campaign_data.json', 'w') as f:
        json.dump(result, f, indent=2)
    
    print(f"\n=== SAVED TO parsed_campaign_data.json ===")

if __name__ == "__main__":
    main()