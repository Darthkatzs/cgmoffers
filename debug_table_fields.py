#!/usr/bin/env python3
"""
Debug Table Fields
Check what's happening with the table field mappings.
"""

from content_control_processor import ContentControlProcessor

def main():
    """Debug the table field mappings."""
    
    processor = ContentControlProcessor()
    
    test_data = {
        "companyName": "Table Fields Test",
        "contactName": "Test User",
        "address": "Test Address 123",
        "postalCode": "1234AB",
        "city": "Test City", 
        "companyId": "TEST123",
        "description": "Testing table field population",
        "oneTimeCosts": [
            {
                "material": "Server Setup",
                "quantity": 2,
                "unitPrice": 750.00,
                "total": 1500.00
            }
        ],
        "recurringCosts": [
            {
                "material": "Monthly Support",
                "quantity": 12,
                "unitPrice": 85.00,
                "total": 1020.00
            }
        ]
    }
    
    # Get the calculations and mappings
    calculations = processor.calculate_values(test_data)
    mappings = processor.build_control_mappings(test_data, calculations)
    
    print("🔍 DEBUG: Control mappings generated:")
    print(f"📊 Total mappings: {len(mappings)}")
    
    table_fields = ['Module', 'Aantal', 'éénmalige setupkost', 'calctotaalsetup', 'Jaarlijks', 'calctotaaljaarlijks']
    
    print(f"\\n📋 Table field mappings:")
    for field in table_fields:
        value = mappings.get(field, '[NOT FOUND]')
        print(f"   {field}: '{value}'")
        
    print(f"\\n📋 Control configuration check:")
    for control_name, config in processor.controls.items():
        if control_name in table_fields:
            print(f"   {control_name}: type='{config.get('type')}', value='{config.get('value')}', formula='{config.get('formula')}'")

if __name__ == "__main__":
    main()