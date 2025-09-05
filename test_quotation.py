#!/usr/bin/env python3
"""
Test Quotation Generation
Simple test to verify the quotation system works.
"""

import requests
import json

def test_quotation_generation():
    """Test the quotation generation endpoint."""
    
    print("🧪 Testing Quotation Generation")
    print("=" * 50)
    
    # Test data
    test_data = {
        "companyName": "Test Company BV",
        "contactName": "Jan Janssen", 
        "address": "Teststraat 123",
        "postalCode": "1234AB",
        "city": "Amsterdam",
        "companyId": "NL123456789B01",
        "oneTimeCosts": [
            {
                "material": "Setup kosten",
                "quantity": 1,
                "unitPrice": 500.00,
                "total": 500.00
            },
            {
                "material": "Hardware installatie",
                "quantity": 2,
                "unitPrice": 250.00,
                "total": 500.00
            }
        ],
        "recurringCosts": [
            {
                "material": "Maandelijks onderhoud",
                "quantity": 12,
                "unitPrice": 75.00,
                "total": 900.00
            }
        ]
    }
    
    try:
        print("📤 Sending test request to quotation server...")
        
        response = requests.post(
            'http://localhost:8001/generate-quotation',
            headers={'Content-Type': 'application/json'},
            json=test_data,
            timeout=30
        )
        
        print(f"📥 Response status: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            print("✅ Quotation generation successful!")
            print(f"📄 Generated file: {result.get('filename', 'Unknown')}")
            
            # Check if file exists
            import os
            filename = result.get('filename', '')
            if filename and os.path.exists(filename):
                print(f"✅ File confirmed on disk: {filename}")
                file_size = os.path.getsize(filename)
                print(f"📊 File size: {file_size} bytes")
            else:
                print("⚠️  File not found on disk")
            
        else:
            print(f"❌ Error: {response.status_code}")
            print(f"Response: {response.text}")
            
    except requests.exceptions.ConnectionError:
        print("❌ Could not connect to quotation server")
        print("   Make sure the server is running on port 8001")
    except requests.exceptions.Timeout:
        print("❌ Request timeout")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

if __name__ == "__main__":
    test_quotation_generation()