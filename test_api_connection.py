#!/usr/bin/env python3
"""
Test script to verify Microsoft Graph API connection for Bookings
Run this first to make sure your credentials work before running the full server
"""

import requests
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication

# UPDATE THESE WITH YOUR ACTUAL VALUES
CONFIG = {
    'client_id': '',
    'client_secret': '', 
    'tenant_id': '',
    'booking_business_id': 'PhoenixSparkUpperConferenceRoom@TravisSpark.onmicrosoft.com'
}

def test_connection():
    """Test the Microsoft Graph API connection"""
    
    print("Testing Microsoft Graph API connection...")
    print("=" * 50)
    
    # Create MSAL app
    app = ConfidentialClientApplication(
        CONFIG['client_id'],
        authority=f"https://login.microsoftonline.com/{CONFIG['tenant_id']}",
        client_credential=CONFIG['client_secret']
    )
    
    # Initialize variables
    access_token = None
    headers = None
    
    # Get access token
    print("1. Getting access token...")
    try:
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "access_token" in result:
            access_token = result["access_token"]
            # Create headers here so they're available throughout the function
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            print("   ✓ Access token obtained successfully")
        else:
            print("   ✗ Failed to get access token:")
            print("   ", result.get("error_description", "Unknown error"))
            return False
    except Exception as e:
        print(f"   ✗ Exception getting token: {e}")
        return False
    
    # Test booking businesses endpoint
    print("\n2. Testing booking businesses endpoint...")
    try:
        url = "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses"
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            businesses = response.json().get('value', [])
            print(f"   ✓ Found {len(businesses)} booking business(es)")
            
            for business in businesses:
                print(f"   - ID: {business.get('id')}")
                print(f"     Name: {business.get('displayName', 'N/A')}")
                print(f"     Email: {business.get('email', 'N/A')}")
                
        else:
            print(f"   ✗ API Error: {response.status_code}")
            print(f"   Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"   ✗ Exception: {e}")
        return False
    
    # Test appointments endpoint
    print("\n3. Testing appointments endpoint...")
    try:
        # Get today's date range
        today = datetime.now()
        start_date = today.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = start_date + timedelta(days=1)
        
        # Format dates for API
        start_str = start_date.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        end_str = end_date.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        
        url = f"https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{CONFIG['booking_business_id']}/appointments"
        
        params = {
            '$filter': f"startDateTime/dateTime ge '{start_str}' and startDateTime/dateTime lt '{end_str}'",
            '$orderby': 'startDateTime/dateTime'
        }
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            appointments = response.json().get('value', [])
            print(f"   ✓ Found {len(appointments)} appointment(s) for today")
            
            if appointments:
                print("   Today's appointments:")
                for i, appointment in enumerate(appointments, 1):
                    start_time = appointment.get('startDateTime', 'N/A')
                    end_time = appointment.get('endDateTime', 'N/A')
                    title = appointment.get('displayName', 'No title')
                    
                    # Convert ISO time to readable format
                    try:
                        start_dt = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
                        end_dt = datetime.fromisoformat(end_time.replace('Z', '+00:00'))
                        start_local = start_dt.astimezone().strftime('%I:%M %p')
                        end_local = end_dt.astimezone().strftime('%I:%M %p')
                        
                        print(f"   {i}. {title}")
                        print(f"      Time: {start_local} - {end_local}")
                    except:
                        print(f"   {i}. {title}")
                        print(f"      Time: {start_time} - {end_time}")
            else:
                print("   No appointments found for today")
                
        else:
            print(f"   ✗ API Error: {response.status_code}")
            print(f"   Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"   ✗ Exception: {e}")
        return False
    
    print("\n" + "=" * 50)
    print("✓ All tests passed! Your API connection is working correctly.")
    print("You can now run the full conference room server.")
    return True

if __name__ == "__main__":
    print("Microsoft Bookings API Connection Test")
    print("Make sure to update CONFIG with your actual values!")
    print()
    
    # Check if config is still placeholder
    if CONFIG['client_id'] == 'YOUR_CLIENT_ID':
        print("⚠️  WARNING: Please update CONFIG with your actual credentials before running!")
        print("   Update the values in this script and try again.")
        exit(1)
    
    success = test_connection()
    
    if success:
        print("\nNext steps:")
        print("1. Update the main server script with the same credentials")
        print("2. Run: python conference_room_server.py")
        print("3. Open browser to: http://localhost:5000")
    else:
        print("\nPlease check your credentials and try again.")
        print("Common issues:")
        print("- Wrong client_id, client_secret, or tenant_id")
        print("- Incorrect booking_business_id")
        print("- Missing API permissions in Azure AD")
        print("- Admin consent not granted")
