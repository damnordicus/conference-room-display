import os
import json
import requests
from datetime import datetime, timedelta, timezone
from flask import Flask, render_template, jsonify
import threading
import time
from msal import ConfidentialClientApplication

app = Flask(__name__)

# Configuration - Update these with your actual values
CONFIG = {
    'client_id': '',
    'client_secret': '', 
    'tenant_id': '',
    'booking_business_id': 'PhoenixSparkUpperConferenceRoom@TravisSpark.onmicrosoft.com',
    'google_sites_url': 'http://www.google.com',  # Optional fallback
    'refresh_interval': 300  # 5 minutes
}

# Set these from environment variables or update directly
# CONFIG['client_id'] = os.getenv('CLIENT_ID', 'your-client-id')
# CONFIG['client_secret'] = os.getenv('CLIENT_SECRET', 'your-client-secret')
# CONFIG['tenant_id'] = os.getenv('TENANT_ID', 'your-tenant-id')
# CONFIG['booking_business_id'] = os.getenv('BOOKING_BUSINESS_ID', 'your-booking-business-id')

# Global variables
current_booking = None
last_updated = None
access_token = None
headers = None
# Microsoft Graph API client
msal_app = ConfidentialClientApplication(
    CONFIG['client_id'],
    authority=f"https://login.microsoftonline.com/{CONFIG['tenant_id']}",
    client_credential=CONFIG['client_secret']
)

def get_access_token():
    """Get access token for Microsoft Graph API"""
    global access_token
    
    try:
        result = msal_app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "access_token" in result:
            access_token = result["access_token"]
            print("Successfully obtained access token")
            return True
        else:
            print("Failed to obtain access token:", result.get("error_description"))
            return False
    except Exception as e:
        print(f"Error getting access token: {e}")
        return False

def parse_graph_datetime(datetime_str):
    """Parse Microsoft Graph datetime string with various formats"""
    if not datetime_str:
        return None
    
    try:
        # Handle different microsecond formats
        # Remove 'Z' and replace with '+00:00' for UTC
        clean_str = datetime_str.replace('Z', '+00:00')
        
        # If it has more than 6 digits of microseconds, truncate to 6
        if '.' in clean_str and '+' in clean_str:
            date_part, tz_part = clean_str.rsplit('+', 1)
            if '.' in date_part:
                main_part, microsec_part = date_part.split('.')
                # Truncate microseconds to 6 digits
                microsec_part = microsec_part[:6].ljust(6, '0')
                clean_str = f"{main_part}.{microsec_part}+{tz_part}"
        
        return datetime.fromisoformat(clean_str)
    except Exception as e:
        print(f"Error parsing datetime '{datetime_str}': {e}")
        # Try alternative parsing methods
        try:
            # Remove microseconds entirely if parsing fails
            if '.' in datetime_str:
                base_str = datetime_str.split('.')[0]
                if base_str.endswith('Z'):
                    base_str = base_str[:-1] + '+00:00'
                elif not ('+' in base_str or '-' in base_str[-6:]):
                    base_str += '+00:00'
                return datetime.fromisoformat(base_str)
        except Exception as e2:
            print(f"Alternative parsing also failed: {e2}")
            return None

def get_today_range():
    """Get today's date range in local timezone"""
    # Use timezone-aware datetime for consistent comparison
    from datetime import timezone
    
    # Get current time in your local timezone (you may want to specify your actual timezone)
    now = datetime.now().astimezone()
    
    # Get start and end of today in local timezone
    start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = start_date + timedelta(days=1)
    
    return start_date, end_date

def compare_datetimes_safely(dt1, dt2, operation='lt'):
    """Safely compare two datetime objects, handling timezone awareness"""
    try:
        # If both are timezone-aware or both are naive, compare directly
        if (dt1.tzinfo is None) == (dt2.tzinfo is None):
            if operation == 'lt':
                return dt1 < dt2
            elif operation == 'le':
                return dt1 <= dt2
            elif operation == 'gt':
                return dt1 > dt2
            elif operation == 'ge':
                return dt1 >= dt2
        
        # If one is timezone-aware and the other is not, make them consistent
        if dt1.tzinfo is None and dt2.tzinfo is not None:
            dt1 = dt1.replace(tzinfo=dt2.tzinfo)
        elif dt2.tzinfo is None and dt1.tzinfo is not None:
            dt2 = dt2.replace(tzinfo=dt1.tzinfo)
        
        if operation == 'lt':
            return dt1 < dt2
        elif operation == 'le':
            return dt1 <= dt2
        elif operation == 'gt':
            return dt1 > dt2
        elif operation == 'ge':
            return dt1 >= dt2
            
    except Exception as e:
        print(f"Error comparing datetimes: {e}")
        return False

def fetch_bookings():
    """Fetch today's bookings from Microsoft Graph API using calendarView"""
    global current_booking, last_updated, access_token
    
    # Get fresh token if needed
    if not access_token:
        if not get_access_token():
            return
    
    try:
        # Get today's date range in local timezone
        start_date, end_date = get_today_range()
        
        # Convert to UTC for the API call
        start_date_utc = start_date.astimezone(timezone.utc)
        end_date_utc = end_date.astimezone(timezone.utc)
        
        # Format dates for API (ISO 8601 format)
        start_str = start_date_utc.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        end_str = end_date_utc.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        
        # Use calendarView endpoint for date range filtering
        url = f"https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{CONFIG['booking_business_id']}/calendarView"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        params = {
            'startDateTime': start_str,
            'endDateTime': end_str,
            '$orderby': 'startDateTime/dateTime'
        }
        
        print(f"Fetching bookings from {start_str} to {end_str}")
        print(f"Local time range: {start_date} to {end_date}")
        
        response = requests.get(url, headers=headers, params=params)

        if response.status_code == 200:
            appointments = response.json().get('value', [])
            print(f"Found {len(appointments)} appointments")
            
            # Debug: Print appointment details
            for apt in appointments:
                start_dt = apt.get('startDateTime', {})
                if isinstance(start_dt, dict):
                    start_time_str = start_dt.get('dateTime', '')
                    print(f"Appointment: {start_time_str}")
                    if start_time_str:
                        parsed_time = parse_graph_datetime(start_time_str)
                        if parsed_time:
                            local_time = parsed_time.astimezone()
                            print(f"  UTC: {parsed_time}")
                            print(f"  Local: {local_time}")
                            print(f"  Date: {local_time.date()}")
                            print(f"  Today: {start_date.date()}")
            
            # Find next upcoming appointment
            now = datetime.now().astimezone()  # Make timezone-aware
            next_appointment = None
            
            for appointment in appointments:
                # Handle both possible response formats
                start_datetime = appointment.get('startDateTime', {})
                end_datetime = appointment.get('endDateTime', {})
                
                # Extract datetime strings
                if isinstance(start_datetime, dict):
                    start_time_str = start_datetime.get('dateTime', '')
                else:
                    start_time_str = start_datetime
                    
                if isinstance(end_datetime, dict):
                    end_time_str = end_datetime.get('dateTime', '')
                else:
                    end_time_str = end_datetime
                
                if not start_time_str or not end_time_str:
                    continue
                
                try:
                    start_time = parse_graph_datetime(start_time_str)
                    end_time = parse_graph_datetime(end_time_str)
                    
                    if not start_time or not end_time:
                        print(f"Failed to parse times: start='{start_time_str}', end='{end_time_str}'")
                        continue
                    
                    # Convert to local time
                    start_local = start_time.astimezone()
                    end_local = end_time.astimezone()
                    
                    # Additional check: ensure this appointment is actually today
                    if start_local.date() != now.date():
                        print(f"Skipping appointment not for today: {start_local.date()} != {now.date()}")
                        continue
                    
                    # Get appointment details
                    customer_name = appointment.get('customerName', 'Unknown Customer')
                    service_name = appointment.get('serviceName', 'Booking')
                    title = f"{customer_name} - {service_name}"
                    
                    # Check if this is the next upcoming appointment
                    if compare_datetimes_safely(now, start_local, 'lt'):
                        next_appointment = {
                            'title': title,
                            'start': start_local.strftime('%I:%M %p'),
                            'end': end_local.strftime('%I:%M %p'),
                            'date': start_local.strftime('%B %d, %Y'),
                            'duration': str(end_local - start_local),
                            'is_current': False
                        }
                        break
                    elif compare_datetimes_safely(start_local, now, 'le') and compare_datetimes_safely(now, end_local, 'le'):
                        # Current ongoing appointment
                        next_appointment = {
                            'title': title,
                            'start': start_local.strftime('%I:%M %p'),
                            'end': end_local.strftime('%I:%M %p'),
                            'date': start_local.strftime('%B %d, %Y'),
                            'duration': str(end_local - start_local),
                            'is_current': True
                        }
                        break
                except Exception as e:
                    print(f"Error parsing appointment time: {e}")
                    continue
            
            current_booking = next_appointment
            last_updated = datetime.now()
            
            if next_appointment:
                status = "Current" if next_appointment['is_current'] else "Next"
                print(f"{status} booking: {next_appointment['title']} at {next_appointment['start']} on {next_appointment['date']}")
            else:
                print("No upcoming bookings found for today")
                
        elif response.status_code == 404:
            print("Booking business not found or calendarView not supported")
            # Fallback to original appointments endpoint
            print("Trying fallback to appointments endpoint...")
            fetch_bookings_fallback()
        else:
            print(f"API Error: {response.status_code} - {response.text}")
            
    except Exception as e:
        print(f"Error fetching bookings: {e}")

def fetch_bookings_fallback():
    """Fallback method using appointments endpoint with client-side filtering"""
    global current_booking, last_updated, access_token
    
    try:
        # Use appointments endpoint without server-side filtering
        url = f"https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{CONFIG['booking_business_id']}/appointments"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        params = {
            '$orderby': 'startDateTime/dateTime'
        }
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            all_appointments = response.json().get('value', [])
            print(f"Fetched {len(all_appointments)} total appointments, filtering for today...")
            
            # Get today's date range for filtering
            start_date, end_date = get_today_range()
            
            # Filter appointments for today
            appointments = []
            for appointment in all_appointments:
                start_datetime = appointment.get('startDateTime', {})
                if isinstance(start_datetime, dict):
                    start_time_str = start_datetime.get('dateTime', '')
                else:
                    start_time_str = start_datetime
                    
                if start_time_str:
                    try:
                        start_dt = parse_graph_datetime(start_time_str)
                        if start_dt:
                            start_dt_local = start_dt.astimezone()
                            # Check if appointment is today
                            if start_dt_local.date() == start_date.date():
                                appointments.append(appointment)
                                print(f"Including appointment: {start_time_str} -> {start_dt_local}")
                            else:
                                print(f"Excluding appointment not for today: {start_time_str} -> {start_dt_local}")
                    except Exception as e:
                        print(f"Error filtering appointment: {e}")
                        continue
            
            print(f"Found {len(appointments)} appointments for today")
            
            # Find next upcoming appointment
            now = datetime.now().astimezone()  # Make timezone-aware
            next_appointment = None
            
            for appointment in appointments:
                start_datetime = appointment.get('startDateTime', {})
                end_datetime = appointment.get('endDateTime', {})
                
                if isinstance(start_datetime, dict):
                    start_time_str = start_datetime.get('dateTime', '')
                else:
                    start_time_str = start_datetime
                    
                if isinstance(end_datetime, dict):
                    end_time_str = end_datetime.get('dateTime', '')
                else:
                    end_time_str = end_datetime
                
                if not start_time_str or not end_time_str:
                    continue
                
                try:
                    start_time = parse_graph_datetime(start_time_str)
                    end_time = parse_graph_datetime(end_time_str)
                    
                    if not start_time or not end_time:
                        print(f"Failed to parse times: start='{start_time_str}', end='{end_time_str}'")
                        continue
                    
                    start_local = start_time.astimezone()
                    end_local = end_time.astimezone()
                    
                    customer_name = appointment.get('customerName', 'Unknown Customer')
                    service_name = appointment.get('serviceName', 'Booking')
                    title = f"{customer_name} - {service_name}"
                    
                    if compare_datetimes_safely(now, start_local, 'lt'):
                        next_appointment = {
                            'title': title,
                            'start': start_local.strftime('%I:%M %p'),
                            'end': end_local.strftime('%I:%M %p'),
                            'date': start_local.strftime('%B %d, %Y'),
                            'duration': str(end_local - start_local),
                            'is_current': False
                        }
                        break
                    elif compare_datetimes_safely(start_local, now, 'le') and compare_datetimes_safely(now, end_local, 'le'):
                        next_appointment = {
                            'title': title,
                            'start': start_local.strftime('%I:%M %p'),
                            'end': end_local.strftime('%I:%M %p'),
                            'date': start_local.strftime('%B %d, %Y'),
                            'duration': str(end_local - start_local),
                            'is_current': True
                        }
                        break
                except Exception as e:
                    print(f"Error parsing appointment time: {e}")
                    continue
            
            current_booking = next_appointment
            last_updated = datetime.now()
            
            if next_appointment:
                status = "Current" if next_appointment['is_current'] else "Next"
                print(f"{status} booking: {next_appointment['title']} at {next_appointment['start']} on {next_appointment['date']}")
            else:
                print("No upcoming bookings found for today")
                
        else:
            print(f"Fallback API Error: {response.status_code} - {response.text}")
            
    except Exception as e:
        print(f"Error in fallback booking fetch: {e}")

def update_bookings_loop():
    """Background thread to update bookings periodically"""
    while True:
        fetch_bookings()
        time.sleep(CONFIG['refresh_interval'])

@app.route('/')
def index():
    """Main display page"""
    return render_template('conference_room.html', 
                                booking=current_booking, 
                                last_updated=last_updated)

@app.route('/google-sites')
def google_sites():
    """Redirect to Google Sites page"""
    if CONFIG['google_sites_url']:
        return f'<script>window.location.href="{CONFIG["google_sites_url"]}"</script>'
    else:
        return '<h1>Google Sites URL not configured</h1><a href="/">Back to main display</a>'

@app.route('/refresh')
def refresh():
    """Manually refresh booking data"""
    fetch_bookings()
    return '<script>window.location.href="/"</script>'

@app.route('/api/booking')
def api_booking():
    """API endpoint for booking data"""
    return jsonify({
        'booking': current_booking,
        'last_updated': last_updated.isoformat() if last_updated else None
    })

if __name__ == '__main__':
    # Start the background update thread
    update_thread = threading.Thread(target=update_bookings_loop, daemon=True)
    update_thread.start()
    
    # Initial fetch
    fetch_bookings()
    
    print("Conference Room Display Server starting...")
    print("Access at: http://localhost:5000")
    print("Google Sites toggle: http://localhost:5000/google-sites")
    
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=False)
