import os
import json
import requests
from datetime import datetime, timedelta
from flask import Flask, render_template_string, jsonify
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
    'google_sites_url': 'https://www.travisspark.com',  # Optional fallback
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
    today = datetime.now().astimezone()
    start_date = today.replace(hour=0, minute=0, second=0, microsecond=0)
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
        # Get today's date range
        today = datetime.now()
        start_date = today.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = start_date + timedelta(days=1)
        
        # Format dates for API (ISO 8601 format)
        start_str = start_date.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        end_str = end_date.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
        
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
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            appointments = response.json().get('value', [])
            print(f"Found {len(appointments)} appointments")
            
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
                print(f"{status} booking: {next_appointment['title']} at {next_appointment['start']}")
            else:
                print("No upcoming bookings found")
                
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
                            if compare_datetimes_safely(start_date, start_dt_local, 'le') and compare_datetimes_safely(start_dt_local, end_date, 'lt'):
                                appointments.append(appointment)
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
                            'duration': str(end_local - start_local),
                            'is_current': False
                        }
                        break
                    elif compare_datetimes_safely(start_local, now, 'le') and compare_datetimes_safely(now, end_local, 'le'):
                        next_appointment = {
                            'title': title,
                            'start': start_local.strftime('%I:%M %p'),
                            'end': end_local.strftime('%I:%M %p'),
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
                print(f"{status} booking: {next_appointment['title']} at {next_appointment['start']}")
            else:
                print("No upcoming bookings found")
                
        else:
            print(f"Fallback API Error: {response.status_code} - {response.text}")
            
    except Exception as e:
        print(f"Error in fallback booking fetch: {e}")

def update_bookings_loop():
    """Background thread to update bookings periodically"""
    while True:
        fetch_bookings()
        time.sleep(CONFIG['refresh_interval'])

# HTML Templates
GENERIC_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Conference Room</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 0;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            height: 100vh;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        .container {
            text-align: center;
            max-width: 800px;
            padding: 40px;
        }
        h1 {
            font-size: 3em;
            margin-bottom: 20px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        .status {
            font-size: 1.5em;
            margin-bottom: 30px;
            opacity: 0.9;
        }
        .booking-info {
            background: rgba(255,255,255,0.1);
            padding: 30px;
            border-radius: 15px;
            margin-top: 30px;
            backdrop-filter: blur(10px);
        }
        .booking-title {
            font-size: 2em;
            margin-bottom: 15px;
            color: #fff;
        }
        .booking-time {
            font-size: 1.3em;
            opacity: 0.9;
        }
        .current-booking {
            background: rgba(255,107,107,0.2);
            border: 2px solid rgba(255,107,107,0.5);
        }
        .next-booking {
            background: rgba(107,255,107,0.2);
            border: 2px solid rgba(107,255,107,0.5);
        }
        .last-updated {
            position: fixed;
            bottom: 20px;
            right: 20px;
            font-size: 0.9em;
            opacity: 0.7;
        }
        .controls {
            position: fixed;
            top: 20px;
            right: 20px;
            background: rgba(0,0,0,0.3);
            padding: 10px;
            border-radius: 10px;
        }
        .controls a {
            color: white;
            text-decoration: none;
            margin: 0 10px;
            padding: 5px 10px;
            background: rgba(255,255,255,0.2);
            border-radius: 5px;
        }
        .controls a:hover {
            background: rgba(255,255,255,0.3);
        }
    </style>
    <script>
        // Auto-refresh every 5 minutes
        setTimeout(() => {
            location.reload();
        }, 300000);
    </script>
</head>
<body>
    <div class="controls">
        <a href="/google-sites">Google Sites</a>
        <a href="/refresh">Refresh</a>
    </div>
    
    <div class="container">
        <h1>Conference Room</h1>
        
        {% if booking %}
            <div class="booking-info {% if booking.is_current %}current-booking{% else %}next-booking{% endif %}">
                <div class="booking-title">{{ booking.title }}</div>
                <div class="booking-date">
                    Date: {{ booking.start }}
                </div>
                <div class="booking-time">
                    {% if booking.is_current %}
                        <strong>Currently in session</strong><br>
                        Started: {{ booking.start }} | Ends: {{ booking.end }}
                    {% else %}
                        <strong>Next booking</strong><br>
                        Time: {{ booking.start }} - {{ booking.end }}
                    {% endif %}
                </div>
            </div>
        {% else %}
            <div class="status">
                <p>🟢 Room Available</p>
                <p>No upcoming bookings</p>
            </div>
        {% endif %}
    </div>
    
    {% if last_updated %}
    <div class="last-updated">
        Last updated: {{ last_updated.strftime('%I:%M %p') }}
    </div>
    {% endif %}
</body>
</html>
"""

@app.route('/')
def index():
    """Main display page"""
    return render_template_string(GENERIC_TEMPLATE, 
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
