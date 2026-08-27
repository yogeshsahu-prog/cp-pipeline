import os
import json
import base64
import io
import datetime
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import pandas as pd

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive.file'
]

def get_gmail_service():
    """Authenticate and return Gmail service"""
    creds = Credentials(
        token=None,
        refresh_token=os.environ.get('GOOGLE_REFRESH_TOKEN'),
        client_id=os.environ.get('GOOGLE_CLIENT_ID'),
        client_secret=os.environ.get('GOOGLE_CLIENT_SECRET'),
        token_uri='https://oauth2.googleapis.com/token',
        scopes=SCOPES
    )
    creds.refresh(google.auth.transport.requests.Request())
    return build('gmail', 'v1', credentials=creds)

def get_drive_service():
    """Authenticate and return Drive service"""
    creds = Credentials(
        token=None,
        refresh_token=os.environ.get('GOOGLE_REFRESH_TOKEN'),
        client_id=os.environ.get('GOOGLE_CLIENT_ID'),
        client_secret=os.environ.get('GOOGLE_CLIENT_SECRET'),
        token_uri='https://oauth2.googleapis.com/token',
        scopes=SCOPES
    )
    creds.refresh(google.auth.transport.requests.Request())
    return build('drive', 'v3', credentials=creds)

def get_latest_email(service, query):
    """Get the absolute newest email matching the query"""
    try:
        # Get all messages matching the query
        result = service.users().messages().list(
            userId='me',
            q=query,
            maxResults=1  # Gmail returns newest first by default
        ).execute()
        
        messages = result.get('messages', [])
        if not messages:
            print("No emails found")
            return None
        
        # Get the first (newest) message
        msg = service.users().messages().get(
            userId='me',
            id=messages[0]['id'],
            format='full'
        ).execute()
        
        # Get the date to verify
        headers = msg.get('payload', {}).get('headers', [])
        for h in headers:
            if h['name'] == 'Date':
                print(f"📧 Processing email from: {h['value']}")
                break
        
        print(f"📧 Message ID: {messages[0]['id']}")
        return msg
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def get_attachment(service, msg):
    """Download the .xlsb attachment"""
    try:
        parts = msg.get('payload', {}).get('parts', [])
        
        def find_attachment(parts):
            for part in parts:
                if 'parts' in part:
                    result = find_attachment(part['parts'])
                    if result:
                        return result
                if part.get('filename') and '.xlsb' in part['filename'].lower():
                    return part
            return None
        
        attachment = find_attachment(parts)
        if not attachment:
            print("No .xlsb attachment found")
            return None
        
        filename = attachment['filename']
        
        if 'data' in attachment['body']:
            data = base64.urlsafe_b64decode(attachment['body']['data'])
        elif 'attachmentId' in attachment['body']:
            att = service.users().messages().attachments().get(
                userId='me',
                messageId=msg['id'],
                id=attachment['body']['attachmentId']
            ).execute()
            data = base64.urlsafe_b64decode(att['data'])
        else:
            return None
        
        print(f"📎 Downloaded: {filename}")
        return data, filename
        
    except Exception as e:
        print(f"❌ Error downloading attachment: {e}")
        return None

def parse_excel(data, filename):
    """Extract the numbers from the Excel file"""
    try:
        df = pd.read_excel(io.BytesIO(data), sheet_name=0)
        print(f"📊 Excel has {len(df)} rows")
        
        # Find the "Current Open" row
        current_open = None
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).str.lower())
            if 'current open' in row_str:
                current_open = idx
                break
        
        if current_open is None:
            print("❌ Could not find 'Current Open'")
            return None
        
        # Get the values
        row = df.iloc[current_open]
        values = [int(v) for v in row if isinstance(v, (int, float))]
        
        print(f"📊 Found values: {values}")
        
        # We need at least 3 values: Today, Tomorrow, Day after
        if len(values) < 3:
            print(f"❌ Expected 3+ values, got {len(values)}")
            return None
        
        today = values[0]
        tomorrow = values[1] if len(values) > 1 else 0
        day_after = values[2] if len(values) > 2 else 0
        
        # Try to find RAD, Intransit, Misroute
        rad = 0
        it = 0
        mis = 0
        
        for idx, row in df.iterrows():
            row_str = ' '.join(row.astype(str).str.lower())
            row_vals = [int(v) for v in row if isinstance(v, (int, float))]
            
            if 'rad' in row_str and row_vals:
                rad = row_vals[0] if row_vals else 0
            elif 'intransit' in row_str and row_vals:
                it = row_vals[0] if row_vals else 0
            elif 'misroute' in row_str and row_vals:
                mis = row_vals[0] if row_vals else 0
        
        result = {
            'meta': {
                'snapshot': filename.replace('.xlsb', ''),
                'open': today + tomorrow + day_after,
                'hubs': 0,
                'states': 0,
                'zones': 0,
                'updated': datetime.now().strftime('%d %b, %I:%M %p IST')
            },
            'overall': {
                'total': today + tomorrow + day_after,
                'rad': rad,
                'it': it,
                'mis': mis,
                'today': today,
                'tom': tomorrow,
                'd2': day_after,
                'noplan': 0,
                'badpin': 0,
                'cod': 0
            },
            'zones': []
        }
        
        print(f"✅ Parsed: Today={today}, Tomorrow={tomorrow}, DayAfter={day_after}")
        print(f"✅ Total = {today + tomorrow + day_after}")
        return result
        
    except Exception as e:
        print(f"❌ Error parsing: {e}")
        import traceback
        traceback.print_exc()
        return None

def update_drive(service, file_id, data):
    """Update the Drive file"""
    try:
        from googleapiclient.http import MediaIoBaseUpload
        
        json_str = json.dumps(data, indent=2)
        media = MediaIoBaseUpload(
            io.BytesIO(json_str.encode('utf-8')),
            mimetype='application/json'
        )
        
        service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        
        print(f"✅ Updated Drive file: {file_id}")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    print("🚀 Starting CP Pipeline...")
    
    # Get environment variables
    query = os.environ.get('GMAIL_QUERY', 'subject:"Flipkart CP Update" has:attachment filename:xlsb newer_than:1d')
    agg_id = os.environ.get('DRIVE_AGG_FILE_ID')
    state_id = os.environ.get('DRIVE_STATE_FILE_ID')
    
    print(f"🔍 Query: {query}")
    
    try:
        # Authenticate
        gmail = get_gmail_service()
        drive = get_drive_service()
        print("✅ Authentication successful")
        
        # Get latest email
        msg = get_latest_email(gmail, query)
        if not msg:
            print("❌ No email found")
            return
        
        # Download attachment
        attachment = get_attachment(gmail, msg)
        if not attachment:
            print("❌ No attachment found")
            return
        
        data, filename = attachment
        
        # Parse data
        parsed = parse_excel(data, filename)
        if not parsed:
            print("❌ Failed to parse")
            return
        
        # Update Drive
        if update_drive(drive, agg_id, parsed):
            print(f"✅ Success! Updated with {filename}")
            
            # Update state
            if state_id:
                state = {
                    'last_msg_id': msg['id'],
                    'file': filename,
                    'open': parsed['meta']['open']
                }
                update_drive(drive, state_id, state)
        else:
            print("❌ Failed to update Drive")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
