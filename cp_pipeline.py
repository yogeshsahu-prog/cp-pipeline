import os
import json
import base64
import io
import datetime
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import pandas as pd
import email.utils
import time

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
    creds.refresh(Request())
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
    creds.refresh(Request())
    return build('drive', 'v3', credentials=creds)

def get_newest_email(service, query):
    """Get the absolute newest email by parsing the Date header"""
    try:
        # Get up to 10 emails matching the query
        result = service.users().messages().list(
            userId='me',
            q=query,
            maxResults=10  # Get more to compare dates
        ).execute()
        
        messages = result.get('messages', [])
        if not messages:
            print("❌ No emails found")
            return None
        
        print(f"🔍 Found {len(messages)} emails, finding newest...")
        
        # Get full details for each email and compare dates
        newest_msg = None
        newest_date = None
        
        for msg_ref in messages:
            msg = service.users().messages().get(
                userId='me',
                id=msg_ref['id'],
                format='metadata',
                metadataHeaders=['Date']
            ).execute()
            
            # Extract the Date header
            headers = msg.get('payload', {}).get('headers', [])
            for h in headers:
                if h['name'] == 'Date':
                    date_str = h['value']
                    # Parse the date
                    try:
                        # Parse RFC 2822 date
                        parsed_date = email.utils.parsedate_to_datetime(date_str)
                        print(f"📧 Found: {parsed_date.strftime('%Y-%m-%d %H:%M:%S')} - {msg['id'][:10]}...")
                        
                        if newest_date is None or parsed_date > newest_date:
                            newest_date = parsed_date
                            newest_msg = msg
                    except:
                        # If parsing fails, just use the raw string
                        print(f"📧 Found: {date_str}")
                        if newest_date is None:
                            newest_date = date_str
                            newest_msg = msg
                    break
        
        if newest_msg:
            print(f"✅ Selected newest email from: {newest_date}")
            # Now get the full message with all parts
            full_msg = service.users().messages().get(
                userId='me',
                id=newest_msg['id'],
                format='full'
            ).execute()
            return full_msg
        else:
            print("❌ Could not determine newest email")
            return None
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
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
            print("❌ No .xlsb attachment found")
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
        print(f"📊 Excel has {len(df)} rows and {len(df.columns)} columns")
        
        # Find the "Current Open" row
        current_open_row = None
        for idx, row in df.iterrows():
            row_str = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
            if 'current open' in row_str:
                current_open_row = idx
                print(f"🔍 Found 'Current Open' at row {idx}")
                break
        
        if current_open_row is None:
            print("❌ Could not find 'Current Open'")
            return None
        
        # Get the row
        row = df.iloc[current_open_row]
        
        # Extract numeric values
        values = []
        for v in row:
            if pd.notna(v) and isinstance(v, (int, float)):
                values.append(int(v))
        
        print(f"📊 Found numeric values: {values}")
        
        if len(values) < 3:
            print(f"❌ Expected 3+ values, got {len(values)}")
            return None
        
        today = values[0]
        tomorrow = values[1] if len(values) > 1 else 0
        day_after = values[2] if len(values) > 2 else 0
        
        total_open = today + tomorrow + day_after
        
        # Find RAD, Intransit, Misroute
        rad = 0
        it = 0
        mis = 0
        noplan = 0
        badpin = 0
        cod = 0
        
        for idx, row in df.iterrows():
            row_str = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
            row_vals = [int(v) for v in row if pd.notna(v) and isinstance(v, (int, float))]
            
            if 'rad' in row_str and not row_vals:
                continue
            elif 'rad' in row_str:
                rad = row_vals[0] if row_vals else 0
                print(f"🔍 Found RAD: {rad}")
            elif 'intransit' in row_str:
                it = row_vals[0] if row_vals else 0
                print(f"🔍 Found Intransit: {it}")
            elif 'misroute' in row_str:
                mis = row_vals[0] if row_vals else 0
                print(f"🔍 Found Misroute: {mis}")
            elif 'not manifested' in row_str or 'no plan' in row_str:
                noplan = row_vals[0] if row_vals else 0
                print(f"🔍 Found No Plan: {noplan}")
            elif 'badpin' in row_str or 'bad pin' in row_str:
                badpin = row_vals[0] if row_vals else 0
                print(f"🔍 Found Badpin: {badpin}")
            elif 'cod' in row_str and len(row_str) < 10:  # Avoid false matches
                cod = row_vals[0] if row_vals else 0
                print(f"🔍 Found COD: {cod}")
        
        result = {
            'meta': {
                'snapshot': filename.replace('.xlsb', ''),
                'open': total_open,
                'hubs': 0,
                'states': 0,
                'zones': 0,
                'updated': datetime.now().strftime('%d %b, %I:%M %p IST')
            },
            'overall': {
                'total': total_open,
                'rad': rad,
                'it': it,
                'mis': mis,
                'today': today,
                'tom': tomorrow,
                'd2': day_after,
                'noplan': noplan,
                'badpin': badpin,
                'cod': cod
            },
            'zones': []
        }
        
        print(f"✅ Parsed: Today={today}, Tomorrow={tomorrow}, DayAfter={day_after}")
        print(f"✅ Total Open = {total_open}")
        print(f"✅ RAD={rad}, Intransit={it}, Misroute={mis}")
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
        print(f"❌ Error updating: {e}")
        return False

def main():
    print("🚀 Starting CP Pipeline...")
    
    # Get environment variables
    query = os.environ.get('GMAIL_QUERY', 'subject:"Flipkart CP Update" has:attachment filename:xlsb')
    agg_id = os.environ.get('DRIVE_AGG_FILE_ID')
    state_id = os.environ.get('DRIVE_STATE_FILE_ID')
    
    print(f"🔍 Query: {query}")
    
    try:
        # Authenticate
        gmail = get_gmail_service()
        drive = get_drive_service()
        print("✅ Authentication successful")
        
        # Get NEWEST email
        msg = get_newest_email(gmail, query)
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
                    'open': parsed['meta']['open'],
                    'updated': datetime.now().isoformat()
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
