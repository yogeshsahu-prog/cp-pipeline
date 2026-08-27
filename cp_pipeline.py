import os
import json
import base64
import re
import io
import pickle
import datetime
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive.file'
]

def get_gmail_service():
    """Authenticate and return Gmail service"""
    creds = None
    
    # Get credentials from environment variables
    client_id = os.environ.get('GOOGLE_CLIENT_ID')
    client_secret = os.environ.get('GOOGLE_CLIENT_SECRET')
    refresh_token = os.environ.get('GOOGLE_REFRESH_TOKEN')
    
    if client_id and client_secret and refresh_token:
        # Use credentials from environment (GitHub Actions)
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            client_id=client_id,
            client_secret=client_secret,
            token_uri='https://oauth2.googleapis.com/token',
            scopes=SCOPES
        )
        # Refresh the token
        creds.refresh(Request())
    else:
        # Local development - use token.pickle
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)

def get_drive_service():
    """Authenticate and return Drive service"""
    creds = None
    
    # Get credentials from environment variables
    client_id = os.environ.get('GOOGLE_CLIENT_ID')
    client_secret = os.environ.get('GOOGLE_CLIENT_SECRET')
    refresh_token = os.environ.get('GOOGLE_REFRESH_TOKEN')
    
    if client_id and client_secret and refresh_token:
        # Use credentials from environment (GitHub Actions)
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            client_id=client_id,
            client_secret=client_secret,
            token_uri='https://oauth2.googleapis.com/token',
            scopes=SCOPES
        )
        # Refresh the token
        creds.refresh(Request())
    else:
        # Local development - use token.pickle
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
    
    return build('drive', 'v3', credentials=creds)

def get_latest_email(service, query):
    """Get the most recent email matching the query"""
    try:
        # Search for emails matching the query, get the most recent one
        result = service.users().messages().list(
            userId='me',
            q=query,
            maxResults=1  # This gets the most recent email (newest first)
        ).execute()
        
        messages = result.get('messages', [])
        if not messages:
            print("No emails found matching query")
            return None
        
        # Get the first (newest) message
        msg = service.users().messages().get(
            userId='me',
            id=messages[0]['id'],
            format='full'
        ).execute()
        
        # Extract email date to verify it's recent
        headers = msg.get('payload', {}).get('headers', [])
        email_date = None
        for header in headers:
            if header['name'] == 'Date':
                email_date = header['value']
                break
        
        print(f"Processing email from: {email_date}")
        return msg
        
    except Exception as e:
        print(f"Error fetching email: {e}")
        return None

def get_attachment(service, msg, filename_pattern='.xlsb'):
    """Download the first attachment matching the pattern"""
    try:
        parts = msg.get('payload', {}).get('parts', [])
        
        # Handle nested parts
        def find_attachment(parts, pattern):
            for part in parts:
                # Check if this part has nested parts
                if 'parts' in part:
                    result = find_attachment(part['parts'], pattern)
                    if result:
                        return result
                
                # Check if this part is an attachment
                if part.get('filename'):
                    filename = part['filename']
                    if pattern in filename.lower():
                        return part
            return None
        
        # Find the attachment
        attachment_part = find_attachment(parts, filename_pattern.lower())
        
        if not attachment_part:
            print(f"No attachment found with pattern: {filename_pattern}")
            return None
        
        # Get the attachment data
        if 'data' in attachment_part['body']:
            # Attachment data is included directly
            data = attachment_part['body']['data']
            file_data = base64.urlsafe_b64decode(data)
            return file_data, attachment_part['filename']
        elif 'attachmentId' in attachment_part['body']:
            # Need to fetch the attachment separately
            attachment_id = attachment_part['body']['attachmentId']
            attachment = service.users().messages().attachments().get(
                userId='me',
                messageId=msg['id'],
                id=attachment_id
            ).execute()
            
            data = attachment['data']
            file_data = base64.urlsafe_b64decode(data)
            return file_data, attachment_part['filename']
        else:
            print("No attachment data found")
            return None
            
    except Exception as e:
        print(f"Error downloading attachment: {e}")
        return None

def parse_excel_data(file_data, filename):
    """Parse the Excel file and extract the dashboard data"""
    try:
        # Read the Excel file
        df = pd.read_excel(io.BytesIO(file_data), sheet_name=None)
        
        # Find the sheet with CP data (usually the first sheet)
        sheet_name = list(df.keys())[0]
        print(f"Processing sheet: {sheet_name}")
        
        df_sheet = df[sheet_name]
        
        # The structure seems to have data in a specific format
        # Based on the email, we need to find the "Current Open" row
        # and extract the values for Today, Tomorrow, Day after
        
        # Look for the row containing "Current Open"
        current_open_row = None
        total_row = None
        
        for idx, row in df_sheet.iterrows():
            row_values = row.astype(str).str.lower().tolist()
            row_str = ' '.join(row_values)
            
            if 'current open' in row_str:
                current_open_row = idx
                print(f"Found 'Current Open' at row {idx}")
            
            if 'grand total' in row_str or 'total' in row_str and idx > 10:
                total_row = idx
                print(f"Found 'Grand Total' at row {idx}")
        
        if current_open_row is None:
            print("Could not find 'Current Open' in the data")
            return None
        
        # Extract the data
        row_data = df_sheet.iloc[current_open_row].values
        
        # Find the columns with AWB numbers (skip text columns)
        numeric_values = []
        for val in row_data:
            if isinstance(val, (int, float)):
                numeric_values.append(int(val))
        
        # Typically the format is: Today, Tomorrow, Day after
        # Based on the email: 1,702 (Today) · 2,477 (Tomorrow) · 9,930 (Day after)
        if len(numeric_values) >= 3:
            today_val = numeric_values[0]
            tomorrow_val = numeric_values[1]
            day_after_val = numeric_values[2]
            
            # Find RAD, Intransit, Misroute values
            rad_row = None
            it_row = None
            mis_row = None
            noplan_row = None
            badpin_row = None
            cod_row = None
            
            for idx, row in df_sheet.iterrows():
                row_values = row.astype(str).str.lower().tolist()
                row_str = ' '.join(row_values)
                
                if 'rad' in row_str and 'ready' in row_str:
                    rad_row = idx
                elif 'intransit' in row_str:
                    it_row = idx
                elif 'misroute' in row_str:
                    mis_row = idx
                elif 'not manifested' in row_str or 'no plan' in row_str:
                    noplan_row = idx
                elif 'badpin' in row_str or 'bad pin' in row_str:
                    badpin_row = idx
                elif 'cod' in row_str:
                    cod_row = idx
            
            # Extract RAD, Intransit, Misroute values
            rad_val = 0
            it_val = 0
            mis_val = 0
            noplan_val = 0
            badpin_val = 0
            cod_val = 0
            
            if rad_row is not None:
                rad_data = df_sheet.iloc[rad_row].values
                rad_vals = [int(v) for v in rad_data if isinstance(v, (int, float))]
                rad_val = rad_vals[0] if rad_vals else 0
            
            if it_row is not None:
                it_data = df_sheet.iloc[it_row].values
                it_vals = [int(v) for v in it_data if isinstance(v, (int, float))]
                it_val = it_vals[0] if it_vals else 0
            
            if mis_row is not None:
                mis_data = df_sheet.iloc[mis_row].values
                mis_vals = [int(v) for v in mis_data if isinstance(v, (int, float))]
                mis_val = mis_vals[0] if mis_vals else 0
            
            if noplan_row is not None:
                noplan_data = df_sheet.iloc[noplan_row].values
                noplan_vals = [int(v) for v in noplan_data if isinstance(v, (int, float))]
                noplan_val = noplan_vals[0] if noplan_vals else 0
            
            if badpin_row is not None:
                badpin_data = df_sheet.iloc[badpin_row].values
                badpin_vals = [int(v) for v in badpin_data if isinstance(v, (int, float))]
                badpin_val = badpin_vals[0] if badpin_vals else 0
            
            if cod_row is not None:
                cod_data = df_sheet.iloc[cod_row].values
                cod_vals = [int(v) for v in cod_data if isinstance(v, (int, float))]
                cod_val = cod_vals[0] if cod_vals else 0
            
            # Also get the total
            total_val = 0
            if total_row is not None:
                total_data = df_sheet.iloc[total_row].values
                total_vals = [int(v) for v in total_data if isinstance(v, (int, float))]
                total_val = total_vals[0] if total_vals else today_val + tomorrow_val + day_after_val
            
            # Build the result structure
            result = {
                'meta': {
                    'snapshot': os.path.basename(filename),
                    'open': today_val + tomorrow_val + day_after_val,
                    'hubs': 0,  # Will be populated from hub data
                    'states': 0,  # Will be populated from state data
                    'zones': 0,  # Will be populated from zone data
                    'updated': datetime.now().strftime('%d %b, %I:%M %p IST')
                },
                'overall': {
                    'total': today_val + tomorrow_val + day_after_val,
                    'rad': rad_val,
                    'it': it_val,
                    'mis': mis_val,
                    'today': today_val,
                    'tom': tomorrow_val,
                    'd2': day_after_val,
                    'noplan': noplan_val,
                    'badpin': badpin_val,
                    'cod': cod_val
                },
                'zones': []  # Will be populated with zone data
            }
            
            print(f"Parsed data: Today={today_val}, Tomorrow={tomorrow_val}, DayAfter={day_after_val}")
            print(f"RAD={rad_val}, Intransit={it_val}, Misroute={mis_val}")
            
            return result
            
        else:
            print(f"Found {len(numeric_values)} numeric values, expected at least 3")
            return None
            
    except Exception as e:
        print(f"Error parsing Excel: {e}")
        import traceback
        traceback.print_exc()
        return None

def update_drive_file(service, file_id, data):
    """Update a file in Google Drive with new data"""
    try:
        # Convert data to JSON string
        json_str = json.dumps(data, indent=2)
        
        # Update the file
        from googleapiclient.http import MediaIoBaseUpload
        import io
        
        media = MediaIoBaseUpload(
            io.BytesIO(json_str.encode('utf-8')),
            mimetype='application/json',
            resumable=True
        )
        
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        
        print(f"Successfully updated file: {updated_file.get('name')}")
        return True
        
    except Exception as e:
        print(f"Error updating Drive file: {e}")
        return False

def main():
    """Main function to run the CP pipeline"""
    print("Starting CP Pipeline...")
    
    # Get environment variables
    gmail_query = os.environ.get('GMAIL_QUERY', 'subject:"Flipkart CP Update" has:attachment filename:xlsb newer_than:2d')
    drive_agg_file_id = os.environ.get('DRIVE_AGG_FILE_ID')
    drive_state_file_id = os.environ.get('DRIVE_STATE_FILE_ID')
    
    if not drive_agg_file_id:
        print("ERROR: DRIVE_AGG_FILE_ID not set")
        return
    
    print(f"Gmail Query: {gmail_query}")
    
    try:
        # Authenticate
        print("Authenticating with Google...")
        gmail_service = get_gmail_service()
        drive_service = get_drive_service()
        print("Authentication successful!")
        
        # Get the latest email
        print(f"Searching for emails with query: {gmail_query}")
        msg = get_latest_email(gmail_service, gmail_query)
        
        if not msg:
            print("No email found matching the query")
            # Update with empty data
            empty_data = {
                'meta': {
                    'snapshot': 'No email found',
                    'open': 0,
                    'hubs': 0,
                    'states': 0,
                    'zones': 0,
                    'updated': datetime.now().strftime('%d %b, %I:%M %p IST')
                },
                'overall': {
                    'total': 0,
                    'rad': 0,
                    'it': 0,
                    'mis': 0,
                    'today': 0,
                    'tom': 0,
                    'd2': 0,
                    'noplan': 0,
                    'badpin': 0,
                    'cod': 0
                },
                'zones': []
            }
            update_drive_file(drive_service, drive_agg_file_id, empty_data)
            return
        
        # Download the attachment
        print("Downloading attachment...")
        attachment = get_attachment(gmail_service, msg, '.xlsb')
        
        if not attachment:
            print("No attachment found")
            return
        
        file_data, filename = attachment
        print(f"Downloaded attachment: {filename}")
        
        # Parse the Excel data
        print("Parsing Excel data...")
        parsed_data = parse_excel_data(file_data, filename)
        
        if not parsed_data:
            print("Failed to parse Excel data")
            return
        
        # Update the agg.json file in Drive
        print("Updating agg.json in Drive...")
        success = update_drive_file(drive_service, drive_agg_file_id, parsed_data)
        
        if success:
            print("✅ CP Pipeline completed successfully!")
            print(f"Updated with data from: {filename}")
            
            # Update state file with timestamp
            if drive_state_file_id:
                state_data = {
                    'last_processed': datetime.now().isoformat(),
                    'filename': filename,
                    'total_open': parsed_data['meta']['open']
                }
                update_drive_file(drive_service, drive_state_file_id, state_data)
        else:
            print("❌ Failed to update Drive file")
            
    except Exception as e:
        print(f"❌ Error in main pipeline: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
