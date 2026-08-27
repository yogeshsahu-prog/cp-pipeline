import os
import json
import base64
import io
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive"
]

IST = ZoneInfo("Asia/Kolkata")


# ============================================================
# GOOGLE AUTHENTICATION
# ============================================================

def get_credentials():
    """Create and refresh Google OAuth credentials."""

    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN")
    client_id = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")

    if not refresh_token:
        raise RuntimeError("GOOGLE_REFRESH_TOKEN is missing")

    if not client_id:
        raise RuntimeError("GOOGLE_CLIENT_ID is missing")

    if not client_secret:
        raise RuntimeError("GOOGLE_CLIENT_SECRET is missing")

    creds = Credentials(
        token=None,
        refresh_token=refresh_token,
        client_id=client_id,
        client_secret=client_secret,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES
    )

    creds.refresh(Request())

    return creds


def get_gmail_service():
    """Authenticate and return Gmail service."""
    creds = get_credentials()
    return build("gmail", "v1", credentials=creds)


def get_drive_service():
    """Authenticate and return Drive service."""
    creds = get_credentials()
    return build("drive", "v3", credentials=creds)


# ============================================================
# GET LATEST EMAIL
# ============================================================

def get_latest_email(service, query):
    """
    Fetch the latest email matching the Gmail query.

    We fetch multiple matching messages and then select the one
    with the highest Gmail internalDate to make sure we truly
    process the newest email.
    """

    print("🔍 Searching Gmail...")
    print(f"🔍 Query: {query}")

    try:
        result = service.users().messages().list(
            userId="me",
            q=query,
            maxResults=20
        ).execute()

        messages = result.get("messages", [])

        if not messages:
            raise RuntimeError(
                f"No emails found matching query: {query}"
            )

        print(f"📧 Found {len(messages)} matching email(s)")

        latest_message = None
        latest_timestamp = -1

        for message in messages:

            msg = service.users().messages().get(
                userId="me",
                id=message["id"],
                format="metadata",
                metadataHeaders=["Date", "Subject", "From"]
            ).execute()

            internal_date = int(
                msg.get("internalDate", "0")
            )

            if internal_date > latest_timestamp:
                latest_timestamp = internal_date
                latest_message = msg

        if not latest_message:
            raise RuntimeError("Could not determine latest email")

        # --------------------------------------------------------
        # Print email information
        # --------------------------------------------------------

        headers = {
            h["name"].lower(): h["value"]
            for h in latest_message
            .get("payload", {})
            .get("headers", [])
        }

        email_date = headers.get("date", "Unknown")
        subject = headers.get("subject", "Unknown")
        sender = headers.get("from", "Unknown")

        internal_dt = datetime.fromtimestamp(
            latest_timestamp / 1000,
            tz=IST
        )

        print("")
        print("==========================================")
        print("📧 LATEST EMAIL SELECTED")
        print("==========================================")
        print(f"From      : {sender}")
        print(f"Subject   : {subject}")
        print(f"Email Date: {email_date}")
        print(
            f"Received  : "
            f"{internal_dt.strftime('%d %b %Y, %I:%M:%S %p IST')}"
        )
        print(f"Message ID: {latest_message['id']}")
        print("==========================================")
        print("")

        # Get complete email because we need attachment data
        full_msg = service.users().messages().get(
            userId="me",
            id=latest_message["id"],
            format="full"
        ).execute()

        return full_msg

    except Exception as e:
        print(f"❌ Gmail error: {e}")
        raise


# ============================================================
# GET XLSB ATTACHMENT
# ============================================================

def get_attachment(service, msg):
    """Find and download the .xlsb attachment."""

    print("📎 Searching for XLSB attachment...")

    try:
        payload = msg.get("payload", {})
        parts = payload.get("parts", [])

        def find_attachment(part_list):

            for part in part_list:

                filename = part.get("filename", "")

                if filename and filename.lower().endswith(".xlsb"):
                    return part

                nested_parts = part.get("parts", [])

                if nested_parts:
                    result = find_attachment(nested_parts)

                    if result:
                        return result

            return None

        attachment = find_attachment(parts)

        if not attachment:
            raise RuntimeError(
                "No .xlsb attachment found in the latest email"
            )

        filename = attachment["filename"]
        body = attachment.get("body", {})

        # --------------------------------------------------------
        # Attachment data directly inside email
        # --------------------------------------------------------

        if body.get("data"):

            data = base64.urlsafe_b64decode(
                body["data"] + "=" * (-len(body["data"]) % 4)
            )

        # --------------------------------------------------------
        # Attachment stored separately by Gmail
        # --------------------------------------------------------

        elif body.get("attachmentId"):

            att = service.users().messages().attachments().get(
                userId="me",
                messageId=msg["id"],
                id=body["attachmentId"]
            ).execute()

            encoded_data = att.get("data")

            if not encoded_data:
                raise RuntimeError(
                    "Gmail attachment response contains no data"
                )

            data = base64.urlsafe_b64decode(
                encoded_data + "=" * (-len(encoded_data) % 4)
            )

        else:
            raise RuntimeError(
                f"Attachment body is empty: {filename}"
            )

        print(f"✅ Downloaded attachment: {filename}")
        print(f"📦 Attachment size: {len(data):,} bytes")

        return data, filename

    except Exception as e:
        print(f"❌ Attachment error: {e}")
        raise


# ============================================================
# PARSE EXCEL
# ============================================================

def parse_excel(data, filename):
    """Read XLSB and extract CP numbers."""

    print("📊 Reading Excel file...")

    try:

        df = pd.read_excel(
            io.BytesIO(data),
            sheet_name=0,
            engine="pyxlsb"
        )

        print(f"📊 Excel rows: {len(df)}")
        print(f"📊 Excel columns: {len(df.columns)}")

        # --------------------------------------------------------
        # Find Current Open row
        # --------------------------------------------------------

        current_open = None

        for idx, row in df.iterrows():

            row_str = " ".join(
                row.astype(str)
                .str.lower()
                .tolist()
            )

            if "current open" in row_str:
                current_open = idx
                break

        if current_open is None:
            raise RuntimeError(
                "Could not find 'Current Open' row in Excel"
            )

        print(f"✅ Current Open row found: {current_open}")

        # --------------------------------------------------------
        # Extract numeric values
        # --------------------------------------------------------

        row = df.iloc[current_open]

        values = (
            pd.to_numeric(row, errors="coerce")
            .dropna()
            .astype(int)
            .tolist()
        )

        print(f"📊 Current Open values: {values}")

        if len(values) < 3:
            raise RuntimeError(
                f"Expected at least 3 numbers in Current Open row, "
                f"but found {len(values)}"
            )

        today = values[0]
        tomorrow = values[1]
        day_after = values[2]

        # --------------------------------------------------------
        # Find RAD / Intransit / Misroute
        # --------------------------------------------------------

        rad = 0
        intransit = 0
        misroute = 0

        for idx, excel_row in df.iterrows():

            row_str = " ".join(
                excel_row.astype(str)
                .str.lower()
                .tolist()
            )

            row_values = (
                pd.to_numeric(
                    excel_row,
                    errors="coerce"
                )
                .dropna()
                .astype(int)
                .tolist()
            )

            if "rad" in row_str and row_values:
                rad = row_values[0]

            elif "intransit" in row_str and row_values:
                intransit = row_values[0]

            elif "misroute" in row_str and row_values:
                misroute = row_values[0]

        # --------------------------------------------------------
        # Calculate total
        # --------------------------------------------------------

        total_open = today + tomorrow + day_after

        now_ist = datetime.now(IST)

        result = {

            "meta": {
                "snapshot": filename.replace(".xlsb", ""),
                "open": total_open,
                "hubs": 0,
                "states": 0,
                "zones": 0,
                "updated": now_ist.strftime(
                    "%d %b, %I:%M %p IST"
                )
            },

            "overall": {
                "total": total_open,
                "rad": rad,
                "it": intransit,
                "mis": misroute,
                "today": today,
                "tom": tomorrow,
                "d2": day_after,
                "noplan": 0,
                "badpin": 0,
                "cod": 0
            },

            "zones": []
        }

        print("")
        print("==========================================")
        print("📊 PARSED DATA")
        print("==========================================")
        print(f"Today     : {today:,}")
        print(f"Tomorrow  : {tomorrow:,}")
        print(f"Day After : {day_after:,}")
        print(f"Total     : {total_open:,}")
        print(f"RAD       : {rad:,}")
        print(f"Intransit : {intransit:,}")
        print(f"Misroute  : {misroute:,}")
        print("==========================================")
        print("")

        return result

    except Exception as e:
        print(f"❌ Excel parsing error: {e}")
        raise


# ============================================================
# UPDATE GOOGLE DRIVE JSON
# ============================================================

def update_drive(service, file_id, data):
    """Update an existing JSON file in Google Drive."""

    if not file_id:
        raise RuntimeError(
            "Drive file ID is missing"
        )

    try:

        json_str = json.dumps(
            data,
            indent=2,
            ensure_ascii=False
        )

        media = MediaIoBaseUpload(
            io.BytesIO(
                json_str.encode("utf-8")
            ),
            mimetype="application/json",
            resumable=False
        )

        service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()

        print(f"✅ Updated Drive file: {file_id}")

        return True

    except Exception as e:
        print(f"❌ Drive update error: {e}")
        raise


# ============================================================
# MAIN PIPELINE
# ============================================================

def main():

    print("")
    print("==========================================")
    print("🚀 CP PENDENCY PIPELINE STARTED")
    print("==========================================")
    print("")

    # --------------------------------------------------------
    # Environment variables
    # --------------------------------------------------------

    query = os.environ.get(
        "GMAIL_QUERY",
        'subject:"Flipkart CP Update" '
        'has:attachment filename:xlsb newer_than:1d'
    )

    agg_id = os.environ.get(
        "DRIVE_AGG_FILE_ID"
    )

    state_id = os.environ.get(
        "DRIVE_STATE_FILE_ID"
    )

    print(f"🔍 Gmail Query: {query}")
    print(f"📁 Aggregate Drive ID: {agg_id}")
    print(f"📁 State Drive ID: {state_id}")
    print("")

    # --------------------------------------------------------
    # Authenticate
    # --------------------------------------------------------

    print("🔐 Authenticating with Google...")

    gmail = get_gmail_service()
    drive = get_drive_service()

    print("✅ Google authentication successful")
    print("")

    # --------------------------------------------------------
    # Find latest email
    # --------------------------------------------------------

    msg = get_latest_email(
        gmail,
        query
    )

    # --------------------------------------------------------
    # Download attachment
    # --------------------------------------------------------

    attachment = get_attachment(
        gmail,
        msg
    )

    data, filename = attachment

    # --------------------------------------------------------
    # Parse XLSB
    # --------------------------------------------------------

    parsed = parse_excel(
        data,
        filename
    )

    # --------------------------------------------------------
    # Update aggregate JSON
    # --------------------------------------------------------

    update_drive(
        drive,
        agg_id,
        parsed
    )

    # --------------------------------------------------------
    # Update state JSON
    # --------------------------------------------------------

    if state_id:

        state = {
            "last_msg_id": msg["id"],
            "file": filename,
            "open": parsed["meta"]["open"],
            "updated": parsed["meta"]["updated"]
        }

        update_drive(
            drive,
            state_id,
            state
        )

    # --------------------------------------------------------
    # Complete
    # --------------------------------------------------------

    print("")
    print("==========================================")
    print("✅ CP PIPELINE COMPLETED SUCCESSFULLY")
    print("==========================================")
    print(f"📧 Email Message ID: {msg['id']}")
    print(f"📎 File: {filename}")
    print(f"📊 CP Open: {parsed['meta']['open']:,}")
    print("==========================================")
    print("")


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    try:
        main()

    except Exception as e:
        print("")
        print("==========================================")
        print("❌ CP PIPELINE FAILED")
        print("==========================================")
        print(str(e))
        print("==========================================")
        print("")

        raise
