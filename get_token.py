#!/usr/bin/env python3
"""
ONE-TIME, run on your own laptop. Opens a browser, you log in as the account
that RECEIVES the CP emails, and it prints a refresh token. Paste that token
into the GitHub secret GOOGLE_REFRESH_TOKEN. You never run this again.

Setup before running:
  1. Google Cloud Console -> create/select a project.
  2. APIs & Services -> Enable: Gmail API, Google Drive API, Google Sheets API.
  3. OAuth consent screen -> External -> add your own email as a Test user.
  4. Credentials -> Create OAuth client ID -> type "Desktop app" -> download
     the JSON, save it next to this file as client_secret.json.
  5. pip install google-auth-oauthlib
  6. python get_token.py
"""
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly",
          "https://www.googleapis.com/auth/drive",
          "https://www.googleapis.com/auth/spreadsheets"]

flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")

print("\n================  COPY THESE INTO GITHUB SECRETS  ================")
print("GOOGLE_CLIENT_ID     =", creds.client_id)
print("GOOGLE_CLIENT_SECRET =", creds.client_secret)
print("GOOGLE_REFRESH_TOKEN =", creds.refresh_token)
print("=================================================================")
