"""
Run this ONCE to authorize Gmail access.
It will open your browser -> log in with your Google/work account -> saves token.json

Usage:
    python authorize_gmail.py
"""

import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/drive.file"
]
_DIR        = os.path.dirname(os.path.abspath(__file__))
CREDS_FILE  = os.path.join(_DIR, "credentials.json")
TOKEN_FILE  = os.path.join(_DIR, "token.json")

def main():
    if not os.path.exists(CREDS_FILE):
        print("ERROR: credentials.json not found in this folder.")
        print("Download it from Google Cloud Console and rename it credentials.json")
        return

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Refresh failed ({e}). Re-authenticating...")
                flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    print("\nAuthorization successful! token.json saved.")
    print("You can now use the dashboard to send emails.")

if __name__ == "__main__":
    main()
