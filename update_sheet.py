import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Google Sheets API Scope
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_sheet_client():
    # GitHub Secrets မှ Google Credentials များကို ချိတ်ဆက်ခြင်း
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("Google Credentials JSON not found in environment variables.")
    
    import json
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def map_myanmar_odds(ah):
    """ LOCKED Myanmar Odds Mapping Rules """
    if pd.isna(ah):
        return "-"
    abs_ah = abs(ah)
    if abs_ah == 0.0:
        return "0.0 (D)"
    elif abs_ah == 0.25:
        return "0.25 (L-50)" if ah > 0 else "0.25 (1+50)"
    elif abs_ah == 0.5:
        return "0.5 (L-100)"
    elif abs_ah == 0.75:
        return "0.75 (1+50)"
    elif abs_ah == 1.0:
        return "1.0 (1D)"
    elif abs_ah == 1.25:
        return "1.25 (1-50)"
    elif abs_ah == 1.5:
        return "1.5 (1-100)"
    elif abs_ah == 1.75:
        return "1.75 (2+50)"
    elif abs_ah == 2.0:
        return "2.0 (2D)"
    return str(ah)

def main():
    print("Starting automated football data update...")
    client = get_sheet_client()
    
    # Google Sheet ကို Spreadsheet Key ဖြင့် ချိတ်ဆက်ခြင်း
    spreadsheet_id = os.environ.get("GOOGLE_SHEET_ID")
    sheet = client.open_by_key(spreadsheet_id)
    
    print("Successfully connected to Google Sheet.")

if __name__ == "__main__":
    main()
