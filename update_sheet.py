import os
import requests
import pandas as pd
from datetime import datetime

# API Configuration from GitHub Secrets
API_KEY = os.environ.get("FOOTBALL_API_KEY")
BASE_HEADERS = {
    'x-rapidapi-key': API_KEY,
    'x-rapidapi-host': 'v3.football.api-sports.io'
}

# League mapping with respective exact Excel filenames in your repository
LEAGUES = {
    "EPL": {"id": 39, "file": "EPL_auto_sheet.xlsx"},
    "LaLiga": {"id": 140, "file": "LaLiga_auto_sheet_2.xlsx"},
    "SerieA": {"id": 135, "file": "SerieA_auto_sheet_2.xlsx"},
    "Bundesliga": {"id": 78, "file": "Bundesliga_auto_sheet.xlsx"},
    "Ligue1": {"id": 61, "file": "Ligue1_auto_sheet.xlsx"}
}

def fetch_api_data(league_id, season=2026):
    # Fetching finished fixtures (FT) to ensure only authentic data is used for 2026/2027 season
    url = f"https://v3.football.api-sports.io/fixtures?league={league_id}&season={season}&status=FT"
    response = requests.get(url, headers=BASE_HEADERS)
    if response.status_code != 200:
        print(f"API Error for league {league_id}: {response.text}")
        return []
    return response.json().get('response', [])

def process_and_update():
    if not API_KEY:
        raise ValueError("FOOTBALL_API_KEY environment variable is missing in GitHub Secrets!")

    # Check execution day (Allowing Tuesday, Wednesday, and Thursday updates)
    current_day = datetime.now().strftime('%A')
    print(f"Current execution day: {current_day}")
    
    if current_day not in ['Tuesday', 'Wednesday', 'Thursday']:
        print("Today is not an update day (Tuesday/Wednesday/Thursday). Skipping scheduled update.")
        return

    for league_name, info in LEAGUES.items():
        print(f"Processing {league_name} on {current_day}...")
        file_path = info["file"]
        
        if not os.path.exists(file_path):
            print(f"Warning: File {file_path} not found in repository. Skipping...")
            continue
            
        fixtures = fetch_api_data(info["id"])
        if not fixtures:
            print(f"No completed fixtures found for {league_name}. Holding / Pending (No Fake data).")
            continue
            
        # Parse API fixtures into a lookup dictionary
        finished_matches = {}
        for fix in fixtures:
            match_date = fix['fixture']['date'][:10] # YYYY-MM-DD
            home = fix['teams']['home']['name'].strip().lower()
            away = fix['teams']['away']['name'].strip().lower()
            h_goals = fix['goals']['home']
            a_goals = fix['goals']['away']
            
            if h_goals is not None and a_goals is not None:
                finished_matches[(match_date, home, away)] = f"{h_goals}-{a_goals}"

        if not finished_matches:
            print(f"No verified FT scores available for {league_name}.")
            continue

        # Read Excel workbook while preserving all sheets and structure
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            continue

        updated_any = False

        # Open ExcelWriter to update sheets without losing structure or formatting
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            for sheet in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                
                sheet_team = sheet.strip().lower()
                
                # Iterate through rows to match Date, Opponent and update Score (Column I / index 8)
                for idx, row in df.iterrows():
                    date_val = row.iloc[1] # Column B (Date)
                    opp_val = str(row.iloc[4]).strip().lower()  # Column E (Opponent)
                    
                    if pd.isna(date_val):
                        continue
                    
                    match_date_str = str(date_val)[:10]
                    
                    score_found = None
                    key_home = (match_date_str, sheet_team, opp_val)
                    key_away = (match_date_str, opp_val, sheet_team)
                    
                    if key_home in finished_matches:
                        score_found = finished_matches[key_home]
                    elif key_away in finished_matches:
                        score_found = finished_matches[key_away]
                        
                    if score_found:
                        current_score = str(row.iloc[8]) if pd.notna(row.iloc[8]) else ""
                        if current_score != score_found:
                            df.iat[idx, 8] = score_found
                            updated_any = True

                df.to_excel(writer, sheet_name=sheet, index=False)
                
        if updated_any:
            print(f"Successfully updated {file_path} with verified match scores for {current_day}.")
        else:
            print(f"No new matching scores to update for {file_path}.")

if __name__ == "__main__":
    process_and_update()
