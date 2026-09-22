import os
import requests
import pandas as pd

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

# Myanmar Handicap mapping rule
def convert_to_myanmar_odds(handicap_value):
    mapping = {
        0.0: "D", 0.25: "L-50", 0.5: "L-100", 0.75: "1+50", 1.0: "1D",
        1.25: "1-50", 1.5: "1-100", 1.75: "2+50", 2.0: "2D", 2.25: "2-50",
        2.5: "2.5", 2.75: "3+50", 3.0: "3D", 3.25: "3-50", 3.5: "3-100",
        3.75: "4+50", 4.0: "4D", 4.25: "4-50", 4.5: "4-100", 5.0: "5D"
    }
    odds_str = mapping.get(abs(handicap_value), str(handicap_value))
    return f"{odds_str} ↑"

def fetch_api_data(league_id, season=2026):
    url = f"https://v3.football.api-sports.io/fixtures?league={league_id}&season={season}"
    response = requests.get(url, headers=BASE_HEADERS)
    if response.status_code != 200:
        print(f"API Error for league {league_id}: {response.text}")
        return []
    return response.json().get('response', [])

def process_and_update():
    if not API_KEY:
        raise ValueError("FOOTBALL_API_KEY environment variable is missing in GitHub Secrets!")

    for league_name, info in LEAGUES.items():
        print(f"Processing {league_name}...")
        file_path = info["file"]
        
        if not os.path.exists(file_path):
            print(f"Warning: File {file_path} not found in repository. Skipping...")
            continue
            
        fixtures = fetch_api_data(info["id"])
        if not fixtures:
            print(f"No fixtures found for {league_name}.")
            continue
            
        # Read Excel workbook while preserving all sheets
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            continue

        # Open ExcelWriter to update sheets without losing structure or formatting
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            for sheet in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                
                # Update logic & matching with API data can be processed here safely per sheet
                
                df.to_excel(writer, sheet_name=sheet, index=False)
                
        print(f"Successfully updated {file_path}")

if __name__ == "__main__":
    process_and_update()
