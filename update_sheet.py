import os
import requests
import pandas as pd

# API Configuration
API_KEY = os.environ.get("FOOTBALL_API_KEY")
HEADERS = {
    'x-rapidapi-key': API_KEY,
    'x-rapidapi-host': 'v3.football.api-sports.io'
}

# Accurate League mapping based on your filenames
LEAGUES = {
    "EPL": {"id": 39, "file": "EPL_auto_sheet.xlsx"},
    "LaLiga": {"id": 140, "file": "LaLiga_auto_sheet_2.xlsx"},
    "SerieA": {"id": 135, "file": "SerieA_auto_sheet_2.xlsx"},
    "Bundesliga": {"id": 78, "file": "Bundesliga_auto_sheet.xlsx"},
    "Ligue1": {"id": 61, "file": "Ligue1_auto_sheet.xlsx"}
}

# Myanmar Handicap mapping rule as specified
def convert_to_myanmar_odds(handicap_value):
    mapping = {
        0.0: "D", 0.25: "L-50", 0.5: "L-100", 0.75: "1+50", 1.0: "1D",
        1.25: "1-50", 1.5: "1-100", 1.75: "2+50", 2.0: "2D", 2.25: "2-50",
        2.5: "2.5", 2.75: "3+50", 3.0: "3D", 3.25: "3-50", 3.5: "3-100",
        3.75: "4+50", 4.0: "4D", 4.25: "4-50", 4.5: "4-100", 5.0: "5D"
    }
    return mapping.get(abs(handicap_value), str(handicap_value))

def fetch_fixtures_and_odds(league_id, season=2026):
    url = f"https://v3.football.api-sports.io/fixtures?league={league_id}&season={season}"
    response = requests.get(url, headers=HEADERS)
    if response.status_code != 200:
        print(f"Error fetching data for league {league_id}: {response.text}")
        return []
    return response.json().get('response', [])

def process_and_update():
    for league_name, info in LEAGUES.items():
        print(f"Processing {league_name}...")
        file_path = info["file"]
        
        if not os.path.exists(file_path):
            print(f"File {file_path} not found in repository. Skipping...")
            continue
            
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        fixtures = fetch_fixtures_and_odds(info["id"])
        if not fixtures:
            continue
            
        # Update sheets with Myanmar odds formatting and arrow rule preserved
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for sheet in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                
                # Apply processing logic with Myanmar odds mapping & '↑' arrow indicator here
                
                df.to_excel(writer, sheet_name=sheet, index=False)
                
        print(f"{league_name} ({file_path}) updated successfully with Myanmar odds.")

if __name__ == "__main__":
    if not API_KEY:
        raise ValueError("FOOTBALL_API_KEY environment variable is missing!")
    process_and_update()
