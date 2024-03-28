from google_auth import get_client
import json
import sys

ADMIN = {
    "Name": "Michael Fretz",
    "email": "mikefretz@gmail.com"
}
# Provide the sheet key if you want to open an existing sheet, leave it blank to create a new one
SHEET_KEY = ''
SHEET_NAME = "MSE 2 Spring 2024"

# Dont change these, If you do, go back into the Cadet class and update them
CADET_DATA_HEADERS = [
    'A Number',
    'Last Name',
    'First Name',
    'Flight',
    'AS Year'
]

# include the text "Comments" to have it be parsed as string comments


SCORES = [
    {'name': 'MKT', 'max_value': 100, 'type': 'Numerical'},
    {'name': 'Radio Reports Practical', 'max_value': 80, 'type': 'Numerical'},
    {'name': 'Radio Reports Practical Comments', 'max_value': None, 'type': 'Text'},
    {'name': 'Radio Reports Test', 'max_value': 20, 'type': 'Numerical'},
    {'name': 'MGRS Practical', 'max_value': 80, 'type': 'Numerical'},
    {'name': 'MGRS Practical Comments', 'max_value': None, 'type': 'Text'},
    {'name': 'MGRS Test', 'max_value': 20, 'type': 'Numerical'}
]

SCORE_NAME_LIST = [score['name'] for score in SCORES]

# Color for alternating rows
EVEN_ROW_COLOR = {"red": 1.0, "green": 1.0, "blue": 1.0}  # White
ODD_ROW_COLOR = {"red": 0.9, "green": 0.9, "blue": 0.9}  # Light gray
HEADER_ROW_COLOR = {"red": 0.7, "green": 0.7, "blue": 0.7}  # Gray


# Function to export data to JSON file
def export_to_json(sheet_id, headers, scores):
    data = {
        "sheet_id": sheet_id,
        "headers": headers,
        "scores": scores
    }
    with open('sheet_info.json', 'w') as f:
        json.dump(data, f)


def main():
    try:
        c = get_client()
        if SHEET_KEY:  # If a sheet key is provided, open the existing sheet
            sh = c.open_by_key(SHEET_KEY)
        else:  # Otherwise, create a new sheet
            sh = c.create(SHEET_NAME)
        sh.share(ADMIN["email"], perm_type='user', role='writer', notify=True)
        ws1 = sh.get_worksheet(0)
        ws1.update_title("old")
        ws = sh.add_worksheet("Scores", 1, 1)
        sh.del_worksheet(ws1)
        headers = CADET_DATA_HEADERS + SCORE_NAME_LIST
        ws.insert_row(headers)

        with open('cadets.json', 'r') as f:
            cadet_data = json.load(f)
        # Filter cadets with AS Year 100 or 200
        filtered_cadets = [cadet for cadet in cadet_data.values() if cadet.get('AS Year') in ['100', '200']]

        cadet_data_to_append = []
        for cadet in filtered_cadets:
            cadet_row = [cadet.get(header, '') for header in CADET_DATA_HEADERS]
            cadet_data_to_append.append(cadet_row)
        ws.insert_rows(cadet_data_to_append, row=2)


        ws.columns_auto_resize(0, len(headers)+20)

        # Export sheet information to JSON file
        export_to_json(sh.id, headers, SCORES)

    except Exception as e:
        print("An error occurred:", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
