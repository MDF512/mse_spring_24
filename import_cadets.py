import csv
import json
from datetime import datetime
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

LIST_OF_ROSTERS = ['rosters/100s.csv', 'rosters/200s.csv', 'rosters/300s.csv', 'rosters/400s.csv']
JSON_FILE = 'cadets.json'

def update_cadets(csv_file, json_file):
    cadet_data = {}

    # Load existing cadet data from JSON file if it exists and is not empty
    if os.path.exists(json_file) and os.stat(json_file).st_size != 0:
        try:
            with open(json_file, 'r') as f:
                cadet_data = json.load(f)
        except json.decoder.JSONDecodeError:
            logging.error("Failed to decode JSON file: %s", json_file)

    # Read CSV and update cadet data
    with open(csv_file, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            section = row[" Section"]
            if "11" in section:
                as_year = '100'
            elif "21" in section:
                as_year = '200'
            elif "31" in section:
                as_year = '300'
            elif "41" in section:
                as_year = '400'
            else:
                as_year = '0'
            cadet = {
                "Last Name": row[" Last Name"].strip(),
                "First Name": row[" First Name"].strip(),
                "A Number": row[" A Number"].strip(),
                "Email": row[" Email"].strip(),
                "AS Year": as_year,
                "Flight": "",
                "Last Edited": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            a_number = row[" A Number"].strip()

            # Check if cadet already exists and update if necessary
            if a_number in cadet_data:
                cadet_data[a_number].update(cadet)
            else:
                cadet_data[a_number] = cadet

    # Write updated cadet data back to JSON file
    with open(json_file, 'w') as f:
        json.dump(cadet_data, f, indent=4)
    logging.info("Updated cadet data written to JSON file: %s", json_file)

if __name__ == '__main__':
    for roster in LIST_OF_ROSTERS:
        logging.info("Processing roster: %s", roster)
        update_cadets(roster, JSON_FILE)
