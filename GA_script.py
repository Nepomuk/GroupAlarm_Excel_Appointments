import pandas as pd
import requests
import json
import argparse
from datetime import datetime, timedelta
import pytz

# personal API key
API_PERSONAL_KEY = "123456789"  # Replace with your personal API key (keep this secret!)

# organization ID
ORGANIZATION_ID = 12345  # Replace with your organization

# list of used labels and their corresponding IDs in GroupAlarm
#
# This list maps the names of labels in the Excel sheet to the labels used 
# in GroupAlarm. This makes it easier to read in the Excel sheet.
#
# Expand the list based on this example. The number in append(x) corresponds 
# to the ID of the label, the text in "x" corresponds to the label name you've
# used in the Excel file.
def assign_label_IDs(row):
    labelIDs = []
    labelIDs.append(1) if label_is_set(row, "OV") else 0
    labelIDs.append(2) if label_is_set(row, "TZ") else 0
    labelIDs.append(3) if label_is_set(row, "Küche") else 0
    labelIDs.append(4) if label_is_set(row, "Jugend") else 0
    return labelIDs

# -------------------------------------------------------------
#  no further changes should be required below this point
#

# API configuration
API_BASE_URL = "https://app.groupalarm.com/api/v1"
HEADERS = {
    "Personal-Access-Token": f"{API_PERSONAL_KEY}",
    "Content-Type": "application/json"
}
# Timezone information
LOCAL_TZ = pytz.timezone("Europe/Berlin")

# Read in excel file
def read_excel(file_path):
    return pd.read_excel(file_path)

# Enter appointments through API
def create_appointment(args, appointment_data):
    url = f"{API_BASE_URL}/appointment"
    if args.verbose:
        print(f"Accessing the GroupAlarm API through {url}")

    if args.dry_run:
        print(f"DRY RUN ... skipping new appointment {appointment_data['name']}.")
    else:
        if args.verbose:
            print(f"Creating appointment {appointment_data['name']} ...")
        response = requests.post(url, headers=HEADERS, json=appointment_data)
        if response.status_code == 201:
            print(f"Appointment created successfully: {appointment_data['name']} am {appointment_data['startDate']}")
        else:
            print(f"Error while creating the appointment: {response.status_code}, {response.text}")

# Convert data and time to ISO-8601
def combine_datetime_and_time(date_obj, time_str):
    time_obj = datetime.strptime(time_str, "%H:%M:%S").time()
    naive_dt = datetime.combine(date_obj, time_obj)
    local_dt = LOCAL_TZ.localize(naive_dt, is_dst=None)
    utc_dt = local_dt.astimezone(pytz.utc)
    return utc_dt.isoformat()

# Substract a number of days from a ISO-8601 date and retults the in ISO format
def subtract_days_from_datetime(datetime_iso, days):
    datetime_obj = datetime.fromisoformat(datetime_iso)  # Entferne 'Z' für UTC
    new_datetime = datetime_obj - timedelta(days=days)
    return new_datetime.isoformat()  # Zurück ins ISO-Format mit UTC-Suffix

def label_is_set(row, label):
    if label in row:
        if row[label]:
            return True
    return False

# check if multiple sheets are available and ask which one to read
def select_sheet(args, appointments):
    sheets = list(appointments.keys())
    if len(sheets) > 1:
        print(f"There are {len(sheets)} sheets in this document:")
        sheet_counter = 0
        for sheet in sheets:
            print(f"{sheet_counter}: {sheet}")
            sheet_counter = sheet_counter + 1
        
        while True:
            selected_sheet = input("Please enter the number of the sheet you like to read: ")
            try:
                selected_sheet = int(selected_sheet)
            except ValueError:
                print("Invalid number")
                continue
            if 0 <= selected_sheet <= len(sheets)-1:
                break
            else:
                print("Invalid range, please only select an available number")

        print(f"Continue parsing sheet '{sheets[selected_sheet]}' from the file.\n")
        appointments = appointments[sheets[selected_sheet]]
    else:
        if args.verbose:
            print(f"Only one sheet detected, reading sheet {sheet[0]}")
        appointments = appointments[sheets[0]]

    return appointments


# Main function
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("file", type=str, help="Excel file to load")
    parser.add_argument("-n", "--dry_run", help="Just parse input but don't access the GroupAlarm API", action="store_true")
    parser.add_argument("-v", "--verbose", help="Increase output verbosity", action="store_true")
    #parser.add_argument("-r", "--reminder", type=int, help="Sets an optional reminder for all entries x days before the event.")
    args = parser.parse_args()

    if args.verbose:
        print(f"Loading file {args.file} and checking for multiple sheets ...")
    appointments = pd.read_excel(args.file, sheet_name=None)
    appointments = select_sheet(args, appointments)
    
    if args.verbose:
        print("\n\nCollecting data from rows")

    row_counter = 0
    for _, row in appointments.iterrows():
        row_counter = row_counter+1
        datetime_start = combine_datetime_and_time(row["Start Tag"], str(row["Start Zeit"]))
        datetime_end = combine_datetime_and_time(row["Ende Tag"], str(row["Ende Zeit"]))
        if row["Einladung x Tage vorher"] > 0:
            datetime_notify = subtract_days_from_datetime(datetime_start, row["Einladung x Tage vorher"])
        else:
            datetime_notify = ""
        labelIDs = assign_label_IDs(row)
        appointment_data = {
            "organizationID": ORGANIZATION_ID,
            "name": row["Titel"],
            "description": row["Beschreibung"],
            "startDate": datetime_start,  # Format: ISO 8601 (z. B. "2024-11-18T10:00:00Z")
            "endDate": datetime_end,      # Format: ISO 8601
            "notificationDate": datetime_notify,
            "labelIDs": labelIDs,
            "participants": [],
            "isPublic": row["Öffentlich"],
            "keepLabelParticipantsInSync": True,
        }
        print(f"{datetime_start}: {row["Titel"]}")
        if args.verbose:
            print(f"\nParsed appointment data of row {row_counter}")
            print(appointment_data)

        if args.verbose:
            print("\n\nParsing completed, uploading data...")
        create_appointment(args, appointment_data)

if __name__ == "__main__":
    main()

