"""
Interact with the Google Sheets API to get and post data to the database.
"""
import pytz, asyncio, json, uuid
from aiogoogle import Aiogoogle
from datetime import datetime, date
import google_secrets

# EDIT THE FOLLOWING VALUES 
service_account_creds = {
    "scopes": [
        "https://www.googleapis.com/auth/spreadsheets"
    ],
    **json.load(open(google_secrets.service_account_creds))
}

SHEET_ID = google_secrets.sheet_id
TIMEZONE = "America/Los_Angeles"
ADMIN = "@reyanlee"

# register a user in the database
def register_user_handler(hkn_handle, text):
    info = text.split(" ")
    if len(info) != 4:
        return {
            "body": "Please use the following format: `register <First name> <Last name> <Officer | AssistantOfficer | Member>`",
            "header": "Error registering user: Invalid format"
        }
    return asyncio.run(register_user(hkn_handle, info[1], info[2], info[3]))

async def register_user(hkn_handle, first_name, last_name, member_type):
    async with Aiogoogle(service_account_creds=service_account_creds) as google:
        sheets_api = await google.discover("sheets", "v4")
        name_match = await find_all_column(google, sheets_api, 'Hkners', 'A', hkn_handle)
        if name_match:
            return {
                'body': f"You are already registered in the attendance system. Please notify {ADMIN} if you think this is a mistake.", 
                'header': "Error: User already exists"
            }
        else:
            #member_type = member_type.split("|")[1][:-1]
            await insert_row(google, sheets_api, 'Hkners', [hkn_handle, first_name, last_name, member_type])
            return {
                'body': f"Registered {first_name} into attendance system. You can now check into events.",
                'header': "Success"
            }

# create a new event and password (admin only)
def create_event_handler(hkn_handle, text):
    info = text.split(" ")[1:]
    if len(info) != 3:
        return {
            "body": "Please use the following format: `newevent \"<event name>\" <event type> <password>`",
            "header": "Error creating event: Invalid format"
        }
    #info = ["", " ".join(info[:-1])[1:-1], info[-1]]
    if info[1] not in ["HM", "GM", "CM", "SnackAttack", "ReviewSession", "Intercommittee", "QSM", "MSM"]:
        return {
            "body": "Please use one of the following event types: HM, GM, CM, SnackAttack, ReviewSession, Intercommittee, QSM, MSM",
            "header": "Error creating event: Invalid type"
        }
    return asyncio.run(create_event(hkn_handle, info[1], info[2], info[3]))

async def create_event(hkn_handle, event_name, event_type, password):
    async with Aiogoogle(service_account_creds=service_account_creds) as google:
        sheets_api = await google.discover("sheets", "v4")
        admin_match = await find_all_column(google, sheets_api, 'Hkners', 'A', hkn_handle)
        if not admin_match:
            return {
                'body': f"Unable to create event. Please contact {ADMIN} if you think this is a mistake.", 
                'header': "Error: No permissions for event creation"
            }
        event_match = await find_all(google, sheets_api, 'Events', 'B', password.upper())
        if event_match:
            return {
                'body': "This password already exists. Please choose a new one.", 
                'header': "Error: Password already exists"
            }
        await insert_row(google, sheets_api, 'Events', [event_name, password.upper(), event_type, hkn_handle])
        return {
            'body': f"Created {event_type} event {event_name}. Use password \"{password}\" to check in.",
            'header': "Success"
        }

# check in to an event with a code
def checkin_handler(hkn_handle, text):
    info = text.split(" ")
    if len(info) != 2:
        return {
            "body": "Please use the following format: `checkin <password>`",
            "header": "Error checking in: Invalid format"
        }
    return asyncio.run(checkin(hkn_handle, info[1].upper()))

async def checkin(hkn_handle, password):
    async with Aiogoogle(service_account_creds=service_account_creds) as google:
        sheets_api = await google.discover("sheets", "v4")
        event_match = await find_all(google, sheets_api, 'Events', 'B', password.upper())
        if not event_match:
            return {
                'body': "Unable to check in, event not found.", 
                'header': "Error: Event not found"
            }
        event_name = event_match[0][0]
        event_password = event_match[0][1]
        event_type = event_match[0][2]
        checkin_time = datetime.now(pytz.timezone(TIMEZONE)).strftime("%m/%d/%Y %H:%M")
        attendance_match = await find_all(google, sheets_api, 'Attendance', 'A', hkn_handle)
        for row in attendance_match:
            if row[3] == event_password:
                return {
                    'body': f"You have already checked into {event_name}.", 
                    'header': "Error: Already checked in"
                }
        await insert_row(google, sheets_api, 'Attendance', [hkn_handle, event_name, event_type, event_password, checkin_time])
        return {
            'body': f"Checked in to event \"{event_name}\".",
            'header': "Success"
        }

# send a summary of events to a user
def update_user_handler(hkn_handle):
    return asyncio.run(update_user(hkn_handle))

async def update_user(hkn_handle):
    async with Aiogoogle(service_account_creds=service_account_creds) as google:
        sheets_api = await google.discover("sheets", "v4")
        total_attendance = await google.as_service_account(sheets_api.spreadsheets.values.get(spreadsheetId=SHEET_ID, range=f"Summary!A2:G2", valueRenderOption='FORMATTED_VALUE', dateTimeRenderOption='FORMATTED_STRING'))
        total_attendance = total_attendance['values'][0]
        total_hm = total_attendance[2]
        total_gm = total_attendance[3]
        total_cm = total_attendance[4]
        total_qsm = total_attendance[8]
        total_msm = total_attendance[9]
        name_match = await find_all(google, sheets_api, 'Progress', 'A', hkn_handle)
        if not name_match:
            return {
                'body': "You are not registered in the attendance system. Please use `register` to register.", 
                'header': "Error: User not registered"
            }
        else:
            hm = name_match[0][2]
            gm = name_match[0][3]
            cm = name_match[0][4]
            snackattack = name_match[0][5]
            intercommittee = name_match[0][7]
            qsm = name_match[0][8]
            msm = name_match[0][9]
            return {
                'header': f"{name_match[0][1]}'s attendance.",
                'body': f"_HM:_ {hm} out of {total_hm}\n_GM:_ {gm} out of {total_gm}\n_CM:_ {cm} out of {total_cm}\n_Snack_Attacks:_ {snackattack} out of 2\n_Inter-committee_Duties:_ {intercommittee} out of 1\n_QSM_: {qsm} out of {total_qsm}\n_MSM_: {msm} out of {total_msm}"
            }

# send a summary of the event to a user
def event_status_handler(hkn_handle, text):
    info = text.split(" ")
    if len(info) != 2:
        return {
            "body": "Please use the following format: `eventstatus <password>`",
            "header": "Error checking event status: Invalid format"
        }
    return asyncio.run(event_status(hkn_handle, info[1].upper()))

async def event_status(hkn_handle, password):
    async with Aiogoogle(service_account_creds=service_account_creds) as google:
        sheets_api = await google.discover("sheets", "v4")
        admin_match = await find_all_column(google, sheets_api, 'Hkners', 'A', hkn_handle)
        if not admin_match:
            return {
                'body': f"Unable to see event status. Please contact {ADMIN} if you think this is a mistake.", 
                'header': "Error: No access"
            }
        event_match = await find_all(google, sheets_api, 'Events', 'B', password)
        if not event_match:
            return {
                'body': "Unable to check event status, event not found.", 
                'header': "Error: Event not found"
            }
        event_name = event_match[0][0]
        event_type = event_match[0][1]
        attendance_match = await find_all(google, sheets_api, 'Attendance', 'D', password)
        if not attendance_match:
            return {
                'body': f"No one has checked into {event_name} yet. Use code {password} to check in.", 
                'header': "Error: Event not found"
            }
        if len(attendance_match) == 1:
            return {
                'body': f"1 person has checked into {event_type} event {event_name} (code: {password}): {attendance_match[0][5]}.", 
                'header': "Success"
            }
        people = ""
        for row in attendance_match:
            people += f"- {row[5]}\n"
        people = people[:-1]
        return {
            'body': f"{len(attendance_match)} people have checked into {event_type} event {event_name} (code: {password}).\n{people}",
            'header': "Success"
        }


''' HELPER FUNCTIONS '''
# gsheet post and get functions

# get a single column's matching values
async def find_all_column(google, sheets_api, sheet_name, column_letter, text):
    all_values_request = sheets_api.spreadsheets.values.get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!{column_letter}2:{column_letter}", valueRenderOption='FORMATTED_VALUE', dateTimeRenderOption='FORMATTED_STRING')
    all_values = await google.as_service_account(all_values_request)
    matches = []
    if 'values' in all_values:
        for v in all_values['values']:
            if v[0] == text:
                matches.append(v[0])
    return matches

# get all rows with matching values
async def find_all(google, sheets_api, sheet_name, column_letter, text):
    all_values_request = sheets_api.spreadsheets.values.get(spreadsheetId=SHEET_ID, range=f"{sheet_name}", valueRenderOption='FORMATTED_VALUE', dateTimeRenderOption='FORMATTED_STRING')
    all_values = await google.as_service_account(all_values_request)
    value_index = ord(column_letter) - 65 # this will break if column letter is beyond Z
    matches = []
    if 'values' in all_values:
        for v in all_values['values']:
            if v[value_index] == text:
                matches.append(v)
    return matches

# insert a new row at the bottom of sheet
async def insert_row(google, sheets_api, sheet_name, data):
    value_range_body = {
        "range": sheet_name,
        "majorDimension": 'ROWS',
        "values": [data]
    }
    request = sheets_api.spreadsheets.values.append(range=sheet_name, spreadsheetId=SHEET_ID, valueInputOption='RAW', insertDataOption='OVERWRITE', json=value_range_body)
    return await google.as_service_account(request)