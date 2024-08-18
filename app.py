import streamlit as st
import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from datetime import datetime
from dateutil import parser

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']

def format_date(date):
    return date.strftime('%b %d, %Y')

def get_google_services():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret_318451619057-79npennvsludigngm43mf4c68egp96kk.apps.googleusercontent.com.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return sheets_service, drive_service


def get_column_values(service, spreadsheet_id, sheet_name, start_column, end_column=None, max_rows=1000):
    try:
        if end_column:
            range_name = f"'{sheet_name}'!{start_column}2:{end_column}{max_rows}"
        else:
            range_name = f"'{sheet_name}'!{start_column}2:{start_column}{max_rows}"
        
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        values = result.get('values', [])
        if not values:
            return []
        
        if end_column:
            return values
        else:
            return list(set([item[0] for item in values if item]))  # Get unique values
    except HttpError as error:
        st.error(f"An error occurred: {error}")
        return []
    
def find_associated_helpers(service, spreadsheet_id, knowledge_file):
    helpers_data = get_column_values(service, spreadsheet_id, 'Current Helpers', 'A', 'F')
    associated_helpers = set()  # Using a set instead of a list to avoid duplicates
    
    for row in helpers_data:
        if len(row) >= 6:  # Ensure the row has enough columns
            helper_name = row[0]
            knowledge_file_list = row[5]
            if knowledge_file_list:
                files = [file.strip() for file in knowledge_file_list.split(',')]
                if knowledge_file in files:
                    associated_helpers.add(helper_name)  # Add to set instead of appending to list
    
    return sorted(list(associated_helpers))  # Convert set to sorted list before returning

def get_latest_entry(service, spreadsheet_id, sheet_name, helper_name=None, helper_type=None, gai=None):
    range_name = f"'{sheet_name}'!A2:I1000"  # Adjust the range as needed
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    values = result.get('values', [])
    if not values:
        return None
    
    # Filter the values based on the selected criteria
    filtered_values = [
        row for row in values
        if (helper_name is None or row[0] == helper_name) and
           (helper_type is None or row[1] == helper_type) and
           (gai is None or row[2] == gai)
    ]
    
    if not filtered_values:
        return None
    
    # Sort by UpdatedDate (assuming it's the 8th column, index 7)
    sorted_values = sorted(filtered_values, key=lambda x: parser.parse(x[7]) if len(x) > 7 else parser.parse('1900-01-01'), reverse=True)
    
    if sorted_values:
        latest_entry = sorted_values[0]
        # Format the dates in the latest entry
        if len(latest_entry) > 7:
            latest_entry[7] = format_date(parser.parse(latest_entry[7]))
        if len(latest_entry) > 8:
            latest_entry[8] = format_date(parser.parse(latest_entry[8]))
        # Ensure the entry has all 9 fields, fill with None if missing
        return latest_entry + [None] * (9 - len(latest_entry))
    return None

def get_filtered_values(service, spreadsheet_id, sheet_name, column_letter, filter_dict):
    range_name = f"'{sheet_name}'!A2:I1000"  # Adjust the range as needed
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    values = result.get('values', [])
    
    filtered_values = []
    for row in values:
        if all(row[i] == filter_dict[col] for i, col in enumerate(['A', 'B', 'C']) if col in filter_dict):
            if len(row) > ord(column_letter) - ord('A'):
                filtered_values.append(row[ord(column_letter) - ord('A')])
    
    return list(set(filtered_values))  # Return unique values

def display_confirmation(data):
    st.subheader("Confirmation Screen")
    st.write("Please confirm the following information:")
    for key, value in data.items():
        if key in ['UpdatedDate', 'CreatedDate']:
            if value:
                try:
                    # Try parsing with the format '%Y-%m-%d %H:%M:%S'
                    date_obj = datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    try:
                        # If that fails, try parsing with the format '%b %d, %Y'
                        date_obj = datetime.strptime(value, '%b %d, %Y')
                    except ValueError:
                        # If both fail, just display the original string
                        st.text(f"{key}: {value}")
                        continue
                # Format the date to 'Aug 18, 2024' format
                value = date_obj.strftime('%b %d, %Y')
        st.text(f"{key}: {value}")
    col1, col2 = st.columns(2)
    with col1:
        confirm = st.button("Confirm Submission")
    with col2:
        cancel = st.button("Cancel")
    return confirm, cancel


def get_index(options, value):
    try:
        return options.index(value)
    except ValueError:
        return 0

def add_new_row(service, spreadsheet_id, sheet_name, helper_name, helper_type, gai, custom_instructions, knowledge_file_num, knowledge_file_list, modified_file_list, created_date):
    current_time = datetime.now()
    formatted_time = format_date(current_time)
    values = [[
        helper_name, 
        helper_type, 
        gai, 
        custom_instructions, 
        knowledge_file_num, 
        knowledge_file_list, 
        modified_file_list, 
        formatted_time,  # UpdatedDate
        created_date     # CreatedDate (retained from the last entry)
    ]]
    body = {'values': values}
    range_name = f"'{sheet_name}'!A:I"
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()
    return result

def search_file_in_folder(drive_service, folder_id, file_name):
    try:
        query = f"'{folder_id}' in parents and name = '{file_name}' and mimeType = 'application/vnd.google-apps.document'"
        results = drive_service.files().list(q=query, fields="files(id, webViewLink)").execute()
        files = results.get('files', [])
        if files:
            return files[0]['webViewLink']
        return None
    except HttpError as error:
        st.error(f"An error occurred while searching for the file: {error}")
        return None

def add_knowledge_file_row(service, spreadsheet_id, sheet_name, file_name, change):
    current_time = datetime.now()
    formatted_time = format_date(current_time)
    values = [[file_name, change, formatted_time]]
    body = {'values': values}
    range_name = f"'{sheet_name}'!A:C"
    try:
        result = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        return result
    except HttpError as error:
        st.error(f"An error occurred while adding the row: {error}")
        return None

def main():
    st.title('Google Sheet Updater')

    try:
        sheets_service, drive_service = get_google_services()
        spreadsheet_id = '1PjH8MJ__WhWtjXN3gbB0ZRFLF2sLocs1sVxFVpLZG5c'
        
        # Sheet selection
        sheet_options = ['Current Helpers', 'Knowledge Files']
        selected_sheet = st.radio("Select Sheet", sheet_options)

        if selected_sheet == 'Current Helpers':
            sheet_name = 'Current Helpers'

            # State management
            if 'confirmation_state' not in st.session_state:
                st.session_state.confirmation_state = False
            if 'submission_data' not in st.session_state:
                st.session_state.submission_data = {}
            if 'user_inputs' not in st.session_state:
                st.session_state.user_inputs = {
                    'helper': '',
                    'helper_type': '',
                    'gai': '',
                    'custom_instructions': '',
                    'knowledge_file_num': '',
                    'knowledge_file_list': '',
                    'modified_file_list': '',
                    'created_date': ''
                }

            if not st.session_state.confirmation_state:
                # Hierarchical dropdowns
                helper_names = [''] + get_column_values(sheets_service, spreadsheet_id, sheet_name, 'A')
                selected_helper = st.selectbox('Select Helper Name', helper_names, index=get_index(helper_names, st.session_state.user_inputs['helper']))

                helper_types = [''] + get_filtered_values(sheets_service, spreadsheet_id, sheet_name, 'B', {'A': selected_helper})
                selected_helper_type = st.selectbox('Select Helper Type', helper_types, index=get_index(helper_types, st.session_state.user_inputs['helper_type']))

                gai_values = [''] + get_filtered_values(sheets_service, spreadsheet_id, sheet_name, 'C', {'A': selected_helper, 'B': selected_helper_type})
                selected_gai = st.selectbox('Select GAI', gai_values, index=get_index(gai_values, st.session_state.user_inputs['gai']))

                # Fetch latest entry based on selections
                if selected_helper and selected_helper_type and selected_gai:
                    latest_entry = get_latest_entry(sheets_service, spreadsheet_id, sheet_name, selected_helper, selected_helper_type, selected_gai)
                    if latest_entry and not st.session_state.user_inputs['custom_instructions']:
                        st.session_state.user_inputs.update({
                            'custom_instructions': latest_entry[3] if latest_entry[3] is not None else "",
                            'knowledge_file_num': latest_entry[4] if latest_entry[4] is not None else "",
                            'knowledge_file_list': latest_entry[5] if latest_entry[5] is not None else "",
                            'modified_file_list': latest_entry[6] if latest_entry[6] is not None else "",
                            'created_date': latest_entry[8] if latest_entry[8] is not None else ""
                        })
                    elif not latest_entry:
                        st.info("No matching entry found for the selected combination.")

                # Text inputs with retained values
                custom_instructions = st.text_area('Enter Custom Instructions', value=st.session_state.user_inputs['custom_instructions'])
                knowledge_file_num = st.text_input('Enter Knowledge File Number', value=st.session_state.user_inputs['knowledge_file_num'])
                knowledge_file_list = st.text_input('Enter Knowledge File List', value=st.session_state.user_inputs['knowledge_file_list'])
                modified_file_list = st.text_input('Enter Modified File List', value=st.session_state.user_inputs['modified_file_list'])

                # Update user_inputs with current values
                st.session_state.user_inputs.update({
                    'helper': selected_helper,
                    'helper_type': selected_helper_type,
                    'gai': selected_gai,
                    'custom_instructions': custom_instructions,
                    'knowledge_file_num': knowledge_file_num,
                    'knowledge_file_list': knowledge_file_list,
                    'modified_file_list': modified_file_list
                })

                # Submit button
                if st.button('Submit'):
                    if all([selected_helper, selected_helper_type, selected_gai, custom_instructions, knowledge_file_num, knowledge_file_list, modified_file_list]):
                        st.session_state.submission_data = {
                            'Helper Name': selected_helper,
                            'Helper Type': selected_helper_type,
                            'GAI': selected_gai,
                            'Custom Instructions': custom_instructions,
                            'Knowledge File Number': knowledge_file_num,
                            'Knowledge File List': knowledge_file_list,
                            'Modified File List': modified_file_list,
                            'CreatedDate': st.session_state.user_inputs['created_date']
                        }
                        st.session_state.confirmation_state = True
                        st.rerun()
                    else:
                        st.error('Please fill in all fields.')
            else:
                confirm, cancel = display_confirmation(st.session_state.submission_data)
                if confirm:
                    result = add_new_row(
                        sheets_service,
                        spreadsheet_id,
                        sheet_name,
                        st.session_state.submission_data['Helper Name'],
                        st.session_state.submission_data['Helper Type'],
                        st.session_state.submission_data['GAI'],
                        st.session_state.submission_data['Custom Instructions'],
                        st.session_state.submission_data['Knowledge File Number'],
                        st.session_state.submission_data['Knowledge File List'],
                        st.session_state.submission_data['Modified File List'],
                        st.session_state.submission_data['CreatedDate']
                    )
                    st.success('New row added successfully!')
                    st.session_state.confirmation_state = False
                    st.session_state.submission_data = {}
                    st.session_state.user_inputs = {key: '' for key in st.session_state.user_inputs}
                    st.rerun()
                elif cancel:
                    st.session_state.confirmation_state = False
                    st.rerun()

        elif selected_sheet == 'Knowledge Files':
            st.subheader("Knowledge Files")

            # Get file names from the Knowledge Files sheet
            file_names = get_column_values(sheets_service, spreadsheet_id, 'Knowledge Files', 'A')
            selected_file = st.selectbox('Select Knowledge File', [''] + file_names)

            if selected_file:
                folder_id = "1A3z87xj15KlVYQC87xFIbw-R23QU62PB"  # Replace with the actual folder ID
                file_url = search_file_in_folder(drive_service, folder_id, selected_file)
                
                if file_url:
                    st.markdown(f"[Open {selected_file} in Google Docs]({file_url})")
                else:
                    st.warning(f"File '{selected_file}' not found in the specified folder.")

                # Find and display associated helpers
                associated_helpers = find_associated_helpers(sheets_service, spreadsheet_id, selected_file)
                if associated_helpers:
                    st.subheader("Associated Helpers:")
                    for helper in associated_helpers:
                        st.write(f"- {helper}")
                else:
                    st.info("No helpers are currently associated with this knowledge file.")

                change = st.text_area('Enter Change Description')

                if st.button('Submit'):
                    if change:
                        result = add_knowledge_file_row(sheets_service, spreadsheet_id, 'Knowledge Files', selected_file, change)
                        if result:
                            st.success('New row added successfully to Knowledge Files sheet!')
                        else:
                            st.error('Failed to add new row.')
                    else:
                        st.error('Please enter a change description.')

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    main()