#!/usr/bin/env python3
"""
Test script for Google Sheets API integration.

This script tests the basic functionality of the Google Sheets API
without using the MCP server.
"""

import os
import pickle
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Define scopes for Google Sheets API
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

def get_credentials():
    """
    Get or refresh credentials for Google API access.
    
    Returns:
        The credentials object.
    """
    # Get credentials directory from environment or use default
    token_dir = os.environ.get('GDRIVE_CREDS_DIR', '.')
    token_path = Path(token_dir) / 'token.pickle'
    
    # Get client ID and secret from environment variables
    client_id = os.environ.get('CLIENT_ID')
    client_secret = os.environ.get('CLIENT_SECRET')
    
    # Path to the credentials JSON file
    creds_file = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'gcp-oauth.keys.json')
    
    creds = None
    
    # Try to load existing token
    if token_path.exists():
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    
    # If credentials don't exist or are invalid, refresh or create new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials")
            creds.refresh(Request())
        else:
            print("Initiating new OAuth flow")
            # Load client config from file or environment variables
            if Path(creds_file).exists():
                # Load from credentials file
                flow = InstalledAppFlow.from_client_secrets_file(
                    creds_file, SCOPES)
            else:
                # Create config from environment variables
                client_config = {
                    "installed": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                        "redirect_uris": ["http://localhost"]
                    }
                }
                flow = InstalledAppFlow.from_client_config(
                    client_config, SCOPES)
            
            # Run the OAuth flow
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for future use
        os.makedirs(os.path.dirname(token_path), exist_ok=True)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def create_spreadsheet(title):
    """
    Create a new Google Sheets spreadsheet.
    
    Args:
        title: The title of the spreadsheet.
    
    Returns:
        A dictionary containing information about the created spreadsheet.
    """
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    spreadsheet_body = {
        'properties': {
            'title': title
        }
    }
    
    request = service.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()
    
    print(f"Created new spreadsheet with ID: {response['spreadsheetId']}")
    
    return {
        'spreadsheet_id': response['spreadsheetId'],
        'title': response['properties']['title'],
        'sheets': [
            {
                'sheet_id': sheet['properties']['sheetId'],
                'title': sheet['properties']['title']
            }
            for sheet in response.get('sheets', [])
        ],
        'spreadsheet_url': response['spreadsheetUrl']
    }

def write_to_spreadsheet(spreadsheet_id, range_name, values):
    """
    Write data to a Google Sheets spreadsheet.
    
    Args:
        spreadsheet_id: The ID of the spreadsheet.
        range_name: The A1 notation range to write to.
        values: The values to write (2D array).
    
    Returns:
        A dictionary containing information about the update.
    """
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    body = {
        'values': values
    }
    
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    
    print(f"Updated {result.get('updatedCells')} cells")
    
    return result

def read_from_spreadsheet(spreadsheet_id, range_name):
    """
    Read data from a Google Sheets spreadsheet.
    
    Args:
        spreadsheet_id: The ID of the spreadsheet.
        range_name: The A1 notation range to read from.
    
    Returns:
        A dictionary containing the data from the range.
    """
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    
    values = result.get('values', [])
    
    print(f"Read {len(values)} rows from range {range_name}")
    
    return {
        'values': values,
        'num_rows': len(values),
        'num_columns': max([len(row) for row in values]) if values else 0
    }

def format_cells(spreadsheet_id, range_name, format_options):
    """
    Apply formatting to cells in a Google Sheets spreadsheet.
    
    Args:
        spreadsheet_id: The ID of the spreadsheet.
        range_name: The A1 notation range to format.
        format_options: The formatting options to apply.
    
    Returns:
        A dictionary containing information about the update.
    """
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    # Parse the range to get sheet ID and grid coordinates
    sheet_name = range_name.split('!')[0] if '!' in range_name else 'Sheet1'
    
    # Get the sheet ID
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()
    
    sheet_id = None
    for sheet in spreadsheet.get('sheets', []):
        if sheet.get('properties', {}).get('title') == sheet_name:
            sheet_id = sheet.get('properties', {}).get('sheetId')
            break
    
    if not sheet_id:
        print(f"Warning: Sheet '{sheet_name}' not found by name, using first sheet")
        # Use the first sheet as fallback
        if spreadsheet.get('sheets'):
            sheet_id = spreadsheet.get('sheets')[0].get('properties', {}).get('sheetId')
        else:
            raise ValueError("No sheets found in the spreadsheet")
    
    # Parse the cell range
    cell_range = range_name.split('!')[1] if '!' in range_name else range_name
    
    # Convert A1 notation to grid coordinates
    # This is a simplified version and may not work for all cases
    def a1_to_grid_range(a1_range):
        # Split the range into start and end cells
        if ':' in a1_range:
            start, end = a1_range.split(':')
        else:
            start = end = a1_range
        
        # Convert column letters to numbers
        def col_to_num(col):
            num = 0
            for c in col:
                if c.isalpha():
                    num = num * 26 + (ord(c.upper()) - ord('A') + 1)
            return num - 1  # 0-based index
        
        # Extract column letters and row numbers
        def cell_to_coords(cell):
            col = ''.join(c for c in cell if c.isalpha())
            row = ''.join(c for c in cell if c.isdigit())
            return col_to_num(col), int(row) - 1  # 0-based index
        
        start_col, start_row = cell_to_coords(start)
        end_col, end_row = cell_to_coords(end)
        
        return {
            'sheetId': sheet_id,
            'startRowIndex': start_row,
            'endRowIndex': end_row + 1,  # exclusive
            'startColumnIndex': start_col,
            'endColumnIndex': end_col + 1  # exclusive
        }
    
    grid_range = a1_to_grid_range(cell_range)
    
    # Prepare the request
    request = {
        'requests': [
            {
                'repeatCell': {
                    'range': grid_range,
                    'cell': {
                        'userEnteredFormat': format_options
                    },
                    'fields': 'userEnteredFormat'
                }
            }
        ]
    }
    
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=request
    ).execute()
    
    print(f"Formatted cells in range {range_name}")
    
    return response

def main():
    """Main function to test Google Sheets API integration."""
    print("Google Sheets API Test")
    print("=====================")
    
    # Create a new spreadsheet
    print("\n1. Creating a new spreadsheet...")
    spreadsheet = create_spreadsheet("Test Spreadsheet")
    spreadsheet_id = spreadsheet['spreadsheet_id']
    print(f"Spreadsheet URL: {spreadsheet['spreadsheet_url']}")
    
    # Write data to the spreadsheet
    print("\n2. Writing data to the spreadsheet...")
    data = [
        ["Name", "Age", "City"],
        ["Alice", 30, "New York"],
        ["Bob", 25, "San Francisco"],
        ["Charlie", 35, "Chicago"]
    ]
    write_to_spreadsheet(spreadsheet_id, "Sheet1!A1:C4", data)
    
    # Read data from the spreadsheet
    print("\n3. Reading data from the spreadsheet...")
    result = read_from_spreadsheet(spreadsheet_id, "Sheet1!A1:C4")
    print("Data:")
    for row in result['values']:
        print(row)
    
    # Format the header row
    print("\n4. Formatting the header row...")
    format_options = {
        "textFormat": {
            "bold": True
        },
        "backgroundColor": {
            "red": 0.8,
            "green": 0.8,
            "blue": 0.8
        },
        "horizontalAlignment": "CENTER"
    }
    format_cells(spreadsheet_id, "Sheet1!A1:C1", format_options)
    
    print("\nTest completed successfully!")
    print(f"View your spreadsheet at: {spreadsheet['spreadsheet_url']}")

if __name__ == "__main__":
    main()
