#!/usr/bin/env python3
# server.py

import os
from typing import Dict, List
from fastmcp import FastMCP
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json

# Create an MCP server
mcp = FastMCP("GoogleSheets")

# Define scopes for Google Sheets API and Drive API (for listing sheets)
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'  # Added for listing sheets
]

# Environment variable for credentials
CREDENTIALS_FILE = os.environ.get('GOOGLE_CREDENTIALS_FILE', 'credentials.json')
TOKEN_FILE = os.environ.get('GOOGLE_TOKEN_FILE', 'token.json')

def get_credentials():
    """Get Google API credentials from environment or file"""
    creds = None
    
    # Check if we have a token file
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'r', encoding='utf-8') as token_file:
            creds = Credentials.from_authorized_user_info(json.load(token_file))
    
    # If credentials don't exist or are invalid
    if not creds or not creds.valid:
        # Check if we can refresh
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        # Check if we have service account credentials
        elif os.path.exists(CREDENTIALS_FILE):
            # Check if it's a service account key
            try:
                creds = service_account.Credentials.from_service_account_file(
                    CREDENTIALS_FILE, scopes=SCOPES)
            except ValueError:
                # If not a service account, it's a client secret
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
                
            # Save the credentials for the next run
            with open(TOKEN_FILE, 'w', encoding='utf-8') as token:
                token.write(creds.to_json())
        else:
            raise ValueError("No valid credentials found. Please set GOOGLE_CREDENTIALS_FILE environment variable.")
    
    return creds

def get_sheets_service():
    """Build and return the Google Sheets service"""
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    return service.spreadsheets()

def get_drive_service():
    """Build and return the Google Drive service"""
    creds = get_credentials()
    return build('drive', 'v3', credentials=creds)

# Data Operations

@mcp.tool()
def read_sheet(spreadsheet_id: str, sheet_range: str) -> List[List[str]]:
    """Read data from a Google Sheet
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_range: The range to read (e.g., 'Sheet1!A1:D10')
        
    Returns:
        The data from the specified range
    """
    sheets = get_sheets_service()
    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet_range
    ).execute()
    
    return result.get('values', [])

@mcp.tool()
def write_sheet(spreadsheet_id: str, sheet_range: str, values: List[List[str]]) -> Dict:
    """Write data to a Google Sheet
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_range: The range to write to (e.g., 'Sheet1!A1:D10')
        values: The data to write
        
    Returns:
        Status of the operation
    """
    sheets = get_sheets_service()
    body = {
        'values': values
    }
    result = sheets.values().update(
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    
    return {
        'updated_cells': result.get('updatedCells'),
        'updated_range': result.get('updatedRange')
    }

@mcp.tool()
def append_sheet(spreadsheet_id: str, sheet_range: str, values: List[List[str]]) -> Dict:
    """Append data to a Google Sheet
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_range: The range to append to (e.g., 'Sheet1!A1')
        values: The data to append
        
    Returns:
        Status of the operation
    """
    sheets = get_sheets_service()
    body = {
        'values': values
    }
    result = sheets.values().append(
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()
    
    return {
        'updated_cells': result.get('updates', {}).get('updatedCells'),
        'updated_range': result.get('updates', {}).get('updatedRange')
    }

@mcp.tool()
def update_cell(spreadsheet_id: str, sheet_name: str, cell: str, value: str) -> Dict:
    """Update a single cell in a Google Sheet
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_name: The name of the sheet
        cell: Cell reference (e.g., "A1")
        value: Value to set
        
    Returns:
        Status of the operation
    """
    sheet_range = f"{sheet_name}!{cell}"
    sheets = get_sheets_service()
    body = {
        'values': [[value]]
    }
    result = sheets.values().update(
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    
    return {
        'updated_cells': result.get('updatedCells'),
        'updated_range': result.get('updatedRange')
    }

# Formula Operations

@mcp.tool()
def add_formula(spreadsheet_id: str, sheet_name: str, cell: str, formula: str) -> Dict:
    """Add a formula to a cell
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_name: The name of the sheet
        cell: Cell reference (e.g., "A1")
        formula: Formula string (e.g., "=SUM(B1:B10)")
        
    Returns:
        Status of the operation
    """
    # Ensure formula starts with =
    if not formula.startswith('='):
        formula = f"={formula}"
        
    sheet_range = f"{sheet_name}!{cell}"
    sheets = get_sheets_service()
    body = {
        'values': [[formula]]
    }
    result = sheets.values().update(
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
        valueInputOption='USER_ENTERED',  # USER_ENTERED interprets formulas
        body=body
    ).execute()
    
    return {
        'updated_cells': result.get('updatedCells'),
        'updated_range': result.get('updatedRange')
    }

@mcp.tool()
def batch_add_formulas(
    spreadsheet_id: str, 
    sheet_name: str, 
    cell_formulas: Dict[str, str]
) -> Dict:
    """Add multiple formulas at once
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_name: The name of the sheet
        cell_formulas: Dictionary mapping cells to formulas (e.g., {"A1": "=SUM(B1:B10)"})
        
    Returns:
        Status of the operation
    """
    sheets = get_sheets_service()
    
    # Group cells by row for efficiency
    row_data = {}
    for cell, formula in cell_formulas.items():
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = f"={formula}"
            
        # Extract row and column from cell reference
        import re
        match = re.match(r'([A-Z]+)(\d+)', cell)
        if not match:
            continue
            
        col, row = match.groups()
        if row not in row_data:
            row_data[row] = {}
        row_data[row][col] = formula
    
    # Create batch update request
    batch_data = []
    for row, cols in row_data.items():
        # Find min and max columns
        col_indices = [ord(col) - ord('A') for col in cols.keys()]
        min_col = chr(min(col_indices) + ord('A'))
        max_col = chr(max(col_indices) + ord('A'))
        
        # Create range and values
        range_str = f"{sheet_name}!{min_col}{row}:{max_col}{row}"
        values = [''] * (max(col_indices) - min(col_indices) + 1)
        
        # Fill in values
        for col, formula in cols.items():
            col_idx = ord(col) - ord(min_col)
            values[col_idx] = formula
            
        batch_data.append({
            'range': range_str,
            'values': [values]
        })
    
    # Execute batch update
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': batch_data
    }
    result = sheets.values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()
    
    return {
        'total_updated_cells': result.get('totalUpdatedCells'),
        'total_updated_sheets': result.get('totalUpdatedSheets')
    }

# Formatting Operations

@mcp.tool()
def format_range(
    spreadsheet_id: str, 
    sheet_name: str, 
    cell_range: str, 
    formatting: Dict
) -> Dict:
    """Apply formatting to a range of cells
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_name: The name of the sheet
        cell_range: Cell range (e.g., "A1:B10")
        formatting: Dictionary of formatting options
        
    Returns:
        Status of the operation
    """
    sheets = get_sheets_service()
    
    # Convert range to GridRange format
    grid_range = {
        'sheetId': get_sheet_id(spreadsheet_id, sheet_name),
        'startRowIndex': None,
        'endRowIndex': None,
        'startColumnIndex': None,
        'endColumnIndex': None
    }
    
    # Parse cell range (e.g., "A1:B10")
    import re
    range_match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', cell_range)
    if range_match:
        start_col, start_row, end_col, end_row = range_match.groups()
        grid_range['startRowIndex'] = int(start_row) - 1
        grid_range['endRowIndex'] = int(end_row)
        grid_range['startColumnIndex'] = column_letter_to_index(start_col)
        grid_range['endColumnIndex'] = column_letter_to_index(end_col) + 1
    else:
        # Single cell (e.g., "A1")
        single_match = re.match(r'([A-Z]+)(\d+)', cell_range)
        if single_match:
            col, row = single_match.groups()
            grid_range['startRowIndex'] = int(row) - 1
            grid_range['endRowIndex'] = int(row)
            grid_range['startColumnIndex'] = column_letter_to_index(col)
            grid_range['endColumnIndex'] = column_letter_to_index(col) + 1
    
    # Build format request
    requests = []
    
    # Background color
    if 'backgroundColor' in formatting:
        requests.append({
            'repeatCell': {
                'range': grid_range,
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': parse_color(formatting['backgroundColor'])
                    }
                },
                'fields': 'userEnteredFormat.backgroundColor'
            }
        })
    
    # Text format (bold, italic, etc.)
    text_format = {}
    if 'bold' in formatting:
        text_format['bold'] = formatting['bold']
    if 'italic' in formatting:
        text_format['italic'] = formatting['italic']
    if 'fontFamily' in formatting:
        text_format['fontFamily'] = formatting['fontFamily']
    if 'fontSize' in formatting:
        text_format['fontSize'] = formatting['fontSize']
    if 'foregroundColor' in formatting:
        text_format['foregroundColor'] = parse_color(formatting['foregroundColor'])
    
    if text_format:
        requests.append({
            'repeatCell': {
                'range': grid_range,
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': text_format
                    }
                },
                'fields': f"userEnteredFormat.textFormat({','.join(text_format.keys())})"
            }
        })
    
    # Horizontal alignment
    if 'horizontalAlignment' in formatting:
        requests.append({
            'repeatCell': {
                'range': grid_range,
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': formatting['horizontalAlignment'].upper()
                    }
                },
                'fields': 'userEnteredFormat.horizontalAlignment'
            }
        })
    
    # Vertical alignment
    if 'verticalAlignment' in formatting:
        requests.append({
            'repeatCell': {
                'range': grid_range,
                'cell': {
                    'userEnteredFormat': {
                        'verticalAlignment': formatting['verticalAlignment'].upper()
                    }
                },
                'fields': 'userEnteredFormat.verticalAlignment'
            }
        })
    
    # Number format
    if 'numberFormat' in formatting:
        requests.append({
            'repeatCell': {
                'range': grid_range,
                'cell': {
                    'userEnteredFormat': {
                        'numberFormat': {
                            'type': formatting['numberFormat'].upper(),
                            'pattern': formatting.get('numberPattern', '')
                        }
                    }
                },
                'fields': 'userEnteredFormat.numberFormat'
            }
        })
    
    # Borders
    if 'borders' in formatting:
        borders = {}
        for position, style in formatting['borders'].items():
            borders[position] = {
                'style': style.get('style', 'SOLID'),
                'width': style.get('width', 1),
                'color': parse_color(style.get('color', 'black'))
            }
        
        requests.append({
            'updateBorders': {
                'range': grid_range,
                **{f"{pos}": border for pos, border in borders.items()}
            }
        })
    
    # Execute the formatting request
    body = {
        'requests': requests
    }
    result = sheets.batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()
    
    return {
        'replies': len(result.get('replies', [])),
        'status': 'success'
    }

@mcp.tool()
def add_conditional_formatting(
    spreadsheet_id: str,
    sheet_name: str,
    cell_range: str,
    condition_type: str,
    condition_values: List[str],
    format_settings: Dict
) -> Dict:
    """Add conditional formatting to a range of cells
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_name: The name of the sheet
        cell_range: Cell range (e.g., "A1:B10")
        condition_type: Type of condition (e.g., "NUMBER_GREATER", "TEXT_CONTAINS")
        condition_values: Values for the condition
        format_settings: Formatting to apply when condition is met
        
    Returns:
        Status of the operation
    """
    sheets = get_sheets_service()
    
    # Convert range to GridRange format
    grid_range = {
        'sheetId': get_sheet_id(spreadsheet_id, sheet_name),
        'startRowIndex': None,
        'endRowIndex': None,
        'startColumnIndex': None,
        'endColumnIndex': None
    }
    
    # Parse cell range
    import re
    range_match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', cell_range)
    if range_match:
        start_col, start_row, end_col, end_row = range_match.groups()
        grid_range['startRowIndex'] = int(start_row) - 1
        grid_range['endRowIndex'] = int(end_row)
        grid_range['startColumnIndex'] = column_letter_to_index(start_col)
        grid_range['endColumnIndex'] = column_letter_to_index(end_col) + 1
    
    # Build conditional format
    condition = {
        'type': condition_type,
        'values': [{'userEnteredValue': val} for val in condition_values]
    }
    
    # Build format
    format_rule = {}
    
    # Background color
    if 'backgroundColor' in format_settings:
        if 'backgroundColor' not in format_rule:
            format_rule['backgroundColor'] = {}
        format_rule['backgroundColor'] = parse_color(format_settings['backgroundColor'])
    
    # Text format
    if any(k in format_settings for k in ['bold', 'italic', 'fontFamily', 'fontSize', 'foregroundColor']):
        format_rule['textFormat'] = {}
        
        if 'bold' in format_settings:
            format_rule['textFormat']['bold'] = format_settings['bold']
        if 'italic' in format_settings:
            format_rule['textFormat']['italic'] = format_settings['italic']
        if 'fontFamily' in format_settings:
            format_rule['textFormat']['fontFamily'] = format_settings['fontFamily']
        if 'fontSize' in format_settings:
            format_rule['textFormat']['fontSize'] = format_settings['fontSize']
        if 'foregroundColor' in format_settings:
            format_rule['textFormat']['foregroundColor'] = parse_color(format_settings['foregroundColor'])
    
    # Create request
    request = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [grid_range],
                'booleanRule': {
                    'condition': condition,
                    'format': format_rule
                }
            },
            'index': 0  # Insert at the beginning
        }
    }
    
    # Execute the request
    body = {
        'requests': [request]
    }
    result = sheets.batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()
    
    return {
        'replies': len(result.get('replies', [])),
        'status': 'success'
    }

# Helper functions

def get_sheet_id(spreadsheet_id: str, sheet_name: str) -> int:
    """Get the sheet ID from the sheet name"""
    sheets = get_sheets_service()
    spreadsheet = sheets.get(spreadsheetId=spreadsheet_id).execute()
    
    for sheet in spreadsheet.get('sheets', []):
        if sheet.get('properties', {}).get('title') == sheet_name:
            return sheet.get('properties', {}).get('sheetId')
    
    raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")

def column_letter_to_index(column: str) -> int:
    """Convert column letter (A, B, AA, etc.) to index (0, 1, 26, etc.)"""
    result = 0
    for char in column:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def parse_color(color_str: str) -> Dict:
    """Parse color string to Google Sheets color format"""
    # Handle named colors
    named_colors = {
        'black': {'red': 0, 'green': 0, 'blue': 0},
        'white': {'red': 1, 'green': 1, 'blue': 1},
        'red': {'red': 1, 'green': 0, 'blue': 0},
        'green': {'red': 0, 'green': 1, 'blue': 0},
        'blue': {'red': 0, 'green': 0, 'blue': 1},
        'yellow': {'red': 1, 'green': 1, 'blue': 0},
        'purple': {'red': 0.5, 'green': 0, 'blue': 0.5},
        'orange': {'red': 1, 'green': 0.65, 'blue': 0},
        'gray': {'red': 0.5, 'green': 0.5, 'blue': 0.5}
    }
    
    if color_str.lower() in named_colors:
        return named_colors[color_str.lower()]
    
    # Handle hex colors
    if color_str.startswith('#'):
        hex_color = color_str[1:]
        if len(hex_color) == 3:
            # Expand shorthand hex
            hex_color = ''.join([c*2 for c in hex_color])
        
        r = int(hex_color[0:2], 16) / 255.0
        g = int(hex_color[2:4], 16) / 255.0
        b = int(hex_color[4:6], 16) / 255.0
        
        return {'red': r, 'green': g, 'blue': b}
    
    # Handle rgb colors
    import re
    rgb_match = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color_str)
    if rgb_match:
        r, g, b = map(int, rgb_match.groups())
        return {'red': r/255.0, 'green': g/255.0, 'blue': b/255.0}
    
    # Default to black
    return {'red': 0, 'green': 0, 'blue': 0}

# Sheet Listing Operations

@mcp.tool()
def list_sheets(filter_term: str = None) -> List[Dict]:
    """List Google Sheets the user has access to
    
    Args:
        filter_term: Optional term to filter sheet names (case-insensitive)
        
    Returns:
        List of sheets with their names and IDs
    """
    drive = get_drive_service()
    query = "mimeType='application/vnd.google-apps.spreadsheet'"
    
    # Add name filter if provided
    if filter_term:
        query += f" and name contains '{filter_term}'"
        
    results = drive.files().list(
        q=query,
        fields="files(id, name, createdTime, modifiedTime, webViewLink)",
        orderBy="modifiedTime desc"
    ).execute()
    
    return results.get('files', [])

# Resource for accessing sheet data directly
@mcp.resource("sheets://{spreadsheet_id}/{sheet_name}")
def get_sheet_data(spreadsheet_id: str, sheet_name: str) -> str:
    """Get data from a Google Sheet
    
    Args:
        spreadsheet_id: The ID of the spreadsheet
        sheet_name: The name of the sheet
        
    Returns:
        JSON string with sheet data
    """
    sheets = get_sheets_service()
    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet_name
    ).execute()
    
    return json.dumps(result.get('values', []), indent=2)

def main():
    """Main entry point for the MCP server"""
    print("Starting Google Sheets MCP server...")
    print(f"Looking for credentials in: {CREDENTIALS_FILE}")
    print("Available tools:")
    # List all the functions decorated with @mcp.tool()
    tool_functions = [
        "list_sheets",  # Added new tool for listing sheets
        "read_sheet", "write_sheet", "append_sheet", "update_cell",
        "add_formula", "batch_add_formulas", "format_range", "add_conditional_formatting"
    ]
    for tool_name in tool_functions:
        print(f"  - {tool_name}")
    print("Available resources:")
    # List all the functions decorated with @mcp.resource()
    resource_functions = ["sheets://{spreadsheet_id}/{sheet_name}"]
    for resource in resource_functions:
        print(f"  - {resource}")
    print("Server running. Press Ctrl+C to exit.")
    mcp.run()

if __name__ == "__main__":
    main()
