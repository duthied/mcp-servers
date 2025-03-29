#!/usr/bin/env python3
"""
Google Sheets Core Module for MCP Server

This module provides core functionality for interacting with Google Sheets,
including spreadsheet and worksheet management.
"""

import logging
from typing import Dict, Any, List, Optional, Union

from ..auth import GoogleAuthManager

# Configure logging
logger = logging.getLogger("google_sheets_mcp.sheets.core")


class SheetsCore:
    """
    Core functionality for interacting with Google Sheets.
    
    This class provides base operations for spreadsheet and worksheet management,
    as well as utility methods for working with the Sheets API.
    """
    
    def __init__(self, auth_manager: GoogleAuthManager):
        """
        Initialize the Sheets Core module.
        
        Args:
            auth_manager: An instance of GoogleAuthManager for authentication.
        """
        self.auth_manager = auth_manager
    
    def _get_sheets_service(self):
        """
        Get the Google Sheets API service.
        
        Returns:
            The Google Sheets API service object.
        """
        self.auth_manager.refresh_if_needed()
        return self.auth_manager.get_sheets_service()
    
    def _get_drive_service(self):
        """
        Get the Google Drive API service.
        
        Returns:
            The Google Drive API service object.
        """
        self.auth_manager.refresh_if_needed()
        return self.auth_manager.get_drive_service()
    
    async def create_sheet(
        self, 
        title: str, 
        sheet_type: str, 
        parent_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Create a new spreadsheet or worksheet.
        
        Args:
            title: The title of the spreadsheet or worksheet.
            sheet_type: The type of sheet to create ('spreadsheet' or 'worksheet').
            parent_id: For worksheet creation, the ID of the parent spreadsheet.
                       Required if sheet_type is 'worksheet'.
        
        Returns:
            A dictionary containing information about the created sheet.
        """
        try:
            if sheet_type == 'spreadsheet':
                # Create a new spreadsheet
                sheets_service = self._get_sheets_service()
                spreadsheet_body = {
                    'properties': {
                        'title': title
                    }
                }
                request = sheets_service.spreadsheets().create(body=spreadsheet_body)
                response = request.execute()
                
                logger.info("Created new spreadsheet with ID: %s", response['spreadsheetId'])
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
                
            elif sheet_type == 'worksheet':
                # Create a new worksheet in an existing spreadsheet
                if not parent_id:
                    raise ValueError("parent_id is required for worksheet creation")
                
                sheets_service = self._get_sheets_service()
                request_body = {
                    'requests': [
                        {
                            'addSheet': {
                                'properties': {
                                    'title': title
                                }
                            }
                        }
                    ]
                }
                
                request = sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=parent_id,
                    body=request_body
                )
                response = request.execute()
                
                # Extract the new sheet's properties
                new_sheet = response['replies'][0]['addSheet']['properties']
                
                logger.info(
                    "Created new worksheet '%s' in spreadsheet %s", 
                    title, parent_id
                )
                
                return {
                    'spreadsheet_id': parent_id,
                    'sheet_id': new_sheet['sheetId'],
                    'title': new_sheet['title'],
                    'index': new_sheet.get('index', 0)
                }
            else:
                raise ValueError(f"Invalid sheet_type: {sheet_type}. Must be 'spreadsheet' or 'worksheet'")
                
        except Exception as e:
            logger.error("Error creating sheet: %s", str(e))
            raise
    
    async def get_sheet_info(
        self, 
        spreadsheet_id: str, 
        include_sheets: bool = True
    ) -> Dict[str, Any]:
        """
        Get metadata about a spreadsheet.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            include_sheets: Whether to include information about individual worksheets.
        
        Returns:
            A dictionary containing metadata about the spreadsheet.
        """
        try:
            sheets_service = self._get_sheets_service()
            
            # Determine which fields to include
            fields = "properties.title,properties.locale,properties.timeZone,spreadsheetUrl"
            if include_sheets:
                fields += ",sheets.properties"
            
            request = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields=fields
            )
            response = request.execute()
            
            # Format the response
            result = {
                'spreadsheet_id': spreadsheet_id,
                'title': response['properties']['title'],
                'locale': response['properties'].get('locale', 'en_US'),
                'time_zone': response['properties'].get('timeZone', 'GMT'),
                'url': response.get('spreadsheetUrl', '')
            }
            
            if include_sheets and 'sheets' in response:
                result['sheets'] = [
                    {
                        'sheet_id': sheet['properties']['sheetId'],
                        'title': sheet['properties']['title'],
                        'index': sheet['properties'].get('index', 0),
                        'sheet_type': sheet['properties'].get('sheetType', 'GRID'),
                        'row_count': sheet['properties'].get('gridProperties', {}).get('rowCount', 0),
                        'column_count': sheet['properties'].get('gridProperties', {}).get('columnCount', 0)
                    }
                    for sheet in response['sheets']
                ]
            
            return result
            
        except Exception as e:
            logger.error("Error getting sheet info: %s", str(e))
            raise
    
    async def manage_permissions(
        self, 
        spreadsheet_id: str, 
        permissions: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Control access and sharing for a spreadsheet.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            permissions: A dictionary containing permission settings.
                Expected format:
                {
                    "role": "reader|writer|commenter|owner",
                    "type": "user|group|domain|anyone",
                    "emailAddress": "user@example.com",  # Only for user or group
                    "domain": "example.com",  # Only for domain
                    "allowFileDiscovery": true|false,  # Only for anyone or domain
                    "transferOwnership": true|false,  # Only for owner role
                    "sendNotificationEmail": true|false,
                    "emailMessage": "Optional message"
                }
        
        Returns:
            A dictionary containing information about the updated permissions.
        """
        try:
            drive_service = self._get_drive_service()
            
            # Prepare permission body
            permission_body = {
                'role': permissions.get('role', 'reader'),
                'type': permissions.get('type', 'user')
            }
            
            # Add type-specific fields
            if permission_body['type'] == 'user' or permission_body['type'] == 'group':
                if 'emailAddress' not in permissions:
                    raise ValueError("emailAddress is required for user or group permission type")
                permission_body['emailAddress'] = permissions['emailAddress']
            elif permission_body['type'] == 'domain':
                if 'domain' not in permissions:
                    raise ValueError("domain is required for domain permission type")
                permission_body['domain'] = permissions['domain']
            
            # Add optional fields
            if 'allowFileDiscovery' in permissions and permission_body['type'] in ['domain', 'anyone']:
                permission_body['allowFileDiscovery'] = permissions['allowFileDiscovery']
            
            # Prepare request parameters
            request_params = {
                'fileId': spreadsheet_id,
                'body': permission_body,
                'sendNotificationEmail': permissions.get('sendNotificationEmail', False)
            }
            
            if 'emailMessage' in permissions and request_params['sendNotificationEmail']:
                request_params['emailMessage'] = permissions['emailMessage']
            
            if permissions.get('transferOwnership', False) and permission_body['role'] == 'owner':
                request_params['transferOwnership'] = True
            
            # Execute the request
            response = drive_service.permissions().create(**request_params).execute()
            
            logger.info(
                "Updated permissions for spreadsheet %s: %s %s", 
                spreadsheet_id, response['role'], response['type']
            )
            
            return {
                'permission_id': response['id'],
                'role': response['role'],
                'type': response['type'],
                'email_address': response.get('emailAddress'),
                'domain': response.get('domain'),
                'allow_file_discovery': response.get('allowFileDiscovery')
            }
            
        except Exception as e:
            logger.error("Error managing permissions: %s", str(e))
            raise
    
    async def get_sheet_id_by_name(
        self, 
        spreadsheet_id: str, 
        sheet_name: str
    ) -> Optional[int]:
        """
        Get the sheet ID for a worksheet by name.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_name: The name of the worksheet.
        
        Returns:
            The sheet ID if found, None otherwise.
        """
        try:
            sheet_info = await self.get_sheet_info(spreadsheet_id, include_sheets=True)
            
            for sheet in sheet_info.get('sheets', []):
                if sheet['title'] == sheet_name:
                    return sheet['sheet_id']
            
            return None
            
        except Exception as e:
            logger.error("Error getting sheet ID by name: %s", str(e))
            raise
    
    async def parse_range_notation(
        self, 
        spreadsheet_id: str, 
        range_notation: str
    ) -> Dict[str, Any]:
        """
        Parse an A1 notation range into its components.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_notation: The A1 notation range (e.g., 'Sheet1!A1:B10').
        
        Returns:
            A dictionary containing the parsed range components.
        """
        try:
            # Check if range includes sheet name
            if '!' in range_notation:
                sheet_name, cell_range = range_notation.split('!', 1)
                
                # Remove quotes if present
                if (sheet_name.startswith("'") and sheet_name.endswith("'")) or \
                   (sheet_name.startswith('"') and sheet_name.endswith('"')):
                    sheet_name = sheet_name[1:-1]
            else:
                # Use default sheet if not specified
                sheet_info = await self.get_sheet_info(spreadsheet_id, include_sheets=True)
                sheet_name = sheet_info['sheets'][0]['title']
                cell_range = range_notation
            
            # Get sheet ID
            sheet_id = await self.get_sheet_id_by_name(spreadsheet_id, sheet_name)
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'sheet_name': sheet_name,
                'sheet_id': sheet_id,
                'cell_range': cell_range,
                'full_range': range_notation
            }
            
        except Exception as e:
            logger.error("Error parsing range notation: %s", str(e))
            raise
