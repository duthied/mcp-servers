#!/usr/bin/env python3
"""
Google Sheets Formulas Module for MCP Server

This module provides functionality for working with formulas in Google Sheets.
"""

import logging
from typing import Dict, Any, List, Optional

from .core import SheetsCore

# Configure logging
logger = logging.getLogger("google_sheets_mcp.sheets.formulas")


class FormulaOperations:
    """
    Formula operations for Google Sheets.
    
    This class provides methods for working with formulas in Google Sheets,
    including applying formulas to cells, applying array formulas, and creating named ranges.
    """
    
    def __init__(self, sheets_core: SheetsCore):
        """
        Initialize the Formula Operations module.
        
        Args:
            sheets_core: An instance of SheetsCore for base sheet operations.
        """
        self.sheets_core = sheets_core
    
    async def apply_formula(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        formula: str
    ) -> Dict[str, Any]:
        """
        Apply a formula to a cell or range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to apply formula to.
            formula: Formula to apply (without the leading '=').
        
        Returns:
            A dictionary containing information about the formula application.
        """
        try:
            # Ensure formula starts with '='
            if not formula.startswith('='):
                formula = '=' + formula
            
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range_str)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'values': [[formula]]  # Single cell formula
            }
            
            # Execute the request
            request = sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_str,
                valueInputOption="USER_ENTERED",  # Interpret as formula
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Applied formula '%s' to range %s in spreadsheet %s", 
                formula, range_str, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'formula': formula,
                'updated_cells': response.get('updatedCells', 0),
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error applying formula: %s", str(e))
            raise
    
    async def apply_array_formula(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        formula: str
    ) -> Dict[str, Any]:
        """
        Apply an array formula to a range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to apply array formula to.
            formula: Array formula to apply (without the leading '=').
        
        Returns:
            A dictionary containing information about the array formula application.
        """
        try:
            # Ensure formula starts with '='
            if not formula.startswith('='):
                formula = '=' + formula
            
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range_str)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'requests': [
                    {
                        'updateCells': {
                            'range': {
                                'sheetId': range_info['sheet_id'],
                                'startRowIndex': None,  # Will be determined from A1 notation
                                'endRowIndex': None,    # Will be determined from A1 notation
                                'startColumnIndex': None,  # Will be determined from A1 notation
                                'endColumnIndex': None     # Will be determined from A1 notation
                            },
                            'rows': [
                                {
                                    'values': [
                                        {
                                            'userEnteredValue': {
                                                'formulaValue': formula
                                            }
                                        }
                                    ]
                                }
                            ],
                            'fields': 'userEnteredValue.formulaValue'
                        }
                    }
                ]
            }
            
            # Convert A1 notation to grid range
            grid_range = self._a1_to_grid_range(range_info)
            request_body['requests'][0]['updateCells']['range'].update(grid_range)
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Applied array formula '%s' to range %s in spreadsheet %s", 
                formula, range_str, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'formula': formula,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error applying array formula: %s", str(e))
            raise
    
    async def create_named_range(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        name: str, 
        scope: str = "WORKBOOK"
    ) -> Dict[str, Any]:
        """
        Create a named range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to name.
            name: Name for the range.
            scope: Scope of the named range.
                Options: "WORKBOOK", "SHEET"
        
        Returns:
            A dictionary containing information about the named range creation.
        """
        try:
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range_str)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'requests': [
                    {
                        'addNamedRange': {
                            'namedRange': {
                                'name': name,
                                'range': {
                                    'sheetId': range_info['sheet_id'],
                                    'startRowIndex': None,  # Will be determined from A1 notation
                                    'endRowIndex': None,    # Will be determined from A1 notation
                                    'startColumnIndex': None,  # Will be determined from A1 notation
                                    'endColumnIndex': None     # Will be determined from A1 notation
                                }
                            }
                        }
                    }
                ]
            }
            
            # Convert A1 notation to grid range
            grid_range = self._a1_to_grid_range(range_info)
            request_body['requests'][0]['addNamedRange']['namedRange']['range'].update(grid_range)
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            # Extract the named range ID
            named_range_id = response['replies'][0]['addNamedRange']['namedRange']['namedRangeId']
            
            logger.info(
                "Created named range '%s' for range %s in spreadsheet %s", 
                name, range_str, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'name': name,
                'named_range_id': named_range_id,
                'scope': scope,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error creating named range: %s", str(e))
            raise
    
    async def get_named_ranges(
        self, 
        spreadsheet_id: str
    ) -> Dict[str, Any]:
        """
        Get all named ranges in a spreadsheet.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
        
        Returns:
            A dictionary containing information about the named ranges.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Execute the request
            request = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields='namedRanges,sheets.properties'
            )
            response = request.execute()
            
            # Extract named ranges
            named_ranges = response.get('namedRanges', [])
            sheets = {sheet['properties']['sheetId']: sheet['properties']['title'] 
                     for sheet in response.get('sheets', [])}
            
            # Format the response
            result = {
                'spreadsheet_id': spreadsheet_id,
                'named_ranges': [
                    {
                        'name': nr['name'],
                        'named_range_id': nr['namedRangeId'],
                        'range': {
                            'sheet_id': nr['range']['sheetId'],
                            'sheet_name': sheets.get(nr['range']['sheetId'], ''),
                            'start_row_index': nr['range'].get('startRowIndex', 0),
                            'end_row_index': nr['range'].get('endRowIndex', 0),
                            'start_column_index': nr['range'].get('startColumnIndex', 0),
                            'end_column_index': nr['range'].get('endColumnIndex', 0)
                        }
                    }
                    for nr in named_ranges
                ]
            }
            
            logger.info(
                "Retrieved %d named ranges from spreadsheet %s", 
                len(result['named_ranges']), spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error getting named ranges: %s", str(e))
            raise
    
    async def delete_named_range(
        self, 
        spreadsheet_id: str, 
        named_range_id: str
    ) -> Dict[str, Any]:
        """
        Delete a named range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            named_range_id: The ID of the named range to delete.
        
        Returns:
            A dictionary containing information about the deletion.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'requests': [
                    {
                        'deleteNamedRange': {
                            'namedRangeId': named_range_id
                        }
                    }
                ]
            }
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Deleted named range with ID %s from spreadsheet %s", 
                named_range_id, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'named_range_id': named_range_id,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error deleting named range: %s", str(e))
            raise
    
    def _a1_to_grid_range(self, range_info: Dict[str, Any]) -> Dict[str, int]:
        """
        Convert A1 notation cell range to grid range indices.
        
        This is a simplified implementation that assumes the cell_range is in the format
        'A1:B2' or similar. For a complete implementation, you would need to handle
        more complex cases.
        
        Args:
            range_info: Range information from parse_range_notation.
        
        Returns:
            A dictionary with grid range indices.
        """
        cell_range = range_info['cell_range']
        
        # Handle single cell case
        if ':' not in cell_range:
            col, row = self._a1_to_row_col(cell_range)
            return {
                'startRowIndex': row,
                'endRowIndex': row + 1,
                'startColumnIndex': col,
                'endColumnIndex': col + 1
            }
        
        # Handle range case
        start, end = cell_range.split(':')
        start_col, start_row = self._a1_to_row_col(start)
        end_col, end_row = self._a1_to_row_col(end)
        
        # In grid range, end indices are exclusive
        return {
            'startRowIndex': start_row,
            'endRowIndex': end_row + 1,
            'startColumnIndex': start_col,
            'endColumnIndex': end_col + 1
        }
    
    def _a1_to_row_col(self, a1_notation: str) -> tuple:
        """
        Convert A1 notation cell reference to row and column indices.
        
        Args:
            a1_notation: A1 notation cell reference (e.g., 'A1').
        
        Returns:
            A tuple of (column_index, row_index).
        """
        # Extract column letters and row number
        col_str = ''.join(c for c in a1_notation if c.isalpha())
        row_str = ''.join(c for c in a1_notation if c.isdigit())
        
        # Convert column letters to index (0-based)
        col_index = 0
        for c in col_str:
            col_index = col_index * 26 + (ord(c.upper()) - ord('A') + 1)
        col_index -= 1  # Adjust to 0-based index
        
        # Convert row number to index (0-based)
        row_index = int(row_str) - 1
        
        return (col_index, row_index)
