#!/usr/bin/env python3
"""
Google Sheets Formatting Module for MCP Server

This module provides functionality for formatting cells in Google Sheets,
including styling, conditional formatting, merging cells, and number formats.
"""

import logging
from typing import Dict, Any, List, Optional

from .core import SheetsCore

# Configure logging
logger = logging.getLogger("google_sheets_mcp.sheets.formatting")


class FormattingOperations:
    """
    Formatting operations for Google Sheets.
    
    This class provides methods for applying formatting to cells in Google Sheets,
    including styling, conditional formatting, merging cells, and number formats.
    """
    
    def __init__(self, sheets_core: SheetsCore):
        """
        Initialize the Formatting Operations module.
        
        Args:
            sheets_core: An instance of SheetsCore for base sheet operations.
        """
        self.sheets_core = sheets_core
    
    async def format_cells(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        format_options: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Apply formatting to cells.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to format.
            format_options: Formatting options to apply.
                Example: {
                    "backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0},
                    "textFormat": {"bold": true, "fontSize": 12}
                }
        
        Returns:
            A dictionary containing information about the formatting operation.
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
                        'repeatCell': {
                            'range': {
                                'sheetId': range_info['sheet_id'],
                                'startRowIndex': None,  # Will be determined from A1 notation
                                'endRowIndex': None,    # Will be determined from A1 notation
                                'startColumnIndex': None,  # Will be determined from A1 notation
                                'endColumnIndex': None     # Will be determined from A1 notation
                            },
                            'cell': {
                                'userEnteredFormat': format_options
                            },
                            'fields': 'userEnteredFormat(' + ','.join(format_options.keys()) + ')'
                        }
                    }
                ]
            }
            
            # Convert A1 notation to grid range
            grid_range = self._a1_to_grid_range(range_info)
            request_body['requests'][0]['repeatCell']['range'].update(grid_range)
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Applied formatting to range %s in spreadsheet %s", 
                range_str, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'format_options': format_options,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error formatting cells: %s", str(e))
            raise
    
    async def add_conditional_format(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        rule_type: str, 
        rule_values: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Add conditional formatting to a range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to format.
            rule_type: Type of conditional formatting rule.
                Options: "BOOLEAN", "GRADIENT", "TEXT", "DATE", "CUSTOM_FORMULA"
            rule_values: Values for the conditional formatting rule.
                Example for BOOLEAN: {
                    "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "5"}]},
                    "format": {"backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}
                }
        
        Returns:
            A dictionary containing information about the conditional formatting operation.
        """
        try:
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range_str)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body based on rule type
            request_body = {
                'requests': [
                    {
                        'addConditionalFormatRule': {
                            'rule': {
                                'ranges': [
                                    {
                                        'sheetId': range_info['sheet_id'],
                                        'startRowIndex': None,  # Will be determined from A1 notation
                                        'endRowIndex': None,    # Will be determined from A1 notation
                                        'startColumnIndex': None,  # Will be determined from A1 notation
                                        'endColumnIndex': None     # Will be determined from A1 notation
                                    }
                                ],
                                'booleanRule' if rule_type == 'BOOLEAN' else 'gradientRule' if rule_type == 'GRADIENT' else 'customRule': rule_values
                            },
                            'index': 0  # Insert at the beginning
                        }
                    }
                ]
            }
            
            # Convert A1 notation to grid range
            grid_range = self._a1_to_grid_range(range_info)
            request_body['requests'][0]['addConditionalFormatRule']['rule']['ranges'][0].update(grid_range)
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Added conditional formatting to range %s in spreadsheet %s", 
                range_str, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'rule_type': rule_type,
                'rule_values': rule_values,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error adding conditional formatting: %s", str(e))
            raise
    
    async def merge_cells(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        merge_type: str = "MERGE_ALL"
    ) -> Dict[str, Any]:
        """
        Merge cells in a range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to merge.
            merge_type: Type of merge to perform.
                Options: "MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"
        
        Returns:
            A dictionary containing information about the merge operation.
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
                        'mergeCells': {
                            'range': {
                                'sheetId': range_info['sheet_id'],
                                'startRowIndex': None,  # Will be determined from A1 notation
                                'endRowIndex': None,    # Will be determined from A1 notation
                                'startColumnIndex': None,  # Will be determined from A1 notation
                                'endColumnIndex': None     # Will be determined from A1 notation
                            },
                            'mergeType': merge_type
                        }
                    }
                ]
            }
            
            # Convert A1 notation to grid range
            grid_range = self._a1_to_grid_range(range_info)
            request_body['requests'][0]['mergeCells']['range'].update(grid_range)
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Merged cells in range %s in spreadsheet %s with merge type %s", 
                range_str, spreadsheet_id, merge_type
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'merge_type': merge_type,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error merging cells: %s", str(e))
            raise
    
    async def set_number_format(
        self, 
        spreadsheet_id: str, 
        range_str: str, 
        number_format: str
    ) -> Dict[str, Any]:
        """
        Set number format for cells.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range_str: A1 notation range to format.
            number_format: Number format pattern to apply.
                Examples: "0.00", "$0.00", "0%", "MM/DD/YYYY"
        
        Returns:
            A dictionary containing information about the number format operation.
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
                        'repeatCell': {
                            'range': {
                                'sheetId': range_info['sheet_id'],
                                'startRowIndex': None,  # Will be determined from A1 notation
                                'endRowIndex': None,    # Will be determined from A1 notation
                                'startColumnIndex': None,  # Will be determined from A1 notation
                                'endColumnIndex': None     # Will be determined from A1 notation
                            },
                            'cell': {
                                'userEnteredFormat': {
                                    'numberFormat': {
                                        'type': self._determine_number_format_type(number_format),
                                        'pattern': number_format
                                    }
                                }
                            },
                            'fields': 'userEnteredFormat.numberFormat'
                        }
                    }
                ]
            }
            
            # Convert A1 notation to grid range
            grid_range = self._a1_to_grid_range(range_info)
            request_body['requests'][0]['repeatCell']['range'].update(grid_range)
            
            # Execute the request
            request = sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            logger.info(
                "Set number format '%s' for range %s in spreadsheet %s", 
                number_format, range_str, spreadsheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'range': range_str,
                'number_format': number_format,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error setting number format: %s", str(e))
            raise
    
    def _determine_number_format_type(self, pattern: str) -> str:
        """
        Determine the number format type based on the pattern.
        
        Args:
            pattern: Number format pattern.
        
        Returns:
            The number format type.
        """
        if '%' in pattern:
            return 'PERCENT'
        elif '$' in pattern or '€' in pattern or '£' in pattern:
            return 'CURRENCY'
        elif any(date_part in pattern.upper() for date_part in ['Y', 'M', 'D']):
            return 'DATE'
        elif any(time_part in pattern.upper() for time_part in ['H', 'MM', 'SS']):
            return 'TIME'
        elif '.' in pattern or '#' in pattern or '0' in pattern:
            return 'NUMBER'
        else:
            return 'TEXT'
    
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
