#!/usr/bin/env python3
"""
Google Sheets Data Operations Module for MCP Server

This module provides functionality for reading and writing data to Google Sheets.
"""

import logging
from typing import Dict, Any, List, Optional, Union

from .core import SheetsCore

# Configure logging
logger = logging.getLogger("google_sheets_mcp.sheets.data_ops")


class DataOperations:
    """
    Data operations for Google Sheets.
    
    This class provides methods for reading and writing data to Google Sheets,
    including reading ranges, writing ranges, appending rows, and clearing ranges.
    """
    
    def __init__(self, sheets_core: SheetsCore):
        """
        Initialize the Data Operations module.
        
        Args:
            sheets_core: An instance of SheetsCore for base sheet operations.
        """
        self.sheets_core = sheets_core
    
    async def read_range(
        self, 
        spreadsheet_id: str, 
        range: str, 
        value_render_option: str = "FORMATTED_VALUE"
    ) -> Dict[str, Any]:
        """
        Read data from a specified range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range: A1 notation range to read.
            value_render_option: How values should be rendered.
                Options: "FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"
        
        Returns:
            A dictionary containing the data from the range.
        """
        try:
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Execute the request
            request = sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueRenderOption=value_render_option
            )
            response = request.execute()
            
            # Format the response
            values = response.get('values', [])
            
            result = {
                'spreadsheet_id': spreadsheet_id,
                'range': response.get('range', range),
                'values': values,
                'num_rows': len(values),
                'num_columns': max([len(row) for row in values]) if values else 0
            }
            
            logger.info(
                "Read %d rows from range %s in spreadsheet %s", 
                result['num_rows'], range, spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error reading range: %s", str(e))
            raise
    
    async def write_range(
        self, 
        spreadsheet_id: str, 
        range: str, 
        values: List[List[Any]], 
        value_input_option: str = "USER_ENTERED"
    ) -> Dict[str, Any]:
        """
        Write data to a specified range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range: A1 notation range to write.
            values: 2D array of values to write.
            value_input_option: How input values should be interpreted.
                Options: "RAW", "USER_ENTERED"
        
        Returns:
            A dictionary containing information about the update.
        """
        try:
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'values': values
            }
            
            # Execute the request
            request = sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueInputOption=value_input_option,
                body=request_body
            )
            response = request.execute()
            
            # Format the response
            result = {
                'spreadsheet_id': spreadsheet_id,
                'range': response.get('updatedRange', range),
                'updated_rows': response.get('updatedRows', 0),
                'updated_columns': response.get('updatedColumns', 0),
                'updated_cells': response.get('updatedCells', 0)
            }
            
            logger.info(
                "Updated %d cells in range %s in spreadsheet %s", 
                result['updated_cells'], range, spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error writing to range: %s", str(e))
            raise
    
    async def append_rows(
        self, 
        spreadsheet_id: str, 
        range: str, 
        values: List[List[Any]], 
        insert_data_option: str = "INSERT_ROWS"
    ) -> Dict[str, Any]:
        """
        Append rows to a sheet.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range: A1 notation range to append to.
            values: 2D array of values to append.
            insert_data_option: How the input data should be inserted.
                Options: "OVERWRITE", "INSERT_ROWS"
        
        Returns:
            A dictionary containing information about the append operation.
        """
        try:
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'values': values
            }
            
            # Execute the request
            request = sheets_service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueInputOption="USER_ENTERED",
                insertDataOption=insert_data_option,
                body=request_body
            )
            response = request.execute()
            
            # Format the response
            result = {
                'spreadsheet_id': spreadsheet_id,
                'range': response.get('tableRange', range),
                'updates': {
                    'updated_range': response.get('updates', {}).get('updatedRange', ''),
                    'updated_rows': response.get('updates', {}).get('updatedRows', 0),
                    'updated_columns': response.get('updates', {}).get('updatedColumns', 0),
                    'updated_cells': response.get('updates', {}).get('updatedCells', 0)
                }
            }
            
            logger.info(
                "Appended %d rows to range %s in spreadsheet %s", 
                len(values), range, spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error appending rows: %s", str(e))
            raise
    
    async def clear_range(
        self, 
        spreadsheet_id: str, 
        range: str
    ) -> Dict[str, Any]:
        """
        Clear data from a range.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            range: A1 notation range to clear.
        
        Returns:
            A dictionary containing information about the clear operation.
        """
        try:
            # Parse the range notation
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, range)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Execute the request
            request = sheets_service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=range,
                body={}
            )
            response = request.execute()
            
            # Format the response
            result = {
                'spreadsheet_id': spreadsheet_id,
                'cleared_range': response.get('clearedRange', range)
            }
            
            logger.info(
                "Cleared range %s in spreadsheet %s", 
                range, spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error clearing range: %s", str(e))
            raise
    
    async def batch_get_values(
        self, 
        spreadsheet_id: str, 
        ranges: List[str], 
        value_render_option: str = "FORMATTED_VALUE"
    ) -> Dict[str, Any]:
        """
        Get values from multiple ranges in a single request.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            ranges: List of A1 notation ranges to read.
            value_render_option: How values should be rendered.
                Options: "FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"
        
        Returns:
            A dictionary containing the data from the ranges.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Execute the request
            request = sheets_service.spreadsheets().values().batchGet(
                spreadsheetId=spreadsheet_id,
                ranges=ranges,
                valueRenderOption=value_render_option
            )
            response = request.execute()
            
            # Format the response
            value_ranges = response.get('valueRanges', [])
            
            result = {
                'spreadsheet_id': spreadsheet_id,
                'value_ranges': [
                    {
                        'range': vr.get('range', ''),
                        'values': vr.get('values', []),
                        'num_rows': len(vr.get('values', [])),
                        'num_columns': max([len(row) for row in vr.get('values', [])]) if vr.get('values', []) else 0
                    }
                    for vr in value_ranges
                ]
            }
            
            logger.info(
                "Batch read %d ranges from spreadsheet %s", 
                len(ranges), spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error batch getting values: %s", str(e))
            raise
    
    async def batch_update_values(
        self, 
        spreadsheet_id: str, 
        data: List[Dict[str, Any]], 
        value_input_option: str = "USER_ENTERED"
    ) -> Dict[str, Any]:
        """
        Update values in multiple ranges in a single request.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            data: List of dictionaries with 'range' and 'values' keys.
                Example: [{'range': 'Sheet1!A1:B2', 'values': [[1, 2], [3, 4]]}]
            value_input_option: How input values should be interpreted.
                Options: "RAW", "USER_ENTERED"
        
        Returns:
            A dictionary containing information about the updates.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'valueInputOption': value_input_option,
                'data': [
                    {
                        'range': item['range'],
                        'values': item['values']
                    }
                    for item in data
                ]
            }
            
            # Execute the request
            request = sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=request_body
            )
            response = request.execute()
            
            # Format the response
            responses = response.get('responses', [])
            
            result = {
                'spreadsheet_id': spreadsheet_id,
                'total_updated_cells': response.get('totalUpdatedCells', 0),
                'total_updated_rows': response.get('totalUpdatedRows', 0),
                'total_updated_columns': response.get('totalUpdatedColumns', 0),
                'total_updated_sheets': response.get('totalUpdatedSheets', 0),
                'responses': [
                    {
                        'updated_range': r.get('updatedRange', ''),
                        'updated_rows': r.get('updatedRows', 0),
                        'updated_columns': r.get('updatedColumns', 0),
                        'updated_cells': r.get('updatedCells', 0)
                    }
                    for r in responses
                ]
            }
            
            logger.info(
                "Batch updated %d ranges in spreadsheet %s", 
                len(data), spreadsheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error batch updating values: %s", str(e))
            raise
