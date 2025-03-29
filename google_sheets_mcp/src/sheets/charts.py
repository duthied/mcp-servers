#!/usr/bin/env python3
"""
Google Sheets Charts Module for MCP Server

This module provides functionality for creating and managing charts in Google Sheets.
"""

import logging
from typing import Dict, Any, List, Optional

from .core import SheetsCore

# Configure logging
logger = logging.getLogger("google_sheets_mcp.sheets.charts")


class ChartOperations:
    """
    Chart operations for Google Sheets.
    
    This class provides methods for creating and managing charts in Google Sheets,
    including creating charts, updating charts, and positioning charts.
    """
    
    def __init__(self, sheets_core: SheetsCore):
        """
        Initialize the Chart Operations module.
        
        Args:
            sheets_core: An instance of SheetsCore for base sheet operations.
        """
        self.sheets_core = sheets_core
    
    async def create_chart(
        self, 
        spreadsheet_id: str, 
        sheet_id: int, 
        chart_type: str, 
        data_range: str, 
        options: Dict[str, Any] = None
    ) -> Dict[str, Any]:
        """
        Create a chart in a spreadsheet.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The ID of the sheet to add the chart to.
            chart_type: Type of chart to create.
                Options: "BAR", "LINE", "PIE", "COLUMN", "AREA", "SCATTER"
            data_range: A1 notation range for chart data.
            options: Chart options.
                Example: {
                    "title": "Sales Data",
                    "legend": {"position": "BOTTOM"},
                    "hAxis": {"title": "Month"},
                    "vAxis": {"title": "Sales"}
                }
        
        Returns:
            A dictionary containing information about the created chart.
        """
        try:
            # Parse the data range
            range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, data_range)
            
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare chart spec based on chart type
            chart_spec = self._create_chart_spec(chart_type, range_info, options or {})
            
            # Prepare the request body
            request_body = {
                'requests': [
                    {
                        'addChart': {
                            'chart': {
                                'spec': chart_spec,
                                'position': {
                                    'overlayPosition': {
                                        'anchorCell': {
                                            'sheetId': sheet_id,
                                            'rowIndex': 0,
                                            'columnIndex': 0
                                        },
                                        'offsetXPixels': 0,
                                        'offsetYPixels': 0,
                                        'widthPixels': 600,
                                        'heightPixels': 400
                                    }
                                }
                            }
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
            
            # Extract the chart ID
            chart_id = response['replies'][0]['addChart']['chart']['chartId']
            
            logger.info(
                "Created %s chart in spreadsheet %s, sheet %s with chart ID %s", 
                chart_type, spreadsheet_id, sheet_id, chart_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'sheet_id': sheet_id,
                'chart_id': chart_id,
                'chart_type': chart_type,
                'data_range': data_range,
                'options': options or {}
            }
            
        except Exception as e:
            logger.error("Error creating chart: %s", str(e))
            raise
    
    async def update_chart(
        self, 
        spreadsheet_id: str, 
        sheet_id: int, 
        chart_id: int, 
        update_options: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Update an existing chart.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The ID of the sheet containing the chart.
            chart_id: The ID of the chart to update.
            update_options: Chart options to update.
                Example: {
                    "title": "Updated Title",
                    "data_range": "Sheet1!A1:B10",  # New data range
                    "chart_type": "LINE",  # New chart type
                    "options": {  # New chart options
                        "legend": {"position": "TOP"},
                        "hAxis": {"title": "Updated X-Axis"}
                    }
                }
        
        Returns:
            A dictionary containing information about the updated chart.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Get current chart info
            current_chart = await self.get_chart_data(spreadsheet_id, sheet_id, chart_id)
            
            # Prepare the update request
            requests = []
            
            # Update chart properties
            if 'title' in update_options or 'options' in update_options:
                chart_properties = {}
                
                if 'title' in update_options:
                    chart_properties['title'] = update_options['title']
                
                if 'options' in update_options:
                    options = update_options['options']
                    
                    if 'legend' in options:
                        chart_properties['legend'] = options['legend']
                    
                    # Add other options as needed
                
                requests.append({
                    'updateChartSpec': {
                        'chartId': chart_id,
                        'spec': {
                            'title': chart_properties.get('title', current_chart.get('title', '')),
                            'basicChart': {
                                'legendPosition': chart_properties.get('legend', {}).get('position', 'RIGHT')
                                # Add other properties as needed
                            }
                        }
                    }
                })
            
            # Update data range if provided
            if 'data_range' in update_options:
                data_range = update_options['data_range']
                range_info = await self.sheets_core.parse_range_notation(spreadsheet_id, data_range)
                
                # Create a new chart spec with the updated data range
                chart_type = update_options.get('chart_type', current_chart.get('chart_type', 'BAR'))
                chart_options = update_options.get('options', {})
                
                chart_spec = self._create_chart_spec(chart_type, range_info, chart_options)
                
                requests.append({
                    'updateChartSpec': {
                        'chartId': chart_id,
                        'spec': chart_spec
                    }
                })
            
            # Execute the request if there are updates
            if requests:
                request_body = {'requests': requests}
                
                request = sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=request_body
                )
                response = request.execute()
                
                logger.info(
                    "Updated chart %s in spreadsheet %s, sheet %s", 
                    chart_id, spreadsheet_id, sheet_id
                )
            else:
                logger.info("No updates provided for chart %s", chart_id)
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'sheet_id': sheet_id,
                'chart_id': chart_id,
                'updated_options': update_options,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error updating chart: %s", str(e))
            raise
    
    async def get_chart_data(
        self, 
        spreadsheet_id: str, 
        sheet_id: int, 
        chart_id: int
    ) -> Dict[str, Any]:
        """
        Get information about a chart.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The ID of the sheet containing the chart.
            chart_id: The ID of the chart.
        
        Returns:
            A dictionary containing information about the chart.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Get the spreadsheet
            request = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields='sheets.charts,sheets.properties.sheetId'
            )
            response = request.execute()
            
            # Find the sheet with the specified sheet_id
            sheet = None
            for s in response.get('sheets', []):
                if s['properties']['sheetId'] == sheet_id:
                    sheet = s
                    break
            
            if not sheet:
                raise ValueError(f"Sheet with ID {sheet_id} not found")
            
            # Find the chart with the specified chart_id
            chart = None
            for c in sheet.get('charts', []):
                if c['chartId'] == chart_id:
                    chart = c
                    break
            
            if not chart:
                raise ValueError(f"Chart with ID {chart_id} not found in sheet {sheet_id}")
            
            # Extract chart information
            chart_spec = chart.get('spec', {})
            chart_position = chart.get('position', {})
            
            # Determine chart type
            chart_type = 'UNKNOWN'
            for chart_type_key in ['basicChart', 'pieChart', 'bubbleChart', 'candlestickChart', 'orgChart', 'histogramChart']:
                if chart_type_key in chart_spec:
                    chart_type = chart_type_key.replace('Chart', '').upper()
                    break
            
            # Extract data source
            data_sources = []
            if 'basicChart' in chart_spec:
                for series in chart_spec['basicChart'].get('series', []):
                    if 'series' in series and 'sourceRange' in series['series']:
                        source_range = series['series']['sourceRange']
                        data_sources.append(source_range)
            
            result = {
                'spreadsheet_id': spreadsheet_id,
                'sheet_id': sheet_id,
                'chart_id': chart_id,
                'chart_type': chart_type,
                'title': chart_spec.get('title', ''),
                'data_sources': data_sources,
                'position': chart_position
            }
            
            logger.info(
                "Retrieved data for chart %s in spreadsheet %s, sheet %s", 
                chart_id, spreadsheet_id, sheet_id
            )
            
            return result
            
        except Exception as e:
            logger.error("Error getting chart data: %s", str(e))
            raise
    
    async def position_chart(
        self, 
        spreadsheet_id: str, 
        sheet_id: int, 
        chart_id: int, 
        position: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Adjust the position of a chart.
        
        Args:
            spreadsheet_id: The ID of the spreadsheet.
            sheet_id: The ID of the sheet containing the chart.
            chart_id: The ID of the chart to position.
            position: Position information.
                Example: {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": 0,
                            "rowIndex": 10,
                            "columnIndex": 5
                        },
                        "offsetXPixels": 10,
                        "offsetYPixels": 10,
                        "widthPixels": 400,
                        "heightPixels": 300
                    }
                }
        
        Returns:
            A dictionary containing information about the position update.
        """
        try:
            # Get the sheets service
            sheets_service = self.sheets_core._get_sheets_service()
            
            # Prepare the request body
            request_body = {
                'requests': [
                    {
                        'updateChartSpec': {
                            'chartId': chart_id,
                            'position': position
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
                "Updated position for chart %s in spreadsheet %s, sheet %s", 
                chart_id, spreadsheet_id, sheet_id
            )
            
            return {
                'spreadsheet_id': spreadsheet_id,
                'sheet_id': sheet_id,
                'chart_id': chart_id,
                'position': position,
                'status': 'success'
            }
            
        except Exception as e:
            logger.error("Error positioning chart: %s", str(e))
            raise
    
    def _create_chart_spec(
        self, 
        chart_type: str, 
        range_info: Dict[str, Any], 
        options: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Create a chart specification based on chart type and options.
        
        Args:
            chart_type: Type of chart to create.
            range_info: Range information from parse_range_notation.
            options: Chart options.
        
        Returns:
            A dictionary containing the chart specification.
        """
        # Base chart spec
        chart_spec = {
            'title': options.get('title', ''),
        }
        
        # Convert sheet ID and range to source ranges
        source_range = {
            'sources': [
                {
                    'sheetId': range_info['sheet_id'],
                    'startRowIndex': 0,  # Will be determined from A1 notation
                    'endRowIndex': 0,    # Will be determined from A1 notation
                    'startColumnIndex': 0,  # Will be determined from A1 notation
                    'endColumnIndex': 0     # Will be determined from A1 notation
                }
            ]
        }
        
        # Convert A1 notation to grid range
        grid_range = self._a1_to_grid_range(range_info)
        source_range['sources'][0].update(grid_range)
        
        # Create chart spec based on chart type
        if chart_type in ['BAR', 'LINE', 'COLUMN', 'AREA', 'SCATTER']:
            # Basic chart types
            chart_spec['basicChart'] = {
                'chartType': chart_type,
                'legendPosition': options.get('legend', {}).get('position', 'RIGHT'),
                'axis': [
                    {
                        'position': 'BOTTOM_AXIS',
                        'title': options.get('hAxis', {}).get('title', '')
                    },
                    {
                        'position': 'LEFT_AXIS',
                        'title': options.get('vAxis', {}).get('title', '')
                    }
                ],
                'domains': [
                    {
                        'domain': {
                            'sourceRange': source_range
                        }
                    }
                ],
                'series': [
                    {
                        'series': {
                            'sourceRange': source_range
                        },
                        'targetAxis': 'LEFT_AXIS'
                    }
                ],
                'headerCount': 1
            }
        elif chart_type == 'PIE':
            # Pie chart
            chart_spec['pieChart'] = {
                'legendPosition': options.get('legend', {}).get('position', 'RIGHT'),
                'domain': {
                    'sourceRange': source_range
                },
                'series': {
                    'sourceRange': source_range
                },
                'pieHole': options.get('pieHole', 0)
            }
        
        return chart_spec
    
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
