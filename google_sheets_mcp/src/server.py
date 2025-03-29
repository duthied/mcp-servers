#!/usr/bin/env python3
"""
Google Sheets MCP Server

This module implements an MCP server that provides tools for interacting with Google Sheets,
including reading/writing data, formatting, charts, and formula management.
"""

import os
import json
import logging
from typing import Dict, Any, List, Optional

from modelcontextprotocol.sdk.server import Server
from modelcontextprotocol.sdk.server.stdio import StdioServerTransport
from modelcontextprotocol.sdk.types import (
    CallToolRequestSchema,
    ErrorCode,
    ListResourcesRequestSchema,
    ListResourceTemplatesRequestSchema,
    ListToolsRequestSchema,
    McpError,
    ReadResourceRequestSchema,
)

from .auth import GoogleAuthManager
from .sheets.core import SheetsCore
from .sheets.data_ops import DataOperations
from .sheets.formatting import FormattingOperations
from .sheets.charts import ChartOperations
from .sheets.formulas import FormulaOperations

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("google_sheets_mcp")


class GoogleSheetsMcpServer:
    """
    MCP Server implementation for Google Sheets integration.
    
    This server provides tools for interacting with Google Sheets, including
    reading/writing data, formatting, charts, and formula management.
    """

    def __init__(self):
        """Initialize the Google Sheets MCP server."""
        self.server = Server(
            {
                "name": "google-sheets-mcp",
                "version": "0.1.0",
            },
            {
                "capabilities": {
                    "resources": {},
                    "tools": {},
                },
            },
        )

        # Initialize Google Auth Manager
        self.auth_manager = GoogleAuthManager()
        
        # Initialize Sheets modules
        self.sheets_core = SheetsCore(self.auth_manager)
        self.data_ops = DataOperations(self.sheets_core)
        self.formatting = FormattingOperations(self.sheets_core)
        self.charts = ChartOperations(self.sheets_core)
        self.formulas = FormulaOperations(self.sheets_core)

        # Set up request handlers
        self._setup_tool_handlers()
        self._setup_resource_handlers()
        
        # Error handling
        self.server.onerror = self._handle_error
        
    def _handle_error(self, error: Exception) -> None:
        """Handle errors that occur during server operation."""
        logger.error(f"MCP Server Error: {error}")

    def _setup_tool_handlers(self) -> None:
        """Set up handlers for MCP tool requests."""
        # List available tools
        self.server.set_request_handler(ListToolsRequestSchema, self._handle_list_tools)
        
        # Call tool
        self.server.set_request_handler(CallToolRequestSchema, self._handle_call_tool)

    def _setup_resource_handlers(self) -> None:
        """Set up handlers for MCP resource requests."""
        # List available resources
        self.server.set_request_handler(
            ListResourcesRequestSchema, self._handle_list_resources
        )
        
        # List resource templates
        self.server.set_request_handler(
            ListResourceTemplatesRequestSchema, self._handle_list_resource_templates
        )
        
        # Read resource
        self.server.set_request_handler(
            ReadResourceRequestSchema, self._handle_read_resource
        )

    async def _handle_list_tools(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle request to list available tools."""
        return {
            "tools": [
                # Sheet Management Tools
                {
                    "name": "create_sheet",
                    "description": "Create a new spreadsheet or worksheet",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "title": {
                                "type": "string",
                                "description": "Title of the spreadsheet or worksheet",
                            },
                            "sheet_type": {
                                "type": "string",
                                "description": "Type of sheet to create (spreadsheet or worksheet)",
                                "enum": ["spreadsheet", "worksheet"],
                            },
                            "parent_id": {
                                "type": "string",
                                "description": "ID of the parent spreadsheet (for worksheet creation)",
                            },
                        },
                        "required": ["title", "sheet_type"],
                    },
                },
                {
                    "name": "get_sheet_info",
                    "description": "Get metadata about a spreadsheet",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "include_sheets": {
                                "type": "boolean",
                                "description": "Whether to include worksheet information",
                                "default": True,
                            },
                        },
                        "required": ["spreadsheet_id"],
                    },
                },
                {
                    "name": "manage_permissions",
                    "description": "Control access and sharing for a spreadsheet",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "permissions": {
                                "type": "object",
                                "description": "Permissions to set",
                            },
                        },
                        "required": ["spreadsheet_id", "permissions"],
                    },
                },
                
                # Data Operation Tools
                {
                    "name": "read_range",
                    "description": "Read data from a specified range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to read",
                            },
                            "value_render_option": {
                                "type": "string",
                                "description": "How values should be rendered",
                                "enum": ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"],
                                "default": "FORMATTED_VALUE",
                            },
                        },
                        "required": ["spreadsheet_id", "range"],
                    },
                },
                {
                    "name": "write_range",
                    "description": "Write data to a specified range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to write",
                            },
                            "values": {
                                "type": "array",
                                "description": "Values to write (2D array)",
                                "items": {
                                    "type": "array",
                                    "items": {}
                                },
                            },
                            "value_input_option": {
                                "type": "string",
                                "description": "How input values should be interpreted",
                                "enum": ["RAW", "USER_ENTERED"],
                                "default": "USER_ENTERED",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "values"],
                    },
                },
                {
                    "name": "append_rows",
                    "description": "Append rows to a sheet",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to append to",
                            },
                            "values": {
                                "type": "array",
                                "description": "Values to append (2D array)",
                                "items": {
                                    "type": "array",
                                    "items": {}
                                },
                            },
                            "insert_data_option": {
                                "type": "string",
                                "description": "How the input data should be inserted",
                                "enum": ["OVERWRITE", "INSERT_ROWS"],
                                "default": "INSERT_ROWS",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "values"],
                    },
                },
                {
                    "name": "clear_range",
                    "description": "Clear data from a range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to clear",
                            },
                        },
                        "required": ["spreadsheet_id", "range"],
                    },
                },
                
                # Formatting Tools
                {
                    "name": "format_cells",
                    "description": "Apply formatting to cells",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to format",
                            },
                            "format_options": {
                                "type": "object",
                                "description": "Formatting options to apply",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "format_options"],
                    },
                },
                {
                    "name": "add_conditional_format",
                    "description": "Add conditional formatting to a range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to format",
                            },
                            "rule_type": {
                                "type": "string",
                                "description": "Type of conditional formatting rule",
                            },
                            "rule_values": {
                                "type": "object",
                                "description": "Values for the conditional formatting rule",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "rule_type", "rule_values"],
                    },
                },
                {
                    "name": "merge_cells",
                    "description": "Merge cells in a range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to merge",
                            },
                            "merge_type": {
                                "type": "string",
                                "description": "Type of merge to perform",
                                "enum": ["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"],
                                "default": "MERGE_ALL",
                            },
                        },
                        "required": ["spreadsheet_id", "range"],
                    },
                },
                {
                    "name": "set_number_format",
                    "description": "Set number format for cells",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to format",
                            },
                            "number_format": {
                                "type": "string",
                                "description": "Number format pattern to apply",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "number_format"],
                    },
                },
                
                # Chart Tools
                {
                    "name": "create_chart",
                    "description": "Create a chart in a spreadsheet",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "sheet_id": {
                                "type": "integer",
                                "description": "ID of the sheet to add the chart to",
                            },
                            "chart_type": {
                                "type": "string",
                                "description": "Type of chart to create",
                                "enum": ["BAR", "LINE", "PIE", "COLUMN", "AREA", "SCATTER"],
                            },
                            "data_range": {
                                "type": "string",
                                "description": "A1 notation range for chart data",
                            },
                            "options": {
                                "type": "object",
                                "description": "Chart options",
                            },
                        },
                        "required": ["spreadsheet_id", "sheet_id", "chart_type", "data_range"],
                    },
                },
                {
                    "name": "update_chart",
                    "description": "Update an existing chart",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "sheet_id": {
                                "type": "integer",
                                "description": "ID of the sheet containing the chart",
                            },
                            "chart_id": {
                                "type": "integer",
                                "description": "ID of the chart to update",
                            },
                            "update_options": {
                                "type": "object",
                                "description": "Chart options to update",
                            },
                        },
                        "required": ["spreadsheet_id", "sheet_id", "chart_id", "update_options"],
                    },
                },
                {
                    "name": "get_chart_data",
                    "description": "Get information about a chart",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "sheet_id": {
                                "type": "integer",
                                "description": "ID of the sheet containing the chart",
                            },
                            "chart_id": {
                                "type": "integer",
                                "description": "ID of the chart",
                            },
                        },
                        "required": ["spreadsheet_id", "sheet_id", "chart_id"],
                    },
                },
                {
                    "name": "position_chart",
                    "description": "Adjust the position of a chart",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "sheet_id": {
                                "type": "integer",
                                "description": "ID of the sheet containing the chart",
                            },
                            "chart_id": {
                                "type": "integer",
                                "description": "ID of the chart to position",
                            },
                            "position": {
                                "type": "object",
                                "description": "Position information",
                                "properties": {
                                    "overlayPosition": {
                                        "type": "object",
                                        "properties": {
                                            "anchorCell": {
                                                "type": "object",
                                                "properties": {
                                                    "sheetId": {"type": "integer"},
                                                    "rowIndex": {"type": "integer"},
                                                    "columnIndex": {"type": "integer"},
                                                },
                                            },
                                            "offsetXPixels": {"type": "integer"},
                                            "offsetYPixels": {"type": "integer"},
                                            "widthPixels": {"type": "integer"},
                                            "heightPixels": {"type": "integer"},
                                        },
                                    },
                                },
                            },
                        },
                        "required": ["spreadsheet_id", "sheet_id", "chart_id", "position"],
                    },
                },
                
                # Formula Tools
                {
                    "name": "apply_formula",
                    "description": "Apply a formula to a cell",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to apply formula to",
                            },
                            "formula": {
                                "type": "string",
                                "description": "Formula to apply",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "formula"],
                    },
                },
                {
                    "name": "apply_array_formula",
                    "description": "Apply an array formula to a range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to apply array formula to",
                            },
                            "formula": {
                                "type": "string",
                                "description": "Array formula to apply",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "formula"],
                    },
                },
                {
                    "name": "create_named_range",
                    "description": "Create a named range",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "spreadsheet_id": {
                                "type": "string",
                                "description": "ID of the spreadsheet",
                            },
                            "range": {
                                "type": "string",
                                "description": "A1 notation range to name",
                            },
                            "name": {
                                "type": "string",
                                "description": "Name for the range",
                            },
                            "scope": {
                                "type": "string",
                                "description": "Scope of the named range",
                                "enum": ["WORKBOOK", "SHEET"],
                                "default": "WORKBOOK",
                            },
                        },
                        "required": ["spreadsheet_id", "range", "name"],
                    },
                },
            ]
        }

    async def _handle_call_tool(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle request to call a tool."""
        tool_name = request["params"]["name"]
        arguments = request["params"]["arguments"]
        
        logger.info(f"Calling tool: {tool_name} with arguments: {arguments}")
        
        try:
            # Sheet Management Tools
            if tool_name == "create_sheet":
                result = await self.sheets_core.create_sheet(
                    title=arguments["title"],
                    sheet_type=arguments["sheet_type"],
                    parent_id=arguments.get("parent_id"),
                )
            elif tool_name == "get_sheet_info":
                result = await self.sheets_core.get_sheet_info(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    include_sheets=arguments.get("include_sheets", True),
                )
            elif tool_name == "manage_permissions":
                result = await self.sheets_core.manage_permissions(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    permissions=arguments["permissions"],
                )
                
            # Data Operation Tools
            elif tool_name == "read_range":
                result = await self.data_ops.read_range(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    value_render_option=arguments.get("value_render_option", "FORMATTED_VALUE"),
                )
            elif tool_name == "write_range":
                result = await self.data_ops.write_range(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    values=arguments["values"],
                    value_input_option=arguments.get("value_input_option", "USER_ENTERED"),
                )
            elif tool_name == "append_rows":
                result = await self.data_ops.append_rows(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    values=arguments["values"],
                    insert_data_option=arguments.get("insert_data_option", "INSERT_ROWS"),
                )
            elif tool_name == "clear_range":
                result = await self.data_ops.clear_range(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                )
                
            # Formatting Tools
            elif tool_name == "format_cells":
                result = await self.formatting.format_cells(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    format_options=arguments["format_options"],
                )
            elif tool_name == "add_conditional_format":
                result = await self.formatting.add_conditional_format(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    rule_type=arguments["rule_type"],
                    rule_values=arguments["rule_values"],
                )
            elif tool_name == "merge_cells":
                result = await self.formatting.merge_cells(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    merge_type=arguments.get("merge_type", "MERGE_ALL"),
                )
            elif tool_name == "set_number_format":
                result = await self.formatting.set_number_format(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    number_format=arguments["number_format"],
                )
                
            # Chart Tools
            elif tool_name == "create_chart":
                result = await self.charts.create_chart(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    sheet_id=arguments["sheet_id"],
                    chart_type=arguments["chart_type"],
                    data_range=arguments["data_range"],
                    options=arguments.get("options", {}),
                )
            elif tool_name == "update_chart":
                result = await self.charts.update_chart(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    sheet_id=arguments["sheet_id"],
                    chart_id=arguments["chart_id"],
                    update_options=arguments["update_options"],
                )
            elif tool_name == "get_chart_data":
                result = await self.charts.get_chart_data(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    sheet_id=arguments["sheet_id"],
                    chart_id=arguments["chart_id"],
                )
            elif tool_name == "position_chart":
                result = await self.charts.position_chart(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    sheet_id=arguments["sheet_id"],
                    chart_id=arguments["chart_id"],
                    position=arguments["position"],
                )
                
            # Formula Tools
            elif tool_name == "apply_formula":
                result = await self.formulas.apply_formula(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    formula=arguments["formula"],
                )
            elif tool_name == "apply_array_formula":
                result = await self.formulas.apply_array_formula(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    formula=arguments["formula"],
                )
            elif tool_name == "create_named_range":
                result = await self.formulas.create_named_range(
                    spreadsheet_id=arguments["spreadsheet_id"],
                    range=arguments["range"],
                    name=arguments["name"],
                    scope=arguments.get("scope", "WORKBOOK"),
                )
            else:
                raise McpError(
                    ErrorCode.MethodNotFound,
                    f"Unknown tool: {tool_name}",
                )
                
            return {
                "content": [
                    {
                        "type": "text",
                        "text": json.dumps(result, indent=2),
                    }
                ]
            }
            
        except Exception as e:
            logger.error(f"Error calling tool {tool_name}: {e}")
            return {
                "content": [
                    {
                        "type": "text",
                        "text": f"Error: {str(e)}",
                    }
                ],
                "isError": True,
            }

    async def _handle_list_resources(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle request to list available resources."""
        return {
            "resources": [
                # We could add static resources here if needed
            ]
        }

    async def _handle_list_resource_templates(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle request to list available resource templates."""
        return {
            "resourceTemplates": [
                {
                    "uriTemplate": "sheets://{spreadsheet_id}/{range}",
                    "name": "Google Sheets Range",
                    "mimeType": "application/json",
                    "description": "Data from a specific range in a Google Sheet",
                }
            ]
        }

    async def _handle_read_resource(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle request to read a resource."""
        uri = request["params"]["uri"]
        
        # Parse the URI to extract spreadsheet_id and range
        match = uri.match(r"^sheets://([^/]+)/(.+)$")
        if not match:
            raise McpError(
                ErrorCode.InvalidRequest,
                f"Invalid URI format: {uri}",
            )
            
        spreadsheet_id = match.group(1)
        range_notation = match.group(2)
        
        try:
            # Use the data_ops module to read the range
            data = await self.data_ops.read_range(
                spreadsheet_id=spreadsheet_id,
                range=range_notation,
            )
            
            return {
                "contents": [
                    {
                        "uri": uri,
                        "mimeType": "application/json",
                        "text": json.dumps(data, indent=2),
                    }
                ]
            }
        except Exception as e:
            logger.error(f"Error reading resource {uri}: {e}")
            raise McpError(
                ErrorCode.InternalError,
                f"Error reading resource: {str(e)}",
            )

    async def run(self) -> None:
        """Run the MCP server."""
        transport = StdioServerTransport()
        await self.server.connect(transport)
        logger.info("Google Sheets MCP server running on stdio")


def main() -> None:
    """Main entry point for the Google Sheets MCP server."""
    import asyncio
    
    server = GoogleSheetsMcpServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    main()
