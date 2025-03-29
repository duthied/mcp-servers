#!/usr/bin/env python3
"""
Tests for the Google Sheets MCP Server.

This module contains tests for the Google Sheets MCP server functionality.
"""

import asyncio
import json
import os
import pytest
from unittest.mock import MagicMock, patch

# Add the parent directory to the path so we can import the package
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from google_sheets_mcp.src.server import GoogleSheetsMcpServer
from modelcontextprotocol.sdk.types import (
    CallToolRequestSchema,
    ListToolsRequestSchema,
)


@pytest.fixture
def mock_auth_manager():
    """Create a mock authentication manager."""
    mock = MagicMock()
    mock.get_sheets_service.return_value = MagicMock()
    mock.get_drive_service.return_value = MagicMock()
    return mock


@pytest.fixture
def server(mock_auth_manager):
    """Create a server instance with mocked dependencies."""
    with patch('google_sheets_mcp.src.auth.GoogleAuthManager', return_value=mock_auth_manager):
        server = GoogleSheetsMcpServer()
        yield server


@pytest.mark.asyncio
async def test_list_tools(server):
    """Test that the server can list available tools."""
    # Create a request to list tools
    request = {
        "jsonrpc": "2.0",
        "id": "1",
        "method": "list_tools",
        "params": {}
    }
    
    # Call the handler
    response = await server._handle_list_tools(request)
    
    # Check that the response contains tools
    assert "tools" in response
    assert len(response["tools"]) > 0
    
    # Check that each tool has the required fields
    for tool in response["tools"]:
        assert "name" in tool
        assert "description" in tool
        assert "inputSchema" in tool


@pytest.mark.asyncio
async def test_create_sheet(server):
    """Test creating a new spreadsheet."""
    # Mock the sheets service
    sheets_service = server.sheets_core._get_sheets_service()
    spreadsheets = sheets_service.spreadsheets.return_value
    create = spreadsheets.create.return_value
    
    # Mock the execute method to return a sample response
    create.execute.return_value = {
        "spreadsheetId": "test_spreadsheet_id",
        "properties": {"title": "Test Spreadsheet"},
        "sheets": [{"properties": {"sheetId": 0, "title": "Sheet1"}}],
        "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/test_spreadsheet_id"
    }
    
    # Create a request to create a sheet
    request = {
        "jsonrpc": "2.0",
        "id": "1",
        "method": "call_tool",
        "params": {
            "name": "create_sheet",
            "arguments": {
                "title": "Test Spreadsheet",
                "sheet_type": "spreadsheet"
            }
        }
    }
    
    # Call the handler
    response = await server._handle_call_tool(request)
    
    # Check that the response contains the expected content
    assert "content" in response
    assert len(response["content"]) == 1
    assert response["content"][0]["type"] == "text"
    
    # Parse the JSON response
    result = json.loads(response["content"][0]["text"])
    
    # Check the result
    assert result["spreadsheet_id"] == "test_spreadsheet_id"
    assert result["title"] == "Test Spreadsheet"
    assert len(result["sheets"]) == 1
    assert result["sheets"][0]["title"] == "Sheet1"
    assert result["spreadsheet_url"] == "https://docs.google.com/spreadsheets/d/test_spreadsheet_id"


@pytest.mark.asyncio
async def test_read_range(server):
    """Test reading data from a range."""
    # Mock the sheets service
    sheets_service = server.sheets_core._get_sheets_service()
    spreadsheets = sheets_service.spreadsheets.return_value
    values = spreadsheets.values.return_value
    get = values.get.return_value
    
    # Mock the execute method to return a sample response
    get.execute.return_value = {
        "range": "Sheet1!A1:B2",
        "values": [
            ["Header 1", "Header 2"],
            ["Value 1", "Value 2"]
        ]
    }
    
    # Create a request to read a range
    request = {
        "jsonrpc": "2.0",
        "id": "1",
        "method": "call_tool",
        "params": {
            "name": "read_range",
            "arguments": {
                "spreadsheet_id": "test_spreadsheet_id",
                "range": "Sheet1!A1:B2"
            }
        }
    }
    
    # Call the handler
    response = await server._handle_call_tool(request)
    
    # Check that the response contains the expected content
    assert "content" in response
    assert len(response["content"]) == 1
    assert response["content"][0]["type"] == "text"
    
    # Parse the JSON response
    result = json.loads(response["content"][0]["text"])
    
    # Check the result
    assert result["spreadsheet_id"] == "test_spreadsheet_id"
    assert result["range"] == "Sheet1!A1:B2"
    assert len(result["values"]) == 2
    assert result["values"][0][0] == "Header 1"
    assert result["values"][0][1] == "Header 2"
    assert result["values"][1][0] == "Value 1"
    assert result["values"][1][1] == "Value 2"
    assert result["num_rows"] == 2
    assert result["num_columns"] == 2


@pytest.mark.asyncio
async def test_write_range(server):
    """Test writing data to a range."""
    # Mock the sheets service
    sheets_service = server.sheets_core._get_sheets_service()
    spreadsheets = sheets_service.spreadsheets.return_value
    values = spreadsheets.values.return_value
    update = values.update.return_value
    
    # Mock the execute method to return a sample response
    update.execute.return_value = {
        "spreadsheetId": "test_spreadsheet_id",
        "updatedRange": "Sheet1!A1:B2",
        "updatedRows": 2,
        "updatedColumns": 2,
        "updatedCells": 4
    }
    
    # Create a request to write to a range
    request = {
        "jsonrpc": "2.0",
        "id": "1",
        "method": "call_tool",
        "params": {
            "name": "write_range",
            "arguments": {
                "spreadsheet_id": "test_spreadsheet_id",
                "range": "Sheet1!A1:B2",
                "values": [
                    ["Header 1", "Header 2"],
                    ["Value 1", "Value 2"]
                ]
            }
        }
    }
    
    # Call the handler
    response = await server._handle_call_tool(request)
    
    # Check that the response contains the expected content
    assert "content" in response
    assert len(response["content"]) == 1
    assert response["content"][0]["type"] == "text"
    
    # Parse the JSON response
    result = json.loads(response["content"][0]["text"])
    
    # Check the result
    assert result["spreadsheet_id"] == "test_spreadsheet_id"
    assert result["range"] == "Sheet1!A1:B2"
    assert result["updated_rows"] == 2
    assert result["updated_columns"] == 2
    assert result["updated_cells"] == 4


if __name__ == "__main__":
    pytest.main(["-xvs", __file__])
