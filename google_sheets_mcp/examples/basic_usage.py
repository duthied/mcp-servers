#!/usr/bin/env python3
"""
Basic usage example for the Google Sheets MCP server.

This example demonstrates how to:
1. Create a new spreadsheet
2. Write data to the spreadsheet
3. Format cells
4. Create a chart
5. Apply formulas

To run this example:
1. Make sure the Google Sheets MCP server is configured in your MCP settings
2. Run this script: python basic_usage.py
"""

import asyncio
import json
import os
import sys
from typing import Dict, Any

# Add the parent directory to the path so we can import the package
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Import the MCP client
try:
    from modelcontextprotocol.client import Client
except ImportError:
    print("Error: modelcontextprotocol package not found.")
    print("Please install it with: pip install modelcontextprotocol")
    sys.exit(1)


async def main():
    """Main function to demonstrate Google Sheets MCP server usage."""
    print("Google Sheets MCP Server - Basic Usage Example")
    print("=============================================")
    
    # Connect to the MCP client
    client = Client()
    
    # Get the Google Sheets MCP server
    try:
        server = client.get_server("google_sheets")
        print("✓ Connected to Google Sheets MCP server")
    except Exception as e:
        print(f"Error connecting to Google Sheets MCP server: {e}")
        print("Make sure the server is configured in your MCP settings.")
        sys.exit(1)
    
    # Step 1: Create a new spreadsheet
    print("\n1. Creating a new spreadsheet...")
    try:
        response = await server.call_tool(
            "create_sheet",
            {
                "title": "MCP Example Spreadsheet",
                "sheet_type": "spreadsheet"
            }
        )
        
        spreadsheet_id = response["spreadsheet_id"]
        print(f"✓ Created spreadsheet with ID: {spreadsheet_id}")
        print(f"✓ Spreadsheet URL: {response['spreadsheet_url']}")
    except Exception as e:
        print(f"Error creating spreadsheet: {e}")
        sys.exit(1)
    
    # Step 2: Write data to the spreadsheet
    print("\n2. Writing data to the spreadsheet...")
    try:
        data = [
            ["Product", "Q1", "Q2", "Q3", "Q4", "Total"],
            ["Widgets", 100, 120, 130, 110, "=SUM(B2:E2)"],
            ["Gadgets", 200, 220, 210, 230, "=SUM(B3:E3)"],
            ["Doohickeys", 340, 310, 360, 390, "=SUM(B4:E4)"],
            ["Thingamajigs", 150, 160, 170, 180, "=SUM(B5:E5)"],
            ["Total", "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)", "=SUM(F2:F5)"]
        ]
        
        response = await server.call_tool(
            "write_range",
            {
                "spreadsheet_id": spreadsheet_id,
                "range": "Sheet1!A1:F6",
                "values": data,
                "value_input_option": "USER_ENTERED"
            }
        )
        
        print(f"✓ Updated {response['updated_cells']} cells")
    except Exception as e:
        print(f"Error writing data: {e}")
        sys.exit(1)
    
    # Step 3: Format the header row
    print("\n3. Formatting the header row...")
    try:
        response = await server.call_tool(
            "format_cells",
            {
                "spreadsheet_id": spreadsheet_id,
                "range": "Sheet1!A1:F1",
                "format_options": {
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
            }
        )
        
        print("✓ Formatted header row")
    except Exception as e:
        print(f"Error formatting cells: {e}")
        sys.exit(1)
    
    # Step 4: Format the total row and column
    print("\n4. Formatting the total row and column...")
    try:
        # Format total row
        await server.call_tool(
            "format_cells",
            {
                "spreadsheet_id": spreadsheet_id,
                "range": "Sheet1!A6:F6",
                "format_options": {
                    "textFormat": {
                        "bold": True
                    },
                    "backgroundColor": {
                        "red": 0.9,
                        "green": 0.9,
                        "blue": 0.9
                    }
                }
            }
        )
        
        # Format total column
        await server.call_tool(
            "format_cells",
            {
                "spreadsheet_id": spreadsheet_id,
                "range": "Sheet1!F1:F6",
                "format_options": {
                    "textFormat": {
                        "bold": True
                    },
                    "backgroundColor": {
                        "red": 0.9,
                        "green": 0.9,
                        "blue": 0.9
                    }
                }
            }
        )
        
        print("✓ Formatted total row and column")
    except Exception as e:
        print(f"Error formatting cells: {e}")
        sys.exit(1)
    
    # Step 5: Create a chart
    print("\n5. Creating a chart...")
    try:
        # Get the sheet ID
        sheet_info = await server.call_tool(
            "get_sheet_info",
            {
                "spreadsheet_id": spreadsheet_id,
                "include_sheets": True
            }
        )
        
        sheet_id = sheet_info["sheets"][0]["sheet_id"]
        
        # Create a column chart
        response = await server.call_tool(
            "create_chart",
            {
                "spreadsheet_id": spreadsheet_id,
                "sheet_id": sheet_id,
                "chart_type": "COLUMN",
                "data_range": "Sheet1!A1:E5",
                "options": {
                    "title": "Quarterly Sales by Product",
                    "legend": {
                        "position": "BOTTOM"
                    },
                    "hAxis": {
                        "title": "Quarter"
                    },
                    "vAxis": {
                        "title": "Units Sold"
                    }
                }
            }
        )
        
        print(f"✓ Created chart with ID: {response['chart_id']}")
    except Exception as e:
        print(f"Error creating chart: {e}")
        sys.exit(1)
    
    # Step 6: Add conditional formatting
    print("\n6. Adding conditional formatting...")
    try:
        response = await server.call_tool(
            "add_conditional_format",
            {
                "spreadsheet_id": spreadsheet_id,
                "range": "Sheet1!B2:E5",
                "rule_type": "BOOLEAN",
                "rule_values": {
                    "condition": {
                        "type": "NUMBER_GREATER_THAN_EQ",
                        "values": [
                            {
                                "userEnteredValue": "300"
                            }
                        ]
                    },
                    "format": {
                        "backgroundColor": {
                            "red": 0.7,
                            "green": 0.9,
                            "blue": 0.7
                        }
                    }
                }
            }
        )
        
        print("✓ Added conditional formatting")
    except Exception as e:
        print(f"Error adding conditional formatting: {e}")
        sys.exit(1)
    
    print("\nExample completed successfully!")
    print(f"View your spreadsheet at: {sheet_info['url']}")


if __name__ == "__main__":
    asyncio.run(main())
