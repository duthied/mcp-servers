# Google Sheets MCP Server

A Python-based MCP (Model Context Protocol) server that provides advanced integration with Google Sheets, including reading/writing data, formatting, charts, and formula management.

## Overview

This MCP server allows AI assistants to interact with Google Sheets through a standardized interface. It provides a comprehensive set of tools for working with spreadsheets, including:

- **Sheet Management**: Create spreadsheets, manage worksheets, and control permissions
- **Data Operations**: Read and write data to ranges, append rows, and clear ranges
- **Formatting**: Apply cell styling, conditional formatting, merge cells, and set number formats
- **Charts**: Create and manage various chart types
- **Formulas**: Apply formulas, array formulas, and create named ranges

## Installation

### Prerequisites

- Python 3.9 or higher
- Google Cloud Platform account with Google Sheets API enabled
- OAuth 2.0 credentials for Google API access

### Setup

1. Clone the repository:

```bash
git clone https://github.com/yourusername/google-sheets-mcp.git
cd google-sheets-mcp
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Set up environment variables:

Create a `.env` file in the project root with the following variables:

```
GDRIVE_CREDS_DIR=/path/to/credentials/directory
CLIENT_ID=your-oauth-client-id
CLIENT_SECRET=your-oauth-client-secret
```

4. Place your OAuth credentials JSON file in the project root as `gcp-oauth.keys.json` or set the `GOOGLE_APPLICATION_CREDENTIALS` environment variable to point to your credentials file.

## Usage

### Running the Server

To start the MCP server:

```bash
python -m google_sheets_mcp.src.server
```

The server will run on stdio, allowing it to communicate with MCP clients.

### Connecting to the Server

To connect to the server from an MCP client, add the following configuration to your MCP settings:

```json
{
  "mcpServers": {
    "google_sheets": {
      "command": "python",
      "args": ["-m", "google_sheets_mcp.src.server"],
      "env": {
        "GDRIVE_CREDS_DIR": "/path/to/credentials/directory",
        "CLIENT_ID": "your-oauth-client-id",
        "CLIENT_SECRET": "your-oauth-client-secret"
      }
    }
  }
}
```

### Available Tools

The server provides the following tools:

#### Sheet Management

- `create_sheet`: Create a new spreadsheet or worksheet
- `get_sheet_info`: Get metadata about a spreadsheet
- `manage_permissions`: Control access and sharing for a spreadsheet

#### Data Operations

- `read_range`: Read data from a specified range
- `write_range`: Write data to a specified range
- `append_rows`: Append rows to a sheet
- `clear_range`: Clear data from a range
- `batch_get_values`: Get values from multiple ranges in a single request
- `batch_update_values`: Update values in multiple ranges in a single request

#### Formatting

- `format_cells`: Apply formatting to cells
- `add_conditional_format`: Add conditional formatting to a range
- `merge_cells`: Merge cells in a range
- `set_number_format`: Set number format for cells

#### Charts

- `create_chart`: Create a chart in a spreadsheet
- `update_chart`: Update an existing chart
- `get_chart_data`: Get information about a chart
- `position_chart`: Adjust the position of a chart

#### Formulas

- `apply_formula`: Apply a formula to a cell
- `apply_array_formula`: Apply an array formula to a range
- `create_named_range`: Create a named range
- `get_named_ranges`: Get all named ranges in a spreadsheet
- `delete_named_range`: Delete a named range

### Example Usage

Here's an example of how to use the MCP server from a client:

```python
from modelcontextprotocol.client import Client

# Connect to the MCP server
client = Client()
server = client.get_server("google_sheets")

# Create a new spreadsheet
response = server.call_tool(
    "create_sheet",
    {
        "title": "My Spreadsheet",
        "sheet_type": "spreadsheet"
    }
)

# Get the spreadsheet ID
spreadsheet_id = response["spreadsheet_id"]

# Write data to the spreadsheet
server.call_tool(
    "write_range",
    {
        "spreadsheet_id": spreadsheet_id,
        "range": "Sheet1!A1:B2",
        "values": [
            ["Name", "Score"],
            ["Alice", 95]
        ]
    }
)

# Apply formatting to the header row
server.call_tool(
    "format_cells",
    {
        "spreadsheet_id": spreadsheet_id,
        "range": "Sheet1!A1:B1",
        "format_options": {
            "textFormat": {
                "bold": True
            },
            "backgroundColor": {
                "red": 0.8,
                "green": 0.8,
                "blue": 0.8
            }
        }
    }
)

# Create a chart
server.call_tool(
    "create_chart",
    {
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": 0,
        "chart_type": "COLUMN",
        "data_range": "Sheet1!A1:B2",
        "options": {
            "title": "Scores"
        }
    }
)
```

## Authentication

The server uses OAuth 2.0 for authentication with Google APIs. On first run, it will initiate an OAuth flow that opens a browser window for the user to grant permission. After authentication, the token is stored in the specified credentials directory for future use.

## Development

### Project Structure

```
google_sheets_mcp/
├── src/
│   ├── __init__.py
│   ├── server.py           # Main MCP server implementation
│   ├── auth.py             # Authentication module
│   ├── sheets/
│   │   ├── __init__.py
│   │   ├── core.py         # Core sheets functionality
│   │   ├── data_ops.py     # Data operations
│   │   ├── formatting.py   # Formatting operations
│   │   ├── charts.py       # Chart operations
│   │   └── formulas.py     # Formula operations
│   └── utils/
│       ├── __init__.py
│       ├── validators.py   # Input validation
│       └── converters.py   # Data conversion utilities
├── tests/                  # Unit tests
├── examples/               # Usage examples
├── README.md               # Project documentation
└── setup.py                # Package configuration
```

### Running Tests

To run the tests:

```bash
pytest
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
