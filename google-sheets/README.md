# Google Sheets MCP Server (python)

A Model Context Protocol (MCP) server that provides tools for interacting with Google Sheets. This server allows you to read and write data, update cells, add formulas, and apply formatting to Google Sheets.

## Features

- **Sheet Management**
  - List all Google Sheets the user has access to
  - Filter sheets by name

- **Data Operations**
  - Read data from sheets
  - Write data to sheets
  - Append data to sheets
  - Update individual cells

- **Formula Operations**
  - Add formulas to cells
  - Batch add multiple formulas

- **Formatting Operations**
  - Apply formatting to cell ranges
  - Add conditional formatting
  - Support for colors, fonts, alignment, borders, and number formats

## Installation

This project uses `uv` for dependency management.

```bash
# Create a virtual environment and install dependencies
uv venv
uv pip install -e .
```

## Configuration

### Google API Credentials

To use this MCP server, you need Google API credentials with access to the Google Sheets API. You can provide these credentials in two ways:

1. **Environment Variables**:
   - `GOOGLE_CREDENTIALS_FILE`: Path to your credentials file (service account key or OAuth client ID)
   - `GOOGLE_TOKEN_FILE`: Path to save the OAuth token (default: `token.json`)

2. **Default Locations**:
   - Place your credentials file at `credentials.json` in the same directory as the server

### Authentication Types

This server supports two types of authentication:

1. **Service Account**: For server-to-server applications
2. **OAuth 2.0**: For applications acting on behalf of users

## Setting Up OAuth for Google Sheets API

To use this MCP server with OAuth authentication, you need to set up a Google Cloud project and create OAuth 2.0 credentials. Follow these steps:

### 1. Create a Google Cloud Project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Click on the project dropdown at the top of the page
3. Click on "New Project"
4. Enter a name for your project and click "Create"
5. Wait for the project to be created and then select it from the project dropdown

### 2. Enable Required APIs

1. In your Google Cloud project, navigate to "APIs & Services" > "Library"
2. Search for and enable the following APIs:
   - Google Sheets API
   - Google Drive API (required for listing sheets)

### 3. Configure OAuth Consent Screen

1. Go to "APIs & Services" > "OAuth consent screen"
2. Select "External" user type (unless you have a Google Workspace organization)
3. Click "Create"
4. Fill in the required fields:
   - App name: (e.g., "My Sheets MCP")
   - User support email: Your email address
   - Developer contact information: Your email address
5. Click "Save and Continue"
6. On the "Scopes" page, click "Add or Remove Scopes"
7. Add the following scopes:
   - `https://www.googleapis.com/auth/spreadsheets` (for sheet operations)
   - `https://www.googleapis.com/auth/drive.readonly` (for listing sheets)
8. Click "Save and Continue"
9. On the "Test users" page, click "Add Users" and add your email address
10. Click "Save and Continue" and then "Back to Dashboard"

### 4. Create OAuth 2.0 Credentials

1. Go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "OAuth client ID"
3. Select "Desktop app" as the application type
4. Enter a name for your OAuth client (e.g., "Sheets MCP Client")
5. Click "Create"
6. Click "Download JSON" to download your credentials file
7. Rename the downloaded file to `credentials.json`

### 5. Use the Credentials with the MCP Server

There are two ways to use your OAuth credentials with the MCP server:

#### Option 1: Place in Default Location

Place the `credentials.json` file in the same directory as the server.py file.

#### Option 2: Specify via Environment Variable

Set the `GOOGLE_CREDENTIALS_FILE` environment variable to the path of your credentials file when configuring the MCP server in your Claude Desktop or Cline settings.

### 6. First-Time Authentication Flow

The first time you use the MCP server with OAuth credentials:

1. The server will detect that you don't have a token file
2. A browser window will automatically open
3. You'll be asked to sign in with your Google account
4. You'll see a warning that the app isn't verified - click "Continue"
5. Grant permission to access your Google Sheets and Drive (read-only)
6. After successful authentication, the browser will show "Authentication successful"
7. The token will be saved to `token.json` (or the path specified in `GOOGLE_TOKEN_FILE`)
8. Subsequent uses of the MCP server will use this token without requiring re-authentication

### OAuth vs. Service Account

- **OAuth 2.0**: Best for personal use or when you need to access the user's own sheets. Requires the user to go through the authentication flow.
- **Service Account**: Best for automated systems or when you need to access specific sheets shared with the service account. No user interaction required, but sheets must be explicitly shared with the service account email.

## Usage

### Integrating with Claude

#### Claude Desktop App

To use this MCP server with the Claude Desktop app, you need to add it to the Claude desktop configuration file:

1. Open the Claude desktop configuration file:
   ```bash
   # On macOS
   open ~/Library/Application\ Support/Claude/claude_desktop_config.json
   
   # On Windows
   notepad %APPDATA%\Claude\claude_desktop_config.json
   
   # On Linux
   nano ~/.config/Claude/claude_desktop_config.json
   ```

2. Add the Google Sheets MCP server to the `mcpServers` object in the configuration file:
   ```json
   {
     "mcpServers": {
       "google-sheets": {
         "command": "python",
         "args": ["/path/to/google-sheets/server.py"],
         "env": {
           "GOOGLE_CREDENTIALS_FILE": "/path/to/credentials.json"
         }
       }
     }
   }
   ```

3. Replace `/path/to/google-sheets/server.py` with the actual path to the server.py file.
4. Replace `/path/to/credentials.json` with the path to your Google API credentials file.
5. Save the file and restart Claude Desktop.

#### Cline (VSCode Extension)

To use this MCP server with Cline in VSCode, you need to add it to the Cline MCP settings file:

1. Open the Cline MCP settings file:
   ```bash
   # On macOS
   open ~/Library/Application\ Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json
   
   # On Windows
   notepad %APPDATA%\Code\User\globalStorage\saoudrizwan.claude-dev\settings\cline_mcp_settings.json
   
   # On Linux
   nano ~/.config/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json
   ```

2. Add the Google Sheets MCP server to the `mcpServers` object in the configuration file:
   ```json
   {
     "mcpServers": {
       "google-sheets": {
         "command": "python",
         "args": ["/path/to/google-sheets/server.py"],
         "env": {
           "GOOGLE_CREDENTIALS_FILE": "/path/to/credentials.json"
         },
         "disabled": false,
         "autoApprove": []
       }
     }
   }
   ```

3. Replace `/path/to/google-sheets/server.py` with the actual path to the server.py file.
4. Replace `/path/to/credentials.json` with the path to your Google API credentials file.
5. Save the file and restart VSCode.

Once configured, you can use the Google Sheets MCP tools directly in Claude or Cline:

```
Can you read data from my Google Sheet with ID "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms" and range "Sheet1!A1:D10"?
```

### Available Tools

#### Sheet Management

- `list_sheets`: List all Google Sheets the user has access to
  ```python
  # List all sheets
  list_sheets()
  
  # List sheets with "Budget" in the name
  list_sheets(filter_term="Budget")
  ```

#### Data Operations

- `read_sheet`: Read data from a Google Sheet
  ```python
  read_sheet(spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", sheet_range="Sheet1!A1:D10")
  ```

- `write_sheet`: Write data to a Google Sheet
  ```python
  write_sheet(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_range="Sheet1!A1:B2",
      values=[["Name", "Score"], ["Alice", "95"]]
  )
  ```

- `append_sheet`: Append data to a Google Sheet
  ```python
  append_sheet(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_range="Sheet1!A1",
      values=[["Bob", "87"]]
  )
  ```

- `update_cell`: Update a single cell
  ```python
  update_cell(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_name="Sheet1",
      cell="B3",
      value="92"
  )
  ```

#### Formula Operations

- `add_formula`: Add a formula to a cell
  ```python
  add_formula(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_name="Sheet1",
      cell="D1",
      formula="=SUM(B2:B10)"
  )
  ```

- `batch_add_formulas`: Add multiple formulas at once
  ```python
  batch_add_formulas(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_name="Sheet1",
      cell_formulas={
          "D1": "=SUM(B2:B10)",
          "D2": "=AVERAGE(B2:B10)"
      }
  )
  ```

#### Formatting Operations

- `format_range`: Apply formatting to a range of cells
  ```python
  format_range(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_name="Sheet1",
      cell_range="A1:D1",
      formatting={
          "backgroundColor": "#f3f3f3",
          "bold": True,
          "horizontalAlignment": "center"
      }
  )
  ```

- `add_conditional_formatting`: Add conditional formatting
  ```python
  add_conditional_formatting(
      spreadsheet_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
      sheet_name="Sheet1",
      cell_range="B2:B10",
      condition_type="NUMBER_GREATER",
      condition_values=["90"],
      format_settings={
          "backgroundColor": "green",
          "bold": True
      }
  )
  ```

### Available Resources

- `sheets://{spreadsheet_id}/{sheet_name}`: Access sheet data directly
  ```
  sheets://1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/Sheet1
  ```

## Troubleshooting

### Re-authenticating for New Scopes

If you've previously authenticated with the server and are now adding the new sheet listing functionality, you'll need to re-authenticate to grant the additional Drive API scope. Here's how to do it:

1. **Delete the existing token file**:
   ```bash
   # If using the default location
   rm token.json
   
   # If using a custom location specified in GOOGLE_TOKEN_FILE
   rm /path/to/your/token.json
   ```

2. **Restart the server**:
   The next time you run the server or use a tool that requires authentication, you'll be prompted to authenticate again with the new scopes.

3. **Complete the OAuth flow**:
   - A browser window will open automatically
   - Sign in with your Google account
   - Grant permission to access both Google Sheets and Google Drive (read-only)
   - After successful authentication, the new token will be saved

This process only needs to be done once after adding the new Drive API scope.

### Common Issues

1. **Authentication Errors**:
   - Ensure your credentials file is valid and has the correct permissions
   - For service accounts, make sure the sheet is shared with the service account email

2. **Permission Errors**:
   - Check that your credentials have the necessary scopes:
     - `https://www.googleapis.com/auth/spreadsheets` for sheet operations
     - `https://www.googleapis.com/auth/drive.readonly` for listing sheets
   - Verify that the user or service account has access to the spreadsheet
   - If you've previously authenticated but can't list sheets, you may need to re-authenticate to grant the additional Drive API scope

3. **Rate Limiting**:
   - Google Sheets API has quotas. If you hit rate limits, implement exponential backoff

## Development

### Adding New Features

To add new features to this MCP server:

1. Add new functions to `server.py`
2. Decorate them with `@mcp.tool()` or `@mcp.resource()`
3. Ensure proper type hints and docstrings for good MCP integration

### Running Tests

```bash
# Run tests
pytest
```

## License

MIT
