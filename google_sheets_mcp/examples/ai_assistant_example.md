# Using Google Sheets MCP Server with AI Assistants

This example demonstrates how to use the Google Sheets MCP server with AI assistants like Claude.

## Setup

1. Install the Google Sheets MCP server:
   ```bash
   pip install google-sheets-mcp
   ```

2. Configure the MCP server in your MCP settings file:
   - For Claude Desktop: `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS)
   - For Claude VSCode Extension: `~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json` (macOS)

   Add the following configuration:
   ```json
   {
     "mcpServers": {
       "google_sheets": {
         "command": "python",
         "args": ["-m", "google_sheets_mcp.src.server"],
         "env": {
           "GDRIVE_CREDS_DIR": "/path/to/credentials/directory",
           "CLIENT_ID": "your-oauth-client-id",
           "CLIENT_SECRET": "your-oauth-client-secret",
           "GOOGLE_APPLICATION_CREDENTIALS": "/path/to/gcp-oauth.keys.json"
         },
         "disabled": false,
         "autoApprove": []
       }
     }
   }
   ```

3. Restart your AI assistant application to load the new MCP server.

## Example Prompts

Here are some example prompts you can use with your AI assistant to interact with Google Sheets:

### Creating a New Spreadsheet

```
Create a new Google Sheets spreadsheet called "Monthly Budget" with columns for Category, Budget Amount, Actual Spending, and Difference.
```

### Writing Data to a Spreadsheet

```
I have a Google Sheets spreadsheet with ID "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms". 
Please add the following data to it:

Category | Budget | Actual | Difference
---------|--------|--------|----------
Rent     | 1500   | 1500   | =B2-C2
Groceries| 500    | 520    | =B3-C3
Utilities| 200    | 180    | =B4-C4
Entertainment| 300| 350    | =B5-C5
Transportation| 250| 240   | =B6-C6
Total    | =SUM(B2:B6) | =SUM(C2:C6) | =SUM(D2:D6)
```

### Formatting a Spreadsheet

```
Format my budget spreadsheet with ID "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms" as follows:
1. Make the header row bold with a light gray background
2. Make the Total row bold
3. Format the Budget and Actual columns as currency
4. Add conditional formatting to the Difference column to highlight negative values in red
```

### Creating a Chart

```
Create a bar chart in my budget spreadsheet with ID "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms" that compares Budget vs Actual spending for each category.
```

### Reading Data from a Spreadsheet

```
Read the data from my budget spreadsheet with ID "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms" and tell me which categories are over budget.
```

### Applying Formulas

```
In my spreadsheet with ID "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", add a new column called "Percentage of Budget" that calculates what percentage of the budget was spent for each category.
```

## Expected AI Assistant Response

When you provide these prompts to an AI assistant with access to the Google Sheets MCP server, it should:

1. Recognize the request as a Google Sheets operation
2. Use the appropriate MCP tool to perform the operation
3. Provide feedback on the result

For example, when asked to create a new spreadsheet, the AI might respond:

```
I've created a new Google Sheets spreadsheet called "Monthly Budget" for you.

Spreadsheet ID: 1a2b3c4d5e6f7g8h9i0j
URL: https://docs.google.com/spreadsheets/d/1a2b3c4d5e6f7g8h9i0j/edit

I've set up the following columns:
- Category
- Budget Amount
- Actual Spending
- Difference (with a formula to calculate Budget - Actual)

I've also added a Total row at the bottom with SUM formulas to calculate the totals for each column.

You can access your spreadsheet using the URL above.
```

## Troubleshooting

If the AI assistant is unable to use the Google Sheets MCP server, check the following:

1. Make sure the MCP server is properly configured in your MCP settings file
2. Verify that the OAuth credentials are correct
3. Check that the AI assistant has permission to use MCP servers
4. Look for any error messages in the AI assistant's logs
