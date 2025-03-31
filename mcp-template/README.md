# MCP Template

A starter template for creating new Model Context Protocol (MCP) servers. This template provides the basic structure and examples needed to quickly develop custom MCP servers that extend Claude's capabilities with additional tools and resources.

## Features

- **Basic Server Setup**: Core configuration for an MCP server using FastMCP
- **Example Tool Implementation**: Sample implementation of a simple addition tool
- **Example Resource Implementation**: Sample implementation of a dynamic greeting resource
- **Project Structure**: Properly configured Python project with dependencies

## Installation

This project uses `uv` for dependency management.

```bash
# Create a virtual environment and install dependencies
uv venv
uv pip install -e .
```

## Usage

### Creating a New MCP Server

1. **Copy the Template**:
   ```bash
   # Clone or copy the template to a new directory
   cp -r mcp-template my-new-mcp-server
   cd my-new-mcp-server
   ```

2. **Update Project Configuration**:
   Edit `pyproject.toml` to update the project name, description, and dependencies:
   ```toml
   [project]
   name = "my-new-mcp-server"
   version = "0.1.0"
   description = "My custom MCP server"
   # ... other configuration
   ```

3. **Customize the Server**:
   Edit `server.py` to implement your custom tools and resources.

### Example Code

The template includes the following examples:

#### Tool Example

```python
# Add an addition tool
@mcp.tool()
def add(a: int, b: int) -> int:
    """Add two numbers"""
    return a + b
```

This creates a tool that Claude can use to add two numbers together.

#### Resource Example

```python
# Add a dynamic greeting resource
@mcp.resource("greeting://{name}")
def get_greeting(name: str) -> str:
    """Get a personalized greeting"""
    return f"Hello, {name}!"
```

This creates a dynamic resource that Claude can access to get a personalized greeting.

## Integrating with Claude

### Claude Desktop App

To use your MCP server with the Claude Desktop app:

1. Open the Claude desktop configuration file:
   ```bash
   # On macOS
   open ~/Library/Application\ Support/Claude/claude_desktop_config.json
   
   # On Windows
   notepad %APPDATA%\Claude\claude_desktop_config.json
   
   # On Linux
   nano ~/.config/Claude/claude_desktop_config.json
   ```

2. Add your MCP server to the `mcpServers` object:
   ```json
   {
     "mcpServers": {
       "my-new-mcp-server": {
         "command": "python",
         "args": ["/path/to/my-new-mcp-server/server.py"],
         "env": {
           // Any environment variables your server needs
         }
       }
     }
   }
   ```

3. Save the file and restart Claude Desktop.

### Cline (VSCode Extension)

To use your MCP server with Cline in VSCode:

1. Open the Cline MCP settings file:
   ```bash
   # On macOS
   open ~/Library/Application\ Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json
   
   # On Windows
   notepad %APPDATA%\Code\User\globalStorage\saoudrizwan.claude-dev\settings\cline_mcp_settings.json
   
   # On Linux
   nano ~/.config/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json
   ```

2. Add your MCP server to the `mcpServers` object:
   ```json
   {
     "mcpServers": {
       "my-new-mcp-server": {
         "command": "python",
         "args": ["/path/to/my-new-mcp-server/server.py"],
         "env": {
           // Any environment variables your server needs
         },
         "disabled": false,
         "autoApprove": []
       }
     }
   }
   ```

3. Save the file and restart VSCode.

## Development

### Adding New Tools

To add a new tool to your MCP server:

```python
@mcp.tool()
def my_new_tool(param1: str, param2: int) -> str:
    """Description of what the tool does
    
    Args:
        param1: Description of param1
        param2: Description of param2
        
    Returns:
        Description of the return value
    """
    # Tool implementation
    return f"Result: {param1} {param2}"
```

### Adding New Resources

To add a new resource to your MCP server:

```python
@mcp.resource("myresource://{id}")
def get_my_resource(id: str) -> str:
    """Description of what the resource provides
    
    Args:
        id: Description of the id parameter
        
    Returns:
        Description of the resource content
    """
    # Resource implementation
    return f"Resource content for {id}"
```

### Best Practices

1. **Type Hints**: Always use proper type hints for parameters and return values
2. **Docstrings**: Include clear docstrings for all tools and resources
3. **Error Handling**: Implement proper error handling in your tools and resources
4. **Testing**: Test your MCP server thoroughly before deployment

## License

MIT
