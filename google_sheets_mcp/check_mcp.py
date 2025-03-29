#!/usr/bin/env python3
"""
Test the Firecrawl MCP server by using the firecrawl_scrape tool.
"""

import json
import os
import subprocess
import sys
import time

def test_firecrawl_mcp():
    """Test the Firecrawl MCP server by scraping example.com."""
    print("Testing Firecrawl MCP server...")
    
    # Create a simple JSON request to scrape example.com
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "callTool",
        "params": {
            "name": "firecrawl_scrape",
            "arguments": {
                "url": "https://example.com",
                "formats": ["markdown"],
                "onlyMainContent": True
            }
        }
    }
    
    # Convert the request to a JSON string
    request_json = json.dumps(request)
    
    try:
        # Start the MCP server process
        print("Starting Firecrawl MCP server...")
        node_path = "/Users/devlon/.nvm/versions/node/v18.16.1/bin/node"
        firecrawl_path = "/Users/devlon/.nvm/versions/node/v18.16.1/lib/node_modules/firecrawl-mcp/dist/index.js"
        
        process = subprocess.Popen(
            [node_path, firecrawl_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            env={"FIRECRAWL_API_KEY": "fc-f18dd9d40a274fb2a264e1baf18e4f5e", "PATH": "/Users/devlon/.nvm/versions/node/v18.16.1/bin:" + os.environ.get("PATH", "")}
        )
        
        # Check if the process started successfully
        if process.poll() is not None:
            print(f"Process exited with code {process.returncode}")
            stderr_output = process.stderr.read()
            if stderr_output:
                print(f"Error output: {stderr_output}")
            return False
        
        # Send the request to the server
        print("Sending request to scrape example.com...")
        process.stdin.write(request_json + "\n")
        process.stdin.flush()
        
        # Print stderr in real-time for debugging
        print("Checking for error output...")
        stderr_output = ""
        while process.poll() is None:
            line = process.stderr.readline()
            if line:
                stderr_output += line
                print(f"Server stderr: {line.strip()}")
            else:
                break
        
        # Wait for a response (with timeout)
        start_time = time.time()
        timeout = 30  # seconds
        response = ""
        
        while time.time() - start_time < timeout:
            line = process.stdout.readline()
            if line:
                response = line
                break
            time.sleep(0.1)
        
        # Process the response
        if response:
            try:
                response_json = json.loads(response)
                print("\nResponse received:")
                print(json.dumps(response_json, indent=2))
                
                # Check if the response contains content
                if "result" in response_json and "content" in response_json["result"]:
                    print("\nFirecrawl MCP server is working correctly!")
                else:
                    print("\nResponse received but no content found.")
            except json.JSONDecodeError:
                print(f"\nReceived non-JSON response: {response}")
        else:
            print("\nNo response received within timeout period.")
        
        # Terminate the process
        process.terminate()
        
    except Exception as e:
        print(f"Error: {e}")
        return False
    
    return True

if __name__ == "__main__":
    # Check if modelcontextprotocol package is installed
    try:
        import modelcontextprotocol
        print("modelcontextprotocol package is installed.")
        print(f"Version: {modelcontextprotocol.__version__ if hasattr(modelcontextprotocol, '__version__') else 'Unknown'}")
        print(f"Path: {modelcontextprotocol.__file__}")
    except ImportError:
        print("modelcontextprotocol package is not installed.")
    
    # Test the Firecrawl MCP server
    print("\n" + "="*50 + "\n")
    success = test_firecrawl_mcp()
    
    if not success:
        sys.exit(1)
