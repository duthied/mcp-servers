#!/usr/bin/env python3
"""
Simple test to verify that the Firecrawl MCP server can be started correctly.
"""

import subprocess
import time
import os

def test_firecrawl_server():
    """Test if the Firecrawl MCP server can be started correctly."""
    print("Testing Firecrawl MCP server startup...")
    
    node_path = "/Users/devlon/.nvm/versions/node/v18.16.1/bin/node"
    firecrawl_path = "/Users/devlon/.nvm/versions/node/v18.16.1/lib/node_modules/firecrawl-mcp/dist/index.js"
    
    try:
        # Start the MCP server process
        process = subprocess.Popen(
            [node_path, firecrawl_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            env={
                "FIRECRAWL_API_KEY": "fc-36730e1f03874f038fd50f13e0851b4b",
                "PATH": "/Users/devlon/.nvm/versions/node/v18.16.1/bin:" + os.environ.get("PATH", "")
            }
        )
        
        # Wait a short time for the server to start
        time.sleep(2)
        
        # Check if the process is still running
        if process.poll() is None:
            print("Firecrawl MCP server started successfully!")
            
            # Read any stderr output
            stderr_output = process.stderr.read(1024)  # Read up to 1KB
            if stderr_output:
                print("\nServer output:")
                print(stderr_output)
            
            # Terminate the process
            print("\nTerminating server...")
            process.terminate()
            process.wait(timeout=5)
            print("Server terminated.")
            
            return True
        else:
            print(f"Process exited with code {process.returncode}")
            stderr_output = process.stderr.read()
            if stderr_output:
                print(f"Error output: {stderr_output}")
            return False
        
    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    success = test_firecrawl_server()
    
    if success:
        print("\nFirecrawl MCP server is properly configured and can be used by Claude.")
    else:
        print("\nThere was an issue with the Firecrawl MCP server configuration.")
