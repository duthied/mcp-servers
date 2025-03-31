#!/bin/bash
# run-server.sh - Launcher script for Google Sheets MCP server

# Set working directory to script location
cd "$(dirname "$0")"

# Activate the virtual environment
source .venv/bin/activate

# Run the server
python server.py
