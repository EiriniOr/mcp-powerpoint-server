#!/bin/bash
# Activation script for PowerPoint MCP Server
cd /Users/rena/mcp-powerpoint-server
source venv/bin/activate
exec python server.py
