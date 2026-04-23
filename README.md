# PowerPoint MCP Server

MCP server for creating and editing PowerPoint presentations via `python-pptx`.

## Setup

```bash
cd /Users/rena/mcp-powerpoint-server
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Claude Code Integration

Add to `~/.claude/settings.json` (global) or `.claude/settings.json` (project):

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "/Users/rena/mcp-powerpoint-server/venv/bin/python",
      "args": ["/Users/rena/mcp-powerpoint-server/server.py"],
      "type": "stdio"
    }
  }
}
```

Then restart Claude Code. The server exposes tools for:

- **Slides**: create, open, save, delete, duplicate, merge presentations
- **Content**: title, bullet, two-column, comparison, timeline, agenda slides
- **Media**: images, image grids, QR codes, hyperlinks
- **Charts**: bar, column, line, pie, area, scatter, bubble — or from CSV/JSON/Excel
- **Tables**: with styled headers
- **Shapes**: rectangles, circles, arrows, flowcharts, connectors
- **Formatting**: text styling, slide backgrounds, speaker notes, footers, themes
