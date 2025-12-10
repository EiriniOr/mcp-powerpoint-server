# Setup Guide for PowerPoint MCP Server

## Quick Start

Your PowerPoint MCP server is ready to use! Here's how to set it up with different clients.

## Option 1: Claude Desktop (Recommended)

1. **Find your Claude Desktop config file:**
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

2. **Edit the config file and add:**

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["/Users/rena/mcp-powerpoint-server/server.py"]
    }
  }
}
```

3. **Restart Claude Desktop**

4. **Test it by asking:**
   - "Create a presentation about AI trends"
   - "Add a chart showing sales data"
   - "Create a timeline for our project"

## Option 2: Claude Code (CLI)

1. **Create or edit:** `~/.config/claude-code/mcp_config.json`

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "python",
      "args": ["/Users/rena/mcp-powerpoint-server/server.py"]
    }
  }
}
```

2. **Restart Claude Code**

3. **All PowerPoint tools will be available automatically**

## Option 3: Manual Testing

Test the server directly from command line:

```bash
cd /Users/rena/mcp-powerpoint-server
python server.py
```

The server will wait for MCP protocol messages on stdin.

## What You Can Do

Once set up, you can ask Claude to:

- **Create presentations**: "Create a presentation about machine learning"
- **Add slides**: "Add a timeline slide showing our Q4 roadmap"
- **Analyze data**: "Read sales.csv and create a bar chart"
- **Format content**: "Add a comparison slide between Product A and Product B"
- **Customize**: "Set a blue background on the last slide"
- **And much more!**

## Available Tools

Your server has 36 powerful tools:

### Basic Management (7)
- `create_presentation` - Start a new presentation
- `open_presentation` - Open existing .pptx files
- `save_presentation` - Save to disk
- `list_presentations` - See what's in memory
- `add_title_slide` - Title slides
- `add_content_slide` - Bullet point slides
- `add_two_column_slide` - Two-column layouts

### Images (2)
- `add_image_slide` - Images with various layouts
- `add_image_grid` - Multiple images in grid layout

### Tables & Data (3)
- `add_table_slide` - Formatted tables
- `read_data_file` - Get statistics from data files
- `analyze_and_chart` - Auto-analyze CSV/Excel/JSON and create charts

### Charts (3)
- `add_chart_slide` - Bar, column, line, pie, area charts
- `add_scatter_chart` - XY scatter plots
- `add_bubble_chart` - 3D bubble charts

### Specialized Slides (2)
- `add_comparison_slide` - Side-by-side comparisons
- `add_timeline_slide` - Visual timelines

### Shapes & Diagrams (3)
- `add_shape` - Rectangle, circle, triangle, arrow, star, pentagon, hexagon
- `add_connector` - Straight, elbow, curved connector lines
- `add_flowchart` - Automated flowcharts

### Interactive Elements (2)
- `add_hyperlink` - Clickable links
- `add_qr_code` - QR code generation

### Text Formatting (1)
- `format_text` - Custom text formatting

### Organization (2)
- `add_section` - Section break slides
- `add_agenda_slide` - Table of contents

### Slide Operations (3)
- `duplicate_slide` - Copy slides
- `delete_slide` - Remove slides
- `merge_presentations` - Combine presentations

### Customization (5)
- `set_slide_background` - Custom backgrounds
- `add_speaker_notes` - Presenter notes
- `apply_theme` - Color themes (blue, red, green, purple, orange, professional, modern)
- `add_footer` - Footers with page numbers
- `export_to_pdf` - Export to PDF format

## Example Usage

### Example 1: Quick Presentation
```
User: "Create a presentation about AI Ethics with 3 slides"
Claude: [Creates presentation with title slide and 2 content slides]
```

### Example 2: Data-Driven
```
User: "Analyze sales_data.csv and create a presentation with charts"
Claude: [Reads CSV, analyzes data, creates charts automatically]
```

### Example 3: Professional Deck
```
User: "Create a product launch presentation with:
- Title slide
- Problem/Solution comparison
- Timeline for Q1-Q4
- Sales forecast chart
- Closing slide"
Claude: [Creates comprehensive presentation with all elements]
```

## Troubleshooting

**Server won't start:**
- Check Python is installed: `python --version`
- Verify dependencies: `pip install -r requirements.txt`

**Tools not appearing:**
- Restart Claude Desktop/Code after config changes
- Check config file has correct path
- Verify JSON syntax is valid

**Charts/data analysis failing:**
- Ensure pandas and openpyxl are installed
- Check data file path is absolute
- Verify column names match your data

## File Locations

- **Server**: `/Users/rena/mcp-powerpoint-server/server.py`
- **Default save location**: `~/Downloads/`
- **Config examples**: This file

## Next Steps

1. Set up with Claude Desktop or Claude Code
2. Try creating a simple presentation
3. Test data analysis features with a CSV file
4. Experiment with different slide types
5. Build amazing presentations!

Enjoy your enhanced PowerPoint capabilities! 🎉
