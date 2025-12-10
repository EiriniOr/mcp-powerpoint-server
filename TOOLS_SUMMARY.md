# PowerPoint MCP Server - Complete Tool Reference

## Overview
The PowerPoint MCP Server provides **36 tools** for comprehensive PowerPoint automation through the Model Context Protocol.

## Tool Categories

### Basic Management (7 tools)
1. `create_presentation` - Create new presentation with title slide
2. `open_presentation` - Open existing .pptx file
3. `save_presentation` - Save to disk
4. `list_presentations` - List presentations in memory
5. `add_title_slide` - Add title slide
6. `add_content_slide` - Add bullet point slide
7. `add_two_column_slide` - Add two-column layout

### Images (2 tools)
8. `add_image_slide` - Add image with layout options (centered, title+image, left, right)
9. `add_image_grid` - Add multiple images in grid layout with captions

### Tables & Data (3 tools)
10. `add_table_slide` - Add formatted table
11. `read_data_file` - Read and analyze CSV/Excel/JSON files
12. `analyze_and_chart` - Auto-analyze data file and create chart

### Basic Charts (1 tool)
13. `add_chart_slide` - Bar, column, line, pie, area charts

### Advanced Charts (2 tools)
14. `add_scatter_chart` - XY scatter plots
15. `add_bubble_chart` - 3D bubble charts

### Specialized Slides (2 tools)
16. `add_comparison_slide` - Side-by-side comparison
17. `add_timeline_slide` - Timeline visualization

### Shapes & Diagrams (3 tools)
18. `add_shape` - Add shapes (rectangle, circle, triangle, arrow, star, pentagon, hexagon)
19. `add_connector` - Add connector lines (straight, elbow, curved)
20. `add_flowchart` - Create automated flowcharts

### Interactive Elements (2 tools)
21. `add_hyperlink` - Add clickable links
22. `add_qr_code` - Generate and embed QR codes

### Text Formatting (1 tool)
23. `format_text` - Advanced text formatting (fonts, colors, sizes, bold, italic)

### Organization (2 tools)
24. `add_section` - Add section break slides
25. `add_agenda_slide` - Create table of contents

### Slide Operations (3 tools)
26. `duplicate_slide` - Duplicate existing slide
27. `delete_slide` - Remove slide
28. `merge_presentations` - Combine multiple presentations

### Backgrounds & Notes (2 tools)
29. `set_slide_background` - Set color or image background
30. `add_speaker_notes` - Add presenter notes

### Themes & Styling (2 tools)
31. `apply_theme` - Apply color themes (blue, red, green, purple, orange, professional, modern)
32. `add_footer` - Add footers with page numbers

### Export (1 tool)
33. `export_to_pdf` - Export to PDF (requires external tools)

## Total: 36 Tools

## Usage

### With Claude Code
After installing dependencies and restarting Claude Code:
```
"Create a presentation about AI with timeline, flowchart, and scatter chart"
```

### With ChatGPT
Use the Python wrapper:
```python
from chatgpt_wrapper import PowerPointAPI
# ChatGPT generates code using these methods
```

### Standalone
```python
from server import call_tool
await call_tool("create_presentation", {"title": "My Deck", "filename": "deck.pptx"})
```

## Supported File Formats
- **PowerPoint**: .pptx
- **Data Input**: CSV, Excel (.xlsx, .xls), JSON
- **Images**: PNG, JPG, JPEG, GIF
- **Export**: PDF (via external tools)

## Integration Options
- **Claude Code**: Native MCP integration (36 tools available)
- **ChatGPT**: Python wrapper API
- **Custom Scripts**: Direct Python API calls
- **CI/CD**: Automated presentation generation

## Dependencies
```
python-pptx>=0.6.21
pandas>=2.0.0
openpyxl>=3.1.0
mcp>=0.1.0
qrcode[pil]>=7.4.2
```
