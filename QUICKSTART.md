# Quick Start Guide

## Overview

This PowerPoint MCP server provides automated presentation generation with data analysis capabilities.

## Configuration Status

### 1. Claude Code Integration
**Config**: `~/.config/claude-code/mcp_config.json`

**Activation**:
- Restart Claude Code
- All 36 PowerPoint tools will be available
- Use natural language commands

### 2. Test Results
**Location**: `/Users/rena/Downloads/test_demo.pptx`

Test presentation includes:
- Title slide
- Content slide with bullet points
- Comparison slide
- Chart slide
- Timeline slide
- Table slide

Open it to see what the server can do!

## 🚀 Usage Examples

### With Claude Code (This CLI)
After restarting, just ask me:
- "Create a sales presentation with Q4 data"
- "Make a timeline showing our project milestones"
- "Add a comparison slide between Product A and B"
- "Create charts from my sales.csv file"

### With ChatGPT
See [CHATGPT_SETUP.md](CHATGPT_SETUP.md) for integration options:
- Python wrapper (easiest)
- HTTP API bridge
- LangChain integration
- Custom GPT actions

### Standalone Testing
```bash
cd /Users/rena/mcp-powerpoint-server
python test_mcp.py
```

## 📚 Available Tools (36 total)

**Basic Management** (7 tools)
- create_presentation, open_presentation, save_presentation
- list_presentations, add_title_slide, add_content_slide
- add_two_column_slide

**Charts** (3 tools)
- add_chart_slide (bar, column, line, pie, area)
- add_scatter_chart, add_bubble_chart

**Data Analysis** (2 tools)
- analyze_and_chart (Auto-analyze CSV/Excel/JSON)
- read_data_file

**Images** (2 tools)
- add_image_slide, add_image_grid

**Tables & Specialized** (4 tools)
- add_table_slide, add_comparison_slide
- add_timeline_slide, format_text

**Shapes & Diagrams** (3 tools)
- add_shape, add_connector, add_flowchart

**Interactive** (2 tools)
- add_hyperlink, add_qr_code

**Organization** (2 tools)
- add_section, add_agenda_slide

**Operations** (3 tools)
- duplicate_slide, delete_slide, merge_presentations

**Styling** (5 tools)
- set_slide_background, add_speaker_notes
- apply_theme, add_footer, export_to_pdf

## 📖 Full Documentation

- [README.md](README.md) - Complete API reference
- [SETUP.md](SETUP.md) - Detailed setup instructions
- [CHATGPT_SETUP.md](CHATGPT_SETUP.md) - ChatGPT integration guide
- [test_mcp.py](test_mcp.py) - Test script

## 🎓 Example Workflows

### 1. Data-Driven Presentation
```python
# Analyze CSV and create presentation
python -c "
import asyncio
from server import call_tool

async def create():
    await call_tool('create_presentation', {
        'title': 'Sales Analysis',
        'filename': 'sales.pptx'
    })

    await call_tool('analyze_and_chart', {
        'filename': 'sales.pptx',
        'data_file': 'sales.csv',
        'chart_type': 'column',
        'x_column': 'Month',
        'y_columns': ['Revenue', 'Profit']
    })

    await call_tool('save_presentation', {
        'filename': 'sales.pptx'
    })

asyncio.run(create())
"
```

### 2. With Claude Code
Just restart and ask:
```
"Create a product launch presentation with:
- Title slide
- Problem/solution comparison
- Feature timeline
- Pricing table
- Thank you slide"
```

### 3. Batch Processing
Create multiple presentations:
```bash
for topic in "AI" "ML" "Data Science"; do
    python -c "
import asyncio
from server import call_tool

async def create():
    await call_tool('create_presentation', {
        'title': '$topic Overview',
        'filename': '${topic}.pptx'
    })
    await call_tool('save_presentation', {'filename': '${topic}.pptx'})

asyncio.run(create())
"
done
```

## 🔧 Troubleshooting

**Server not responding?**
```bash
cd /Users/rena/mcp-powerpoint-server
python -c "from server import app; print('✓ Server OK')"
```

**Tools not available in Claude Code?**
1. Check config: `cat ~/.config/claude-code/mcp_config.json`
2. Restart Claude Code
3. Try: "List available tools"

**Python errors?**
```bash
pip install -r requirements.txt
```

## 🎉 Next Steps

1. **Restart Claude Code** to activate the MCP server
2. **Try it out**: Ask me to create a presentation
3. **Experiment**: Try different slide types and formats
4. **Integrate**: Follow [CHATGPT_SETUP.md](CHATGPT_SETUP.md) for other LLMs

## 📞 Getting Help

- Run test: `python test_mcp.py`
- Check logs: Server prints to console
- Review docs: See README.md for full API

---

**Server Location**: `/Users/rena/mcp-powerpoint-server/server.py`
**Default Output**: `~/Downloads/`
**Test File**: `/Users/rena/Downloads/test_demo.pptx`

**Status**: 🟢 All Systems Go!
