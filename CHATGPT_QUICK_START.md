# ChatGPT Quick Start Guide

## Using the Python Wrapper

A simple Python API is available that ChatGPT (or any LLM) can use to generate presentation code.

## Setup

Dependencies are already installed. Test the wrapper:

```bash
cd /Users/rena/mcp-powerpoint-server
python chatgpt_wrapper.py test
```

This will create 3 example presentations to verify everything works.

## Using with ChatGPT

### Method 1: Copy-Paste Code (Simplest)

1. Copy one of the examples from `chatgpt_wrapper.py`
2. Paste it into ChatGPT
3. Ask ChatGPT to modify it for your needs

**Example conversation:**

```
You: I need to create a sales presentation with Q4 data

ChatGPT: I'll help you create that. Here's the code:

[ChatGPT provides modified code]

You: [Run the code on your computer]
```

### Method 2: Iterative Development

Ask ChatGPT to write Python code using the API:

```
You: "Using the PowerPointAPI class from chatgpt_wrapper.py,
     create a presentation about AI trends with:
     - Title slide
     - 3 content slides
     - A timeline
     - Save it"

ChatGPT: [Provides complete code]

You: [Copy and run it]
```

### Method 3: Data Analysis

If you have a CSV file:

```
You: "I have a file sales_data.csv with columns Month, Revenue, Profit.
     Create a presentation with charts showing the trends."

ChatGPT: [Provides code using analyze_data_and_chart()]

You: [Run it]
```

## Quick Examples

### Example 1: Simple Presentation

Ask ChatGPT:
```
"Write Python code using PowerPointAPI to create a presentation
about space exploration with 3 slides"
```

ChatGPT will provide:
```python
import asyncio
from chatgpt_wrapper import PowerPointAPI

async def create_space_presentation():
    ppt = PowerPointAPI()

    await ppt.create_presentation(
        title="Space Exploration",
        subtitle="The Final Frontier",
        filename="space.pptx"
    )

    await ppt.add_content_slide(
        filename="space.pptx",
        title="Why Space?",
        content=[
            "Scientific discovery",
            "Resource exploration",
            "Human curiosity"
        ]
    )

    # ... more slides ...

    await ppt.save_presentation("space.pptx")

asyncio.run(create_space_presentation())
```

### Example 2: Data-Driven Presentation

Ask ChatGPT:
```
"Create a presentation analyzing quarterly_sales.csv"
```

ChatGPT will provide:
```python
import asyncio
from chatgpt_wrapper import PowerPointAPI

async def analyze_sales():
    ppt = PowerPointAPI()

    # Read data first
    data_summary = await ppt.read_data_file("quarterly_sales.csv")
    print(data_summary)

    # Create presentation
    await ppt.create_presentation(
        title="Sales Analysis",
        filename="sales_analysis.pptx"
    )

    # Auto-create charts from data
    await ppt.analyze_data_and_chart(
        filename="sales_analysis.pptx",
        data_file="quarterly_sales.csv",
        chart_type="column",
        x_column="Quarter",
        y_columns=["Revenue", "Profit"]
    )

    await ppt.save_presentation("sales_analysis.pptx")

asyncio.run(analyze_sales())
```

## Available Functions

### Basic
- `create_presentation(title, filename, subtitle="")`
- `save_presentation(filename, output_path=None)`

### Content
- `add_content_slide(filename, title, content)` - Bullet points
- `add_chart(filename, title, chart_type, categories, series)` - Charts
- `add_table(filename, title, headers, rows)` - Tables
- `add_comparison(filename, title, left_title, left_content, right_title, right_content)`
- `add_timeline(filename, title, events)`
- `add_image(filename, image_path, title, caption, layout)`

### Data Analysis
- `analyze_data_and_chart(filename, data_file, chart_type, x_column, y_columns)`
- `read_data_file(data_file, sheet_name=None)`

## Tips for ChatGPT Conversations

### Good Prompts:
✅ "Create a presentation about [topic] with [specific slides]"
✅ "Analyze my data.csv file and create charts"
✅ "Make a timeline showing [events]"
✅ "Create a comparison between [A] and [B]"

### Better Prompts:
✅ "Using PowerPointAPI, create a 5-slide presentation about AI with a timeline and comparison slide"
✅ "Read my sales.csv and create a presentation with column charts for Revenue and Profit by Month"

## Testing the Examples

Try these commands:

```bash
cd /Users/rena/mcp-powerpoint-server

# Run all examples
python chatgpt_wrapper.py test

# Run specific example
python chatgpt_wrapper.py simple
python chatgpt_wrapper.py charts
python chatgpt_wrapper.py comprehensive
```

## Common Workflows

### Workflow 1: Quick Presentation
1. Ask ChatGPT: "Create a presentation about [topic]"
2. ChatGPT provides code
3. Save code to `my_presentation.py`
4. Run: `python my_presentation.py`
5. Find .pptx in Downloads folder

### Workflow 2: Data Analysis
1. Ask ChatGPT: "Analyze my data.csv and create presentation"
2. ChatGPT asks what columns to use
3. You specify columns
4. ChatGPT provides complete code
5. Run it

### Workflow 3: Custom Design
1. Ask ChatGPT: "Create presentation with timeline, comparison, and charts"
2. Review code
3. Ask for modifications: "Make the timeline bigger" or "Add more data points"
4. ChatGPT updates code
5. Run final version

## Real-World Example

**You:** "I have quarterly_revenue.csv with columns: Quarter, ProductA, ProductB, ProductC. Create a presentation showing the trends."

**ChatGPT provides:**
```python
import asyncio
from chatgpt_wrapper import PowerPointAPI

async def revenue_presentation():
    ppt = PowerPointAPI()

    # Create title
    await ppt.create_presentation(
        title="Quarterly Revenue Analysis",
        subtitle="Product Performance Review",
        filename="revenue.pptx"
    )

    # Overview slide
    await ppt.add_content_slide(
        filename="revenue.pptx",
        title="Overview",
        content=[
            "Analyzing Q1-Q4 performance",
            "Three product lines",
            "Identifying growth trends"
        ]
    )

    # Auto-analyze and create chart
    await ppt.analyze_data_and_chart(
        filename="revenue.pptx",
        data_file="quarterly_revenue.csv",
        chart_type="line",
        x_column="Quarter",
        y_columns=["ProductA", "ProductB", "ProductC"],
        title="Product Revenue Trends"
    )

    # Add comparison
    await ppt.add_comparison(
        filename="revenue.pptx",
        title="Top vs Bottom Performers",
        left_title="Product A (Best)",
        left_content=["Consistent growth", "High margins", "Strong demand"],
        right_title="Product C (Needs Attention)",
        right_content=["Declining sales", "Price pressure", "Competition"]
    )

    # Save
    await ppt.save_presentation("revenue.pptx")
    print("✅ Presentation created: ~/Downloads/revenue.pptx")

asyncio.run(revenue_presentation())
```

**You:** [Save and run the code]

Done! Your presentation is ready in ~/Downloads/

## Advantages of This Approach

✅ **Simple** - Just Python code, no complex setup
✅ **Flexible** - ChatGPT can generate any presentation logic
✅ **Iterative** - Easy to modify and re-run
✅ **Powerful** - Full access to all 20 MCP tools
✅ **Data-Driven** - Automatically analyze CSV/Excel files

## Next Steps

1. Try the test examples: `python chatgpt_wrapper.py test`
2. Open ChatGPT and ask it to create a presentation
3. Run the code it provides
4. Iterate and improve!

## Troubleshooting

**Import errors?**
```bash
cd /Users/rena/mcp-powerpoint-server
python -c "from chatgpt_wrapper import PowerPointAPI; print('✓ OK')"
```

**ChatGPT confused?**
- Show it the `chatgpt_wrapper.py` file
- Point to specific functions you want to use
- Provide concrete examples

**Code not working?**
- Make sure you're in the right directory
- Check file paths are absolute
- Run test first: `python chatgpt_wrapper.py test`

---

**File Location**: `/Users/rena/mcp-powerpoint-server/chatgpt_wrapper.py`
**Test Command**: `python chatgpt_wrapper.py test`
**Output Location**: `~/Downloads/`

Happy presenting! 🎉
