# PowerPoint MCP Server

A comprehensive MCP (Model Context Protocol) server for creating and managing PowerPoint presentations with advanced features including charts, tables, images, custom formatting, and data analysis.

## Features

### Core Functionality
- **Create and manage presentations** - Create new or open existing PowerPoint files
- **Multiple slide layouts** - Title slides, content slides, two-column layouts
- **Save presentations** - Save to local disk

### Advanced Features
- **Images** - Add images with multiple layout options (centered, title+image, left/right positioned)
- **Tables** - Create formatted tables with headers and data rows
- **Charts & Graphs** - Bar, column, line, pie, and area charts
- **Data Analysis** - Automatically analyze CSV, JSON, and Excel files and create charts
- **Custom Text Formatting** - Control fonts, sizes, colors, bold, italic
- **Specialized Slides** - Comparison slides, timeline slides
- **Backgrounds** - Set solid colors or image backgrounds
- **Speaker Notes** - Add presenter notes to slides

## Installation

```bash
pip install -r requirements.txt
```

## Available Tools

### Basic Presentation Management

#### `create_presentation`
Creates a new PowerPoint presentation with a title slide.

**Parameters:**
- `title` (string, required): Title for the presentation
- `subtitle` (string, optional): Subtitle for the title slide
- `filename` (string, required): Filename to save as (e.g., 'presentation.pptx')

#### `open_presentation`
Opens an existing PowerPoint presentation from disk.

**Parameters:**
- `file_path` (string, required): Path to the existing PowerPoint file
- `filename` (string, optional): Internal name to reference this presentation

#### `save_presentation`
Saves the presentation to disk.

**Parameters:**
- `filename` (string, required): The presentation filename
- `output_path` (string, optional): Full path where to save (defaults to ~/Downloads)

#### `list_presentations`
Lists all presentations currently in memory.

### Slide Creation

#### `add_title_slide`
Adds a title slide to an existing presentation.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `subtitle` (string, optional): Slide subtitle

#### `add_content_slide`
Adds a content slide with title and bullet points.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `content` (array of strings, required): List of bullet points

#### `add_two_column_slide`
Adds a slide with two columns of content.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `left_content` (array of strings, required): Content for left column
- `right_content` (array of strings, required): Content for right column

### Images

#### `add_image_slide`
Adds a slide with an image and optional title/caption.

**Parameters:**
- `filename` (string, required): The presentation filename
- `image_path` (string, required): Path to the image file
- `title` (string, optional): Slide title
- `caption` (string, optional): Image caption
- `layout` (string, optional): Image layout style - `"centered"` (default), `"title_and_image"`, `"image_left"`, or `"image_right"`

### Tables

#### `add_table_slide`
Adds a slide with a table.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `headers` (array of strings, required): Table column headers
- `rows` (array of arrays, required): Table rows

### Charts & Data Analysis

#### `add_chart_slide`
Adds a slide with a chart/graph.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `chart_type` (string, required): Type of chart - `"bar"`, `"column"`, `"line"`, `"pie"`, or `"area"`
- `categories` (array of strings, required): Chart categories (x-axis labels)
- `series` (array of objects, required): Chart data series (each with `name` and `values`)

#### `analyze_and_chart`
Analyzes a data file (CSV, JSON, Excel) and creates a chart slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `data_file` (string, required): Path to data file
- `chart_type` (string, required): Type of chart
- `x_column` (string, required): Column name for x-axis
- `y_columns` (array of strings, required): Column name(s) for y-axis
- `title` (string, optional): Slide title (auto-generated if not provided)

#### `read_data_file`
Reads and analyzes a data file, returning summary statistics.

**Parameters:**
- `data_file` (string, required): Path to data file
- `sheet_name` (string, optional): Sheet name for Excel files

### Specialized Slides

#### `add_comparison_slide`
Adds a comparison slide with two items side-by-side.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `left_title` (string, required): Title for left side
- `left_content` (array of strings, required): Left side content
- `right_title` (string, required): Title for right side
- `right_content` (array of strings, required): Right side content

#### `add_timeline_slide`
Adds a timeline slide showing events chronologically.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `events` (array of objects, required): Timeline events (each with `date` and `event`)

### Text Formatting

#### `format_text`
Adds a text slide with advanced formatting options.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `text_blocks` (array of objects, required): Text blocks with formatting options (`text`, `font_size`, `bold`, `italic`, `color`, `font_name`)

### Customization

#### `set_slide_background`
Sets the background color or image for a slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, optional): Slide index (0-based, -1 for last slide)
- `color` (string, optional): Hex color code for solid color background
- `image_path` (string, optional): Path to background image

#### `add_speaker_notes`
Adds speaker notes to a slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, optional): Slide index (0-based, -1 for last slide)
- `notes` (string, required): Speaker notes text

## Supported Data File Formats

- **CSV** (.csv) - Comma-separated values
- **Excel** (.xlsx, .xls) - Microsoft Excel files
- **JSON** (.json) - JSON format with tabular structure

## Running the Server

```bash
python server.py
```

The server runs using stdio transport, making it compatible with any MCP client.

## Dependencies

- `python-pptx` - PowerPoint file creation and manipulation
- `pandas` - Data analysis and file reading
- `openpyxl` - Excel file support
- `mcp` - Model Context Protocol framework
