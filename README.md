# PowerPoint MCP Server

A comprehensive MCP (Model Context Protocol) server for creating and managing PowerPoint presentations. Provides 36 powerful tools for automation including charts, tables, images, shapes, flowcharts, QR codes, advanced formatting, and data analysis.

## Features

### Core Functionality
- **Create and manage presentations** - Create new or open existing PowerPoint files
- **Multiple slide layouts** - Title slides, content slides, two-column layouts
- **Save presentations** - Save to local disk

### Advanced Features
- **Images** - Add images with multiple layout options (centered, title+image, left/right positioned) and image grids
- **Tables** - Create formatted tables with headers and data rows
- **Charts & Graphs** - Bar, column, line, pie, area, scatter, and bubble charts
- **Data Analysis** - Automatically analyze CSV, JSON, and Excel files and create charts
- **Shapes & Diagrams** - Add shapes (rectangle, circle, triangle, arrow, star, pentagon, hexagon), connectors, and automated flowcharts
- **Interactive Elements** - Add hyperlinks and QR codes
- **Custom Text Formatting** - Control fonts, sizes, colors, bold, italic
- **Specialized Slides** - Comparison slides, timeline slides, agenda slides, section breaks
- **Slide Operations** - Duplicate, delete, and merge presentations
- **Backgrounds** - Set solid colors or image backgrounds
- **Themes & Styling** - Apply color themes and add footers with page numbers
- **Speaker Notes** - Add presenter notes to slides
- **Export** - Export presentations to PDF format

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

### Advanced Charts

#### `add_scatter_chart`
Adds a scatter (XY) chart slide for visualizing relationships between two variables.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `series` (array of objects, required): Data series (each with `name`, `x_values`, and `y_values`)

#### `add_bubble_chart`
Adds a bubble chart slide for three-dimensional data visualization.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `series` (array of objects, required): Data series (each with `name` and `data_points` containing `x`, `y`, `size`)

### Shapes & Diagrams

#### `add_shape`
Adds a shape to a slide with optional fill color and text.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, optional): Slide index (-1 for last slide)
- `shape_type` (string, required): Shape type - `"rectangle"`, `"circle"`, `"triangle"`, `"arrow"`, `"star"`, `"pentagon"`, or `"hexagon"`
- `left` (number, required): Left position in inches
- `top` (number, required): Top position in inches
- `width` (number, required): Width in inches
- `height` (number, required): Height in inches
- `fill_color` (string, optional): Hex color code (e.g., "#FF0000")
- `text` (string, optional): Text to add inside the shape

#### `add_connector`
Adds a connector line between two points on a slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, optional): Slide index (-1 for last slide)
- `connector_type` (string, required): Connector type - `"straight"`, `"elbow"`, or `"curved"`
- `start_x` (number, required): Start X position in inches
- `start_y` (number, required): Start Y position in inches
- `end_x` (number, required): End X position in inches
- `end_y` (number, required): End Y position in inches

#### `add_flowchart`
Creates an automated flowchart with shapes and connectors.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `steps` (array of objects, required): Flowchart steps (each with `text` and optional `shape_type`)

### Multi-Image Layouts

#### `add_image_grid`
Adds a slide with multiple images arranged in a grid layout.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title
- `images` (array of objects, required): Images with `path` and optional `caption`
- `columns` (number, optional): Number of columns (default: 2)

### Interactive Elements

#### `add_hyperlink`
Adds a hyperlink to text or shape on a slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, optional): Slide index (-1 for last slide)
- `text` (string, required): Link text to display
- `url` (string, required): URL to link to
- `left` (number, required): Left position in inches
- `top` (number, required): Top position in inches

#### `add_qr_code`
Generates and adds a QR code to a slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, optional): Slide index (-1 for last slide)
- `data` (string, required): Data to encode (URL, text, etc.)
- `left` (number, required): Left position in inches
- `top` (number, required): Top position in inches
- `size` (number, optional): QR code size in inches (default: 2.0)

### Organization

#### `add_section`
Adds a section break slide to organize presentation into parts.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Section title

#### `add_agenda_slide`
Creates an agenda/table of contents slide.

**Parameters:**
- `filename` (string, required): The presentation filename
- `title` (string, required): Slide title (e.g., "Agenda")
- `items` (array of strings, required): Agenda items

### Slide Operations

#### `duplicate_slide`
Duplicates an existing slide in the presentation.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, required): Index of slide to duplicate (0-based)

#### `delete_slide`
Deletes a slide from the presentation.

**Parameters:**
- `filename` (string, required): The presentation filename
- `slide_index` (number, required): Index of slide to delete (0-based)

#### `merge_presentations`
Merges multiple presentations into one.

**Parameters:**
- `filename` (string, required): Target presentation filename
- `source_files` (array of strings, required): Paths to presentations to merge

### Themes & Export

#### `apply_theme`
Applies a color theme to the presentation.

**Parameters:**
- `filename` (string, required): The presentation filename
- `theme_name` (string, required): Theme name - `"blue"`, `"red"`, `"green"`, `"purple"`, `"orange"`, `"professional"`, or `"modern"`

#### `add_footer`
Adds footer with text and page numbers to slides.

**Parameters:**
- `filename` (string, required): The presentation filename
- `text` (string, optional): Footer text
- `show_page_number` (boolean, optional): Show page numbers (default: true)
- `show_on_title_slide` (boolean, optional): Show on title slide (default: false)

#### `export_to_pdf`
Exports the presentation to PDF format (requires LibreOffice or PowerPoint).

**Parameters:**
- `filename` (string, required): The presentation filename
- `output_path` (string, optional): PDF output path

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
- `qrcode[pil]` - QR code generation
