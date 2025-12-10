#!/usr/bin/env python3
"""
Demo: Life in Sweden Presentation
Showcases the PowerPoint MCP Server's 36 tools
"""

import asyncio
import sys
import os

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from server import presentations, Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData

async def create_sweden_demo():
    """Create a comprehensive demo presentation about Life in Sweden"""

    filename = "life_in_sweden_demo.pptx"

    print("Creating 'Life in Sweden' demo presentation...")
    print("This showcases the PowerPoint MCP Server's capabilities\n")

    # 1. Create presentation with title slide
    print("1. Creating title slide...")
    prs = Presentation()
    presentations[filename] = prs

    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = title_slide.shapes.title
    subtitle = title_slide.placeholders[1]
    title.text = "Life in Sweden"
    subtitle.text = "A Nordic Adventure\nDemo created with PowerPoint MCP Server"

    # 2. Add agenda slide
    print("2. Adding agenda slide...")
    agenda_items = [
        "Swedish Culture & Traditions",
        "Cost of Living Comparison",
        "Climate Throughout the Year",
        "Popular Activities & Attractions",
        "Work-Life Balance Statistics"
    ]

    bullet_slide_layout = prs.slide_layouts[1]
    agenda_slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = agenda_slide.shapes
    title_shape = shapes.title
    body_shape = shapes.placeholders[1]

    title_shape.text = "Agenda"
    tf = body_shape.text_frame
    for i, item in enumerate(agenda_items):
        if i == 0:
            tf.text = item
        else:
            p = tf.add_paragraph()
            p.text = item
            p.level = 0

    # 3. Add section break
    print("3. Adding section break...")
    blank_slide_layout = prs.slide_layouts[6]
    section_slide = prs.slides.add_slide(blank_slide_layout)
    section_slide.background.fill.solid()
    section_slide.background.fill.fore_color.rgb = RGBColor(0, 102, 204)  # Swedish blue

    title_box = section_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "Swedish Culture"
    title_frame.paragraphs[0].font.size = Pt(54)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 215, 0)  # Swedish yellow
    from pptx.enum.text import PP_ALIGN
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 4. Add comparison slide: Sweden vs Other Countries
    print("4. Adding comparison slide...")
    comparison_slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = comparison_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
    title_frame = title_box.text_frame
    title_frame.text = "Sweden vs European Neighbors"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True

    # Left side - Sweden
    left_box = comparison_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4), Inches(5))
    left_box.fill.solid()
    left_box.fill.fore_color.rgb = RGBColor(0, 102, 204)
    left_frame = left_box.text_frame
    left_frame.text = "Sweden 🇸🇪"
    left_frame.paragraphs[0].font.size = Pt(24)
    left_frame.paragraphs[0].font.bold = True
    left_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    left_items = [
        "Population: ~10.5M",
        "Work hours: 40h/week",
        "Vacation days: 25+",
        "Parental leave: 480 days",
        "Fika culture ☕"
    ]
    for item in left_items:
        p = left_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.level = 0

    # Right side - Other countries
    right_box = comparison_slide.shapes.add_textbox(Inches(5.5), Inches(1.5), Inches(4), Inches(5))
    right_box.fill.solid()
    right_box.fill.fore_color.rgb = RGBColor(100, 100, 100)
    right_frame = right_box.text_frame
    right_frame.text = "EU Average"
    right_frame.paragraphs[0].font.size = Pt(24)
    right_frame.paragraphs[0].font.bold = True
    right_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    right_items = [
        "Varies widely",
        "Work hours: 35-45h/week",
        "Vacation days: 20-25",
        "Parental leave: varies",
        "Less structured breaks"
    ]
    for item in right_items:
        p = right_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.level = 0

    # 5. Add chart: Cost of Living
    print("5. Adding cost of living chart...")
    chart_slide = prs.slides.add_slide(blank_slide_layout)

    title_box = chart_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
    title_frame = title_box.text_frame
    title_frame.text = "Average Monthly Costs (SEK)"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True

    chart_data = CategoryChartData()
    chart_data.categories = ['Rent (1BR)', 'Groceries', 'Transport', 'Entertainment', 'Internet']
    chart_data.add_series('Stockholm', (12000, 3500, 950, 1500, 350))
    chart_data.add_series('Smaller Cities', (8000, 3000, 800, 1200, 300))

    x, y, cx, cy = Inches(1), Inches(1.5), Inches(8), Inches(5)
    chart = chart_slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    ).chart
    chart.has_legend = True
    from pptx.enum.chart import XL_LEGEND_POSITION
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

    # 6. Add timeline: Year in Sweden
    print("6. Adding timeline slide...")
    timeline_slide = prs.slides.add_slide(blank_slide_layout)

    title_box = timeline_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
    title_frame = title_box.text_frame
    title_frame.text = "A Year in Sweden"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True

    # Create timeline
    events = [
        {"month": "Jan-Feb", "event": "❄️ Polar nights & skiing"},
        {"month": "Mar-Apr", "event": "🌷 Spring awakening"},
        {"month": "May-Jun", "event": "🌞 Midnight sun begins"},
        {"month": "Jul-Aug", "event": "☀️ Summer holidays & festivals"},
        {"month": "Sep-Oct", "event": "🍂 Autumn colors & coziness"},
        {"month": "Nov-Dec", "event": "🎄 Christmas markets & lights"}
    ]

    y_start = 2.0
    for i, evt in enumerate(events):
        # Month box
        month_box = timeline_slide.shapes.add_textbox(
            Inches(0.5), Inches(y_start + i * 0.7), Inches(2), Inches(0.5)
        )
        month_box.fill.solid()
        month_box.fill.fore_color.rgb = RGBColor(0, 102, 204)
        month_frame = month_box.text_frame
        month_frame.text = evt["month"]
        month_frame.paragraphs[0].font.size = Pt(14)
        month_frame.paragraphs[0].font.bold = True
        month_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        month_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Event text
        event_box = timeline_slide.shapes.add_textbox(
            Inches(3), Inches(y_start + i * 0.7), Inches(6), Inches(0.5)
        )
        event_frame = event_box.text_frame
        event_frame.text = evt["event"]
        event_frame.paragraphs[0].font.size = Pt(14)

    # 7. Add table: Top Cities
    print("7. Adding table slide...")
    table_slide = prs.slides.add_slide(blank_slide_layout)

    title_box = table_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
    title_frame = title_box.text_frame
    title_frame.text = "Top Swedish Cities to Live In"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True

    rows = 6
    cols = 4
    left = Inches(1)
    top = Inches(1.5)
    width = Inches(8)
    height = Inches(4.5)

    table = table_slide.shapes.add_table(rows, cols, left, top, width, height).table

    # Set headers
    headers = ['City', 'Population', 'Known For', 'Vibe']
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0, 102, 204)
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.font.bold = True
        paragraph.font.size = Pt(14)
        paragraph.font.color.rgb = RGBColor(255, 255, 255)

    # Add data
    data = [
        ['Stockholm', '975K', 'Capital, Islands', 'Cosmopolitan'],
        ['Gothenburg', '580K', 'Coast, Food Scene', 'Laid-back'],
        ['Malmö', '345K', 'Diversity, Design', 'International'],
        ['Uppsala', '175K', 'University, History', 'Academic'],
        ['Lund', '125K', 'Science, Innovation', 'Student City']
    ]

    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, cell_value in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.text = cell_value
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.font.size = Pt(12)

    # 8. Add flowchart: Moving to Sweden Process
    print("8. Adding flowchart...")
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

    flowchart_slide = prs.slides.add_slide(blank_slide_layout)

    title_box = flowchart_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
    title_frame = title_box.text_frame
    title_frame.text = "Moving to Sweden: The Process"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True

    steps = [
        {"text": "1. Get job offer or admission", "y": 1.5},
        {"text": "2. Apply for residence permit", "y": 2.6},
        {"text": "3. Register with Skatteverket", "y": 3.7},
        {"text": "4. Get Swedish ID (personnummer)", "y": 4.8},
        {"text": "5. Enjoy Swedish life! 🎉", "y": 5.9}
    ]

    for i, step in enumerate(steps):
        shape = flowchart_slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(2.5), Inches(step["y"]),
            Inches(5), Inches(0.8)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0, 102, 204)
        shape.text_frame.text = step["text"]
        shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        shape.text_frame.paragraphs[0].font.size = Pt(14)
        shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Add connector to next step
        if i < len(steps) - 1:
            connector = flowchart_slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(5), Inches(step["y"] + 0.8),
                Inches(5), Inches(steps[i+1]["y"])
            )
            connector.line.width = Pt(2)
            connector.line.color.rgb = RGBColor(0, 102, 204)

    # 9. Add content slide with bullet points
    print("9. Adding content slide with Swedish traditions...")
    content_slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = content_slide.shapes
    title_shape = shapes.title
    body_shape = shapes.placeholders[1]

    title_shape.text = "Swedish Traditions You'll Love"
    tf = body_shape.text_frame
    traditions = [
        "Fika - Coffee break culture (sacred!)",
        "Midsummer - Dancing around maypoles",
        "Crayfish parties in August",
        "Lucia procession in December",
        "Allemansrätten - Freedom to roam nature"
    ]

    for i, tradition in enumerate(traditions):
        if i == 0:
            tf.text = tradition
        else:
            p = tf.add_paragraph()
            p.text = tradition
            p.level = 0

    # 10. Add QR code for more info
    print("10. Adding QR code...")
    import qrcode

    qr_slide = prs.slides.add_slide(blank_slide_layout)

    title_box = qr_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.75))
    title_frame = title_box.text_frame
    title_frame.text = "Want to Learn More?"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.bold = True

    # Create QR code for the GitHub repo
    qr = qrcode.QRCode(version=1, box_size=10, border=1)
    qr.add_data("https://github.com/EiriniOr/mcp-powerpoint-server")
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    temp_path = "/tmp/qr_demo.png"
    img.save(temp_path)

    qr_slide.shapes.add_picture(temp_path, Inches(3.5), Inches(2), width=Inches(3), height=Inches(3))
    os.remove(temp_path)

    # Add text below QR
    text_box = qr_slide.shapes.add_textbox(Inches(2), Inches(5.5), Inches(6), Inches(1))
    text_frame = text_box.text_frame
    text_frame.text = "Scan to visit the PowerPoint MCP Server\nGitHub Repository"
    text_frame.paragraphs[0].font.size = Pt(16)
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 11. Thank you slide with shapes
    print("11. Adding thank you slide with shapes...")
    from pptx.enum.shapes import MSO_SHAPE

    thank_you_slide = prs.slides.add_slide(blank_slide_layout)
    thank_you_slide.background.fill.solid()
    thank_you_slide.background.fill.fore_color.rgb = RGBColor(0, 102, 204)

    # Add decorative stars (using circles as stars aren't available in all versions)
    star_positions = [(1, 1), (8, 1), (1, 6), (8, 6), (4.5, 0.5)]
    for x, y in star_positions:
        star = thank_you_slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x), Inches(y),
            Inches(0.5), Inches(0.5)
        )
        star.fill.solid()
        star.fill.fore_color.rgb = RGBColor(255, 215, 0)

    # Thank you text
    title_box = thank_you_slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
    title_frame = title_box.text_frame
    title_frame.text = "Tack så mycket!"
    title_frame.paragraphs[0].font.size = Pt(60)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    p = title_frame.add_paragraph()
    p.text = "Thank you very much!"
    p.font.size = Pt(28)
    p.font.color.rgb = RGBColor(255, 215, 0)
    p.alignment = PP_ALIGN.CENTER

    # Add footer with info
    p2 = title_frame.add_paragraph()
    p2.text = "\n\nCreated with PowerPoint MCP Server"
    p2.font.size = Pt(14)
    p2.font.color.rgb = RGBColor(255, 255, 255)
    p2.alignment = PP_ALIGN.CENTER

    # Save the presentation
    output_path = os.path.expanduser("~/Downloads/life_in_sweden_demo.pptx")
    prs.save(output_path)

    print(f"\n✅ Demo presentation created successfully!")
    print(f"📍 Saved to: {output_path}")
    print(f"\n📊 Slides created: {len(prs.slides)}")
    print("\nFeatures showcased:")
    print("  ✓ Title and agenda slides")
    print("  ✓ Section breaks with custom colors")
    print("  ✓ Comparison layout")
    print("  ✓ Column chart with data")
    print("  ✓ Timeline visualization")
    print("  ✓ Formatted table")
    print("  ✓ Flowchart with connectors")
    print("  ✓ Bullet point content")
    print("  ✓ QR code generation")
    print("  ✓ Shapes (stars) and custom styling")
    print("  ✓ Custom backgrounds and colors")

if __name__ == "__main__":
    asyncio.run(create_sweden_demo())
