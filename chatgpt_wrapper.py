#!/usr/bin/env python3
"""
Simple wrapper to use PowerPoint MCP Server with ChatGPT or any LLM
This provides a clean Python interface that any LLM can call
"""

import asyncio
import sys
import os

# Add server directory to path
sys.path.insert(0, os.path.dirname(__file__))
from server import call_tool

class PowerPointAPI:
    """
    Simple Python API for creating PowerPoint presentations
    Use this with ChatGPT, Claude, or any LLM
    """

    @staticmethod
    async def create_presentation(title: str, filename: str, subtitle: str = ""):
        """Create a new PowerPoint presentation"""
        result = await call_tool("create_presentation", {
            "title": title,
            "subtitle": subtitle,
            "filename": filename
        })
        return result[0].text

    @staticmethod
    async def add_content_slide(filename: str, title: str, content: list):
        """Add a slide with bullet points"""
        result = await call_tool("add_content_slide", {
            "filename": filename,
            "title": title,
            "content": content
        })
        return result[0].text

    @staticmethod
    async def add_chart(filename: str, title: str, chart_type: str,
                       categories: list, series: list):
        """
        Add a chart slide
        chart_type: 'bar', 'column', 'line', 'pie', or 'area'
        series: [{"name": "Series 1", "values": [1,2,3]}, ...]
        """
        result = await call_tool("add_chart_slide", {
            "filename": filename,
            "title": title,
            "chart_type": chart_type,
            "categories": categories,
            "series": series
        })
        return result[0].text

    @staticmethod
    async def analyze_data_and_chart(filename: str, data_file: str,
                                    chart_type: str, x_column: str,
                                    y_columns: list, title: str = None):
        """Analyze CSV/Excel/JSON and create chart automatically"""
        result = await call_tool("analyze_and_chart", {
            "filename": filename,
            "data_file": data_file,
            "chart_type": chart_type,
            "x_column": x_column,
            "y_columns": y_columns,
            "title": title
        })
        return result[0].text

    @staticmethod
    async def add_table(filename: str, title: str, headers: list, rows: list):
        """Add a table slide"""
        result = await call_tool("add_table_slide", {
            "filename": filename,
            "title": title,
            "headers": headers,
            "rows": rows
        })
        return result[0].text

    @staticmethod
    async def add_comparison(filename: str, title: str,
                           left_title: str, left_content: list,
                           right_title: str, right_content: list):
        """Add a comparison slide"""
        result = await call_tool("add_comparison_slide", {
            "filename": filename,
            "title": title,
            "left_title": left_title,
            "left_content": left_content,
            "right_title": right_title,
            "right_content": right_content
        })
        return result[0].text

    @staticmethod
    async def add_timeline(filename: str, title: str, events: list):
        """
        Add a timeline slide
        events: [{"date": "Q1 2024", "event": "Launch"}, ...]
        """
        result = await call_tool("add_timeline_slide", {
            "filename": filename,
            "title": title,
            "events": events
        })
        return result[0].text

    @staticmethod
    async def add_image(filename: str, image_path: str,
                       title: str = None, caption: str = None,
                       layout: str = "centered"):
        """
        Add an image slide
        layout: 'centered', 'title_and_image', 'image_left', 'image_right'
        """
        result = await call_tool("add_image_slide", {
            "filename": filename,
            "image_path": image_path,
            "title": title,
            "caption": caption,
            "layout": layout
        })
        return result[0].text

    @staticmethod
    async def save_presentation(filename: str, output_path: str = None):
        """Save the presentation to disk"""
        result = await call_tool("save_presentation", {
            "filename": filename,
            "output_path": output_path
        })
        return result[0].text

    @staticmethod
    async def read_data_file(data_file: str, sheet_name: str = None):
        """Read and analyze a data file"""
        result = await call_tool("read_data_file", {
            "data_file": data_file,
            "sheet_name": sheet_name
        })
        return result[0].text


# ============================================================================
# EXAMPLE USAGE - Copy these examples to use with ChatGPT
# ============================================================================

async def example_simple_presentation():
    """Example 1: Create a simple presentation"""
    ppt = PowerPointAPI()

    # Create presentation
    print(await ppt.create_presentation(
        title="My Presentation",
        subtitle="Created with Python",
        filename="simple.pptx"
    ))

    # Add content slide
    print(await ppt.add_content_slide(
        filename="simple.pptx",
        title="Key Points",
        content=[
            "First important point",
            "Second important point",
            "Third important point"
        ]
    ))

    # Save
    print(await ppt.save_presentation("simple.pptx"))


async def example_with_charts():
    """Example 2: Presentation with charts"""
    ppt = PowerPointAPI()

    # Create presentation
    await ppt.create_presentation(
        title="Sales Report Q4",
        filename="sales.pptx"
    )

    # Add chart
    print(await ppt.add_chart(
        filename="sales.pptx",
        title="Quarterly Revenue",
        chart_type="column",
        categories=["Q1", "Q2", "Q3", "Q4"],
        series=[
            {"name": "Revenue", "values": [100, 150, 200, 250]},
            {"name": "Costs", "values": [80, 90, 100, 110]}
        ]
    ))

    # Save
    print(await ppt.save_presentation("sales.pptx"))


async def example_data_analysis():
    """Example 3: Analyze data file and create charts"""
    ppt = PowerPointAPI()

    # First, analyze the data
    print(await ppt.read_data_file("data.csv"))

    # Create presentation
    await ppt.create_presentation(
        title="Data Analysis",
        filename="analysis.pptx"
    )

    # Auto-create chart from data
    print(await ppt.analyze_data_and_chart(
        filename="analysis.pptx",
        data_file="data.csv",
        chart_type="line",
        x_column="Month",
        y_columns=["Sales", "Profit"]
    ))

    # Save
    print(await ppt.save_presentation("analysis.pptx"))


async def example_comprehensive():
    """Example 4: Comprehensive presentation"""
    ppt = PowerPointAPI()

    # Create presentation
    await ppt.create_presentation(
        title="Product Launch 2025",
        subtitle="Strategic Overview",
        filename="launch.pptx"
    )

    # Add comparison slide
    await ppt.add_comparison(
        filename="launch.pptx",
        title="Market Position",
        left_title="Current State",
        left_content=[
            "Limited market share",
            "Manual processes",
            "High costs"
        ],
        right_title="Future State",
        right_content=[
            "Market leadership",
            "Automated workflows",
            "Optimized costs"
        ]
    )

    # Add timeline
    await ppt.add_timeline(
        filename="launch.pptx",
        title="Launch Timeline",
        events=[
            {"date": "Jan", "event": "Development"},
            {"date": "Mar", "event": "Beta Testing"},
            {"date": "May", "event": "Marketing"},
            {"date": "Jul", "event": "Launch"}
        ]
    )

    # Add table
    await ppt.add_table(
        filename="launch.pptx",
        title="Feature Comparison",
        headers=["Feature", "Competitor A", "Competitor B", "Our Product"],
        rows=[
            ["Price", "$99", "$149", "$79"],
            ["Support", "Email", "Email", "24/7 Phone"],
            ["Updates", "Yearly", "Quarterly", "Monthly"]
        ]
    )

    # Save
    print(await ppt.save_presentation("launch.pptx"))


# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

def print_help():
    """Print available commands"""
    print("""
PowerPoint API - Command Line Interface

Usage:
    python chatgpt_wrapper.py <command> [arguments]

Commands:
    test              Run all example scripts
    simple            Create a simple presentation
    charts            Create presentation with charts
    comprehensive     Create comprehensive presentation
    help              Show this help message

Examples:
    python chatgpt_wrapper.py simple
    python chatgpt_wrapper.py test
    """)


async def main():
    """Main entry point"""
    if len(sys.argv) < 2:
        print_help()
        return

    command = sys.argv[1].lower()

    if command == "test":
        print("Running all examples...\n")
        print("=" * 60)
        print("Example 1: Simple Presentation")
        print("=" * 60)
        await example_simple_presentation()

        print("\n" + "=" * 60)
        print("Example 2: Presentation with Charts")
        print("=" * 60)
        await example_with_charts()

        print("\n" + "=" * 60)
        print("Example 4: Comprehensive Presentation")
        print("=" * 60)
        await example_comprehensive()

        print("\n✅ All examples completed!")
        print("📂 Check your Downloads folder for the presentations")

    elif command == "simple":
        await example_simple_presentation()

    elif command == "charts":
        await example_with_charts()

    elif command == "comprehensive":
        await example_comprehensive()

    elif command == "help":
        print_help()

    else:
        print(f"Unknown command: {command}")
        print_help()


if __name__ == "__main__":
    asyncio.run(main())
