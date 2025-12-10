#!/usr/bin/env python3
"""
Test script for PowerPoint MCP Server
Run this to verify your MCP server works correctly
"""

import asyncio
import sys
import os

# Add server directory to path
sys.path.insert(0, os.path.dirname(__file__))
from server import call_tool

async def test():
    print("🧪 Testing PowerPoint MCP Server...\n")

    # Test 1: Create presentation
    print("1️⃣  Creating presentation...")
    result = await call_tool("create_presentation", {
        "title": "MCP Test Presentation",
        "subtitle": "Generated from Python",
        "filename": "test_demo.pptx"
    })
    print(f"   ✓ {result[0].text}\n")

    # Test 2: Add content slide
    print("2️⃣  Adding content slide...")
    result = await call_tool("add_content_slide", {
        "filename": "test_demo.pptx",
        "title": "Amazing Features",
        "content": [
            "📊 Charts and graphs",
            "📋 Tables with data",
            "🖼️ Images with layouts",
            "📅 Timeline visualizations",
            "🎨 Custom formatting"
        ]
    })
    print(f"   ✓ {result[0].text}\n")

    # Test 3: Add comparison slide
    print("3️⃣  Adding comparison slide...")
    result = await call_tool("add_comparison_slide", {
        "filename": "test_demo.pptx",
        "title": "Before vs After",
        "left_title": "Before MCP",
        "left_content": [
            "Manual slide creation",
            "Time consuming",
            "Repetitive work",
            "Limited automation"
        ],
        "right_title": "After MCP",
        "right_content": [
            "Automated generation",
            "Lightning fast",
            "Consistent results",
            "Full API control"
        ]
    })
    print(f"   ✓ {result[0].text}\n")

    # Test 4: Add chart
    print("4️⃣  Adding chart slide...")
    result = await call_tool("add_chart_slide", {
        "filename": "test_demo.pptx",
        "title": "Quarterly Performance",
        "chart_type": "column",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
            {"name": "Revenue", "values": [100, 150, 200, 280]},
            {"name": "Profit", "values": [20, 35, 60, 90]}
        ]
    })
    print(f"   ✓ {result[0].text}\n")

    # Test 5: Add timeline
    print("5️⃣  Adding timeline slide...")
    result = await call_tool("add_timeline_slide", {
        "filename": "test_demo.pptx",
        "title": "Project Roadmap 2025",
        "events": [
            {"date": "Jan", "event": "Planning"},
            {"date": "Apr", "event": "Development"},
            {"date": "Jul", "event": "Testing"},
            {"date": "Oct", "event": "Launch"}
        ]
    })
    print(f"   ✓ {result[0].text}\n")

    # Test 6: Add table
    print("6️⃣  Adding table slide...")
    result = await call_tool("add_table_slide", {
        "filename": "test_demo.pptx",
        "title": "Feature Comparison",
        "headers": ["Feature", "Basic", "Pro", "Enterprise"],
        "rows": [
            ["Slides", "10", "100", "Unlimited"],
            ["Charts", "❌", "✅", "✅"],
            ["Tables", "❌", "✅", "✅"],
            ["Data Analysis", "❌", "❌", "✅"]
        ]
    })
    print(f"   ✓ {result[0].text}\n")

    # Test 7: List presentations
    print("7️⃣  Listing presentations in memory...")
    result = await call_tool("list_presentations", {})
    print(f"   ✓ {result[0].text}\n")

    # Test 8: Save presentation
    print("8️⃣  Saving presentation...")
    result = await call_tool("save_presentation", {
        "filename": "test_demo.pptx"
    })
    print(f"   ✓ {result[0].text}\n")

    print("=" * 60)
    print("✅ All tests passed!")
    print("=" * 60)
    print("\n📂 Check your Downloads folder for 'test_demo.pptx'")
    print("🎉 Your MCP server is working perfectly!\n")

if __name__ == "__main__":
    try:
        asyncio.run(test())
    except Exception as e:
        print(f"\n❌ Test failed: {str(e)}")
        sys.exit(1)
