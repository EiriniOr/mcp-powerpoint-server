# Using PowerPoint MCP Server with ChatGPT and Other LLMs

Your MCP server follows the Model Context Protocol standard, but ChatGPT doesn't natively support MCP yet. Here are your options:

## Option 1: MCP-to-OpenAI Bridge (Recommended)

Create a bridge that exposes your MCP tools as OpenAI function calls.

### Step 1: Install Dependencies

```bash
cd /Users/rena/mcp-powerpoint-server
pip install openai fastapi uvicorn
```

### Step 2: Create Bridge Server

Save this as `openai_bridge.py`:

```python
#!/usr/bin/env python3
"""
Bridge server that exposes MCP PowerPoint tools as OpenAI functions
"""

import asyncio
import json
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import subprocess
import sys

app = FastAPI()

# Import the MCP server functions
sys.path.insert(0, '/Users/rena/mcp-powerpoint-server')
from server import presentations
import server as ppt_server

# Map OpenAI function calls to MCP tools
@app.post("/v1/chat/completions")
async def chat_completions(request: Request):
    data = await request.json()

    # Check if there are function calls in the request
    messages = data.get("messages", [])

    # Extract function call if present
    for message in reversed(messages):
        if message.get("function_call"):
            function_name = message["function_call"]["name"]
            arguments = json.loads(message["function_call"]["arguments"])

            # Call the MCP tool
            result = await ppt_server.call_tool(function_name, arguments)

            return JSONResponse({
                "choices": [{
                    "message": {
                        "role": "function",
                        "name": function_name,
                        "content": result[0].text
                    }
                }]
            })

    return JSONResponse({"error": "Not implemented"}, status_code=501)

# List available functions (MCP tools)
@app.get("/v1/tools")
async def list_tools():
    tools = await ppt_server.list_tools()

    # Convert MCP tools to OpenAI function format
    openai_functions = []
    for tool in tools:
        openai_functions.append({
            "name": tool.name,
            "description": tool.description,
            "parameters": tool.inputSchema
        })

    return {"functions": openai_functions}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
```

### Step 3: Run the Bridge

```bash
python openai_bridge.py
```

Then configure ChatGPT to use your local endpoint (requires ChatGPT Plus with plugin/function support).

## Option 2: Custom GPT with Actions

If you have ChatGPT Plus, you can create a Custom GPT:

1. Go to ChatGPT → "Explore GPTs" → "Create a GPT"
2. Add actions that call your bridge server
3. Define the PowerPoint operations as actions

## Option 3: Python Client for Any LLM

Use this Python wrapper to integrate with any LLM:

```python
import asyncio
import sys
sys.path.insert(0, '/Users/rena/mcp-powerpoint-server')
from server import call_tool, presentations, Presentation

async def create_presentation_with_llm(llm_response):
    """
    Parse LLM response and execute PowerPoint commands

    Example:
        llm_response = {
            "action": "create_presentation",
            "params": {
                "title": "AI Overview",
                "filename": "ai.pptx"
            }
        }
    """

    action = llm_response["action"]
    params = llm_response["params"]

    result = await call_tool(action, params)
    return result[0].text

# Example usage with any LLM
async def main():
    # Create presentation
    result = await create_presentation_with_llm({
        "action": "create_presentation",
        "params": {
            "title": "My Presentation",
            "filename": "demo.pptx"
        }
    })
    print(result)

    # Add a slide
    result = await create_presentation_with_llm({
        "action": "add_content_slide",
        "params": {
            "filename": "demo.pptx",
            "title": "Key Points",
            "content": ["Point 1", "Point 2", "Point 3"]
        }
    })
    print(result)

    # Save
    result = await create_presentation_with_llm({
        "action": "save_presentation",
        "params": {
            "filename": "demo.pptx"
        }
    })
    print(result)

if __name__ == "__main__":
    asyncio.run(main())
```

## Option 4: LangChain Integration

```python
from langchain.tools import Tool
from langchain.agents import initialize_agent
from langchain.llms import OpenAI
import asyncio
import sys

sys.path.insert(0, '/Users/rena/mcp-powerpoint-server')
from server import call_tool

def create_powerpoint_tool(tool_name, description):
    async def run_tool(**kwargs):
        result = await call_tool(tool_name, kwargs)
        return result[0].text

    def sync_wrapper(**kwargs):
        return asyncio.run(run_tool(**kwargs))

    return Tool(
        name=tool_name,
        description=description,
        func=sync_wrapper
    )

# Create tools for each PowerPoint operation
tools = [
    create_powerpoint_tool(
        "create_presentation",
        "Creates a new PowerPoint presentation with a title slide. Args: title, filename, subtitle (optional)"
    ),
    create_powerpoint_tool(
        "add_content_slide",
        "Adds a content slide with bullet points. Args: filename, title, content (array)"
    ),
    create_powerpoint_tool(
        "add_chart_slide",
        "Adds a chart slide. Args: filename, title, chart_type, categories, series"
    ),
    create_powerpoint_tool(
        "save_presentation",
        "Saves the presentation. Args: filename, output_path (optional)"
    ),
    # Add more tools as needed...
]

# Initialize agent with ChatGPT
llm = OpenAI(temperature=0)
agent = initialize_agent(
    tools,
    llm,
    agent="zero-shot-react-description",
    verbose=True
)

# Use it
agent.run("Create a presentation about AI with 3 slides and save it")
```

## Option 5: Direct HTTP API

Create a simple REST API wrapper:

```python
from fastapi import FastAPI
import asyncio
import sys

sys.path.insert(0, '/Users/rena/mcp-powerpoint-server')
from server import call_tool

app = FastAPI()

@app.post("/powerpoint/{action}")
async def execute_action(action: str, params: dict):
    result = await call_tool(action, params)
    return {"result": result[0].text}

# Run with: uvicorn api_wrapper:app --reload
```

Then call from any language:

```bash
curl -X POST http://localhost:8000/powerpoint/create_presentation \
  -H "Content-Type: application/json" \
  -d '{"title": "My Presentation", "filename": "demo.pptx"}'
```

## Comparison

| Method | Pros | Cons | Best For |
|--------|------|------|----------|
| MCP-to-OpenAI Bridge | Native ChatGPT integration | Requires local server | ChatGPT Plus users |
| Custom GPT | No coding needed | Requires ChatGPT Plus | Non-technical users |
| Python Client | Simple, direct | Manual integration | Developers |
| LangChain | Framework integration | Complex setup | AI applications |
| HTTP API | Language-agnostic | Additional server | Web apps |

## Quick Test Script

Test your setup with this script:

```python
#!/usr/bin/env python3
import asyncio
import sys
sys.path.insert(0, '/Users/rena/mcp-powerpoint-server')
from server import call_tool

async def test():
    # Create presentation
    result = await call_tool("create_presentation", {
        "title": "Test Presentation",
        "subtitle": "Generated from Python",
        "filename": "test.pptx"
    })
    print(f"✓ {result[0].text}")

    # Add content slide
    result = await call_tool("add_content_slide", {
        "filename": "test.pptx",
        "title": "Features",
        "content": ["Charts", "Tables", "Images", "Timelines"]
    })
    print(f"✓ {result[0].text}")

    # Add chart
    result = await call_tool("add_chart_slide", {
        "filename": "test.pptx",
        "title": "Sample Chart",
        "chart_type": "column",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
            {"name": "Sales", "values": [100, 150, 200, 250]},
            {"name": "Costs", "values": [80, 90, 100, 110]}
        ]
    })
    print(f"✓ {result[0].text}")

    # Save
    result = await call_tool("save_presentation", {
        "filename": "test.pptx"
    })
    print(f"✓ {result[0].text}")

if __name__ == "__main__":
    print("Testing PowerPoint MCP Server...")
    asyncio.run(test())
    print("\n✅ All tests passed!")
```

Save as `test_mcp.py` and run:
```bash
python test_mcp.py
```

## Resources

- [OpenAI Functions Documentation](https://platform.openai.com/docs/guides/function-calling)
- [LangChain Tools](https://python.langchain.com/docs/modules/agents/tools/)
- [FastAPI Documentation](https://fastapi.tiangolo.com/)
- [MCP Specification](https://modelcontextprotocol.io)

## Support

For issues or questions:
- MCP Server: Check `/Users/rena/mcp-powerpoint-server/server.py`
- Test script: Run `python test_mcp.py`
- Logs: Check console output when running the bridge
