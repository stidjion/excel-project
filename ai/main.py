import ollama
import requests
import json
from typing import List, Dict, Any

# ────────────────────────────────────────────────
#  CONFIG
# ────────────────────────────────────────────────

API_URL = "http://127.0.0.1:8000/execute"          # your running FastAPI server
MODEL = "qwen2.5:7b"                               # change if using different tag

# Tool definition — matches your Command model in api.py
TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "execute_excel_action",
            "description": "Perform an action on the Excel file (data.xlsx). Use this for create_sheet, add_row, get_preview, etc.",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "description": "The exact action name (must match your connector)",
                        "enum": ["create_sheet", "add_row", "get_preview"]  # ← only your MVP actions
                    },
                    "params": {
                        "type": "object",
                        "description": "Parameters required by the action",
                        "additionalProperties": True
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["execute", "dry_run"],
                        "description": "Execution mode (default: execute)",
                        "default": "execute"
                    }
                },
                "required": ["action"]
            }
        }
    }
]

SYSTEM_PROMPT = """
You are an Excel assistant working with the file 'data.xlsx'.
You can ONLY use the 'execute_excel_action' tool to perform operations.
Do not invent actions or parameters — stick exactly to the tool description.

Examples:
- User: "Create a new sheet called Sales"
  → Call tool with action: "create_sheet", params: {"sheet_name": "Sales"}

- User: "Add row to Sheet1: Name=Rehan, Age=28"
  → Call tool with action: "add_row", params: {"sheet_name": "Sheet1", "values": {"Name": "Rehan", "Age": 28}}

- User: "Show first 5 rows of Data"
  → Call tool with action: "get_preview", params: {"sheet_name": "Data", "n": 5}

If the request is unclear, ask for more details instead of guessing.
Always use the tool when an Excel operation is needed.
"""

# ────────────────────────────────────────────────
#  MAIN CHAT LOOP
# ────────────────────────────────────────────────

messages: List[Dict[str, Any]] = [
    {"role": "system", "content": SYSTEM_PROMPT}
]

print("\nExcel Assistant ready! Tell me what you'd like to do with data.xlsx")
print("(type 'exit' or 'quit' to stop)\n")

while True:
    user_input = input("You: ").strip()
    if user_input.lower() in ["exit", "quit", "q", "bye"]:
        print("Goodbye!")
        break
    if not user_input:
        continue

    messages.append({"role": "user", "content": user_input})

    # ── Call Ollama with tools ────────────────────────────────
    response = ollama.chat(
        model=MODEL,
        messages=messages,
        tools=TOOLS,
        options={"temperature": 0.2},  # low = more reliable tool use
    )

    assistant_msg = response['message']
    messages.append(assistant_msg)

    # Print assistant's content (explanation or question)
    if assistant_msg.get('content'):
        print(f"Assistant: {assistant_msg['content']}")

        if assistant_msg.get('tool_calls'):
          for tool_call in assistant_msg['tool_calls']:
            try:
                payload = {
                    "action": tool_call["function"]["name"],
                    "params": json.loads(tool_call["function"]["arguments"]),
                    "mode": None
                }

                api_response = requests.post(API_URL, json=payload)
                api_response.raise_for_status()

                result = api_response.json()
                print(f"API result: {result}")

                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.get("id", "call_1"),
                    "name": tool_call["function"]["name"],
                    "content": json.dumps(result)
                })

            except requests.exceptions.RequestException as e:
                error_msg = f"API call failed: {str(e)}"
                print(error_msg)
                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.get("id", "call_1"),
                    "name": tool_call["function"]["name"],
                    "content": error_msg
                })
        followup_response = ollama.chat(
            model=MODEL,
            messages=messages,
            options={"temperature": 0.2},
        )

        final_msg = followup_response['message']

        if final_msg.get('content'):
            print(f"\nAssistant: {final_msg['content']}")
            messages.append(final_msg)
       
