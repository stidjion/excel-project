from google import genai
from google.genai import types
import defined_function



# Configure the client and tools
client = genai.Client()
tools = types.Tool(function_declarations=defined_function.defined_functions)
config = types.GenerateContentConfig(tools=[tools])

# Send request with function declarations
response = client.models.generate_content(
    model="gemini-3-flash-preview",
    contents="Schedule a meeting with Bob and Alice for 03/14/2025 at 10:00 AM about the Q3 planning.",
    config=config,
)

# Check for a function call
if response.candidates[0].content.parts[0].function_call:
    function_call = response.candidates[0].content.parts[0].function_call
    print(f"Function to call: {function_call.name}")
    print(f"Arguments: {function_call.args}")
    #  In a real app, you would call your function here:
    #  result = schedule_meeting(**function_call.args)
else:
    print("No function call found in the response.")
    print(response.text)