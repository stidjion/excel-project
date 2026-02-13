import ollama
from ollama import chat 
from.defined_function import defined_functions





messages = [{'role': 'user', 'content': "What is the temperature in New York?"}]

while True:
  stream = chat(
    model='qwen 2.5:7b',
    messages=messages,
    tools=[defined_functions],
    stream=True,
    think=True,
  )

  thinking = ''
  content = ''
  tool_calls = []

  done_thinking = False
  # accumulate the partial fields
  for chunk in stream:
    if chunk.message.thinking:
      thinking += chunk.message.thinking
      print(chunk.message.thinking, end='', flush=True)
    if chunk.message.content:
      if not done_thinking:
        done_thinking = True
        print('\n')
      content += chunk.message.content
      print(chunk.message.content, end='', flush=True)
    if chunk.message.tool_calls:
      tool_calls.extend(chunk.message.tool_calls)
      print(chunk.message.tool_calls)

  # append accumulated fields to the messages
  if thinking or content or tool_calls:
    messages.append({'role': 'assistant', 'thinking': thinking, 'content': content, 'tool_calls': tool_calls})

  if not tool_calls:
    break

  for call in tool_calls:
    if call.function.name == defined_functions[0].name:
      result = defined_functions[0](**call.function.arguments)
    else:
      result = 'Unknown tool'
    messages.append({'role': 'tool', 'tool_name': call.function.name, 'content': result})
