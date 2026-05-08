import os
from anthropic import Anthropic

client = Anthropic()  # reads ANTHROPIC_API_KEY from environment

resp = client.messages.create(
    model="claude-haiku-4-5",
    max_tokens=50,
    messages=[{"role": "user", "content": "say hi in 5 words"}]
)

print(resp.content[0].text)