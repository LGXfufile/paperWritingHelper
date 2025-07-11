

# Please install OpenAI SDK first: `pip3 install openai`

from openai import OpenAI

client = OpenAI(api_key="sk-0f69d5e4890f450d9b958d6ad9e19c7e", base_url="https://api.deepseek.com")

response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {"role": "system", "content": "You are a helpful assistant"},
        {"role": "user", "content": "中文回复我，如何快速赚到1000万，请详细描述"},
    ],
    stream=False
)

print(response.choices[0].message.content)