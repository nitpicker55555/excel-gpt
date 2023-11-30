import os
import openai
import re
import sys
openai.api_key = ""


response = openai.ChatCompletion.create(
    model="gpt-4",
    messages=[
        {
            "role": "system",
            "content": str("Write a complete vba program for excel on the current worksheet according to prompt, do not output other text. ")
        },
        {
            "role": "user",
            "content": str(sys.argv[1])
        }
    ],
    temperature=1,
    max_tokens=1024,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
)

# print(response["choices"][0]["message"]["content"])
final_query = response["choices"][0]["message"]["content"]
json_str = str(final_query)




result = re.search(r'Sub.*End Sub', json_str, re.DOTALL)
if result:
        extracted_text = result.group()
        extracted_text = bytes(extracted_text, 'utf-8').decode('unicode_escape')
        text=extracted_text.replace(r"\n", "\n")
        text = text.replace(r"\\", "")
        print(text)
else:text="nothing"
y = json_str+text

with open('c:\\bbdata\\file.txt', 'w',encoding='utf-8') as file:
    file.write(y)
