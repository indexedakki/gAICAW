import os
from openai import AzureOpenAI

from dotenv import load_dotenv, find_dotenv
load_dotenv(find_dotenv(), override=True)

api_key=os.getenv("OPENAI_API_KEY")
api_version=os.getenv("OPENAI_API_VERSION")
azure_endpoint=os.getenv("OPENAI_API_BASE")

client = AzureOpenAI(
api_key=api_key,  
api_version=api_version,
azure_endpoint=azure_endpoint 
)

conversation = [{"role": "system", "content": "You are expert python programmer and have expertise in working with xml files."}]

user_input = f"""
    
    I want to convert a word document to xml section wise and then convert each xml back to different word document.
    write whole code for it
    
    """  

conversation.append({"role": "user", "content": user_input})

response = client.chat.completions.create(
    model="LMM_OPENAI",
    messages=conversation
)

conversation.append({"role": "assistant", "content": response.choices[0].message.content})


print("\n" + response.choices[0].message.content + "\n")