import requests

API_URL = "https://api-inference.huggingface.co/models/meta-llama/Llama-2-7b-chat-hf"
headers = {"Authorization": "Bearer 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"}

def query(payload):
	response = requests.post(API_URL, headers=headers, json=payload)
	return response.json()
	
output = query({
	"inputs": "Can you please let us know more details about your name?",
    "wait_for_model" : True
})
print(output)