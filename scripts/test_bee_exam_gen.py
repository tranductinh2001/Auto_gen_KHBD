import os
import sys
import json
import requests
from dotenv import load_dotenv

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()
api_key = os.getenv("BEE_API_KEY")
base_url = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")
model = os.getenv("VISION_MODEL", "google/gemini-3.1-flash-lite")

headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
payload = {
    "model": model,
    "messages": [
        {
            "role": "system",
            "content": "Bạn là chuyên gia khảo thí môn Tin học GDPT 2018. Hãy trả về kết quả dạng JSON thuần túy không có markdown block."
        },
        {
            "role": "user",
            "content": "Hãy tạo 1 câu hỏi trắc nghiệm Tin học 9 về 'Chất lượng thông tin', mức độ 'Thông hiểu' gồm 4 lựa chọn A, B, C, D và đáp án đúng. Định dạng JSON: {\"question\": \"...\", \"level\": \"Thông hiểu\", \"options\": {\"A\": \"...\", \"B\": \"...\", \"C\": \"...\", \"D\": \"...\"}, \"correct\": \"...\"}"
        }
    ],
    "temperature": 0.2,
    "max_tokens": 1000
}

resp = requests.post(f"{base_url}/chat/completions", headers=headers, json=payload, timeout=30)
print("Status:", resp.status_code)
content = resp.json()["choices"][0]["message"]["content"].strip()
print("Raw response:\n", content)

if content.startswith("```"):
    lines = content.split("\n")
    if lines[0].startswith("```"): lines = lines[1:]
    if lines and lines[-1].startswith("```"): lines = lines[:-1]
    content = "\n".join(lines).strip()

data = json.loads(content)
print("\nParsed JSON:", json.dumps(data, ensure_ascii=False, indent=2))
