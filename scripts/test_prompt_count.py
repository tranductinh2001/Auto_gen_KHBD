import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("BEE_API_KEY")
base_url = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")
model = os.getenv("VISION_MODEL", "google/gemini-3.1-flash-lite")

headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
payload = {
    "model": model,
    "messages": [
        {"role": "system", "content": "Ban la chuyen gia khao thi GDPT 2018. Tra ve JSON duy nhat hop le."},
        {"role": "user", "content": """Hay bien soan bo cau hoi thi hoc ky I mon Tin hoc 9.
YEU CAU SO LUONG BAT BUOC (PHAI DU 100%, KHONG RUT GON):
- "mcqs": mang chua CHINH XAC 12 cau hoi trac nghiem (MCQ_01 den MCQ_12)
- "tfs": mang chua CHINH XAC 4 cau dung/sai (TF_01 den TF_04), moi cau du 4 y a, b, c, d
- "essays": mang chua CHINH XAC 3 cau tu luan (ES_01 den ES_03)

Schema JSON:
{
  "mcqs": [
    {"id": "MCQ_01", "unit": "Bai 1", "level": "Nhan biet", "content": "Cau 1...", "options": {"A": "...", "B": "...", "C": "...", "D": "..."}, "correct": "A"},
    {"id": "MCQ_02", "unit": "Bai 1", "level": "Nhan biet", "content": "Cau 2...", "options": {"A": "...", "B": "...", "C": "...", "D": "..."}, "correct": "B"}
  ],
  "tfs": [
    {"id": "TF_01", "unit": "Bai 1", "topic": "Tinh huong...", "statements": [
      {"label": "a", "content": "Y a...", "correct": "D", "reason": "..."},
      {"label": "b", "content": "Y b...", "correct": "S", "reason": "..."},
      {"label": "c", "content": "Y c...", "correct": "D", "reason": "..."},
      {"label": "d", "content": "Y d...", "correct": "S", "reason": "..."}
    ]}
  ],
  "essays": [
    {"id": "ES_01", "unit": "Bai 1", "level": "Thong hieu", "score": 1.0, "content": "Tu luan 1...", "rubric": [{"criterion": "Y 1", "score": 0.5}, {"criterion": "Y 2", "score": 0.5}]}
  ]
}
Luu y: Mau tren chi minh hoa 2 cau mcqs, 1 cau tf, 1 cau essay. KHI TRA VE BAN PHAI SINH DU 12 CAU MCQS, 4 CAU TFS, VA 3 CAU ESSAYS!"""}
    ],
    "temperature": 0.2,
    "max_tokens": 16000
}

resp = requests.post(f"{base_url}/chat/completions", headers=headers, json=payload, timeout=90)
clean = resp.json()["choices"][0]["message"]["content"].strip()
if clean.startswith("```"):
    clean = clean.split("\n", 1)[1].rsplit("\n", 1)[0].strip()

data = json.loads(clean)
print("MCQ count:", len(data.get("mcqs", [])))
print("TF count:", len(data.get("tfs", [])))
print("Essay count:", len(data.get("essays", [])))
