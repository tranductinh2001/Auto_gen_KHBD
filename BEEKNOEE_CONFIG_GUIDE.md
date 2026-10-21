# 📘 HƯỚNG DẪN CẤU HÌNH & TÍCH HỢP BEEKNOEE AI API (CHUẨN OPENAI-COMPATIBLE)

> **Tài liệu chuẩn hóa:** Dùng làm mẫu tham chiếu cấu hình (Cheat-sheet & Reference Guide) để sao chép nhanh sang các dự án AI, Web, Tool tự động hóa khác.

---

## 1. TỔNG QUAN VỀ NỀN TẢNG BEEKNOEE AI PLATFORM

Beeknoee AI Platform cung cấp gateway tương thích 100% theo chuẩn **OpenAI REST API**, cho phép truy cập hàng trăm mô hình trí tuệ nhân tạo hàng đầu thế giới (Google Gemini, OpenAI, DeepSeek, GLM, Claude...) qua một cổng kết nối duy nhất:

- **Base URL:** `https://platform.beeknoee.com/api/v1`
- **Giao thức xác thực:** `Authorization: Bearer <YOUR_BEE_API_KEY>`
- **Chuẩn giao tiếp:** JSON qua HTTP POST `/chat/completions`, GET `/models`
- **Hỗ trợ Đa phương thức (Multimodal):** Nhận diện hình ảnh (OCR, sơ đồ, giao diện) và văn bản cùng lúc.

---

## 2. NGUYÊN NHÂN LỖI 404 VÀ MODEL THAY THẾ CHUẨN XÁC

### ⚠️ Lỗi gặp phải trước đó:
```json
HTTP 404 Not Found:
{"error":{"message":"Model \"google/gemini-2.5-flash-lite\" không tồn tại hoặc đã bị vô hiệu. Dùng GET /v1/models để xem danh sách model."}}
```
**Nguyên nhân:** Nhà cung cấp Beeknoee đã nâng cấp hạ tầng và tắt các bản preview cũ (`gemini-2.5-flash-lite`).

### ✅ Model thay thế chính thức tốt nhất hiện nay:
👉 **`google/gemini-3.1-flash-lite`** (Hoặc `gemini-3.1-flash-lite`)

**Ưu điểm vượt trội:**
1. **Tốc độ cực nhanh:** Độ trễ thấp (phản hồi trong 0.5 – 1.5 giây), phù hợp làm web app thời gian thực.
2. **Đa phương thức (Multimodal):** Đọc trích xuất ảnh SGK, sơ đồ lưu đồ khối, biểu đồ và văn bản cực kỳ chính xác.
3. **Độ ổn định cao:** Đã được kiểm tra kết nối trả về HTTP 200 OK trên Beeknoee Gateway.
4. **Chi phí token tối ưu:** Giá rẻ nhất trong dòng Gemini 3.x.

---

## 3. DANH SÁCH CÁC MODEL KHUYÊN DÙNG TRÊN BEEKNOEE

| Tên Model (ID chuẩn) | Loại hình | Thế mạnh & Trường hợp sử dụng |
| :--- | :---: | :--- |
| **`google/gemini-3.1-flash-lite`** *(Khuyên dùng)* | Multimodal + Text | **Mặc định cho toàn bộ dự án:** OCR ảnh SGK, sinh đề thi, tóm tắt, trích xuất dữ liệu tốc độ cao. |
| **`google/gemini-3.5-flash-lite`** | Multimodal + Text | Phiên bản Flash 3.5 nâng cấp, tốc độ phản hồi cực nhanh. |
| **`google/gemini-3.5-flash`** | Multimodal + Text | Khả năng lập luận sâu sắc hơn, phù hợp bài toán logic và phân tích phức tạp. |
| **`google/gemini-3.1-pro-preview`** | Suy luận cao cấp | Dành cho các tác vụ cần tư duy đa bước, giải toán, viết code chuyên sâu. |
| **`deepseek/deepseek-v4-flash`** | Text LLM | Dòng mã nguồn mở DeepSeek tối ưu hóa tốc độ và lập luận. |

---

## 4. MẪU CẤU HÌNH FILE `.env` (COPY DÙNG NGAY)

Tạo file `.env` tại thư mục gốc của dự án với nội dung:

```env
# ==============================================================================
# CẤU HÌNH BEEKNOEE AI GATEWAY (OPENAI-COMPATIBLE)
# ==============================================================================
BEE_API_KEY=sk-bee-9a4e3ae0de60aaab4298ad268bf3a66ef42e5e59ac4ccd7506750d71d80622ae
BEE_BASE_URL=https://platform.beeknoee.com/api/v1

# Model xử lý chính (Text & Multimodal Vision OCR)
VISION_MODEL=google/gemini-3.1-flash-lite
```

---

## 5. CODE MẪU TÍCH HỢP CHO CÁC DỰ ÁN SAU

### Cách 1: Sử dụng thư viện `requests` (Python cơ bản - không cần cài OpenAI SDK)

```python
import os
import requests
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("BEE_API_KEY")
BASE_URL = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")
MODEL = os.getenv("VISION_MODEL", "google/gemini-3.1-flash-lite")

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

# 1. Chat text thông thường
payload = {
    "model": MODEL,
    "messages": [
        {"role": "system", "content": "Bạn là chuyên gia sư phạm môn Tin học."},
        {"role": "user", "content": "Tạo 1 câu hỏi trắc nghiệm Tin học 9 về Mạng máy tính."}
    ],
    "temperature": 0.2,
    "max_tokens": 1000
}

response = requests.post(f"{BASE_URL}/chat/completions", headers=headers, json=payload, timeout=30)
if response.status_code == 200:
    content = response.json()["choices"][0]["message"]["content"]
    print("AI Response:\n", content)
else:
    print(f"Error {response.status_code}: {response.text}")
```

---

### Cách 2: Sử dụng thư viện `openai` chính thức (Python OpenAI SDK)

```python
import os
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

client = OpenAI(
    api_key=os.getenv("BEE_API_KEY"),
    base_url=os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")
)

response = client.chat.completions.create(
    model=os.getenv("VISION_MODEL", "google/gemini-3.1-flash-lite"),
    messages=[
        {"role": "user", "content": "Xin chào, hãy giới thiệu ngắn gọn trong 1 câu."}
    ],
    temperature=0.3
)

print(response.choices[0].message.content)
```

---

### Cách 3: Gửi ảnh kèm Prompt (Multimodal Vision OCR)

```python
import base64
import requests

def chat_with_image(image_path: str, prompt: str):
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "model": "google/gemini-3.1-flash-lite",
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}}
            ]
        }],
        "temperature": 0.1,
        "max_tokens": 4096
    }
    headers = {
        "Authorization": f"Bearer {os.getenv('BEE_API_KEY')}",
        "Content-Type": "application/json"
    }
    r = requests.post(f"{os.getenv('BEE_BASE_URL')}/chat/completions", headers=headers, json=payload, timeout=60)
    return r.json()["choices"][0]["message"]["content"]
```

---

### Cách 4: Lệnh cURL kiểm tra kết nối nhanh từ Terminal (Windows CMD / PowerShell / Linux)

```bash
curl https://platform.beeknoee.com/api/v1/chat/completions \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer sk-bee-9a4e3ae0de60aaab4298ad268bf3a66ef42e5e59ac4ccd7506750d71d80622ae" \
  -d '{
    "model": "google/gemini-3.1-flash-lite",
    "messages": [{"role": "user", "content": "Ping"}]
  }'
```

---

## 6. SCRIPT KIỂM TRA DANH SÁCH MODEL KHẢ DỤNG BẤT KỲ LÚC NÀO

Khi muốn biết nhà cung cấp đang bật/tắt những model nào, hãy chạy script nhỏ sau:

```python
import os, requests
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("BEE_API_KEY")
base_url = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")

r = requests.get(f"{base_url}/models", headers={"Authorization": f"Bearer {api_key}"})
if r.status_code == 200:
    all_models = [m["id"] for m in r.json().get("data", [])]
    gemini_models = [m for m in all_models if "gemini" in m.lower() or "flash" in m.lower()]
    print("Danh sách Model Gemini / Flash khả dụng:")
    for m in gemini_models:
        print("  ✓", m)
else:
    print("Lỗi kết nối:", r.text)
```

---

## 7. MẸO XỬ LÝ & KINH NGHIỆM THỰC CHIẾN

1. **Chuẩn hóa tiền tố `google/`:** Một số giao diện Beeknoee chấp nhận cả `gemini-3.1-flash-lite` và `google/gemini-3.1-flash-lite`. Tuy nhiên dạng có tiền tố `google/...` luôn được hệ thống ưu tiên định tuyến chính xác nhất.
2. **Cơ chế Fallback thông minh:** Luôn bọc lệnh gọi AI trong khối `try...except`. Nếu API Beeknoee gặp sự cố mạng hoặc hết hạn ngạch, hệ thống tự động fallback về **Local Offline Engine** dựa trên dữ liệu mẫu GDPT 2018 để không bao giờ làm gián đoạn trải nghiệm người dùng.
3. **Tái sử dụng cho các dự án sau:** Bạn có thể copy toàn bộ file này sang thư mục bất kỳ để làm hướng dẫn chuẩn cho đồng nghiệp hoặc các dự án mới mà không cần phải tra cứu lại cấu hình.
