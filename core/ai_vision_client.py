import os
import requests
import time
import logging
from typing import Optional

logger = logging.getLogger(__name__)


DEPRECATED_MODELS_MAP = {
    "gemini-2.5-flash-lite": "google/gemini-3.1-flash-lite",
    "google/gemini-2.5-flash-lite": "google/gemini-3.1-flash-lite",
    "gemini-2.5-flash": "google/gemini-3.5-flash",
    "google/gemini-2.5-flash": "google/gemini-3.5-flash",
    "gemini-2.5-pro": "google/gemini-3.5-flash",
    "google/gemini-2.5-pro": "google/gemini-3.5-flash",
    "gemini-3.1-flash-lite": "google/gemini-3.1-flash-lite",
    "google/gemini-3.1-flash-lite": "google/gemini-3.1-flash-lite",
}


class AIVisionClient:
    """Client giao tiếp với AI Vision API (OpenAI / Gemini compatible endpoint)."""

    def __init__(self, api_key: str, base_url: str = "https://platform.beeknoee.com/api/v1", model: Optional[str] = None):
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        raw_model = model or os.getenv("VISION_MODEL", "google/gemini-3.1-flash-lite")
        normalized = DEPRECATED_MODELS_MAP.get(raw_model, raw_model)
        if "2.5" in normalized:
            normalized = "google/gemini-3.1-flash-lite"
        self.model = normalized
        self.max_retries = 3
        self.retry_delay = 5  # seconds

    def _should_retry(self, status_code: int) -> bool:
        return status_code in [429, 500, 502, 503, 504]

    def chat_with_image(self, image_base64: str, prompt: str, mime_type: str = "image/png") -> str:
        payload = {
            "model": self.model,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_base64}"}}
                ]
            }],
            "temperature": 0.1,
            "max_tokens": 32000
        }

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        for attempt in range(1, self.max_retries + 1):
            try:
                resp = requests.post(
                    f"{self.base_url}/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=180
                )

                if resp.status_code == 200:
                    data = resp.json()
                    return data["choices"][0]["message"]["content"]

                if self._should_retry(resp.status_code):
                    logger.warning(f"API returned status {resp.status_code}. Retry {attempt}/{self.max_retries}")
                    if attempt < self.max_retries:
                        time.sleep(self.retry_delay)
                        continue

                raise Exception(f"API Error {resp.status_code}: {resp.text}")

            except requests.exceptions.Timeout:
                logger.warning(f"API request timeout. Retry {attempt}/{self.max_retries}")
                if attempt < self.max_retries:
                    time.sleep(self.retry_delay)
                    continue
                raise Exception("API Error: Request timeout sau các lần thử lại.")
            except requests.exceptions.RequestException as e:
                logger.error(f"Network error: {e}")
                if attempt < self.max_retries:
                    time.sleep(self.retry_delay)
                    continue
                raise Exception(f"API Error: Lỗi kết nối mạng - {e}")
