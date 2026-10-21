"""
core/ai_exam_generator.py
Bộ sinh đề thi thông minh sử dụng API Beeknoee (Gemini/OpenAI-compatible)
Phân loại chính xác các hạng mục:
1. Mức độ nhận thức (Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao) theo đúng ma trận %
2. Chuẩn Yêu cầu cần đạt GDPT 2018 từng chủ đề/bài học
3. Phân biệt rõ ràng 3 hình thức: TNKQ 4 lựa chọn, Đúng/Sai (4 ý a-b-c-d), và Tự luận (kèm biểu điểm chi tiết)
4. Tự động sinh đồng thời 2 mã đề hoán vị (TH01 & TH02) và Bảng đáp án tương ứng.
"""

import os
import sys
import copy
import json
import random
import logging
import requests
from typing import Dict, Any, Optional, List
from dotenv import load_dotenv

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

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
    "gemini-3.5-flash-lite": "google/gemini-3.5-flash-lite",
    "google/gemini-3.5-flash-lite": "google/gemini-3.5-flash-lite",
    "gemini-3.5-flash": "google/gemini-3.5-flash",
    "google/gemini-3.5-flash": "google/gemini-3.5-flash",
}

DEFAULT_STABLE_MODEL = "google/gemini-3.1-flash-lite"


class AIExamGenerator:
    """Client AI chuyên sâu để ra đề thi môn Tin học chuẩn GDPT 2018 qua Beeknoee API."""

    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None, model: Optional[str] = None):
        load_dotenv(override=True)
        self.api_key = api_key or os.getenv("BEE_API_KEY", "")
        self.base_url = (base_url or os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")).rstrip("/")
        raw_model = model or os.getenv("VISION_MODEL", DEFAULT_STABLE_MODEL)
        
        # Tự động ánh xạ model cũ (2.5) sang model mới nhất tương thích Beeknoee
        normalized_model = DEPRECATED_MODELS_MAP.get(raw_model, raw_model)
        if "2.5" in normalized_model:
            normalized_model = DEFAULT_STABLE_MODEL
        self.model = normalized_model

    def is_available(self) -> bool:
        """Kiểm tra xem API key đã được cấu hình chưa."""
        return bool(self.api_key and self.api_key.startswith("sk-"))

    def generate_pedagogical_exam(
        self,
        grade_level: str,
        exam_title: str,
        school_name: str,
        governing_body: str,
        school_year: str,
        teacher_name: str,
        head_name: str,
        pct_nb: float,
        pct_th: float,
        pct_vd: float,
        pct_vdc: float = 0.0,
        topics_info: Optional[List[Dict[str, Any]]] = None,
        topic_ratios: Optional[Dict[str, float]] = None,
        custom_requirements: Optional[List[Dict[str, Any]]] = None,
        allowed_units: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        Gọi Beeknoee API để sinh bộ hồ sơ đề thi thông minh,
        phân biệt chính xác ma trận, bản đặc tả và các hạng mục câu hỏi bám sát YCCĐ kế hoạch bài dạy.
        """
        if not self.is_available():
            raise RuntimeError("Chưa cấu hình BEE_API_KEY hợp lệ trong file .env!")

        # Chuẩn hóa tỷ lệ %
        total_pct = pct_nb + pct_th + pct_vd + pct_vdc
        if total_pct <= 0: total_pct = 100.0
        nb_norm = round((pct_nb / total_pct) * 100, 1)
        th_norm = round((pct_th / total_pct) * 100, 1)
        vd_norm = round((pct_vd / total_pct) * 100, 1)
        vdc_norm = round(100.0 - (nb_norm + th_norm + vd_norm), 1)

        # Tính toán số lượng câu hỏi theo ma trận
        # Chuẩn đề: 12 câu trắc nghiệm (3.0đ) + 4 câu Đúng/Sai (4.0đ) + 3 câu Tự luận (3.0đ) = 10.0đ
        nb_mcq = max(3, int(round(12 * (nb_norm / 100.0))))
        th_mcq = max(3, int(round(12 * (th_norm / 100.0))))
        vd_mcq = 12 - (nb_mcq + th_mcq)
        if vd_mcq < 0:
            vd_mcq = 1
            if th_mcq > 3: th_mcq -= 1
            else: nb_mcq -= 1

        topic_req_str = ""
        if topic_ratios:
            topic_req_str = "\nPHÂN BỔ TỶ LỆ THEO CÁC CHỦ ĐỀ / CHƯƠNG SGK:\n" + "\n".join([f"- {k}: {v}%" for k, v in topic_ratios.items() if v > 0]) + "\n(Hãy phân chia số lượng câu hỏi Phần I, Phần II và Phần III bám sát tỷ lệ % các chủ đề trên)"

        scope_guard_str = ""
        if allowed_units:
            units_list_txt = "\n".join([f"  + {u}" for u in allowed_units])
            scope_guard_str = (
                f"\n🛑 GIỚI HẠN PHẠM VI BÀI HỌC ĐÃ DẠY THEO KẾ HOẠCH GIÁO DỤC (BẮT BUỘC TUÂN THỦ 100%):\n"
                f"Học sinh chỉ mới học đến các bài sau đây:\n{units_list_txt}\n"
                f"⚠️ NGUYÊN TẮC BẢO VỆ: TUYỆT ĐỐI KHÔNG RA BẤT KỲ CÂU HỎI NÀO VÀO CÁC BÀI CHƯA HỌC NGOÀI DANH SÁCH TRÊN. "
                f"Tất cả câu hỏi trắc nghiệm, đúng/sai và tự luận phải 100% thuộc danh sách các bài học đã học này!\n"
            )

        custom_yccd_str = ""
        if custom_requirements:
            custom_yccd_str = "\nKẾ HOẠCH BÀI DẠY VÀ YÊU CẦU CẦN ĐẠT CỤ THỂ CỦA TỪNG PHẦN (BẮT BUỘC BÁM SÁT 100%):\n"
            for cr_idx, cr in enumerate(custom_requirements, 1):
                ch = cr.get("chapter", "")
                un = cr.get("unit", "")
                req = cr.get("requirement", "")
                custom_yccd_str += f"{cr_idx}. [{ch}] - {un}:\n   {req}\n"
            custom_yccd_str += (
                "\n⚠️ QUY TẮC BẮT BUỘC VỚI BEEKNOEE AI:\n"
                "- TẤT CẢ các câu hỏi trắc nghiệm (Phần I), đúng/sai (Phần II) và tự luận (Phần III) "
                "PHẢI bắt nguồn trực tiếp và ánh xạ chính xác vào các Yêu cầu cần đạt (YCCĐ) nêu trên.\n"
                "- Tuyệt đối không sinh câu hỏi ngoài phạm vi kiến thức kế hoạch bài dạy này.\n"
                "- Sử dụng ngữ pháp, thuật ngữ tin học chuẩn xác, chuẩn mực sư phạm GDPT 2018 Bộ GD&ĐT, không bốc ngữ pháp lung tung.\n"
            )

        prompt = f"""
Bạn là Chuyên gia Khảo thí và Đánh giá Giáo dục môn Tin học Chương trình GDPT 2018 (Bộ Giáo dục và Đào tạo Việt Nam).
Nhiệm vụ của bạn là BIÊN SOẠN BỘ HỒ SƠ ĐỀ THI TIN HỌC {grade_level.upper()} đạt chuẩn 100% theo các quy định sư phạm.

THÔNG TIN KỲ THI:
- Đơn vị chủ quản: {governing_body}
- Trường học: {school_name}
- Kỳ kiểm tra: {exam_title}
- Khối lớp: {grade_level}
- Năm học: {school_year}
- Giáo viên ra đề: {teacher_name}
- Tổ trưởng chuyên môn: {head_name}
- Thời gian làm bài: 45 phút (Trắc nghiệm: 31,5 phút; Tự luận: 13,5 phút)

MA TRẬN NẬN THỨC CẦN PHÂN LOẠI CHÍNH XÁC:
- Nhận biết: {nb_norm}%
- Thông hiểu: {th_norm}%
- Vận dụng: {vd_norm}%
- Vận dụng cao: {vdc_norm}%
- Điểm toàn bài: 10,0 điểm (7,0 điểm Trắc nghiệm - 3,0 điểm Tự luận)
{scope_guard_str}
{topic_req_str}
{custom_yccd_str}

YÊU CẦU CẤU TRÚC ĐỀ THI BẮT BUỘC (PHÂN BIỆT RÕ RÀNG TỪNG HẠNG MỤC):
1. PHẦN I: CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (12 câu x 0,25 điểm = 3,0 điểm)
   - Gồm chính xác {nb_mcq} câu Nhận biết, {th_mcq} câu Thông hiểu, {vd_mcq} câu Vận dụng.
   - Mỗi câu gồm 4 lựa chọn A, B, C, D rõ ràng, chỉ có 1 đáp án đúng duy nhất. Nội dung bám sát SGK Tin học {grade_level} GDPT 2018.

2. PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG / SAI (4 câu x 1,0 điểm = 4,0 điểm)
   - Mỗi câu đưa ra một ngữ cảnh/tình huống thực tiễn số kèm 4 phát biểu a), b), c), d).
   - Mỗi ý a, b, c, d được gán nhãn Đúng (Đ) hoặc Sai (S) kèm lý do/căn cứ giải thích sư phạm.
   - Phân hóa độ khó: có ý Nhận biết, có ý Thông hiểu, có ý Vận dụng.

3. PHẦN III: PHẦN TỰ LUẬN (3 câu x 1,0 điểm = 3,0 điểm)
   - Câu 1 (1,0đ): Mức độ Thông hiểu (giải thích vai trò, phân biệt khái niệm, nêu nguyên nhân).
   - Câu 2 (1,0đ): Mức độ Vận dụng (xử lý tình huống an toàn số, quản lý tệp, bảng tính hoặc thuật toán).
   - Câu 3 (1,0đ): Mức độ Vận dụng (bài toán thực tiễn gắn với đời sống học sinh, ứng xử văn hóa số).
   - Mỗi câu tự luận phải có BIỂU ĐIỂM (Rubric) chi tiết từng tiêu chí (0,25đ; 0,5đ; 1,0đ).

4. BẢNG KHUNG MA TRẬN VÀ BẢNG ĐẶC TẢ SƯ PHẠM:
   - Liệt kê các Chủ đề/Chương và Bài học trọng tâm của học kỳ I môn Tin học {grade_level}.
   - Trình bày rõ Yêu cầu cần đạt chuẩn GDPT 2018 tương ứng với từng mức độ nhận thức (Biết: nêu được..., Hiểu: giải thích được..., Vận dụng: thao tác được...).

🛑 QUY TẮC ĐẶC BIỆT BẮT BUỘC VỀ SỐ LƯỢNG CÂU HỎI (TUYỆT ĐỐI KHÔNG ĐƯỢC RÚT GỌN HOẶC TRẢ VỀ THIẾU):
- "mcqs": BẮT BUỘC PHẢI CHỨA ĐỦ CHÍNH XÁC 12 ĐỐI TƯỢNG CÂU HỎI (từ MCQ_01 đến MCQ_12).
  Gồm đúng: {nb_mcq} câu Nhận biết, {th_mcq} câu Thông hiểu, {vd_mcq} câu Vận dụng.
  Mỗi câu 4 phương án A, B, C, D rõ ràng, chỉ có 1 đáp án đúng.
- "tfs": BẮT BUỘC PHẢI CHỨA ĐỦ CHÍNH XÁC 4 ĐỐI TƯỢNG CÂU HỎI (từ TF_01 đến TF_04).
  Mỗi câu có chủ đề tình huống và ĐỦ 4 PHÁT BIỂU a, b, c, d gán nhãn Đúng (Đ) hoặc Sai (S) kèm căn cứ giải thích.
- "essays": BẮT BUỘC PHẢI CHỨA ĐỦ CHÍNH XÁC 3 ĐỐI TƯỢNG CÂU HỎI (từ ES_01 đến ES_03).
  Mỗi câu 1.0 điểm kèm rubric barem điểm từng ý (0.25đ, 0.5đ, 1.0đ).

TỔNG CỘNG ĐỀ PHẢI ĐỦ 19 CÂU (12 TN + 4 Đ/S + 3 TỰ LUẬN) = 10,0 ĐIỂM.
SCHEMA DƯỚI ĐÂY CHỈ MINH HOẠ CẤU TRÚC PHẦN TỬ. KHI TRẢ VỀ, BẠN BẮT BUỘC PHẢI SINH ĐỦ 12 CÂU MCQS, 4 CÂU TFS VÀ 3 CÂU ESSAYS!

HÃY TRẢ VỀ KẾT QUẢ DƯỚI DẠNG MỘT ĐỐI TƯỢNG JSON DUY NHẤT (không dùng markdown backticks ```json ở đầu và cuối), theo schema sau:
{{
  "meta": {{
    "grade": "{grade_level}",
    "exam_title": "{exam_title}",
    "school_name": "{school_name}",
    "governing_body": "{governing_body}",
    "school_year": "{school_year}",
    "teacher_name": "{teacher_name}",
    "head_name": "{head_name}",
    "duration": 45,
    "pct_nb": {nb_norm},
    "pct_th": {th_norm},
    "pct_vd": {vd_norm},
    "pct_vdc": {vdc_norm}
  }},
  "matrix_rows": {{
    "units": [
      {{
        "tt": "1",
        "chapter": "Tên chủ đề 1",
        "unit": "Tên bài 1",
        "mcq_nb": "2 (0,5)",
        "mcq_th": "2 (0,5)",
        "mcq_vd": "",
        "tf_nb": "1 (1,0)",
        "tf_th": "",
        "tf_vd": "",
        "es_nb": "",
        "es_th": "1 (1,0)",
        "es_vd": "",
        "total_score": "3,0đ (30%)"
      }}
    ],
    "totals": {{
      "total_mcq": "12 câu (3,0đ)",
      "total_tf": "4 câu (4,0đ)",
      "total_es": "3 câu (3,0đ)",
      "sum_score": "10,0 điểm",
      "ratio_str": "{nb_norm}% Biết - {th_norm}% Hiểu - {vd_norm}% Vận dụng"
    }}
  }},
  "spec_rows": [
    {{
      "tt": "1",
      "chapter": "Tên chủ đề 1",
      "unit": "Tên bài 1",
      "requirement": "Nhận biết: ...\\n\\nThông hiểu: ...\\n\\nVận dụng: ...",
      "mcq_nb": "2",
      "mcq_th": "2",
      "mcq_vd": "",
      "tf_nb": "1",
      "tf_th": "",
      "tf_vd": "",
      "es_nb": "",
      "es_th": "1",
      "es_vd": ""
    }}
  ],
  "th01": {{
    "code": "TH01",
    "mcqs": [
      {{
        "id": "MCQ_01",
        "unit": "Tên bài",
        "level": "Nhận biết",
        "content": "Nội dung câu hỏi 1...",
        "options": {{
          "A": "Đáp án A",
          "B": "Đáp án B",
          "C": "Đáp án C",
          "D": "Đáp án D"
        }},
        "correct": "A"
      }},
      {{
        "id": "MCQ_02",
        "unit": "Tên bài",
        "level": "Nhận biết",
        "content": "Nội dung câu hỏi 2...",
        "options": {{
          "A": "Đáp án A",
          "B": "Đáp án B",
          "C": "Đáp án C",
          "D": "Đáp án D"
        }},
        "correct": "B"
      }}
    ],
    "tfs": [
      {{
        "id": "TF_01",
        "unit": "Tên bài",
        "topic": "Mệnh đề dẫn tình huống 1...",
        "statements": [
          {{"label": "a", "content": "Phát biểu a", "correct": "Đ", "reason": "Lý do vì sao đúng"}},
          {{"label": "b", "content": "Phát biểu b", "correct": "S", "reason": "Lý do vì sao sai"}},
          {{"label": "c", "content": "Phát biểu c", "correct": "Đ", "reason": "Lý do vì sao đúng"}},
          {{"label": "d", "content": "Phát biểu d", "correct": "S", "reason": "Lý do vì sao sai"}}
        ]
      }}
    ],
    "essays": [
      {{
        "id": "ES_01",
        "unit": "Tên bài",
        "level": "Thông hiểu",
        "score": 1.0,
        "content": "Nội dung câu hỏi tự luận 1...",
        "rubric": [
          {{"criterion": "Ý 1 cần trả lời", "score": 0.25}},
          {{"criterion": "Ý 2 cần trả lời", "score": 0.25}},
          {{"criterion": "Ý 3 cần trả lời", "score": 0.25}},
          {{"criterion": "Ý 4 cần trả lời", "score": 0.25}}
        ]
      }}
    ]
  }}
}}
"""

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "Bạn là hệ thống AI Khảo thí chuyên sâu của Bộ GD&ĐT Việt Nam. "
                        "Bạn PHẢI trả về duy nhất chuỗi JSON hợp lệ 100%, không được chứa bất kỳ từ ngữ nào ngoài JSON."
                    )
                },
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.2,
            "max_tokens": 32000
        }

        logger.info(f"Đang gọi Beeknoee API ({self.model}) để sinh đề thi {grade_level}...")
        try:
            resp = requests.post(f"{self.base_url}/chat/completions", headers=headers, json=payload, timeout=90)
        except requests.exceptions.Timeout:
            raise RuntimeError("Hết thời gian chờ kết nối Beeknoee AI (Request Timeout sau 90s). Vui lòng thử lại hoặc sử dụng Local Engine.")
        except requests.exceptions.RequestException as req_err:
            raise RuntimeError(f"Lỗi mạng khi kết nối Beeknoee API: {str(req_err)}")

        if resp.status_code != 200:
            err_body = resp.text
            # Tự động thử lại với model ổn định nếu model được chọn bị 404 (bị xóa/vô hiệu hóa)
            if resp.status_code == 404 and self.model != DEFAULT_STABLE_MODEL:
                logger.warning(f"Model {self.model} không tồn tại trên Beeknoee (404). Tự động fallback sang {DEFAULT_STABLE_MODEL}...")
                self.model = DEFAULT_STABLE_MODEL
                payload["model"] = DEFAULT_STABLE_MODEL
                try:
                    resp = requests.post(f"{self.base_url}/chat/completions", headers=headers, json=payload, timeout=90)
                    if resp.status_code == 200:
                        err_body = ""
                    else:
                        err_body = resp.text
                except Exception as ex_retry:
                    raise RuntimeError(f"Lỗi khi retry với model {DEFAULT_STABLE_MODEL}: {ex_retry}")

            if resp.status_code != 200:
                if "<!DOCTYPE html>" in err_body or "<html" in err_body or resp.status_code in [520, 521, 522, 523, 524, 502, 503, 504]:
                    raise RuntimeError(
                        f"Máy chủ Beeknoee AI tạm thời gián đoạn (Cloudflare HTTP {resp.status_code} - Origin Server Timeout/Unavailable)."
                    )
                raise RuntimeError(f"Lỗi Beeknoee API (HTTP {resp.status_code}): {err_body[:250]}")

        raw_content = resp.json()["choices"][0]["message"]["content"].strip()

        # Làm sạch markdown nếu AI lỡ bọc trong ```json
        clean_content = raw_content
        if clean_content.startswith("```"):
            lines = clean_content.split("\n")
            if lines[0].startswith("```"): lines = lines[1:]
            if lines and lines[-1].startswith("```"): lines = lines[:-1]
            clean_content = "\n".join(lines).strip()

        try:
            dossier = json.loads(clean_content)
        except Exception as e:
            logger.error(f"Lỗi parse JSON từ Beeknoee API: {e}\nRaw: {raw_content[:500]}")
            # Trích xuất đoạn JSON giữa { và }
            start_idx = clean_content.find("{")
            end_idx = clean_content.rfind("}")
            if start_idx != -1 and end_idx != -1:
                sub_json = clean_content[start_idx:end_idx + 1]
                dossier = json.loads(sub_json)
            else:
                raise RuntimeError(f"AI không trả về JSON hợp lệ: {str(e)}")

        # ----------------------------------------------------------------------
        # KIỂM TRA & TỰ ĐỘNG BÙ ĐẮP CÂU HỎI (ĐẢM BẢO LUÔN ĐỦ 12 TN, 4 Đ/S, 3 TL)
        # ----------------------------------------------------------------------
        th01 = dossier.get("th01", {})
        mcqs = th01.get("mcqs", [])
        tfs = th01.get("tfs", [])
        essays = th01.get("essays", [])

        from core.exam_builder import ExamMatrixEngine
        fallback_engine = ExamMatrixEngine(
            grade_level=grade_level,
            exam_title=exam_title,
            school_name=school_name,
            governing_body=governing_body,
            school_year=school_year,
            teacher_name=teacher_name,
            head_name=head_name,
            pct_nb=nb_norm,
            pct_th=th_norm,
            pct_vd=vd_norm,
            pct_vdc=vdc_norm,
            topic_ratios=topic_ratios,
            allowed_units=allowed_units
        )
        fallback_dossier = fallback_engine.generate_exam_dossier()
        fb_th01 = fallback_dossier.get("th01", {})

        # 1. Bù MCQ nếu thiếu (cần đúng 12 câu)
        if len(mcqs) < 12:
            existing_mcq_txt = {q.get("content", "").strip() for q in mcqs}
            for fb_q in fb_th01.get("mcqs", []):
                if fb_q.get("content", "").strip() not in existing_mcq_txt:
                    q_add = copy.deepcopy(fb_q)
                    q_add["id"] = f"MCQ_{len(mcqs) + 1:02d}"
                    mcqs.append(q_add)
                    existing_mcq_txt.add(q_add.get("content", "").strip())
                if len(mcqs) >= 12:
                    break

        # 2. Bù Đúng/Sai nếu thiếu (cần đúng 4 câu)
        if len(tfs) < 4:
            existing_tf_txt = {t.get("topic", "").strip() for t in tfs}
            for fb_t in fb_th01.get("tfs", []):
                if fb_t.get("topic", "").strip() not in existing_tf_txt:
                    t_add = copy.deepcopy(fb_t)
                    t_add["id"] = f"TF_{len(tfs) + 1:02d}"
                    tfs.append(t_add)
                    existing_tf_txt.add(t_add.get("topic", "").strip())
                if len(tfs) >= 4:
                    break

        # 3. Bù Tự luận nếu thiếu (cần đúng 3 câu)
        if len(essays) < 3:
            existing_es_txt = {e.get("content", "").strip() for e in essays}
            for fb_e in fb_th01.get("essays", []):
                if fb_e.get("content", "").strip() not in existing_es_txt:
                    e_add = copy.deepcopy(fb_e)
                    e_add["id"] = f"ES_{len(essays) + 1:02d}"
                    essays.append(e_add)
                    existing_es_txt.add(e_add.get("content", "").strip())
                if len(essays) >= 3:
                    break

        th01["mcqs"] = mcqs[:12]
        th01["tfs"] = tfs[:4]
        th01["essays"] = essays[:3]
        dossier["th01"] = th01

        # 4. Đảm bảo Bảng Ma Trận & Bản Đặc Tả luôn tuân thủ 100% chuẩn toán học và bám sát bài học KHGD
        dossier["matrix_rows"] = fallback_dossier["matrix_rows"]
        dossier["spec_rows"] = fallback_dossier["spec_rows"]

        # ----------------------------------------------------------------------
        # TỰ ĐỘNG SINH MÃ ĐỀ HOÁN VỊ TH02 TỪ ĐỀ TH01 DO AI TẠO RA
        # ----------------------------------------------------------------------
        th02_mcqs = []
        for q in th01.get("mcqs", []):
            q_copy = json.loads(json.dumps(q))
            opt_keys = ["A", "B", "C", "D"]
            opt_vals = [q_copy["options"][k] for k in opt_keys]
            correct_val = q_copy["options"][q_copy["correct"]]

            # Xáo trộn phương án cho TH02
            shuffled_vals = list(opt_vals)
            random.seed(len(q_copy["content"]) + 123)
            random.shuffle(shuffled_vals)
            new_opts = {k: v for k, v in zip(opt_keys, shuffled_vals)}

            # Xác định đáp án đúng mới
            new_corr = "A"
            for k, v in new_opts.items():
                if v == correct_val:
                    new_corr = k
                    break
            q_copy["options"] = new_opts
            q_copy["correct"] = new_corr
            th02_mcqs.append(q_copy)

        # Đảo thứ tự các câu hỏi cho mã đề TH02
        random.seed(999)
        th02_mcqs_shuffled = list(th02_mcqs)
        random.shuffle(th02_mcqs_shuffled)

        # Với Đúng/Sai: đảo các phát biểu a, b, c, d
        th02_tfs = []
        for tf_item in th01.get("tfs", []):
            tf_copy = json.loads(json.dumps(tf_item))
            stmts = list(tf_copy.get("statements", []))
            random.seed(len(tf_item.get("topic", "")) + 456)
            random.shuffle(stmts)
            for idx, lbl in enumerate(["a", "b", "c", "d"]):
                if idx < len(stmts):
                    stmts[idx]["label"] = lbl
            tf_copy["statements"] = stmts
            th02_tfs.append(tf_copy)

        dossier["th02"] = {
            "code": "TH02",
            "mcqs": th02_mcqs_shuffled,
            "tfs": th02_tfs,
            "essays": list(th01.get("essays", []))
        }

        return dossier
