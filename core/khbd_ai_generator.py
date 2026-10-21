# -*- coding: utf-8 -*-
"""
core/khbd_ai_generator.py
Bộ sinh Kế hoạch bài dạy (KHBD) chuẩn Công văn 5512/BGDĐT-GDTrH
và Tích hợp Khung Năng lực số (NLS) theo Thông tư 02/2025/TT-BGDĐT
Sử dụng API AI Beekone (OpenAI / Gemini compatible endpoint).
"""

import os
import sys
import json
import logging
import re
import requests
import openpyxl
from typing import Dict, Any, List, Optional
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


class KHBDAIGenerator:
    """Quản lý sinh Kế hoạch bài dạy (KHBD) chuẩn 5512 tích hợp NLS qua Beekone API."""

    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None, model: Optional[str] = None):
        load_dotenv(override=True)
        self.api_key = api_key or os.getenv("BEE_API_KEY", "")
        self.base_url = (base_url or os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")).rstrip("/")
        raw_model = model or os.getenv("VISION_MODEL", DEFAULT_STABLE_MODEL)
        normalized = DEPRECATED_MODELS_MAP.get(raw_model, raw_model)
        if "2.5" in normalized:
            normalized = DEFAULT_STABLE_MODEL
        self.model = normalized

    def is_available(self) -> bool:
        return bool(self.api_key and len(self.api_key.strip()) > 5)

    @staticmethod
    def load_ppct_excel(excel_path: str) -> List[Dict[str, Any]]:
        """
        Đọc danh sách phân phối chương trình từ file Excel (ví dụ _Lop_6A.xlsx).
        Trả về danh sách các tiết học:
        [
            {
                'stt': 1,
                'week': '1',
                'period': '1',
                'lesson_title': 'Bài 1: Thông tin và dữ liệu',
                'date_range': '07/09/2026 - 11/09/2026',
                'output_filename': 'tinhoc6_tuan01_Bai01_Thong-tin-va-du-lieu_Tiet01.docx'
            }, ...
        ]
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Không tìm thấy file Excel: {excel_path}")

        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_name = "PPCT" if "PPCT" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        lessons = []
        # Dò tìm dòng tiêu đề STT, Tuần, Tiết PPCT
        header_row = 9
        for r in range(1, 15):
            val1 = str(ws.cell(r, 1).value or "").strip().upper()
            val2 = str(ws.cell(r, 2).value or "").strip().upper()
            if "STT" in val1 or "TUẦN" in val2 or "TUAN" in val2:
                header_row = r
                break

        for r in range(header_row + 1, ws.max_row + 1):
            stt = ws.cell(r, 1).value
            tuan = ws.cell(r, 2).value
            tiet = ws.cell(r, 3).value
            ten_bai = ws.cell(r, 4).value
            ngay = ws.cell(r, 5).value
            ghi_chu = ws.cell(r, 6).value

            if stt is None and ten_bai is None:
                continue

            ten_bai_str = str(ten_bai).strip() if ten_bai else ""
            if not ten_bai_str:
                continue

            stt_int = int(stt) if stt is not None and str(stt).isdigit() else len(lessons) + 1
            tuan_str = str(tuan).strip() if tuan is not None else str(stt_int)
            tiet_str = str(tiet).strip() if tiet is not None else str(stt_int)
            ngay_str = str(ngay).strip() if ngay is not None else ""

            # Tạo tên file gợi ý nếu ghi chú trống
            if ghi_chu and str(ghi_chu).strip().endswith(".docx"):
                out_fn = str(ghi_chu).strip()
            else:
                clean_title = re.sub(r'[^\w\s-]', '', ten_bai_str).replace(' ', '_')
                out_fn = f"tinhoc6_tuan{tuan_str.zfill(2)}_Tiet{tiet_str.zfill(2)}_{clean_title}.docx"

            lessons.append({
                "stt": stt_int,
                "week": tuan_str,
                "period": tiet_str,
                "lesson_title": ten_bai_str,
                "date_range": ngay_str,
                "output_filename": out_fn
            })

        return lessons

    def generate_khbd_json(
        self,
        lesson_info: Dict[str, Any],
        teacher_name: str = "Trần Đức Tình",
        department: str = "Tổ Tự Nhiên",
        signer_cm: str = "Lưu Văn Chí",
        signer_to: str = "Hoàng Thị Kim Tuyết",
        subject: str = "Tin học",
        grade: str = "6",
        school_name: str = "TRƯỜNG TH & THCS ĐĂK TRÔI"
    ) -> Dict[str, Any]:
        """
        Gọi API Beekone AI để sinh nội dung Kế hoạch bài dạy chuẩn 5512 tích hợp NLS.
        """
        week_str = str(lesson_info.get("week", "1"))
        period_str = str(lesson_info.get("period", "1"))
        date_str = str(lesson_info.get("date_range", "07/09/2026 - 11/09/2026"))
        lesson_title = str(lesson_info.get("lesson_title", "Bài học"))

        system_prompt = (
            "Bạn là Chuyên gia Phương pháp Giảng dạy môn Tin học và Soạn Kế hoạch bài dạy (KHBD) "
            "chuẩn Bộ Giáo dục & Đào tạo Việt Nam.\n"
            "Nhiệm vụ của bạn là biên soạn một Kế hoạch bài dạy hoàn chỉnh, chi tiết và thực tế, "
            "tuân thủ nghiêm ngặt:\n"
            "1. Cấu trúc Công văn 5512/BGDĐT-GDTrH: gồm 5 phần chính:\n"
            "   I. Mục tiêu (Kiến thức, Năng lực, Phẩm chất)\n"
            "   II. Thiết bị dạy học và học liệu (GV, HS)\n"
            "   III. Tiến trình dạy học (4 hoạt động: Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng; "
            "mỗi hoạt động có 4 bước rõ ràng: Bước 1: Chuyển giao nhiệm vụ, Bước 2: Thực hiện nhiệm vụ, "
            "Bước 3: Báo cáo thảo luận, Bước 4: Kết luận nhận định kèm Dự kiến sản phẩm)\n"
            "   IV. Bảng tổng hợp Năng lực số (NLS) tích hợp trong bài dạy\n"
            "   V. Khối phê duyệt chuyên môn\n"
            "2. Tích hợp Khung Năng lực số (NLS) theo Thông tư 02/2025/TT-BGDĐT cấp THCS (Tiêu chuẩn 2 - TC2):\n"
            "   - Chọn lọc 2 đến 4 mã chỉ số NLS phù hợp nhất với bài học (ví dụ: 1.2.TC2a, 2.1.TC2a, 2.2.TC2a, 3.2.TC2a, 4.1.TC2a, 5.2.TC2a, 6.2.TC2a,...)\n"
            "   - Đưa vào mục 2.3 trong Mục tiêu.\n"
            "   - Trong tiến trình dạy học (III), tại các vị trí học sinh/giáo viên thực hiện thao tác số "
            "(sử dụng máy tính, phần mềm, tra cứu Internet, mã QR, chia sẻ học liệu trực tuyến, nộp bài), "
            "hãy đánh dấu bằng thẻ `[NLS-RED]nội dung thao tác số [Mã_Chỉ_Số][/NLS-RED]` để hệ thống tự động bôi đỏ chữ theo chuẩn Thông tư 02/2025.\n"
            "   - Lập Bảng tổng hợp NLS ở Phần IV gồm TT, Nội dung hoạt động số của học sinh, Mã NLS đạt được.\n"
            "3. BẮT BUỘC TRẢ VỀ ĐỊNH DẠNG JSON HỢP LỆ (Không kèm bất kỳ lời giải thích nào khác ngoài chuỗi JSON)."
        )

        user_prompt = f"""Dựa vào thông tin bài dạy sau, hãy biên soạn Kế hoạch bài dạy hoàn chỉnh và trả về dưới dạng JSON thuần túy:

[THÔNG TIN BÀI DẠY]
- Trường: {school_name}
- Môn học: {subject}
- Lớp: {grade}
- Tên bài dạy: {lesson_title}
- Tuần: {week_str}
- Tiết: {period_str}
- Khoảng ngày dạy: {date_str}
- Thời lượng: 1 tiết (45 phút)
- Giáo viên soạn: {teacher_name}
- Tổ chuyên môn: {department}
- Duyệt chuyên môn: {signer_cm}
- Duyệt tổ chuyên môn: {signer_to}

CẤU TRÚC JSON MẪU BẮT BUỘC TUÂN THỦ:
{{
  "header": {{
    "week": "{week_str}",
    "period": "{period_str}",
    "date_range": "{date_str}",
    "teacher_name": "{teacher_name}",
    "department": "{department}",
    "school_name": "{school_name}"
  }},
  "title": "{lesson_title.upper()}",
  "sub_title": "Môn học: {subject}; Lớp: {grade} | Thời gian thực hiện: 1 tiết (45 phút)",
  "objectives": {{
    "knowledge": [
      "Trình bày được khái niệm cơ bản...",
      "Nhận biết và phân biệt được..."
    ],
    "competencies": {{
      "general": [
        "Năng lực tự chủ và tự học: Tự giác tìm hiểu tài liệu SGK, hoàn thành phiếu học tập cá nhân.",
        "Năng lực giao tiếp và hợp tác: Tương tác nhóm, thảo luận phân chia nhiệm vụ hiệu quả."
      ],
      "specific": [
        "Năng lực nhận thức tin học: Giải thích được bản chất...",
        "Năng lực ứng dụng công nghệ: Thao tác chính xác các bước..."
      ],
      "digital": [
        {{
          "code": "1.2.TC2a",
          "desc": "Sử dụng công cụ số/phần mềm chuyên dụng phù hợp để thao tác và xử lý dữ liệu học tập."
        }},
        {{
          "code": "2.1.TC2a",
          "desc": "Tìm kiếm, chọn lọc và khai thác thông tin từ môi trường số an toàn, hiệu quả."
        }}
      ]
    }},
    "qualities": [
      "Chăm chỉ: Tích cực tham gia các hoạt động học tập, kiên trì hoàn thành nhiệm vụ.",
      "Trách nhiệm: Có ý thức bảo quản thiết bị phòng máy và tôn trọng bản quyền dữ liệu."
    ]
  }},
  "equipment": {{
    "teacher": [
      "Kế hoạch bài dạy, bài giảng điện tử PowerPoint/Canva.",
      "Máy chiếu, máy tính kết nối Internet, phiếu học tập số / liên kết chia sẻ."
    ],
    "student": [
      "Sách giáo khoa Tin học {grade}, vở ghi chép.",
      "Máy tính phòng thực hành / thiết bị số để thực hiện nhiệm vụ học tập."
    ]
  }},
  "activities": [
    {{
      "id": 1,
      "name": "Hoạt động 1: Khởi động (Mở đầu / Xác định vấn đề)",
      "time": "5 - 7 phút",
      "goal": "Tạo tâm thế hứng thú, khơi gợi hiểu biết ban đầu của học sinh về...",
      "content": "Học sinh quan sát tình huống/video và trả lời câu hỏi dẫn dắt...",
      "product": "Câu trả lời của học sinh về...",
      "steps": [
        {{
          "step_name": "Bước 1: Chuyển giao nhiệm vụ",
          "teacher_action": "GV trình chiếu câu hỏi khởi động. [NLS-RED]GV yêu cầu HS sử dụng điện thoại/máy tính quét mã QR truy cập câu hỏi nhanh [1.2.TC2a][/NLS-RED].",
          "student_action": "HS chú ý lắng nghe, [NLS-RED]truy cập vào liên kết và làm câu hỏi khởi động [1.2.TC2a][/NLS-RED]."
        }},
        {{
          "step_name": "Bước 2: Thực hiện nhiệm vụ",
          "teacher_action": "GV quan sát, bao quát lớp, hỗ trợ các học sinh gặp khó khăn kỹ thuật.",
          "student_action": "HS suy nghĩ cá nhân, thảo luận nhanh với bạn cùng bàn để hoàn thành."
        }},
        {{
          "step_name": "Bước 3: Báo cáo, thảo luận",
          "teacher_action": "GV mời 2-3 đại diện học sinh chia sẻ ý kiến, khích lệ cả lớp nhận xét.",
          "student_action": "Học sinh đứng dậy phát biểu, các bạn khác lắng nghe và bổ sung ý kiến."
        }},
        {{
          "step_name": "Bước 4: Kết luận, nhận định",
          "teacher_action": "GV nhận xét câu trả lời, dẫn dắt kết nối vào nội dung bài học mới.",
          "student_action": "HS ghi nhớ vấn đề trọng tâm cần giải quyết trong bài."
        }}
      ]
    }},
    {{
      "id": 2,
      "name": "Hoạt động 2: Hình thành kiến thức mới",
      "time": "20 - 25 phút",
      "goal": "Học sinh chiếm lĩnh chuẩn kiến thức về...",
      "content": "Tìm hiểu thông tin SGK, quan sát ví dụ mẫu, làm việc nhóm...",
      "product": "Nội dung ghi chép hoàn chỉnh, bảng kết quả thảo luận nhóm...",
      "steps": [
        {{
          "step_name": "Bước 1: Chuyển giao nhiệm vụ",
          "teacher_action": "GV chia lớp thành các nhóm, giao phiếu học tập số. [NLS-RED]Yêu cầu các nhóm truy cập thư mục tài liệu số và xử lý thông tin [2.1.TC2a][/NLS-RED].",
          "student_action": "Các nhóm nhận nhiệm vụ, [NLS-RED]mở tài liệu trên máy tính để đọc và phân tích [2.1.TC2a][/NLS-RED]."
        }},
        {{
          "step_name": "Bước 2: Thực hiện nhiệm vụ",
          "teacher_action": "GV theo dõi tiến độ các nhóm, giải đáp thắc mắc, hướng dẫn phương pháp.",
          "student_action": "Các thành viên thảo luận sôi nổi, tổng hợp kết quả vào sản phẩm học tập chung."
        }},
        {{
          "step_name": "Bước 3: Báo cáo, thảo luận",
          "teacher_action": "GV điều hành phiên báo cáo, chỉ định nhóm trình bày và nhóm phản biện.",
          "student_action": "Đại diện nhóm trình bày sản phẩm, các nhóm khác theo dõi, đặt câu hỏi phản hồi."
        }},
        {{
          "step_name": "Bước 4: Kết luận, nhận định",
          "teacher_action": "GV chuẩn hóa kiến thức cốt lõi, chốt lại các nội dung bài học quan trọng.",
          "student_action": "HS đối chiếu kết quả, chỉnh sửa và ghi chép nội dung chuẩn vào vở."
        }}
      ]
    }},
    {{
      "id": 3,
      "name": "Hoạt động 3: Luyện tập",
      "time": "8 - 10 phút",
      "goal": "Củng cố và khắc sâu kiến thức vừa học thông qua các câu hỏi/bài tập thực hành...",
      "content": "Giải quyết các bài tập trắc nghiệm và câu hỏi tình huống trong bài...",
      "product": "Kết quả bài làm bài tập của học sinh, đáp án chính xác...",
      "steps": [
        {{
          "step_name": "Bước 1: Chuyển giao nhiệm vụ",
          "teacher_action": "GV đưa ra hệ thống bài tập củng cố. [NLS-RED]Yêu cầu học sinh nộp câu trả lời trên phần mềm học tập [1.2.TC2a][/NLS-RED].",
          "student_action": "HS [NLS-RED]thực hiện bài kiểm tra trắc nghiệm số trên máy tính [1.2.TC2a][/NLS-RED]."
        }},
        {{
          "step_name": "Bước 2: Thực hiện nhiệm vụ",
          "teacher_action": "GV quan sát hệ thống phản hồi kết quả trực tiếp của học sinh.",
          "student_action": "HS độc lập suy nghĩ và lựa chọn đáp án chính xác cho từng câu hỏi."
        }},
        {{
          "step_name": "Bước 3: Báo cáo, thảo luận",
          "teacher_action": "GV hiển thị phổ điểm tổng quan, chỉ rõ các câu hỏi nhiều học sinh chọn nhầm.",
          "student_action": "HS theo dõi phân tích của GV, tự kiểm tra lại phương án của mình."
        }},
        {{
          "step_name": "Bước 4: Kết luận, nhận định",
          "teacher_action": "GV giải thích cặn kẽ các câu khó, tuyên dương học sinh đạt kết quả tốt.",
          "student_action": "HS ghi chú các điểm cần lưu ý để tránh sai sót tương tự."
        }}
      ]
    }},
    {{
      "id": 4,
      "name": "Hoạt động 4: Vận dụng",
      "time": "3 - 5 phút (giao nhiệm vụ về nhà)",
      "goal": "Phát triển năng lực vận dụng kiến thức tin học vào giải quyết tình huống thực tế đời sống...",
      "content": "Tìm hiểu và giải quyết tình huống thực tiễn gắn với nội dung bài học...",
      "product": "Báo cáo hoặc sản phẩm số do học sinh thực hiện tại nhà...",
      "steps": [
        {{
          "step_name": "Bước 1: Chuyển giao nhiệm vụ",
          "teacher_action": "GV giao nhiệm vụ về nhà: [NLS-RED]HS tìm hiểu thông tin trên Internet và lưu trữ dữ liệu an toàn vào thư mục đám mây [2.2.TC2a][/NLS-RED].",
          "student_action": "HS lắng nghe yêu cầu, [NLS-RED]lưu đường dẫn tài liệu học tập vào thiết bị cá nhân [2.2.TC2a][/NLS-RED]."
        }},
        {{
          "step_name": "Bước 2: Thực hiện nhiệm vụ",
          "teacher_action": "GV hướng dẫn nguồn tham khảo đáng tin cậy và kênh nộp bài trực tuyến.",
          "student_action": "HS thực hiện nhiệm vụ ngoài giờ học, trao đổi với GV khi cần hỗ trợ."
        }},
        {{
          "step_name": "Bước 3: Báo cáo, thảo luận",
          "teacher_action": "GV thu sản phẩm qua nhóm học tập và chọn lọc các bài làm sáng tạo để nhận xét ở tiết sau.",
          "student_action": "HS gửi sản phẩm đúng hạn qua kênh thông tin lớp học."
        }},
        {{
          "step_name": "Bước 4: Kết luận, nhận định",
          "teacher_action": "GV dặn dò chuẩn bị cho bài học kế tiếp.",
          "student_action": "HS hoàn thành nhiệm vụ và chuẩn bị bài mới."
        }}
      ]
    }}
  ],
  "nls_summary_table": [
    {{
      "stt": 1,
      "activity": "HS truy cập công cụ/phần mềm số để thực hiện khảo sát và làm bài luyện tập trực tuyến.",
      "code": "1.2.TC2a"
    }},
    {{
      "stt": 2,
      "activity": "HS tìm kiếm, khai thác học liệu số và làm việc nhóm để giải quyết tình huống học tập.",
      "code": "2.1.TC2a\\n2.2.TC2a"
    }}
  ],
  "approvals": {{
    "cm_title": "Duyệt của Chuyên môn",
    "cm_signer": "{signer_cm}",
    "to_title": "Duyệt của tổ chuyên môn",
    "to_signer": "{signer_to}"
  }}
}}
"""

        # Gửi request sang Beekone API
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": 0.2,
            "max_tokens": 16000
        }

        try:
            resp = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=120
            )

            if resp.status_code != 200:
                raise Exception(f"API Error {resp.status_code}: {resp.text}")

            res_data = resp.json()
            raw_content = res_data["choices"][0]["message"]["content"].strip()

            # Làm sạch JSON markdown block nếu có
            clean_json = raw_content
            if clean_json.startswith("```"):
                lines = clean_json.split("\n")
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].startswith("```"):
                    lines = lines[:-1]
                clean_json = "\n".join(lines).strip()

            parsed_data = json.loads(clean_json)
            return parsed_data

        except Exception as e:
            logger.warning(f"Lỗi khi gọi API Beekone ({e}), kích hoạt chế độ Fallback chuẩn hóa nội bộ.")
            return self._build_fallback_khbd(
                lesson_info=lesson_info,
                teacher_name=teacher_name,
                department=department,
                signer_cm=signer_cm,
                signer_to=signer_to,
                subject=subject,
                grade=grade,
                school_name=school_name
            )

    def _build_fallback_khbd(
        self,
        lesson_info: Dict[str, Any],
        teacher_name: str,
        department: str,
        signer_cm: str,
        signer_to: str,
        subject: str,
        grade: str,
        school_name: str
    ) -> Dict[str, Any]:
        """Tạo KHBD dự phòng chuẩn mực đầy đủ theo quy định 5512 và TT 02/2025."""
        week_str = str(lesson_info.get("week", "1"))
        period_str = str(lesson_info.get("period", "1"))
        date_str = str(lesson_info.get("date_range", "07/09/2026 - 11/09/2026"))
        lesson_title = str(lesson_info.get("lesson_title", "Bài học"))

        return {
            "header": {
                "week": week_str,
                "period": period_str,
                "date_range": date_str,
                "teacher_name": teacher_name,
                "department": department,
                "school_name": school_name
            },
            "title": lesson_title.upper(),
            "sub_title": f"Môn học: {subject}; Lớp: {grade} | Thời gian thực hiện: 1 tiết (45 phút)",
            "objectives": {
                "knowledge": [
                    f"Nêu và giải thích được các khái niệm cốt lõi theo chuẩn kiến thức bài {lesson_title}.",
                    "Vận dụng kiến thức bài học để giải quyết các bài tập và tình huống thực tế."
                ],
                "competencies": {
                    "general": [
                        "Năng lực tự chủ và tự học: Chủ động nghiên cứu nội dung sách giáo khoa, hoàn thành các nhiệm vụ được giao.",
                        "Năng lực giao tiếp và hợp tác: Tương tác tích cực với các thành viên trong nhóm, lắng nghe và phản hồi mang tính xây dựng."
                    ],
                    "specific": [
                        "Năng lực nhận thức tin học: Hiểu và diễn đạt chính xác các thuật ngữ công nghệ.",
                        "Năng lực ứng dụng tin học: Thực hiện đúng quy trình kỹ thuật trên thiết bị số."
                    ],
                    "digital": [
                        {
                            "code": "1.2.TC2a",
                            "desc": "Sử dụng công cụ số/phần mềm chuyên dụng phù hợp để thao tác và xử lý dữ liệu học tập."
                        },
                        {
                            "code": "2.1.TC2a",
                            "desc": "Tìm kiếm, chọn lọc và khai thác thông tin từ môi trường số an toàn, hiệu quả."
                        }
                    ]
                },
                "qualities": [
                    "Chăm chỉ: Có tinh thần vượt khó, hoàn thành tốt các bài tập thực hành.",
                    "Trách nhiệm: Tuân thủ nội quy phòng máy, bảo vệ dữ liệu và thiết bị học tập."
                ]
            },
            "equipment": {
                "teacher": [
                    "Kế hoạch bài dạy, bài trình chiếu PowerPoint tương tác.",
                    "Máy chiếu, máy tính kết nối Internet, phiếu học tập số."
                ],
                "student": [
                    f"Sách giáo khoa Tin học {grade}, vở ghi chép bài.",
                    "Máy tính phòng học bộ môn để thao tác học tập."
                ]
            },
            "activities": [
                {
                    "id": 1,
                    "name": "Hoạt động 1: Mở đầu / Khởi động (Xác định vấn đề)",
                    "time": "5 - 7 phút",
                    "goal": "Tạo tâm thế hứng khởi, gợi mở sự tò mò và định hướng vấn đề cần giải quyết trong bài học.",
                    "content": "Học sinh tham gia trò chơi khởi động/quan sát tình huống thực tế để nhận diện vấn đề.",
                    "product": "Ý kiến phát biểu của học sinh, câu trả lời tình huống mở đầu.",
                    "steps": [
                        {
                            "step_name": "Bước 1: Chuyển giao nhiệm vụ",
                            "teacher_action": "GV trình chiếu câu hỏi khởi động. [NLS-RED]Yêu cầu học sinh quét mã QR truy cập câu hỏi trắc nghiệm nhanh trên điện thoại/máy tính [1.2.TC2a][/NLS-RED].",
                            "student_action": "HS quan sát, [NLS-RED]sử dụng thiết bị số quét mã truy cập câu hỏi khởi động [1.2.TC2a][/NLS-RED]."
                        },
                        {
                            "step_name": "Bước 2: Thực hiện nhiệm vụ",
                            "teacher_action": "GV bao quát, hướng dẫn học sinh thao tác đúng trên phần mềm.",
                            "student_action": "HS suy nghĩ độc lập, chọn đáp án và nộp câu trả lời trên ứng dụng."
                        },
                        {
                            "step_name": "Bước 3: Báo cáo, thảo luận",
                            "teacher_action": "GV trình chiếu biểu đồ kết quả trả lời của cả lớp, gọi đại diện giải thích.",
                            "student_action": "HS chia sẻ lý do chọn phương án, các bạn khác nhận xét bổ sung."
                        },
                        {
                            "step_name": "Bước 4: Kết luận, nhận định",
                            "teacher_action": "GV tổng kết, tuyên dương sự tham gia tích cực và dẫn dắt vào bài mới.",
                            "student_action": "HS lắng nghe và ghi nhớ định hướng bài học."
                        }
                    ]
                },
                {
                    "id": 2,
                    "name": "Hoạt động 2: Hình thành kiến thức mới",
                    "time": "20 - 25 phút",
                    "goal": "Học sinh chiếm lĩnh chuẩn kiến thức và kỹ năng trọng tâm của bài học.",
                    "content": "Làm việc với SGK, quan sát ví dụ mẫu, thảo luận nhóm để rút ra kết luận.",
                    "product": "Nội dung ghi chép hoàn chỉnh vào vở, sản phẩm thực hành nhóm.",
                    "steps": [
                        {
                            "step_name": "Bước 1: Chuyển giao nhiệm vụ",
                            "teacher_action": "GV phân chia nhóm học tập, giao phiếu học tập số. [NLS-RED]Yêu cầu các nhóm truy cập liên kết học liệu số và tra cứu kiến thức [2.1.TC2a][/NLS-RED].",
                            "student_action": "Các nhóm nhận nhiệm vụ, [NLS-RED]mở tài liệu trên máy tính, phân công nhiệm vụ cho từng thành viên [2.1.TC2a][/NLS-RED]."
                        },
                        {
                            "step_name": "Bước 2: Thực hiện nhiệm vụ",
                            "teacher_action": "GV quan sát, kịp thời hỗ trợ giải đáp khó khăn cho các nhóm.",
                            "student_action": "HS đọc tài liệu, trao đổi thảo luận và hoàn thành phiếu học tập."
                        },
                        {
                            "step_name": "Bước 3: Báo cáo, thảo luận",
                            "teacher_action": "GV mời đại diện nhóm thuyết trình sản phẩm, các nhóm khác theo dõi phản biện.",
                            "student_action": "Đại diện nhóm báo cáo kết quả, nhóm khác nhận xét, bổ sung ý kiến."
                        },
                        {
                            "step_name": "Bước 4: Kết luận, nhận định",
                            "teacher_action": "GV chuẩn hóa kiến thức, làm rõ những điểm học sinh hay nhầm lẫn.",
                            "student_action": "HS đối chiếu và ghi chép kiến thức cốt lõi vào vở."
                        }
                    ]
                },
                {
                    "id": 3,
                    "name": "Hoạt động 3: Luyện tập",
                    "time": "8 - 10 phút",
                    "goal": "Khắc sâu kiến thức, rèn luyện kỹ năng giải quyết bài tập.",
                    "content": "Hoàn thành hệ thống câu hỏi trắc nghiệm và bài tập vận dụng nhanh.",
                    "product": "Bảng điểm và đáp án bài luyện tập của học sinh.",
                    "steps": [
                        {
                            "step_name": "Bước 1: Chuyển giao nhiệm vụ",
                            "teacher_action": "GV giao bộ câu hỏi luyện tập. [NLS-RED]Yêu cầu học sinh nộp bài tập qua phần mềm học tập trực tuyến [1.2.TC2a][/NLS-RED].",
                            "student_action": "HS [NLS-RED]thực hiện bài làm trực tiếp trên hệ thống số [1.2.TC2a][/NLS-RED]."
                        },
                        {
                            "step_name": "Bước 2: Thực hiện nhiệm vụ",
                            "teacher_action": "GV giám sát quá trình làm bài và theo dõi tiến độ nộp bài của cả lớp.",
                            "student_action": "HS tập trung hoàn thành các câu hỏi theo thời gian quy định."
                        },
                        {
                            "step_name": "Bước 3: Báo cáo, thảo luận",
                            "teacher_action": "GV công bố đáp án, chữa chi tiết các câu hỏi có tỷ lệ sai cao.",
                            "student_action": "HS tự đối chiếu kết quả, phát biểu thắc mắc nếu có."
                        },
                        {
                            "step_name": "Bước 4: Kết luận, nhận định",
                            "teacher_action": "GV đánh giá mức độ tiếp thu của lớp và khen ngợi các bạn đạt điểm xuất sắc.",
                            "student_action": "HS rút kinh nghiệm và sửa lỗi bài làm."
                        }
                    ]
                },
                {
                    "id": 4,
                    "name": "Hoạt động 4: Vận dụng",
                    "time": "3 - 5 phút (giao nhiệm vụ về nhà)",
                    "goal": "Vận dụng kiến thức đã học vào thực tiễn cuộc sống hoặc tình huống mở rộng.",
                    "content": "Thực hiện nhiệm vụ tìm hiểu mở rộng ngoài giờ lên lớp.",
                    "product": "Báo cáo hoặc sản phẩm số nộp trực tuyến cho giáo viên.",
                    "steps": [
                        {
                            "step_name": "Bước 1: Chuyển giao nhiệm vụ",
                            "teacher_action": "GV giao nhiệm vụ về nhà: [NLS-RED]Tìm hiểu tình huống thực tế và nộp sản phẩm qua nhóm học tập [3.2.TC2a][/NLS-RED].",
                            "student_action": "HS [NLS-RED]lưu thông tin nhiệm vụ và hướng dẫn nộp bài vào thiết bị số [3.2.TC2a][/NLS-RED]."
                        },
                        {
                            "step_name": "Bước 2: Thực hiện nhiệm vụ",
                            "teacher_action": "GV giải thích rõ yêu cầu về hình thức và thời hạn nộp bài.",
                            "student_action": "HS thực hiện tại nhà theo hướng dẫn."
                        },
                        {
                            "step_name": "Bước 3: Báo cáo, thảo luận",
                            "teacher_action": "GV thu sản phẩm, chấm và chọn ra các bài tốt để tuyên dương tiết sau.",
                            "student_action": "HS nộp bài đúng hạn và trao đổi khi cần thiết."
                        },
                        {
                            "step_name": "Bước 4: Kết luận, nhận định",
                            "teacher_action": "GV dặn dò chuẩn bị cho tiết học tiếp theo.",
                            "student_action": "HS ghi nhớ dặn dò của giáo viên."
                        }
                    ]
                }
            ],
            "nls_summary_table": [
                {
                    "stt": 1,
                    "activity": "HS sử dụng thiết bị số và ứng dụng trực tuyến để làm bài kiểm tra nhanh và bài tập luyện tập.",
                    "code": "1.2.TC2a"
                },
                {
                    "stt": 2,
                    "activity": "HS truy cập, khai thác học liệu số và trao đổi nhóm qua môi trường mạng.",
                    "code": "2.1.TC2a\n3.2.TC2a"
                }
            ],
            "approvals": {
                "cm_title": "Duyệt của Chuyên môn",
                "cm_signer": signer_cm,
                "to_title": "Duyệt của tổ chuyên môn",
                "to_signer": signer_to
            }
        }
