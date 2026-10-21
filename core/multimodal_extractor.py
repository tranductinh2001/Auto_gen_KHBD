import os
import base64
import json
import time
import re
from typing import List, Dict, Tuple, Optional
from core.ai_vision_client import AIVisionClient

MULTIMODAL_INFORMATICS_PROMPT = """Bạn là chuyên gia sư phạm và kỹ sư tri thức đồ thị (Knowledge Graph Engineer) cho đồ án:
"Xây dựng hệ thống quản lý tri thức, phân tích chương trình và hỗ trợ ra đề thi tự động môn Tin học phổ thông GDPT 2018 sử dụng Neo4j".

Nhiệm vụ của bạn là đọc và phân tích trang Sách giáo khoa (SGK) Tin học trong hình ảnh, tự động phân đoạn sư phạm (Chunking) và trích xuất dữ liệu đa phương thức có cấu trúc chính xác tuyệt đối thành JSON.

Yêu cầu kỹ thuật chi tiết:

1. CẤP ĐỘ SƯ PHẠM VĂN BẢN:
   - grade_level: Khối lớp (Ví dụ: "Lớp 9", "Lớp 6", "Lớp 10"...).
   - chapter_title: Tên chủ đề / chương lớn (Ví dụ: "Chủ đề 1: Máy tính và cộng đồng", "Chủ đề 2: Mạng máy tính và Internet", "Chủ đề E: Ứng dụng tin học", "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính"...).
   - lesson_id: Mã bài học cụ thể (Ví dụ: "Bài 1", "Bài 2", "Bài 9a"...).
   - lesson_title: Tên bài học đầy đủ (Ví dụ: "Thế giới kĩ thuật số", "Thông tin trong giải quyết vấn đề"...).
   - requirements: Danh sách Yêu cầu cần đạt chuẩn GDPT 2018 (từ phần "Sau bài học này em sẽ...", khung mục tiêu đầu bài):
     + requirement_code: Mã yêu cầu tự sinh (Ví dụ: "YC01", "YC02"...)
     + requirement_content: Nội dung cụ thể của yêu cầu cần đạt.
   - knowledges: Danh sách các mục kiến thức cốt lõi / lý thuyết trong bài:
     + knowledge_name: Tên mục hoặc khái niệm cốt lõi (Ví dụ: "1. Khái niệm thiết bị số", "2. Bảng tính điện tử").
     + knowledge_content: Nội dung lý thuyết chi tiết, chuẩn xác, không bị dính chữ, giữ nguyên thuật ngữ Tin học.
   - skills: Danh sách kỹ năng thực hành hình thành trong bài:
     + skill_name: Tên kỹ năng (Ví dụ: "Sử dụng hàm VLOOKUP trong Excel", "Viết thuật toán tìm kiếm tuần tự bằng Python").
     + skill_description: Mô tả thao tác, các bước thực hiện.

2. TRÍCH XUẤT ĐA PHƯƠNG THỨC - HÌNH ẢNH, SƠ ĐỒ, GIAO DIỆN (VISUAL ASSETS):
   - ĐỐI VỚI MỌI hình ảnh minh họa, sơ đồ cấu trúc thuật toán, sơ đồ tư duy, ảnh chụp giao diện phần mềm trong sách (như Excel, Video Editor, Scratch, Python IDLE, Inkscape, CSDL Access/MySQL... có đánh số như Hình 1.1, Hình 2.3, Hình 9a.2, Bảng 1.1...):
   - TUYỆT ĐỐI KHÔNG BỎ QUA. Bạn phải "nhìn" trực tiếp vào bức ảnh và sinh ra:
     + visual_asset_code: Mã hình ảnh trong SGK (Ví dụ: "Hình 1.1", "Hình 2.3", "Bảng 1.1").
     + caption: Tiêu đề hoặc phần chú thích gốc của hình ảnh trong sách.
     + detailed_description: Đoạn văn bản mô tả chi tiết ý nghĩa trực quan, thành phần cấu trúc của hình ảnh đó (Detailed Image Description): phân tích các thành phần bên trong (menu, thanh công cụ, các vùng hiển thị, đoạn mã lệnh, luồng thuật toán, bảng dữ liệu) và nêu rõ ý nghĩa sư phạm mà hình ảnh truyền tải tới học sinh.

3. PHẦN ĐÁNH GIÁ VÀ HỖ TRỢ RA ĐỀ THI TỰ ĐỘNG (QUESTIONS & OPTIONS):
   - questions: Danh sách các câu hỏi luyện tập, ôn tập, bài tập vận dụng cuối bài:
     + question_id: Mã câu hỏi (Ví dụ: "CH01", "LT1", "VD1").
     + content: Nội dung câu hỏi kiểm tra.
     + difficulty: Phân loại mức độ nhận thức chuẩn sư phạm: "Nhận biết" | "Thông hiểu" | "Vận dụng".
     + options: Danh sách các phương án lựa chọn nếu là câu trắc nghiệm (hoặc mảng rỗng [] nếu là tự luận):
       * option_key: "A" | "B" | "C" | "D"
       * content: Nội dung phương án lựa chọn
       * is_correct: true (nếu là đáp án đúng), false (nếu sai). Dùng tri thức môn Tin học để xác định đáp án đúng chính xác.

4. NGUYÊN TẮC ĐẦU RA:
   - Các từ ngữ phải rõ ràng, tuyệt đối không dính chữ.
   - Giữ nguyên các thuật ngữ Tin học (RAM, ROM, CPU, LAN, IP, Python, Scratch, Excel, SQL...).
   - Chỉ trả về chuỗi JSON hợp lệ theo cấu trúc bên dưới, KHÔNG giải thích thêm.

Cấu trúc JSON bắt buộc:
{
  "grade_level": "Lớp 9",
  "chapter_title": "Chủ đề 1: Máy tính và cộng đồng",
  "lessons": [
    {
      "lesson_id": "Bài 1",
      "lesson_title": "Thế giới kĩ thuật số",
      "requirements": [
        {
          "requirement_code": "YC01",
          "requirement_content": "Nhận biết và nêu được vai trò của thiết bị số trong đời sống."
        }
      ],
      "knowledges": [
        {
          "knowledge_name": "1. Thiết bị số thông dụng",
          "knowledge_content": "Thiết bị số là thiết bị có khả năng thu thập, lưu trữ, xử lý và truyền dữ liệu số..."
        }
      ],
      "visual_assets": [
        {
          "visual_asset_code": "Hình 1.1",
          "caption": "Một số thiết bị số thông dụng",
          "detailed_description": "Ảnh chụp chi tiết các thiết bị ngoại vi gồm bàn phím cơ kết nối USB, chuột quang, máy quét và màn hình LED. Hình ảnh giúp học sinh phân biệt rõ thiết bị vào và thiết bị ra."
        }
      ],
      "skills": [
        {
          "skill_name": "Kết nối và thao tác thiết bị số",
          "skill_description": "Thực hiện cắm đúng cổng kết nối USB/HDMI và sử dụng thiết bị."
        }
      ],
      "questions": [
        {
          "question_id": "CH01",
          "content": "Thiết bị nào sau đây là thiết bị vào của máy tính?",
          "difficulty": "Nhận biết",
          "options": [
            {"option_key": "A", "content": "Màn hình", "is_correct": false},
            {"option_key": "B", "content": "Bàn phím", "is_correct": true},
            {"option_key": "C", "content": "Máy in", "is_correct": false},
            {"option_key": "D", "content": "Loa", "is_correct": false}
          ]
        }
      ]
    }
  ]
}
"""


class MultimodalExtractor:
    """Bộ trích xuất tri thức sư phạm đa phương thức từ hình ảnh trang SGK Tin học."""

    def __init__(self, client: AIVisionClient):
        self.client = client

    def _extract_json_from_text(self, text: str) -> dict:
        """Parse JSON linh hoạt, loại bỏ markdown block nếu có."""
        match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', text)
        json_str = match.group(1) if match else text

        if not match:
            obj_match = re.search(r'(\{[\s\S]*\})', text)
            if obj_match:
                json_str = obj_match.group(1)

        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            cleaned = re.sub(r',\s*([\]}])', r'\1', json_str)
            try:
                return json.loads(cleaned)
            except Exception:
                raise Exception(f"Không thể parse JSON từ kết quả AI: {e}\nNội dung text: {text[:400]}...")

    def extract_page(
        self,
        image_path: str,
        source_file: str,
        page_num: int,
        grade_hint: str = "Lớp 9"
    ) -> Tuple[List[dict], dict, dict]:
        """
        Gửi ảnh trang sách lên AI Vision và trích xuất:
        - flat_rows: Các dòng bảng phẳng đầy đủ các cột sư phạm
        - normalized_data: Các tập thực thể phân cấp cho Neo4j Property Graph
        - log_info: Nhật ký xử lý
        """
        start_time = time.time()
        log_info = {
            "source_file": os.path.basename(source_file),
            "page_num": page_num,
            "status": "OK",
            "error": "",
            "rows_extracted": 0,
            "duration_sec": 0.0
        }

        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")

        prompt = MULTIMODAL_INFORMATICS_PROMPT
        if grade_hint and grade_hint != "Tự động":
            prompt += f"\n\nLƯU Ý: Trang sách này thuộc môn Tin học {grade_hint}. Hãy ưu tiên gán grade_level là '{grade_hint}'."

        try:
            response_text = self.client.chat_with_image(
                image_base64=img_b64,
                prompt=prompt,
                mime_type="image/png"
            )
            raw_data = self._extract_json_from_text(response_text)
            flat_rows, norm_data = self.flatten_and_normalize(
                data=raw_data,
                source_file=source_file,
                page_num=page_num,
                fallback_grade=grade_hint
            )
            log_info["rows_extracted"] = len(flat_rows)
            log_info["duration_sec"] = round(time.time() - start_time, 2)
            return flat_rows, norm_data, log_info

        except Exception as e:
            log_info["status"] = "ERROR"
            log_info["error"] = str(e)
            log_info["duration_sec"] = round(time.time() - start_time, 2)
            return [], {
                "lessons": [], "requirements": [], "knowledges": [],
                "visual_assets": [], "skills": [], "questions": [], "answers": []
            }, log_info

    def flatten_and_normalize(
        self,
        data: dict,
        source_file: str,
        page_num: int,
        fallback_grade: str = "Lớp 9"
    ) -> Tuple[List[dict], dict]:
        """Chuẩn hóa dữ liệu sang cả dạng phẳng và dạng quan hệ thực thể cho Neo4j."""
        grade_level = data.get("grade_level") or fallback_grade or "Lớp 9"
        chapter_title = data.get("chapter_title") or "Chủ đề chưa phân loại"
        lessons = data.get("lessons", [])
        if not lessons and (data.get("lesson_id") or data.get("lesson_title")):
            lessons = [data]

        filename = os.path.basename(source_file)

        flat_rows = []
        norm_lessons = []
        norm_requirements = []
        norm_knowledges = []
        norm_visual_assets = []
        norm_skills = []
        norm_questions = []
        norm_answers = []

        for l_idx, lesson in enumerate(lessons, 1):
            l_id = str(lesson.get("lesson_id") or f"Bai_{page_num}_{l_idx}").strip()
            l_title = str(lesson.get("lesson_title") or "Bài học không tên").strip()
            l_chapter = str(lesson.get("chapter_title") or chapter_title).strip()
            l_grade = str(lesson.get("grade_level") or grade_level).strip()

            norm_lessons.append({
                "grade_level": l_grade,
                "chapter_title": l_chapter,
                "lesson_id": l_id,
                "lesson_title": l_title,
                "page_num": page_num,
                "source_file": filename
            })

            reqs = lesson.get("requirements", [])
            knows = lesson.get("knowledges", [])
            visual_assets = lesson.get("visual_assets", [])
            skills = lesson.get("skills", [])
            questions = lesson.get("questions", [])

            # 1. Requirements
            for r_idx, req in enumerate(reqs, 1):
                r_code = str(req.get("requirement_code") or f"YC{r_idx:02d}").strip()
                r_content = str(req.get("requirement_content") or "").strip()
                norm_requirements.append({
                    "grade_level": l_grade,
                    "chapter_title": l_chapter,
                    "lesson_id": l_id,
                    "lesson_title": l_title,
                    "requirement_code": r_code,
                    "requirement_content": r_content,
                    "page_num": page_num
                })

            # 2. Knowledges
            for k_idx, kn in enumerate(knows, 1):
                k_name = str(kn.get("knowledge_name") or f"Mục {k_idx}").strip()
                k_content = str(kn.get("knowledge_content") or "").strip()
                norm_knowledges.append({
                    "grade_level": l_grade,
                    "chapter_title": l_chapter,
                    "lesson_id": l_id,
                    "lesson_title": l_title,
                    "knowledge_name": k_name,
                    "knowledge_content": k_content,
                    "page_num": page_num
                })

            # 3. Visual Assets (Đa phương thức)
            for v_idx, va in enumerate(visual_assets, 1):
                va_code = str(va.get("visual_asset_code") or f"Hình_{v_idx}").strip()
                caption = str(va.get("caption") or "").strip()
                detailed_desc = str(va.get("detailed_description") or "").strip()
                norm_visual_assets.append({
                    "grade_level": l_grade,
                    "chapter_title": l_chapter,
                    "lesson_id": l_id,
                    "lesson_title": l_title,
                    "visual_asset_code": va_code,
                    "caption": caption,
                    "detailed_description": detailed_desc,
                    "page_num": page_num
                })

            # 4. Skills
            for s_idx, sk in enumerate(skills, 1):
                sk_name = str(sk.get("skill_name") or f"Kỹ năng {s_idx}").strip()
                sk_desc = str(sk.get("skill_description") or "").strip()
                norm_skills.append({
                    "grade_level": l_grade,
                    "chapter_title": l_chapter,
                    "lesson_id": l_id,
                    "lesson_title": l_title,
                    "skill_name": sk_name,
                    "skill_description": sk_desc,
                    "page_num": page_num
                })

            # 5. Questions & Options (Hỗ trợ ra đề thi tự động)
            for q_idx, q in enumerate(questions, 1):
                q_id = str(q.get("question_id") or f"CH{q_idx:02d}").strip()
                q_content = str(q.get("content") or "").strip()
                difficulty = str(q.get("difficulty") or "Thông hiểu").strip()
                opts = q.get("options", [])

                norm_questions.append({
                    "grade_level": l_grade,
                    "chapter_title": l_chapter,
                    "lesson_id": l_id,
                    "lesson_title": l_title,
                    "question_id": q_id,
                    "content": q_content,
                    "difficulty": difficulty,
                    "page_num": page_num
                })

                if opts:
                    for opt in opts:
                        opt_key = str(opt.get("option_key") or opt.get("key") or "").strip()
                        opt_content = str(opt.get("content") or "").strip()
                        is_corr = opt.get("is_correct", False)
                        norm_answers.append({
                            "grade_level": l_grade,
                            "lesson_id": l_id,
                            "question_id": q_id,
                            "option_key": opt_key,
                            "content": opt_content,
                            "is_correct": is_corr,
                            "page_num": page_num
                        })
                else:
                    norm_answers.append({
                        "grade_level": l_grade,
                        "lesson_id": l_id,
                        "question_id": q_id,
                        "option_key": "",
                        "content": "Câu hỏi tự luận / Vận dụng",
                        "is_correct": True,
                        "page_num": page_num
                    })

            # 6. Sinh Flat Rows (Bảng phẳng 21 cột)
            first_req_code = reqs[0].get("requirement_code", "") if reqs else ""
            first_req_content = reqs[0].get("requirement_content", "") if reqs else ""
            first_kn_name = knows[0].get("knowledge_name", "") if knows else ""
            first_kn_content = knows[0].get("knowledge_content", "") if knows else ""
            first_sk_name = skills[0].get("skill_name", "") if skills else ""
            first_sk_desc = skills[0].get("skill_description", "") if skills else ""
            first_va_code = visual_assets[0].get("visual_asset_code", "") if visual_assets else ""
            first_va_caption = visual_assets[0].get("caption", "") if visual_assets else ""
            first_va_desc = visual_assets[0].get("detailed_description", "") if visual_assets else ""

            has_flat_rows = False

            if questions:
                for q in questions:
                    q_id = str(q.get("question_id") or "").strip()
                    q_content = str(q.get("content") or "").strip()
                    difficulty = str(q.get("difficulty") or "Thông hiểu").strip()
                    opts = q.get("options", [])
                    if opts:
                        for opt in opts:
                            has_flat_rows = True
                            flat_rows.append({
                                "grade_level": l_grade,
                                "chapter_title": l_chapter,
                                "lesson_id": l_id,
                                "lesson_title": l_title,
                                "requirement_code": first_req_code,
                                "requirement_content": first_req_content,
                                "knowledge_name": first_kn_name,
                                "knowledge_content": first_kn_content,
                                "skill_name": first_sk_name,
                                "skill_description": first_sk_desc,
                                "visual_asset_code": first_va_code,
                                "caption": first_va_caption,
                                "detailed_description": first_va_desc,
                                "question_id": q_id,
                                "content": q_content,
                                "difficulty": difficulty,
                                "option_key": str(opt.get("option_key") or opt.get("key") or ""),
                                "option_content": str(opt.get("content") or ""),
                                "is_correct": opt.get("is_correct", False),
                                "page_num": page_num,
                                "source_file": filename
                            })
                    else:
                        has_flat_rows = True
                        flat_rows.append({
                            "grade_level": l_grade,
                            "chapter_title": l_chapter,
                            "lesson_id": l_id,
                            "lesson_title": l_title,
                            "requirement_code": first_req_code,
                            "requirement_content": first_req_content,
                            "knowledge_name": first_kn_name,
                            "knowledge_content": first_kn_content,
                            "skill_name": first_sk_name,
                            "skill_description": first_sk_desc,
                            "visual_asset_code": first_va_code,
                            "caption": first_va_caption,
                            "detailed_description": first_va_desc,
                            "question_id": q_id,
                            "content": q_content,
                            "difficulty": difficulty,
                            "option_key": "",
                            "option_content": "",
                            "is_correct": "",
                            "page_num": page_num,
                            "source_file": filename
                        })

            if not has_flat_rows:
                max_items = max(len(reqs), len(knows), len(skills), len(visual_assets), 1)
                for i in range(max_items):
                    cur_r = reqs[i] if i < len(reqs) else {}
                    cur_k = knows[i] if i < len(knows) else {}
                    cur_s = skills[i] if i < len(skills) else {}
                    cur_va = visual_assets[i] if i < len(visual_assets) else {}

                    flat_rows.append({
                        "grade_level": l_grade,
                        "chapter_title": l_chapter,
                        "lesson_id": l_id,
                        "lesson_title": l_title,
                        "requirement_code": str(cur_r.get("requirement_code") or (first_req_code if i > 0 else "")),
                        "requirement_content": str(cur_r.get("requirement_content") or (first_req_content if i > 0 else "")),
                        "knowledge_name": str(cur_k.get("knowledge_name") or (first_kn_name if i > 0 else "")),
                        "knowledge_content": str(cur_k.get("knowledge_content") or (first_kn_content if i > 0 else "")),
                        "skill_name": str(cur_s.get("skill_name") or (first_sk_name if i > 0 else "")),
                        "skill_description": str(cur_s.get("skill_description") or (first_sk_desc if i > 0 else "")),
                        "visual_asset_code": str(cur_va.get("visual_asset_code") or (first_va_code if i > 0 else "")),
                        "caption": str(cur_va.get("caption") or (first_va_caption if i > 0 else "")),
                        "detailed_description": str(cur_va.get("detailed_description") or (first_va_desc if i > 0 else "")),
                        "question_id": "",
                        "content": "",
                        "difficulty": "",
                        "option_key": "",
                        "option_content": "",
                        "is_correct": "",
                        "page_num": page_num,
                        "source_file": filename
                    })

        norm_data = {
            "lessons": norm_lessons,
            "requirements": norm_requirements,
            "knowledges": norm_knowledges,
            "visual_assets": norm_visual_assets,
            "skills": norm_skills,
            "questions": norm_questions,
            "answers": norm_answers
        }

        return flat_rows, norm_data
