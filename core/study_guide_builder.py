"""
core/study_guide_builder.py
Bộ sinh Đề Cương Ôn Tập (Study Guide / Practice Review Outline) môn Tin học GDPT 2018.
Tính năng:
1. Chứa 100% tất cả các câu hỏi trong Đề thi (TH01 / TH02).
2. Tự động lấy thêm các câu hỏi mở rộng / tương tự từ ngân hàng câu hỏi (CSV / CURRICULUM_DATA)
   thuộc đúng các bài học trong phạm vi Kế hoạch giáo dục (KHGD) đã dạy.
3. Xáo trộn ngẫu nhiên thứ tự các câu hỏi và các phương án A, B, C, D để học sinh ôn luyện thực chất.
4. Cung cấp đầy đủ 3 phần: Trắc nghiệm nhiều lựa chọn, Trắc nghiệm Đúng/Sai, Tự luận & Thực hành.
5. Kèm theo Bảng đáp án & Hướng dẫn giải chi tiết cho giáo viên và học sinh đối chiếu.
6. Hỗ trợ xuất file Word (.DOCX) chuẩn thể thức và file Text (.TXT).
"""

import os
import copy
import random
import re
from typing import Dict, Any, List, Optional
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import pandas as pd


class StudyGuideBuilder:
    """Xây dựng dữ liệu Đề Cương Ôn Tập dựa trên Đề thi và Ngân hàng câu hỏi bổ trợ."""

    @staticmethod
    def build_study_guide(
        dossier: Dict[str, Any],
        extra_mcq_count: int = 12,
        extra_tf_count: int = 4,
        extra_essay_count: int = 3,
        shuffle: bool = True
    ) -> Dict[str, Any]:
        """
        Tạo dữ liệu đề cương ôn tập bao gồm:
        - 100% câu hỏi trong đề thi TH01
        - Các câu hỏi bổ sung tương ứng từ ngân hàng câu hỏi
        - Xáo trộn câu hỏi và đáp án
        - Bảng đáp án và hướng dẫn giải
        """
        meta = dossier.get("meta", {})
        th01 = dossier.get("th01", {})
        grade = meta.get("grade", "Lớp 6")

        base_mcqs = copy.deepcopy(th01.get("mcqs", []))
        base_tfs = copy.deepcopy(th01.get("tfs", []))
        base_essays = copy.deepcopy(th01.get("essays", []))

        # Lấy danh sách câu hỏi bổ sung từ ngân hàng
        extra_mcqs = StudyGuideBuilder._fetch_extra_mcqs(grade, base_mcqs, count=extra_mcq_count)
        extra_tfs = StudyGuideBuilder._fetch_extra_tfs(grade, base_tfs, count=extra_tf_count)
        extra_essays = StudyGuideBuilder._fetch_extra_essays(grade, base_essays, count=extra_essay_count)

        # Hợp nhất danh sách
        combined_mcqs = base_mcqs + extra_mcqs
        combined_tfs = base_tfs + extra_tfs
        combined_essays = base_essays + extra_essays

        # Xáo trộn câu hỏi và phương án lựa chọn
        if shuffle:
            # 1. Shuffle MCQs và hoán vị đáp án A, B, C, D
            random.seed(random.randint(1, 99999))
            random.shuffle(combined_mcqs)
            for q in combined_mcqs:
                opt_keys = ["A", "B", "C", "D"]
                opts = q.get("options", {})
                curr_correct = q.get("correct", "A")
                curr_correct_val = opts.get(curr_correct, "")

                vals = [opts.get(k, "") for k in opt_keys]
                shuffled_vals = list(vals)
                random.shuffle(shuffled_vals)

                new_opts = {k: v for k, v in zip(opt_keys, shuffled_vals)}
                new_correct = "A"
                for k, v in new_opts.items():
                    if v == curr_correct_val:
                        new_correct = k
                        break
                q["options"] = new_opts
                q["correct"] = new_correct

            # 2. Shuffle TFs và các ý a, b, c, d
            random.shuffle(combined_tfs)
            for tf in combined_tfs:
                stmts = list(tf.get("statements", []))
                random.shuffle(stmts)
                for idx, lbl in enumerate(["a", "b", "c", "d"]):
                    if idx < len(stmts):
                        stmts[idx]["label"] = lbl
                tf["statements"] = stmts

            # 3. Shuffle Essays
            random.shuffle(combined_essays)

        # Đánh lại số thứ tự chuẩn cho đề cương
        for idx, q in enumerate(combined_mcqs, 1):
            q["guide_num"] = idx
        for idx, tf in enumerate(combined_tfs, 1):
            tf["guide_num"] = idx
        for idx, es in enumerate(combined_essays, 1):
            es["guide_num"] = idx

        return {
            "meta": meta,
            "title": f"ĐỀ CƯƠNG ÔN TẬP KIỂM TRA {meta.get('exam_title', 'ĐỊNH KỲ').upper()}",
            "grade": grade,
            "mcqs": combined_mcqs,
            "tfs": combined_tfs,
            "essays": combined_essays,
            "counts": {
                "total_mcq": len(combined_mcqs),
                "total_tf": len(combined_tfs),
                "total_essay": len(combined_essays),
                "total_questions": len(combined_mcqs) + len(combined_tfs) + len(combined_essays)
            }
        }

    @staticmethod
    def _fetch_extra_mcqs(grade: str, base_mcqs: List[Dict[str, Any]], count: int = 12) -> List[Dict[str, Any]]:
        """Lấy thêm câu hỏi trắc nghiệm từ ngân hàng câu hỏi CSV hoặc CURRICULUM_DATA."""
        existing_contents = {q.get("content", "").strip().lower() for q in base_mcqs}
        extra_list = []

        # 1. Thử lấy từ output/neo4j_import/neo4j_questions.csv & neo4j_answers.csv
        try:
            curr_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            q_csv = os.path.join(curr_dir, "output", "neo4j_import", "neo4j_questions.csv")
            a_csv = os.path.join(curr_dir, "output", "neo4j_import", "neo4j_answers.csv")
            if os.path.exists(q_csv) and os.path.exists(a_csv):
                df_q = pd.read_csv(q_csv)
                df_a = pd.read_csv(a_csv)
                sub_q = df_q[df_q["grade_level"] == grade]
                
                # Gom nhóm đáp án theo question_content
                q_groups = df_a[df_a["grade_level"] == grade].groupby("question_content")
                for q_content, grp in q_groups:
                    c_clean = str(q_content).strip().lower()
                    if c_clean in existing_contents or len(c_clean) < 15:
                        continue
                    
                    opts = {}
                    correct_key = "A"
                    for _, r in grp.iterrows():
                        k = str(r.get("answer_option_key", "")).strip().upper()
                        if k in ["A", "B", "C", "D"]:
                            opts[k] = str(r.get("answer_content", "")).strip()
                            if r.get("is_correct") is True or str(r.get("is_correct")).lower() == "true":
                                correct_key = k

                    if len(opts) >= 4:
                        extra_list.append({
                            "id": f"EXTRA_CSV_{len(extra_list)+1}",
                            "unit": str(grp.iloc[0].get("lesson_title", "Ôn tập kiến thức")),
                            "level": str(grp.iloc[0].get("difficulty", "Thông hiểu")),
                            "content": str(q_content).strip(),
                            "options": opts,
                            "correct": correct_key
                        })
                        existing_contents.add(c_clean)
                        if len(extra_list) >= count:
                            break
        except Exception:
            pass

        # 2. Nếu chưa đủ, lấy từ CURRICULUM_DATA
        if len(extra_list) < count:
            try:
                from core.exam_builder import CURRICULUM_DATA
                c_data = CURRICULUM_DATA.get(grade, {})
                for q in c_data.get("mcq_bank", []):
                    c_clean = q.get("content", "").strip().lower()
                    if c_clean not in existing_contents:
                        extra_list.append(copy.deepcopy(q))
                        existing_contents.add(c_clean)
                        if len(extra_list) >= count:
                            break
            except Exception:
                pass

        # 3. Nếu vẫn chưa đủ, nhân bản biến thể xoay đáp án từ base_mcqs
        if len(extra_list) < count and base_mcqs:
            for q in base_mcqs:
                q_var = copy.deepcopy(q)
                q_var["id"] = f"{q.get('id', '')}_VAR"
                q_var["content"] = f"[Mở rộng rèn luyện] {q.get('content', '')}"
                extra_list.append(q_var)
                if len(extra_list) >= count:
                    break

        return extra_list[:count]

    @staticmethod
    def _fetch_extra_tfs(grade: str, base_tfs: List[Dict[str, Any]], count: int = 4) -> List[Dict[str, Any]]:
        """Lấy thêm câu hỏi Đúng/Sai từ ngân hàng hoặc tạo biến thể."""
        existing_topics = {tf.get("topic", "").strip().lower() for tf in base_tfs}
        extra_list = []

        try:
            from core.exam_builder import CURRICULUM_DATA
            c_data = CURRICULUM_DATA.get(grade, {})
            for tf in c_data.get("tf_bank", []):
                t_clean = tf.get("topic", "").strip().lower()
                if t_clean not in existing_topics:
                    extra_list.append(copy.deepcopy(tf))
                    existing_topics.add(t_clean)
                    if len(extra_list) >= count:
                        break
        except Exception:
            pass

        # Nếu chưa đủ, tạo biến thể đổi ý
        if len(extra_list) < count and base_tfs:
            for tf in base_tfs:
                tf_var = copy.deepcopy(tf)
                tf_var["topic"] = f"[Tình huống luyện tập] {tf.get('topic', '')}"
                # Đảo một số ý để học sinh phân tích
                for st in tf_var.get("statements", []):
                    st["content"] = f"Học sinh khẳng định: {st.get('content', '')}"
                extra_list.append(tf_var)
                if len(extra_list) >= count:
                    break

        return extra_list[:count]

    @staticmethod
    def _fetch_extra_essays(grade: str, base_essays: List[Dict[str, Any]], count: int = 3) -> List[Dict[str, Any]]:
        """Lấy thêm câu hỏi Tự luận từ ngân hàng hoặc tạo câu hỏi rèn luyện."""
        existing_essays = {es.get("content", "").strip().lower() for es in base_essays}
        extra_list = []

        try:
            from core.exam_builder import CURRICULUM_DATA
            c_data = CURRICULUM_DATA.get(grade, {})
            for es in c_data.get("essay_bank", []):
                e_clean = es.get("content", "").strip().lower()
                if e_clean not in existing_essays:
                    extra_list.append(copy.deepcopy(es))
                    existing_essays.add(e_clean)
                    if len(extra_list) >= count:
                        break
        except Exception:
            pass

        if len(extra_list) < count and base_essays:
            for es in base_essays:
                es_var = copy.deepcopy(es)
                es_var["content"] = f"[Bài tập vận dụng thêm] {es.get('content', '')}"
                extra_list.append(es_var)
                if len(extra_list) >= count:
                    break

        return extra_list[:count]


# ==============================================================================
# BỘ XUẤT TỆP ĐỀ CƯƠNG ÔN TẬP CHUẨN THỂ THỨC (DOCX & TXT)
# ==============================================================================

class StudyGuideDocxExporter:
    """Xuất Đề Cương Ôn Tập ra file Word (.docx) chuyên nghiệp chuẩn thể thức Bộ GD&ĐT."""

    @staticmethod
    def export(study_guide: Dict[str, Any], output_path: str):
        """Xuất file DOCX Đề cương ôn tập kèm đáp án hướng dẫn chấm."""
        doc = docx.Document()

        # Cấu hình trang A4 dọc (Portrait), lề chuẩn: Top 2cm, Bottom 2cm, Left 2.5cm, Right 2cm
        for section in doc.sections:
            section.top_margin = Inches(0.79)
            section.bottom_margin = Inches(0.79)
            section.left_margin = Inches(0.98)
            section.right_margin = Inches(0.79)
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)

        meta = study_guide.get("meta", {})
        grade = study_guide.get("grade", "Lớp 6")

        # ----------------- 1. HEADER BẢNG THÔNG TIN SƯ PHẠM -----------------
        table_hdr = doc.add_table(rows=1, cols=2)
        table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_hdr.autofit = False

        cell_left = table_hdr.cell(0, 0)
        cell_right = table_hdr.cell(0, 1)
        cell_left.width = Inches(3.2)
        cell_right.width = Inches(3.3)

        # Xóa viền bảng header
        StudyGuideDocxExporter._remove_table_borders(table_hdr)

        p_l = cell_left.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.paragraph_format.space_after = Pt(2)
        r = p_l.add_run(f"{meta.get('governing_body', 'UBND HUYỆN').upper()}\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r = p_l.add_run(f"{meta.get('school_name', 'TRƯỜNG THCS').upper()}\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.bold = True
        r = p_l.add_run("TỔ CHUYÊN MÔN: TIN HỌC")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r.italic = True

        p_r = cell_right.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r.paragraph_format.space_after = Pt(2)
        r = p_r.add_run(f"ĐỀ CƯƠNG ÔN TẬP {meta.get('exam_title', 'KIỂM TRA ĐỊNH KỲ').upper()}\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11.5)
        r.bold = True
        r = p_r.add_run(f"MÔN: TIN HỌC {grade.upper()} — NĂM HỌC: {meta.get('school_year', '2025 - 2026')}\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.bold = True
        r = p_r.add_run("TÀI LIỆU ÔN TẬP LƯU HÀNH NỘI BỘ")
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.italic = True

        # Đường kẻ phân cách
        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.space_before = Pt(2)
        p_line.paragraph_format.space_after = Pt(6)
        r_line = p_line.add_run("—" * 38)
        r_line.font.name = "Times New Roman"
        r_line.font.size = Pt(10)
        r_line.font.color.rgb = RGBColor(100, 116, 139)

        # Khung thông tin học sinh
        p_stu = doc.add_paragraph()
        p_stu.paragraph_format.space_before = Pt(2)
        p_stu.paragraph_format.space_after = Pt(4)
        r = p_stu.add_run("Họ và tên học sinh: ................................................................................. Lớp: ................. Phòng thi: ................")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        # Lời dặn dò ôn tập
        p_dan = doc.add_paragraph()
        p_dan.paragraph_format.space_before = Pt(2)
        p_dan.paragraph_format.space_after = Pt(10)
        r = p_dan.add_run("📌 HƯỚNG DẪN ÔN TẬP:\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r.bold = True
        r = p_dan.add_run(
            "1. Học sinh ôn kỹ các khái niệm, quy tắc, sơ đồ thuật toán và thao tác thực hành trong các bài học đã học.\n"
            "2. Làm đầy đủ các câu hỏi trắc nghiệm nhiều lựa chọn, đúng/sai và tự luận dưới đây trước khi đến lớp.\n"
            "3. Đề cương bao gồm toàn bộ các dạng bài sẽ xuất hiện trong bài kiểm tra định kỳ."
        )
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.italic = True

        # ----------------- PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -----------------
        mcqs = study_guide.get("mcqs", [])
        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(10)
        p_p1.paragraph_format.space_after = Pt(4)
        p_p1.paragraph_format.keep_with_next = True
        r = p_p1.add_run(f"PHẦN I. CÂU HỎI TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (Gồm {len(mcqs)} câu)")
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
        r.font.color.rgb = RGBColor(15, 44, 89)

        p_p1_sub = doc.add_paragraph()
        p_p1_sub.paragraph_format.space_after = Pt(6)
        r = p_p1_sub.add_run("Học sinh chọn duy nhất một phương án đúng nhất (A, B, C hoặc D) cho mỗi câu hỏi dưới đây:")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r.italic = True

        for q in mcqs:
            q_num = q.get("guide_num", 1)
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(4)
            p_q.paragraph_format.space_after = Pt(2)
            p_q.paragraph_format.line_spacing = 1.15
            p_q.paragraph_format.keep_with_next = True

            r = p_q.add_run(f"Câu {q_num}: ")
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)
            r.bold = True
            r = p_q.add_run(q.get("content", ""))
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)

            # Các phương án
            opts = q.get("options", {})
            p_opt = doc.add_paragraph()
            p_opt.paragraph_format.space_before = Pt(1)
            p_opt.paragraph_format.space_after = Pt(4)
            p_opt.paragraph_format.left_indent = Inches(0.2)
            p_opt.paragraph_format.line_spacing = 1.15

            r_opt = p_opt.add_run(
                f"A. {opts.get('A', '')}\n"
                f"B. {opts.get('B', '')}\n"
                f"C. {opts.get('C', '')}\n"
                f"D. {opts.get('D', '')}"
            )
            r_opt.font.name = "Times New Roman"
            r_opt.font.size = Pt(10.5)

        # ----------------- PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI -----------------
        tfs = study_guide.get("tfs", [])
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(14)
        p_p2.paragraph_format.space_after = Pt(4)
        p_p2.paragraph_format.keep_with_next = True
        r = p_p2.add_run(f"PHẦN II. CÂU HỎI TRẮC NGHIỆM ĐÚNG / SAI (Gồm {len(tfs)} câu tình huống)")
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
        r.font.color.rgb = RGBColor(15, 44, 89)

        p_p2_sub = doc.add_paragraph()
        p_p2_sub.paragraph_format.space_after = Pt(6)
        r = p_p2_sub.add_run("Trong mỗi ý a), b), c), d) ở từng câu hỏi tình huống, học sinh chọn Đúng (Đ) hoặc Sai (S):")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r.italic = True

        for tf in tfs:
            tf_num = tf.get("guide_num", 1)
            p_tf = doc.add_paragraph()
            p_tf.paragraph_format.space_before = Pt(6)
            p_tf.paragraph_format.space_after = Pt(2)
            p_tf.paragraph_format.keep_with_next = True

            r = p_tf.add_run(f"Câu {tf_num}: ")
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)
            r.bold = True
            r = p_tf.add_run(tf.get("topic", ""))
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)

            # Các mệnh đề
            for st in tf.get("statements", []):
                p_st = doc.add_paragraph()
                p_st.paragraph_format.space_before = Pt(1)
                p_st.paragraph_format.space_after = Pt(2)
                p_st.paragraph_format.left_indent = Inches(0.2)
                p_st.paragraph_format.line_spacing = 1.15

                lbl = st.get("label", "a")
                r = p_st.add_run(f"{lbl}) {st.get('content', '')} ")
                r.font.name = "Times New Roman"
                r.font.size = Pt(10.5)
                r_box = p_st.add_run("[   ]")
                r_box.font.name = "Consolas"
                r_box.font.size = Pt(10)
                r_box.bold = True

        # ----------------- PHẦN III: TỰ LUẬN & BÀI TẬP VẬN DỤNG -----------------
        essays = study_guide.get("essays", [])
        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(14)
        p_p3.paragraph_format.space_after = Pt(4)
        p_p3.paragraph_format.keep_with_next = True
        r = p_p3.add_run(f"PHẦN III. CÂU HỎI TỰ LUẬN VÀ BÀI TẬP THỰC HÀNH (Gồm {len(essays)} bài)")
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
        r.font.color.rgb = RGBColor(15, 44, 89)

        p_p3_sub = doc.add_paragraph()
        p_p3_sub.paragraph_format.space_after = Pt(6)
        r = p_p3_sub.add_run("Học sinh đọc kỹ câu hỏi và trình bày chi tiết lời giải / các bước thực hiện vào phần dòng kẻ bên dưới:")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r.italic = True

        for es in essays:
            es_num = es.get("guide_num", 1)
            p_es = doc.add_paragraph()
            p_es.paragraph_format.space_before = Pt(6)
            p_es.paragraph_format.space_after = Pt(3)
            p_es.paragraph_format.keep_with_next = True

            r = p_es.add_run(f"Bài tập {es_num}: ")
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)
            r.bold = True
            r = p_es.add_run(es.get("content", ""))
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)

            # Dòng kẻ chấm chấm full khổ giấy để học sinh làm bài
            p_bl_label = doc.add_paragraph()
            p_bl_label.paragraph_format.space_before = Pt(2)
            p_bl_label.paragraph_format.space_after = Pt(1)
            r_bl = p_bl_label.add_run("Bài làm:")
            r_bl.font.name = "Times New Roman"
            r_bl.font.size = Pt(10)
            r_bl.font.italic = True

            for _ in range(3):
                p_dot = doc.add_paragraph()
                p_dot.paragraph_format.space_before = Pt(0)
                p_dot.paragraph_format.space_after = Pt(2)
                p_dot.paragraph_format.line_spacing = 1.35
                p_dot.paragraph_format.tab_stops.add_tab_stop(Inches(6.50), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                r_dot = p_dot.add_run("\t")
                r_dot.font.name = "Times New Roman"
                r_dot.font.size = Pt(10)
                r_dot.font.color.rgb = RGBColor(148, 163, 184)

        # ----------------- 4. BẢNG ĐÁP ÁN & HƯỚNG DẪN GIẢI CHI TIẾT (NGẮT TRANG) -----------------
        doc.add_page_break()

        p_ans_head = doc.add_paragraph()
        p_ans_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ans_head.paragraph_format.space_before = Pt(10)
        p_ans_head.paragraph_format.space_after = Pt(2)
        r = p_ans_head.add_run("ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI CHI TIẾT ĐỀ CƯƠNG ÔN TẬP")
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.bold = True
        r.font.color.rgb = RGBColor(15, 44, 89)

        p_ans_sub = doc.add_paragraph()
        p_ans_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ans_sub.paragraph_format.space_after = Pt(10)
        r = p_ans_sub.add_run(f"MÔN: TIN HỌC {grade.upper()} — HỌC KỲ I / NĂM HỌC 2025 - 2026")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.italic = True

        # 4.1. Bảng đáp án Trắc nghiệm Nhiều lựa chọn
        p_ans_mcq = doc.add_paragraph()
        p_ans_mcq.paragraph_format.space_before = Pt(6)
        p_ans_mcq.paragraph_format.space_after = Pt(4)
        r = p_ans_mcq.add_run("1. BẢNG ĐÁP ÁN PHẦN I — TRẮC NGHIỆM NHIỀU LỰA CHỌN:")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11.5)
        r.bold = True

        # Tạo bảng 10 cột cho đáp án MCQ
        num_mcq = len(mcqs)
        cols_per_row = 10
        num_rows = (num_mcq + cols_per_row - 1) // cols_per_row * 2
        tbl_ans_mcq = doc.add_table(rows=num_rows, cols=cols_per_row)
        tbl_ans_mcq.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_ans_mcq.autofit = False

        for row_block in range(num_rows // 2):
            r_head = tbl_ans_mcq.rows[row_block * 2]
            r_val = tbl_ans_mcq.rows[row_block * 2 + 1]
            for c_idx in range(cols_per_row):
                q_idx = row_block * cols_per_row + c_idx
                if q_idx < num_mcq:
                    q_item = mcqs[q_idx]
                    # Cell Câu
                    c_h = r_head.cells[c_idx]
                    c_h.text = f"C{q_item.get('guide_num', q_idx+1)}"
                    StudyGuideDocxExporter._format_cell(c_h, bold=True, bg_color="E2E8F0")
                    # Cell Đáp án
                    c_v = r_val.cells[c_idx]
                    c_v.text = q_item.get("correct", "A")
                    StudyGuideDocxExporter._format_cell(c_v, bold=True, bg_color="F8FAFC", color_rgb=RGBColor(14, 116, 144))
                else:
                    r_head.cells[c_idx].text = "-"
                    r_val.cells[c_idx].text = "-"

        # 4.2. Đáp án Phần II: Đúng/Sai
        p_ans_tf = doc.add_paragraph()
        p_ans_tf.paragraph_format.space_before = Pt(12)
        p_ans_tf.paragraph_format.space_after = Pt(4)
        r = p_ans_tf.add_run("2. BẢNG ĐÁP ÁN PHẦN II — TRẮC NGHIỆM ĐÚNG / SAI:")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11.5)
        r.bold = True

        tbl_ans_tf = doc.add_table(rows=len(tfs) + 1, cols=5)
        tbl_ans_tf.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_ans_tf.autofit = False

        hdr_cells = tbl_ans_tf.rows[0].cells
        hdr_cells[0].text = "Câu hỏi tình huống"
        hdr_cells[1].text = "Ý a)"
        hdr_cells[2].text = "Ý b)"
        hdr_cells[3].text = "Ý c)"
        hdr_cells[4].text = "Ý d)"
        for c in hdr_cells:
            StudyGuideDocxExporter._format_cell(c, bold=True, bg_color="E2E8F0")

        for idx, tf in enumerate(tfs, 1):
            r_cells = tbl_ans_tf.rows[idx].cells
            r_cells[0].text = f"Câu {tf.get('guide_num', idx)}"
            StudyGuideDocxExporter._format_cell(r_cells[0], bold=True)
            for st_idx, st in enumerate(tf.get("statements", [])[:4]):
                c_target = r_cells[st_idx + 1]
                c_target.text = str(st.get("correct", "Đ")).upper()
                bg = "DCFCE7" if c_target.text == "Đ" else "FEE2E2"
                StudyGuideDocxExporter._format_cell(c_target, bold=True, bg_color=bg)

        # 4.3. Hướng dẫn giải chi tiết Tự luận
        p_ans_es = doc.add_paragraph()
        p_ans_es.paragraph_format.space_before = Pt(12)
        p_ans_es.paragraph_format.space_after = Pt(4)
        r = p_ans_es.add_run("3. GỢI Ý TRẢ LỜI VÀ BIỂU ĐIỂM CHI TIẾT PHẦN III — TỰ LUẬN:")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11.5)
        r.bold = True

        for es in essays:
            p_e = doc.add_paragraph()
            p_e.paragraph_format.space_before = Pt(4)
            p_e.paragraph_format.space_after = Pt(2)
            r = p_e.add_run(f"Bài tập {es.get('guide_num', 1)}: ")
            r.font.name = "Times New Roman"
            r.font.size = Pt(10.5)
            r.bold = True
            r = p_e.add_run(es.get("content", ""))
            r.font.name = "Times New Roman"
            r.font.size = Pt(10.5)
            r.italic = True

            # Các tiêu chí chấm
            for rub in es.get("rubric", []):
                p_rub = doc.add_paragraph(style="List Bullet")
                p_rub.paragraph_format.space_before = Pt(1)
                p_rub.paragraph_format.space_after = Pt(2)
                p_rub.paragraph_format.left_indent = Inches(0.25)
                r = p_rub.add_run(f"{rub.get('criterion', '')} ")
                r.font.name = "Times New Roman"
                r.font.size = Pt(10)
                r_sc = p_rub.add_run(f"({rub.get('score', 0.25)} điểm)")
                r_sc.font.name = "Times New Roman"
                r_sc.font.size = Pt(10)
                r_sc.bold = True

        doc.save(output_path)

    @staticmethod
    def export_text(study_guide: Dict[str, Any]) -> str:
        """Xuất Đề Cương Ôn Tập ra dạng văn bản (.txt) dễ chia sẻ và copy."""
        meta = study_guide.get("meta", {})
        grade = study_guide.get("grade", "Lớp 6")
        lines = []

        lines.append("=" * 65)
        lines.append(f"{meta.get('governing_body', 'UBND HUYỆN').upper()} - {meta.get('school_name', 'TRƯỜNG THCS').upper()}")
        lines.append(f"ĐỀ CƯƠNG ÔN TẬP {meta.get('exam_title', 'KIỂM TRA ĐỊNH KỲ').upper()}")
        lines.append(f"MÔN: TIN HỌC {grade.upper()} — NĂM HỌC: {meta.get('school_year', '2025 - 2026')}")
        lines.append("=" * 65)
        lines.append("")

        mcqs = study_guide.get("mcqs", [])
        lines.append(f"PHẦN I. CÂU HỎI TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN ({len(mcqs)} câu)")
        lines.append("-" * 65)
        for q in mcqs:
            lines.append(f"Câu {q.get('guide_num', 1)}: {q.get('content', '')}")
            opts = q.get("options", {})
            lines.append(f"  A. {opts.get('A', '')}")
            lines.append(f"  B. {opts.get('B', '')}")
            lines.append(f"  C. {opts.get('C', '')}")
            lines.append(f"  D. {opts.get('D', '')}")
            lines.append("")

        tfs = study_guide.get("tfs", [])
        lines.append(f"PHẦN II. CÂU HỎI TRẮC NGHIỆM ĐÚNG / SAI ({len(tfs)} câu)")
        lines.append("-" * 65)
        for tf in tfs:
            lines.append(f"Câu {tf.get('guide_num', 1)}: {tf.get('topic', '')}")
            for st in tf.get("statements", []):
                lines.append(f"  {st.get('label', '')}) {st.get('content', '')}")
            lines.append("")

        essays = study_guide.get("essays", [])
        lines.append(f"PHẦN III. CÂU HỎI TỰ LUẬN VÀ BÀI TẬP THỰC HÀNH ({len(essays)} bài)")
        lines.append("-" * 65)
        for es in essays:
            lines.append(f"Bài tập {es.get('guide_num', 1)}: {es.get('content', '')}")
            lines.append("")

        lines.append("=" * 65)
        lines.append("ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI CHI TIẾT ĐỀ CƯƠNG ÔN TẬP")
        lines.append("=" * 65)
        lines.append("1. ĐÁP ÁN TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN:")
        ans_str = " | ".join([f"C{q.get('guide_num')}: {q.get('correct')}" for q in mcqs])
        lines.append(ans_str)
        lines.append("")

        lines.append("2. ĐÁP ÁN TRẮC NGHIỆM ĐÚNG / SAI:")
        for tf in tfs:
            st_str = ", ".join([f"{st.get('label')}: {st.get('correct')}" for st in tf.get("statements", [])])
            lines.append(f"  Câu {tf.get('guide_num')}: {st_str}")
        lines.append("")

        lines.append("3. GỢI Ý TRẢ LỜI VÀ BIỂU ĐIỂM TỰ LUẬN:")
        for es in essays:
            lines.append(f"  * Bài tập {es.get('guide_num')}:")
            for rub in es.get("rubric", []):
                lines.append(f"    - {rub.get('criterion')}: {rub.get('score')}đ")
            lines.append("")

        return "\n".join(lines)

    @staticmethod
    def _remove_table_borders(table):
        """Xóa viền bảng để làm header vô hình."""
        tblPr = table._tbl.tblPr
        tblBorders = tblPr.first_child_found_in("w:tblBorders")
        if tblBorders is not None:
            tblPr.remove(tblBorders)
        new_borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            '<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
            '<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>'
            '</w:tblBorders>'
        )
        tblPr.append(new_borders)

    @staticmethod
    def _format_cell(cell, bold: bool = False, bg_color: Optional[str] = None, color_rgb: Optional[RGBColor] = None):
        """Định dạng ô bảng nhanh."""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(9.5)
            r.bold = bold
            if color_rgb:
                r.font.color.rgb = color_rgb
        if bg_color:
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            cell._tc.get_or_add_tcPr().append(shd)
