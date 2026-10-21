"""
core/exam_exporters.py
Cung cấp các định dạng xuất tải xuống đa dạng cho đề thi môn Tin học:
1. Xuất file Text (.txt) chuẩn Azota/Google Forms/Quizizz.
2. Xuất file Excel (.xlsx) gồm Ma Trận & Bản Đặc Tả 18 cột và 13 cột.
3. Xuất file JSON (.json) phục vụ tích hợp LMS hoặc Neo4j Knowledge Graph.
"""

import io
import json
from typing import Dict, Any
import pandas as pd


class ExamAuxiliaryExporter:
    """Bộ xuất tài liệu bổ trợ ngoài file Word (.docx)."""

    @staticmethod
    def export_text_format(dossier: Dict[str, Any], exam_code: str = "th01") -> str:
        """Xuất đề thi và bảng đáp án dưới dạng văn bản (.txt) dễ copy/paste."""
        meta = dossier.get("meta", {})
        paper = dossier.get(exam_code, dossier.get("th01", {}))
        code = paper.get("code", "TH01")

        lines = []
        lines.append(f"UBND XÃ / ĐƠN VỊ: {meta.get('governing_body', '')}")
        lines.append(f"TRƯỜNG: {meta.get('school_name', '')}")
        lines.append(f"ĐỀ THI: {meta.get('exam_title', '')} - NĂM HỌC: {meta.get('school_year', '')}")
        lines.append(f"MÔN: TIN HỌC {meta.get('grade', '').upper()} - THỜI GIAN: {meta.get('duration', 45)} PHÚT")
        lines.append(f"MÃ ĐỀ THI: {code}")
        lines.append("=" * 60)
        lines.append("")

        # PHẦN I
        lines.append("PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 điểm)")
        lines.append("Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi chỉ chọn một phương án.")
        lines.append("-" * 60)
        mcqs = paper.get("mcqs", [])
        for idx, q in enumerate(mcqs, 1):
            lines.append(f"Câu {idx}: {q.get('content', '')}")
            opts = q.get("options", {})
            lines.append(f"  A. {opts.get('A', '')}")
            lines.append(f"  B. {opts.get('B', '')}")
            lines.append(f"  C. {opts.get('C', '')}")
            lines.append(f"  D. {opts.get('D', '')}")
            lines.append("")

        # PHẦN II
        lines.append("PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG / SAI (4,0 điểm)")
        lines.append("Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, chọn Đúng hoặc Sai.")
        lines.append("-" * 60)
        tfs = paper.get("tfs", [])
        for idx, tf in enumerate(tfs, 1):
            lines.append(f"Câu {idx}: {tf.get('topic', '')}")
            for st in tf.get("statements", []):
                lines.append(f"  {st.get('label', '')}) {st.get('content', '')}")
            lines.append("")

        # PHẦN III
        lines.append("PHẦN III. TỰ LUẬN (3,0 điểm)")
        lines.append("Thí sinh làm bài tự luận theo yêu cầu của từng câu hỏi.")
        lines.append("-" * 60)
        essays = paper.get("essays", [])
        for idx, es in enumerate(essays, 1):
            num_order = 16 + idx  # Câu 17, 18, 19
            lines.append(f"Câu {num_order} ({es.get('score', 1.0)} điểm): {es.get('content', '')}")
            lines.append("")

        # ĐÁP ÁN & HƯỚNG DẪN CHẤM
        lines.append("=" * 60)
        lines.append(f"ĐÁP ÁN VÀ THANG ĐIỂM CHI TIẾT - MÃ ĐỀ: {code}")
        lines.append("=" * 60)
        lines.append("1. ĐÁP ÁN PHẦN I (0,25 điểm / câu):")
        ans_i_str = " | ".join([f"Câu {i}: {q.get('correct', 'A')}" for i, q in enumerate(mcqs, 1)])
        lines.append(ans_i_str)
        lines.append("")

        lines.append("2. ĐÁP ÁN PHẦN II (Đúng/Sai):")
        for idx, tf in enumerate(tfs, 1):
            stmts_ans = ", ".join([f"{st.get('label', '')}: {st.get('correct', 'Đ')}" for st in tf.get("statements", [])])
            lines.append(f"  - Câu {idx}: {stmts_ans}")
        lines.append("")

        lines.append("3. HƯỚNG DẪN CHẤM PHẦN III (TỰ LUẬN):")
        for idx, es in enumerate(essays, 1):
            num_order = 16 + idx
            lines.append(f"  * Câu {num_order} ({es.get('score', 1.0)} điểm):")
            for rub in es.get("rubric", []):
                lines.append(f"    + {rub.get('criterion', '')}: {rub.get('score', 0.25)}đ")
            lines.append("")

        return "\n".join(lines)

    @staticmethod
    def export_excel_matrix_spec(dossier: Dict[str, Any]) -> bytes:
        """Xuất Khung Ma Trận và Bản Đặc Tả thành file Excel (.xlsx) chuẩn."""
        matrix_rows = dossier.get("matrix_rows", {}).get("units", [])
        spec_rows = dossier.get("spec_rows", [])

        df_matrix = pd.DataFrame(matrix_rows)
        if not df_matrix.empty:
            df_matrix = df_matrix.rename(columns={
                "tt": "TT",
                "chapter": "Chương / Chủ đề",
                "unit": "Nội dung / Đơn vị kiến thức",
                "mcq_nb": "TNKQ (Biết)",
                "mcq_th": "TNKQ (Hiểu)",
                "mcq_vd": "TNKQ (Vận dụng)",
                "tf_nb": "Đúng/Sai (Biết)",
                "tf_th": "Đúng/Sai (Hiểu)",
                "tf_vd": "Đúng/Sai (Vận dụng)",
                "es_nb": "Tự luận (Biết)",
                "es_th": "Tự luận (Hiểu)",
                "es_vd": "Tự luận (Vận dụng)",
                "total_score": "Tổng điểm / Tỷ lệ"
            })

        df_spec = pd.DataFrame(spec_rows)
        if not df_spec.empty:
            df_spec = df_spec.rename(columns={
                "tt": "TT",
                "chapter": "Chương / Chủ đề",
                "unit": "Nội dung / Đơn vị kiến thức",
                "requirement": "Mức độ đánh giá / Yêu cầu cần đạt",
                "mcq_nb": "TNKQ (Biết)",
                "mcq_th": "TNKQ (Hiểu)",
                "mcq_vd": "TNKQ (Vận dụng)",
                "tf_nb": "Đúng/Sai (Biết)",
                "tf_th": "Đúng/Sai (Hiểu)",
                "tf_vd": "Đúng/Sai (Vận dụng)",
                "es_nb": "Tự luận (Biết)",
                "es_th": "Tự luận (Hiểu)",
                "es_vd": "Tự luận (Vận dụng)"
            })

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            if not df_matrix.empty:
                df_matrix.to_excel(writer, index=False, sheet_name="Khung_Ma_Tran")
            if not df_spec.empty:
                df_spec.to_excel(writer, index=False, sheet_name="Ban_Dac_Ta")
        return output.getvalue()
