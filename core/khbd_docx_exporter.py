# -*- coding: utf-8 -*-
"""
core/khbd_docx_exporter.py
Bộ xuất file Word (.docx) cho Kế hoạch bài dạy (KHBD) chuẩn Công văn 5512/BGDĐT-GDTrH
và Tích hợp Khung Năng lực số (NLS) theo Thông tư 02/2025/TT-BGDĐT.
Tuân thủ nghiêm ngặt kỹ thuật format:
- Bảng tiêu đề (Table 0): 190pt / 300pt, tên GV nằm trên 1 dòng duy nhất
- Đánh dấu chữ đỏ RGB(255, 0, 0) cho NLS trong mục tiêu, tiến trình và bảng tổng hợp
- Khối phê duyệt 2 cột chống tách trang (cantSplit, keep_with_next, space_after=36pt)
"""

import os
import re
from typing import Dict, Any, List, Optional
import docx
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

COLOR_RED = RGBColor(255, 0, 0)
COLOR_DARK_BLUE = RGBColor(0, 51, 102)
COLOR_DARK = RGBColor(30, 41, 59)
COLOR_MUTED = RGBColor(100, 116, 139)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Thiết lập padding cho ô trong bảng (đơn vị dxa: 1pt = 20 dxa)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_width_dxa(cell, width_pt: float):
    """Thiết lập độ rộng tuyệt đối cho ô"""
    cell.width = Pt(width_pt)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(width_pt * 20)))
    tcW.set(qn('w:type'), 'dxa')


def set_cell_shading(cell, color_hex: str):
    """Tô màu nền cho ô (vd: '#F1F5F9' -> 'F1F5F9')"""
    hex_clean = color_hex.replace("#", "")
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_clean}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))


def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    """Đặt đường viền nét mỏng đẹp cho bảng"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def set_table_header_borders(table, color="94A3B8"):
    """Đặt viền chuyên nghiệp cho bảng tiến trình và bảng NLS"""
    set_table_borders(table, color=color, sz="4", val="single")


def make_row_cant_split(row):
    """Chống ngắt trang ngang thân dòng trong bảng"""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None:
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))


def add_formatted_runs_with_red_nls(paragraph, text: str, default_font_size: float = 12.0, default_bold: bool = False):
    """
    Phân tích chuỗi văn bản và tự động tô màu đỏ RGB(255, 0, 0)
    cho các đoạn được bọc thẻ [NLS-RED]...[/NLS-RED] hoặc các mã NLS chuẩn [X.X.TC2x].
    """
    if not text:
        return

    # Nếu có thẻ [NLS-RED]...[/NLS-RED]
    if "[NLS-RED]" in text:
        parts = re.split(r'(\[NLS-RED\].*?\[/NLS-RED\])', text, flags=re.DOTALL)
        for part in parts:
            if part.startswith("[NLS-RED]") and part.endswith("[/NLS-RED]"):
                clean = part.replace("[NLS-RED]", "").replace("[/NLS-RED]", "")
                r = paragraph.add_run(clean)
                r.font.name = "Times New Roman"
                r.font.size = Pt(default_font_size)
                r.font.color.rgb = COLOR_RED
                r.bold = True
            else:
                # Kiểm tra tiếp mã [X.X.TC2x] bên trong phần thường
                _add_sub_runs_with_nls_codes(paragraph, part, default_font_size, default_bold)
    else:
        _add_sub_runs_with_nls_codes(paragraph, text, default_font_size, default_bold)


def _add_sub_runs_with_nls_codes(paragraph, text: str, default_font_size: float, default_bold: bool):
    """Tự động tìm mã NLS theo quy chuẩn dạng [1.2.TC2a] hoặc 1.2.TC2a để bôi đỏ"""
    pattern = r'(\[?\b\d\.\d\.TC2[a-z]?\]?)'
    tokens = re.split(pattern, text)
    for token in tokens:
        if re.match(r'^\[?\b\d\.\d\.TC2[a-z]?\]?$', token):
            r = paragraph.add_run(token)
            r.font.name = "Times New Roman"
            r.font.size = Pt(default_font_size)
            r.font.color.rgb = COLOR_RED
            r.bold = True
        else:
            if token:
                r = paragraph.add_run(token)
                r.font.name = "Times New Roman"
                r.font.size = Pt(default_font_size)
                r.bold = default_bold


class KHBDDocxExporter:
    """Class phụ trách kết xuất Kế hoạch bài dạy ra tài liệu Word (.docx) chuẩn mẫu."""

    @classmethod
    def export_khbd_to_docx(cls, khbd_data: Dict[str, Any], output_path: str) -> str:
        """
        Tạo tài liệu Word KHBD hoàn chỉnh từ dữ liệu JSON.
        """
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc = docx.Document()

        # 1. Cấu hình lề trang theo quy chuẩn (Khổ A4 đứng)
        section = doc.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(1.5)

        # Cấu hình Normal style
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.space_after = Pt(2)

        header_data = khbd_data.get("header", {})
        week_str = str(header_data.get("week", "1"))
        period_str = str(header_data.get("period", "1"))
        date_str = str(header_data.get("date_range", ""))
        teacher_name = str(header_data.get("teacher_name", "Trần Đức Tình"))
        department = str(header_data.get("department", "Tổ Tự Nhiên"))

        # 2. BẢNG TIÊU ĐỀ (Table 0 - Header Box): 190pt / 300pt chống rớt dòng tên GV
        cls._render_table_0(doc, week_str, period_str, date_str, teacher_name, department)

        # 3. TÊN BÀI DẠY (Căn giữa, in hoa, màu xanh đậm)
        title_text = str(khbd_data.get("title", "KẾ HOẠCH BÀI DẠY"))
        sub_title_text = str(khbd_data.get("sub_title", "Môn Tin học; Lớp 6 | Thời gian thực hiện: 1 tiết"))

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(2)
        r_t = p_title.add_run(f"TÊN BÀI DẠY: {title_text}")
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(14)
        r_t.bold = True
        r_t.font.color.rgb = COLOR_DARK_BLUE

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(10)
        r_sub = p_sub.add_run(sub_title_text)
        r_sub.font.name = "Times New Roman"
        r_sub.font.size = Pt(12)
        r_sub.italic = True

        # 4. I. MỤC TIÊU
        cls._render_section_i(doc, khbd_data.get("objectives", {}))

        # 5. II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
        cls._render_section_ii(doc, khbd_data.get("equipment", {}))

        # 6. III. TIẾN TRÌNH DẠY HỌC (4 HOẠT ĐỘNG CHUẨN 5512)
        cls._render_section_iii(doc, khbd_data.get("activities", []))

        # 7. IV. BẢNG TỔNG HỢP NĂNG LỰC SỐ (NLS) TÍCH HỢP TRONG BÀI DẠY
        cls._render_section_iv(doc, khbd_data.get("nls_summary_table", []))

        # 8. V. KHỐI PHÊ DUYỆT (CHỐNG TÁCH TRANG)
        approvals = khbd_data.get("approvals", {})
        cls._render_section_v(
            doc,
            cm_title=approvals.get("cm_title", "Duyệt của Chuyên môn"),
            cm_signer=approvals.get("cm_signer", "Lưu Văn Chí"),
            to_title=approvals.get("to_title", "Duyệt của tổ chuyên môn"),
            to_signer=approvals.get("to_signer", "Hoàng Thị Kim Tuyết")
        )

        doc.save(output_path)
        return output_path

    @classmethod
    def _render_table_0(cls, doc, week_str, period_str, date_str, teacher_name, department):
        """Bảng Table 0: Cột trái 190pt, Cột phải 300pt, có khung viền mỏng"""
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        make_row_cant_split(tbl.rows[0])
        set_table_borders(tbl, color="CBD5E1", sz="4", val="single")

        c0 = tbl.cell(0, 0)
        c1 = tbl.cell(0, 1)
        set_cell_width_dxa(c0, 190)
        set_cell_width_dxa(c1, 300)
        set_cell_margins(c0, top=100, bottom=100, left=140, right=140)
        set_cell_margins(c1, top=100, bottom=100, left=140, right=140)

        # Cột trái: Tuần, Tiết, Ngày dạy
        p0 = c0.paragraphs[0]
        p0.paragraph_format.line_spacing = 1.15
        p0.paragraph_format.space_after = Pt(2)
        p0.add_run("Tuần: ").font.size = Pt(12)
        r_w = p0.add_run(week_str)
        r_w.bold = True
        r_w.font.size = Pt(12)

        p1 = c0.add_paragraph()
        p1.paragraph_format.line_spacing = 1.15
        p1.paragraph_format.space_after = Pt(2)
        p1.add_run("Tiết: ").font.size = Pt(12)
        r_p = p1.add_run(period_str)
        r_p.bold = True
        r_p.font.size = Pt(12)

        if date_str:
            p2 = c0.add_paragraph()
            p2.paragraph_format.line_spacing = 1.15
            p2.paragraph_format.space_after = Pt(0)
            p2.add_run("Ngày dạy: ").font.size = Pt(11.5)
            r_d = p2.add_run(date_str)
            r_d.italic = True
            r_d.font.size = Pt(11.5)

        # Cột phải: Họ và tên giáo viên (1 dòng duy nhất) & Tổ chuyên môn
        p0_1 = c1.paragraphs[0]
        p0_1.paragraph_format.line_spacing = 1.15
        p0_1.paragraph_format.space_after = Pt(2)
        p0_1.add_run("Họ và tên giáo viên: ").font.size = Pt(12)
        r_t = p0_1.add_run(teacher_name)
        r_t.bold = True
        r_t.font.size = Pt(12)

        p1_1 = c1.add_paragraph()
        p1_1.paragraph_format.line_spacing = 1.15
        p1_1.paragraph_format.space_after = Pt(0)
        p1_1.add_run("Tổ chuyên môn: ").font.size = Pt(12)
        r_to = p1_1.add_run(department)
        r_to.bold = True
        r_to.font.size = Pt(12)

    @classmethod
    def _render_section_i(cls, doc, objectives: Dict[str, Any]):
        """Render Phần I. MỤC TIÊU với NLS màu đỏ"""
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(8)
        p_sec.paragraph_format.space_after = Pt(2)
        r_sec = p_sec.add_run("I. MỤC TIÊU")
        r_sec.bold = True
        r_sec.font.size = Pt(12.5)

        # 1. Kiến thức
        p_k_head = doc.add_paragraph()
        p_k_head.paragraph_format.space_before = Pt(2)
        p_k_head.paragraph_format.space_after = Pt(2)
        r_kh = p_k_head.add_run("1. Về kiến thức:")
        r_kh.bold = True

        knowledge_list = objectives.get("knowledge", [])
        if isinstance(knowledge_list, list):
            for k in knowledge_list:
                p_item = doc.add_paragraph(style='List Bullet')
                p_item.paragraph_format.space_after = Pt(1.5)
                add_formatted_runs_with_red_nls(p_item, str(k), 12.0)
        elif knowledge_list:
            p_item = doc.add_paragraph()
            add_formatted_runs_with_red_nls(p_item, str(knowledge_list), 12.0)

        # 2. Năng lực
        p_c_head = doc.add_paragraph()
        p_c_head.paragraph_format.space_before = Pt(4)
        p_c_head.paragraph_format.space_after = Pt(2)
        r_ch = p_c_head.add_run("2. Về năng lực:")
        r_ch.bold = True

        comp_data = objectives.get("competencies", {})
        # 2.1 Năng lực chung
        p_g = doc.add_paragraph()
        p_g.paragraph_format.left_indent = Cm(0.5)
        r_g = p_g.add_run("2.1. Năng lực chung:")
        r_g.bold = True
        for item in comp_data.get("general", []):
            p_item = doc.add_paragraph(style='List Bullet')
            p_item.paragraph_format.left_indent = Cm(1.0)
            p_item.paragraph_format.space_after = Pt(1.5)
            add_formatted_runs_with_red_nls(p_item, str(item), 12.0)

        # 2.2 Năng lực đặc thù
        p_s = doc.add_paragraph()
        p_s.paragraph_format.left_indent = Cm(0.5)
        r_s = p_s.add_run("2.2. Năng lực đặc thù:")
        r_s.bold = True
        for item in comp_data.get("specific", []):
            p_item = doc.add_paragraph(style='List Bullet')
            p_item.paragraph_format.left_indent = Cm(1.0)
            p_item.paragraph_format.space_after = Pt(1.5)
            add_formatted_runs_with_red_nls(p_item, str(item), 12.0)

        # 2.3 Năng lực số (NLS) - BẮT BUỘC CHỮ ĐỎ
        p_nls = doc.add_paragraph()
        p_nls.paragraph_format.left_indent = Cm(0.5)
        p_nls.paragraph_format.space_before = Pt(2)
        r_nls = p_nls.add_run("2.3. Năng lực số (NLS) tích hợp theo Thông tư 02/2025/TT-BGDĐT:")
        r_nls.bold = True
        r_nls.font.color.rgb = COLOR_RED

        digital_list = comp_data.get("digital", [])
        if digital_list:
            for d in digital_list:
                p_item = doc.add_paragraph(style='List Bullet')
                p_item.paragraph_format.left_indent = Cm(1.0)
                p_item.paragraph_format.space_after = Pt(1.5)

                if isinstance(d, dict):
                    code = d.get("code", "")
                    desc = d.get("desc", "")
                    r_code = p_item.add_run(f"[{code}]: ")
                    r_code.bold = True
                    r_code.font.color.rgb = COLOR_RED
                    r_desc = p_item.add_run(desc)
                    r_desc.font.color.rgb = COLOR_RED
                else:
                    r_txt = p_item.add_run(str(d))
                    r_txt.font.color.rgb = COLOR_RED

        # 3. Phẩm chất
        p_q_head = doc.add_paragraph()
        p_q_head.paragraph_format.space_before = Pt(4)
        p_q_head.paragraph_format.space_after = Pt(2)
        r_qh = p_q_head.add_run("3. Về phẩm chất:")
        r_qh.bold = True

        qualities_list = objectives.get("qualities", [])
        if isinstance(qualities_list, list):
            for q in qualities_list:
                p_item = doc.add_paragraph(style='List Bullet')
                p_item.paragraph_format.space_after = Pt(1.5)
                add_formatted_runs_with_red_nls(p_item, str(q), 12.0)
        elif qualities_list:
            p_item = doc.add_paragraph()
            add_formatted_runs_with_red_nls(p_item, str(qualities_list), 12.0)

    @classmethod
    def _render_section_ii(cls, doc, equipment: Dict[str, Any]):
        """Render Phần II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU"""
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(8)
        p_sec.paragraph_format.space_after = Pt(2)
        r_sec = p_sec.add_run("II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU")
        r_sec.bold = True
        r_sec.font.size = Pt(12.5)

        teacher_eq = equipment.get("teacher", [])
        p_t = doc.add_paragraph()
        r_th = p_t.add_run("1. Giáo viên: ")
        r_th.bold = True
        if isinstance(teacher_eq, list):
            p_t.add_run("; ".join(teacher_eq) + ".")
        else:
            p_t.add_run(str(teacher_eq))

        student_eq = equipment.get("student", [])
        p_s = doc.add_paragraph()
        r_sh = p_s.add_run("2. Học sinh: ")
        r_sh.bold = True
        if isinstance(student_eq, list):
            p_s.add_run("; ".join(student_eq) + ".")
        else:
            p_s.add_run(str(student_eq))

    @classmethod
    def _render_section_iii(cls, doc, activities: List[Dict[str, Any]]):
        """Render Phần III. TIẾN TRÌNH DẠY HỌC (BẢNG 4 HOẠT ĐỘNG CHUẨN 5512)"""
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(8)
        p_sec.paragraph_format.space_after = Pt(4)
        r_sec = p_sec.add_run("III. TIẾN TRÌNH DẠY HỌC")
        r_sec.bold = True
        r_sec.font.size = Pt(12.5)

        for act in activities:
            act_name = str(act.get("name", "Hoạt động"))
            act_time = str(act.get("time", "")).strip()
            if act_time:
                full_act_name = f"{act_name} [{act_time}]"
            else:
                full_act_name = act_name

            p_act = doc.add_paragraph()
            p_act.paragraph_format.space_before = Pt(6)
            p_act.paragraph_format.space_after = Pt(2)
            p_act.paragraph_format.keep_with_next = True
            r_an = p_act.add_run(full_act_name)
            r_an.bold = True
            r_an.font.size = Pt(12)
            r_an.font.color.rgb = COLOR_DARK_BLUE

            # a) Mục tiêu
            goal = str(act.get("goal", "")).strip()
            if goal:
                p_g = doc.add_paragraph()
                p_g.paragraph_format.left_indent = Cm(0.3)
                p_g.paragraph_format.space_after = Pt(1.5)
                r_gh = p_g.add_run("a) Mục tiêu: ")
                r_gh.bold = True
                add_formatted_runs_with_red_nls(p_g, goal, 12.0)

            # b) Nội dung
            content = str(act.get("content", "")).strip()
            if content:
                p_c = doc.add_paragraph()
                p_c.paragraph_format.left_indent = Cm(0.3)
                p_c.paragraph_format.space_after = Pt(1.5)
                r_ch = p_c.add_run("b) Nội dung: ")
                r_ch.bold = True
                add_formatted_runs_with_red_nls(p_c, content, 12.0)

            # c) Sản phẩm
            product = str(act.get("product", "")).strip()
            if product:
                p_p = doc.add_paragraph()
                p_p.paragraph_format.left_indent = Cm(0.3)
                p_p.paragraph_format.space_after = Pt(1.5)
                r_ph = p_p.add_run("c) Sản phẩm: ")
                r_ph.bold = True
                add_formatted_runs_with_red_nls(p_p, product, 12.0)

            # d) Tổ chức thực hiện: Bảng 2 cột (310pt / 190pt)
            p_d = doc.add_paragraph()
            p_d.paragraph_format.left_indent = Cm(0.3)
            p_d.paragraph_format.space_before = Pt(2)
            p_d.paragraph_format.space_after = Pt(3)
            p_d.paragraph_format.keep_with_next = True
            r_dh = p_d.add_run("d) Tổ chức thực hiện:")
            r_dh.bold = True

            steps = act.get("steps", [])
            cls._render_activity_table(doc, steps, default_product=product)

    @classmethod
    def _render_activity_table(cls, doc, steps: List[Dict[str, Any]], default_product: str = ""):
        """Tạo bảng 2 cột cho các bước trong hoạt động: Cột trái 310pt, Cột phải 190pt"""
        # Số hàng: 1 hàng header + số bước
        rows_count = 1 + max(1, len(steps))
        tbl = doc.add_table(rows=rows_count, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_borders(tbl, color="CBD5E1", sz="4", val="single")

        # Header hàng 0
        h_row = tbl.rows[0]
        make_row_cant_split(h_row)
        c0 = h_row.cells[0]
        c1 = h_row.cells[1]
        set_cell_width_dxa(c0, 310)
        set_cell_width_dxa(c1, 190)
        set_cell_shading(c0, "F1F5F9")
        set_cell_shading(c1, "F1F5F9")
        set_cell_margins(c0, top=100, bottom=100, left=140, right=140)
        set_cell_margins(c1, top=100, bottom=100, left=140, right=140)

        p_h0 = c0.paragraphs[0]
        p_h0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h0.paragraph_format.space_after = Pt(0)
        r_h0 = p_h0.add_run("Hoạt động của GV và HS")
        r_h0.bold = True
        r_h0.font.size = Pt(11.5)

        p_h1 = c1.paragraphs[0]
        p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h1.paragraph_format.space_after = Pt(0)
        r_h1 = p_h1.add_run("Dự kiến sản phẩm")
        r_h1.bold = True
        r_h1.font.size = Pt(11.5)

        if not steps:
            # Fallback 1 dòng trống nếu không có steps
            row = tbl.rows[1]
            make_row_cant_split(row)
            set_cell_width_dxa(row.cells[0], 310)
            set_cell_width_dxa(row.cells[1], 190)
            return

        for idx, step in enumerate(steps, start=1):
            row = tbl.rows[idx]
            make_row_cant_split(row)
            cell_left = row.cells[0]
            cell_right = row.cells[1]
            set_cell_width_dxa(cell_left, 310)
            set_cell_width_dxa(cell_right, 190)
            set_cell_margins(cell_left, top=80, bottom=80, left=120, right=120)
            set_cell_margins(cell_right, top=80, bottom=80, left=120, right=120)

            step_name = step.get("step_name", f"Bước {idx}")
            teacher_action = step.get("teacher_action", "")
            student_action = step.get("student_action", "")
            step_prod = step.get("step_product", "")

            # Cột trái: Tên bước + Hoạt động GV + Hoạt động HS
            p_sname = cell_left.paragraphs[0]
            p_sname.paragraph_format.space_before = Pt(1)
            p_sname.paragraph_format.space_after = Pt(1)
            r_sn = p_sname.add_run(f"* {step_name}:")
            r_sn.bold = True
            r_sn.font.size = Pt(11.5)

            if teacher_action:
                p_t = cell_left.add_paragraph()
                p_t.paragraph_format.left_indent = Cm(0.3)
                p_t.paragraph_format.space_after = Pt(1.5)
                r_tl = p_t.add_run("- GV: ")
                r_tl.bold = True
                r_tl.font.size = Pt(11)
                add_formatted_runs_with_red_nls(p_t, teacher_action, 11.0)

            if student_action:
                p_s = cell_left.add_paragraph()
                p_s.paragraph_format.left_indent = Cm(0.3)
                p_s.paragraph_format.space_after = Pt(1.5)
                r_sl = p_s.add_run("- HS: ")
                r_sl.bold = True
                r_sl.font.size = Pt(11)
                add_formatted_runs_with_red_nls(p_s, student_action, 11.0)

            # Cột phải: Dự kiến sản phẩm của bước (hoặc sản phẩm chung ở bước 2/3)
            p_r = cell_right.paragraphs[0]
            p_r.paragraph_format.space_after = Pt(1.5)
            if step_prod:
                add_formatted_runs_with_red_nls(p_r, step_prod, 11.0)
            elif idx == 2 and default_product:
                add_formatted_runs_with_red_nls(p_r, default_product, 11.0)
            elif idx == 1:
                p_r.add_run("Học sinh tiếp nhận nhiệm vụ học tập rõ ràng.").font.size = Pt(11)
            elif idx == 3:
                p_r.add_run("Kết quả làm việc, báo cáo thảo luận của học sinh.").font.size = Pt(11)
            elif idx == 4:
                p_r.add_run("Nội dung kiến thức cốt lõi được chuẩn hóa vào vở.").font.size = Pt(11)

    @classmethod
    def _render_section_iv(cls, doc, nls_table: List[Dict[str, Any]]):
        """Render Phần IV. BẢNG TỔNG HỢP NĂNG LỰC SỐ (NLS) TÍCH HỢP TRONG BÀI DẠY"""
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(10)
        p_sec.paragraph_format.space_after = Pt(3)
        p_sec.paragraph_format.keep_with_next = True
        r_sec = p_sec.add_run("IV. BẢNG TỔNG HỢP NĂNG LỰC SỐ (NLS) TÍCH HỢP TRONG BÀI DẠY")
        r_sec.bold = True
        r_sec.font.size = Pt(12.5)

        # Bảng 3 cột: TT (40pt), Hoạt động (320pt), Mã NLS (140pt)
        rows_count = 1 + max(1, len(nls_table))
        tbl = doc.add_table(rows=rows_count, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_borders(tbl, color="CBD5E1", sz="4", val="single")

        # Header hàng
        h_row = tbl.rows[0]
        make_row_cant_split(h_row)
        c0, c1, c2 = h_row.cells[0], h_row.cells[1], h_row.cells[2]
        set_cell_width_dxa(c0, 40)
        set_cell_width_dxa(c1, 320)
        set_cell_width_dxa(c2, 140)
        for c in (c0, c1, c2):
            set_cell_shading(c, "F1F5F9")
            set_cell_margins(c, top=100, bottom=100, left=120, right=120)

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run("TT").bold = True

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.add_run("Nội dung hoạt động số của học sinh").bold = True

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_code_head = p2.add_run("Mã NLS đạt được")
        r_code_head.bold = True
        r_code_head.font.color.rgb = COLOR_RED

        if not nls_table:
            # Fallback 1 dòng mẫu nếu trống
            row = tbl.rows[1]
            make_row_cant_split(row)
            set_cell_width_dxa(row.cells[0], 40)
            set_cell_width_dxa(row.cells[1], 320)
            set_cell_width_dxa(row.cells[2], 140)
            row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[0].paragraphs[0].add_run("1")
            row.cells[1].paragraphs[0].add_run("HS sử dụng thiết bị số khai thác tài liệu học tập trực tuyến.")
            p_code = row.cells[2].paragraphs[0]
            p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_c = p_code.add_run("1.2.TC2a")
            r_c.bold = True
            r_c.font.color.rgb = COLOR_RED
            return

        for idx, item in enumerate(nls_table, start=1):
            row = tbl.rows[idx]
            make_row_cant_split(row)
            cell_stt = row.cells[0]
            cell_act = row.cells[1]
            cell_code = row.cells[2]

            set_cell_width_dxa(cell_stt, 40)
            set_cell_width_dxa(cell_act, 320)
            set_cell_width_dxa(cell_code, 140)
            set_cell_margins(cell_stt, top=80, bottom=80, left=80, right=80)
            set_cell_margins(cell_act, top=80, bottom=80, left=120, right=120)
            set_cell_margins(cell_code, top=80, bottom=80, left=100, right=100)

            # Cột TT
            p_stt = cell_stt.paragraphs[0]
            p_stt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_stt.paragraph_format.space_after = Pt(0)
            p_stt.add_run(str(item.get("stt", idx)))

            # Cột Hoạt động số
            p_act = cell_act.paragraphs[0]
            p_act.paragraph_format.space_after = Pt(1.5)
            p_act.paragraph_format.line_spacing = 1.15
            add_formatted_runs_with_red_nls(p_act, str(item.get("activity", "")), 11.5)

            # Cột Mã NLS - BẮT BUỘC CHỮ ĐỎ ĐẬM CĂN GIỮA
            p_code = cell_code.paragraphs[0]
            p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_code.paragraph_format.space_after = Pt(1.5)
            codes = str(item.get("code", "")).replace("\\n", "\n").split("\n")
            for i_c, c_str in enumerate(codes):
                if i_c > 0:
                    p_code.add_run("\n")
                r_c = p_code.add_run(c_str.strip())
                r_c.bold = True
                r_c.font.color.rgb = COLOR_RED
                r_c.font.size = Pt(11.5)

    @classmethod
    def _render_section_v(cls, doc, cm_title="Duyệt của Chuyên môn", cm_signer="Lưu Văn Chí", to_title="Duyệt của tổ chuyên môn", to_signer="Hoàng Thị Kim Tuyết"):
        """Render Phần V. KHỐI PHÊ DUYỆT (Chống tách trang cantSplit, keep_with_next, space_after=36pt)"""
        p_sig = doc.add_paragraph()
        p_sig.paragraph_format.space_before = Pt(10)
        p_sig.paragraph_format.space_after = Pt(0)
        p_sig.paragraph_format.keep_with_next = True

        tbl_s = doc.add_table(rows=1, cols=2)
        tbl_s.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_s.autofit = False
        make_row_cant_split(tbl_s.rows[0])

        c0 = tbl_s.cell(0, 0)
        c1 = tbl_s.cell(0, 1)
        set_cell_width_dxa(c0, 245)
        set_cell_width_dxa(c1, 245)
        set_cell_margins(c0, top=100, bottom=100, left=100, right=100)
        set_cell_margins(c1, top=100, bottom=100, left=100, right=100)

        # Cột trái: Duyệt của Chuyên môn - Lưu Văn Chí
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_after = Pt(2)
        p0.paragraph_format.keep_with_next = True
        r0 = p0.add_run(cm_title)
        r0.bold = True
        r0.font.size = Pt(12.5)

        p1 = c0.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_after = Pt(36)  # Khoảng ký tay chuẩn đẹp
        p1.paragraph_format.keep_with_next = True
        r1 = p1.add_run("(Ký và ghi rõ họ tên)")
        r1.italic = True
        r1.font.size = Pt(11)

        p2 = c0.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(cm_signer)
        r2.bold = True
        r2.font.size = Pt(12.5)

        # Cột phải: Duyệt của tổ chuyên môn - Hoàng Thị Kim Tuyết
        p0_1 = c1.paragraphs[0]
        p0_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0_1.paragraph_format.space_after = Pt(2)
        p0_1.paragraph_format.keep_with_next = True
        r0_1 = p0_1.add_run(to_title)
        r0_1.bold = True
        r0_1.font.size = Pt(12.5)

        p1_1 = c1.add_paragraph()
        p1_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1_1.paragraph_format.space_after = Pt(36)
        p1_1.paragraph_format.keep_with_next = True
        r1_1 = p1_1.add_run("(Ký và ghi rõ họ tên)")
        r1_1.italic = True
        r1_1.font.size = Pt(11)

        p2_1 = c1.add_paragraph()
        p2_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2_1.paragraph_format.space_after = Pt(0)
        r2_1 = p2_1.add_run(to_signer)
        r2_1.bold = True
        r2_1.font.size = Pt(12.5)
