#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
scripts/generate_standard_exam.py
Công cụ CLI sinh hồ sơ đề thi Tin học GDPT 2018 theo ma trận và bản đặc tả chuẩn Bộ GD&ĐT.
"""

import os
import sys
import argparse

# Thiết lập thư mục gốc
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter, CURRICULUM_DATA
from core.ai_exam_generator import AIExamGenerator


def main():
    parser = argparse.ArgumentParser(
        description="Tool sinh đề thi tự động theo ma trận và bản đặc tả chuẩn GDPT 2018 (Neo4j Edu Graph)"
    )
    parser.add_argument("--grade", type=str, default="Lớp 9", choices=list(CURRICULUM_DATA.keys()),
                        help="Khối lớp (Lớp 6, Lớp 7, Lớp 8, Lớp 9)")
    parser.add_argument("--exam", type=str, default="KIỂM TRA GIỮA HỌC KỲ I",
                        help="Tên kỳ kiểm tra")
    parser.add_argument("--school", type=str, default="TRƯỜNG TH&THCS ĐĂK TRÔI",
                        help="Tên trường học")
    parser.add_argument("--gov", type=str, default="UBND XÃ KON CHIÊNG",
                        help="Đơn vị chủ quản")
    parser.add_argument("--year", type=str, default="2025 - 2026",
                        help="Năm học")
    parser.add_argument("--teacher", type=str, default="Lê Thị Phương Thảo",
                        help="Họ tên giáo viên ra đề")
    parser.add_argument("--head", type=str, default="Trần Thị Tuyết Ngân",
                        help="Họ tên tổ trưởng chuyên môn")
    parser.add_argument("--nb", type=float, default=30.0,
                        help="Tỷ lệ % câu hỏi mức Nhận biết")
    parser.add_argument("--th", type=float, default=40.0,
                        help="Tỷ lệ % câu hỏi mức Thông hiểu")
    parser.add_argument("--vd", type=float, default=30.0,
                        help="Tỷ lệ % câu hỏi mức Vận dụng")
    parser.add_argument("--vdc", type=float, default=0.0,
                        help="Tỷ lệ % câu hỏi mức Vận dụng cao")
    parser.add_argument("--output", type=str, default="",
                        help="Đường dẫn file Word đầu ra (.docx)")
    parser.add_argument("--local", action="store_true",
                        help="Bắt buộc dùng Local Offline Engine thay vì Beeknoee AI")

    args = parser.parse_args()

    ai_gen = AIExamGenerator()
    use_ai = not args.local and ai_gen.is_available()

    print("=" * 70)
    print("🎓 KHỞI CHẠY BỘ SINH ĐỀ THI TIN HỌC THEO MA TRẬN & BẢN ĐẶC TẢ CHUẨN")
    print("=" * 70)
    print(f"• Động cơ khảo thí:     {'🤖 Beeknoee AI Engine (' + ai_gen.model + ')' if use_ai else '📁 Local Offline Engine'}")
    print(f"• Khối lớp:             {args.grade}")
    print(f"• Kỳ kiểm tra:          {args.exam}")
    print(f"• Trường học:           {args.school}")
    print(f"• Năm học:              {args.year}")
    print(f"• Ma trận nhận thức:    {args.nb}% Biết | {args.th}% Hiểu | {args.vd}% Vận dụng | {args.vdc}% VDC")
    print(f"• Giáo viên ra đề:      {args.teacher}")
    print(f"• Tổ trưởng thẩm định:  {args.head}")

    dossier = None
    if use_ai:
        print("\n🤖 Đang kết nối Beeknoee AI để phân loại ma trận nhận thức & biên soạn đề thi...")
        try:
            dossier = ai_gen.generate_pedagogical_exam(
                grade_level=args.grade,
                exam_title=args.exam,
                school_name=args.school,
                governing_body=args.gov,
                school_year=args.year,
                teacher_name=args.teacher,
                head_name=args.head,
                pct_nb=args.nb,
                pct_th=args.th,
                pct_vd=args.vd,
                pct_vdc=args.vdc
            )
            print("✔️ Beeknoee AI đã hoàn thành biên soạn và phân loại các hạng mục câu hỏi!")
        except Exception as e:
            print(f"⚠️ Beeknoee AI trả về lỗi: {e}. Đang chuyển sang Local Engine dự phòng...")
            use_ai = False

    if not use_ai or dossier is None:
        print("\n📁 Đang sử dụng Local Offline Engine biên soạn hồ sơ đề thi...")
        engine = ExamMatrixEngine(
            grade_level=args.grade,
            exam_title=args.exam,
            school_name=args.school,
            governing_body=args.gov,
            school_year=args.year,
            teacher_name=args.teacher,
            head_name=args.head,
            pct_nb=args.nb,
            pct_th=args.th,
            pct_vd=args.vd,
            pct_vdc=args.vdc
        )
        dossier = engine.generate_exam_dossier()

    if not args.output:
        safe_grade = args.grade.lower().replace(" ", "_")
        prefix = "Beeknoee_AI" if use_ai else "Local"
        args.output = os.path.join(PROJECT_ROOT, "output", f"Ho_So_De_Thi_{prefix}_{safe_grade}.docx")

    print(f"💾 Đang xuất tài liệu Word (.docx) chuẩn mẫu sở GD&ĐT ra: {args.output}...")
    final_path = ExamDocxExporter.export(dossier, args.output)

    print("\n✅ HOÀN TẤT XUẤT ĐỀ THI!")
    print(f"📁 Tệp đề thi đã lưu: {final_path}")
    print(f"📊 Dung lượng tệp:   {os.path.getsize(final_path):,} bytes")
    print("=" * 70)


if __name__ == "__main__":
    main()
