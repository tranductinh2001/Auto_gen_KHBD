# -*- coding: utf-8 -*-
"""
scripts/batch_generate_khbd.py
Công cụ CLI sinh hàng loạt Kế hoạch bài dạy (KHBD) chuẩn Công văn 5512/BGDĐT-GDTrH
và Tích hợp Khung Năng lực số (NLS) theo Thông tư 02/2025/TT-BGDĐT sử dụng Beekone AI API.
"""

import os
import sys
import argparse
import zipfile
import time
from typing import List

# Đảm bảo đường dẫn import
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.khbd_ai_generator import KHBDAIGenerator
from core.khbd_docx_exporter import KHBDDocxExporter


def batch_generate(
    excel_path: str = "_Lop_6A.xlsx",
    out_dir: str = "output/khbd_lop_6a",
    from_week: int = 1,
    to_week: int = 5,
    teacher_name: str = "Trần Đức Tình",
    department: str = "Tổ Tự Nhiên",
    signer_cm: str = "Lưu Văn Chí",
    signer_to: str = "Hoàng Thị Kim Tuyết",
    create_zip: bool = True
) -> List[str]:
    print("=" * 70)
    print("🚀 BỘ SINH KHBD TỰ ĐỘNG CHUẨN 5512 & TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025)")
    print(f"📁 Nguồn phân phối chương trình: {excel_path}")
    print(f"👨‍🏫 Giáo viên: {teacher_name} | Tổ: {department}")
    print(f"✍️ Duyệt chuyên môn: {signer_cm} | Duyệt tổ: {signer_to}")
    print(f"📅 Khoảng tuần: Tuần {from_week} đến Tuần {to_week}")
    print("=" * 70)

    ai_gen = KHBDAIGenerator()
    if not ai_gen.is_available():
        print("⚠️ CẢNH BÁO: Chưa cấu hình BEE_API_KEY trong .env. Sẽ dùng bộ biên soạn chuẩn hóa cục bộ.")

    lessons = ai_gen.load_ppct_excel(excel_path)
    print(f"📖 Tổng số tiết học tìm thấy trong PPCT: {len(lessons)}")

    # Lọc danh sách theo tuần
    selected_lessons = []
    for l in lessons:
        try:
            w_num = int(l["week"])
            if from_week <= w_num <= to_week:
                selected_lessons.append(l)
        except ValueError:
            selected_lessons.append(l)

    print(f"🎯 Số tiết cần sinh theo yêu cầu: {len(selected_lessons)}")
    os.makedirs(out_dir, exist_ok=True)

    generated_files = []
    start_time = time.time()

    for idx, l_info in enumerate(selected_lessons, start=1):
        stt = l_info["stt"]
        week = l_info["week"]
        period = l_info["period"]
        title = l_info["lesson_title"]
        out_fn = l_info["output_filename"]
        out_path = os.path.join(out_dir, out_fn)

        print(f"\n[{idx}/{len(selected_lessons)}] Đang xử lý: Tuần {week} - Tiết {period}: {title}...")

        try:
            # Gọi AI sinh JSON
            t0 = time.time()
            data = ai_gen.generate_khbd_json(
                lesson_info=l_info,
                teacher_name=teacher_name,
                department=department,
                signer_cm=signer_cm,
                signer_to=signer_to,
                subject="Tin học",
                grade="6"
            )
            dur_ai = time.time() - t0

            # Xuất Word (.docx)
            KHBDDocxExporter.export_khbd_to_docx(data, out_path)
            file_sz = os.path.getsize(out_path) / 1024
            print(f"   ✅ Hoàn tất trong {dur_ai:.1f}s | File: {out_fn} ({file_sz:.1f} KB)")
            generated_files.append(out_path)

        except Exception as e:
            print(f"   ❌ Lỗi khi sinh bài '{title}': {e}")

    total_time = time.time() - start_time
    print("\n" + "=" * 70)
    print(f"🎉 ĐÃ HOÀN TẤT SINH {len(generated_files)}/{len(selected_lessons)} GIÁO ÁN TRONG {total_time:.1f} GIÂY!")

    # Đóng gói ZIP nếu có yêu cầu
    if create_zip and generated_files:
        zip_path = os.path.join(os.path.dirname(out_dir), f"KHBD_Lop_6A_Tuan_{from_week:02d}_den_{to_week:02d}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in generated_files:
                zf.write(f, os.path.basename(f))
        print(f"📦 Đã đóng gói toàn bộ vào file ZIP: {zip_path} ({os.path.getsize(zip_path)/1024:.1f} KB)")

    print("=" * 70)
    return generated_files


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sinh hàng loạt KHBD chuẩn 5512 tích hợp NLS từ Excel PPCT")
    parser.add_argument("--excel", default="_Lop_6A.xlsx", help="Đường dẫn file Excel PPCT (mặc định _Lop_6A.xlsx)")
    parser.add_argument("--out-dir", default="output/khbd_lop_6a", help="Thư mục xuất file Word")
    parser.add_argument("--from-week", type=int, default=1, help="Tuần bắt đầu (mặc định 1)")
    parser.add_argument("--to-week", type=int, default=5, help="Tuần kết thúc (mặc định 5)")
    parser.add_argument("--week", type=int, default=None, help="Chỉ sinh 1 tuần cụ thể")
    parser.add_argument("--all", action="store_true", help="Sinh toàn bộ 35 tuần")
    parser.add_argument("--teacher", default="Trần Đức Tình", help="Tên giáo viên soạn")
    parser.add_argument("--dept", default="Tổ Tự Nhiên", help="Tổ chuyên môn")
    parser.add_argument("--signer-cm", default="Lưu Văn Chí", help="Người duyệt chuyên môn")
    parser.add_argument("--signer-to", default="Hoàng Thị Kim Tuyết", help="Người duyệt tổ chuyên môn")
    parser.add_argument("--no-zip", action="store_true", help="Không tạo file zip")

    args = parser.parse_args()

    f_week = args.from_week
    t_week = args.to_week

    if args.week is not None:
        f_week = args.week
        t_week = args.week
    elif args.all:
        f_week = 1
        t_week = 35

    batch_generate(
        excel_path=args.excel,
        out_dir=args.out_dir,
        from_week=f_week,
        to_week=t_week,
        teacher_name=args.teacher,
        department=args.dept,
        signer_cm=args.signer_cm,
        signer_to=args.signer_to,
        create_zip=not args.no_zip
    )
