import sys
import os

# Cấu hình UTF-8 cho console Windows tránh UnicodeEncodeError cp1252
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import argparse
import subprocess
from dotenv import load_dotenv

# Đảm bảo đường dẫn import tương đối từ thư mục project
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)


def run_cli_mode(args):
    """Chạy trực tiếp từ dòng lệnh CLI không cần giao diện đồ họa."""
    from core.ai_vision_client import AIVisionClient
    from core.pdf_processor import PDFPageProcessor
    from core.multimodal_extractor import MultimodalExtractor
    from exporters.graph_exporter import GraphDataExporter

    env_path = os.path.join(PROJECT_DIR, ".env")
    load_dotenv(env_path)
    api_key = os.getenv("BEE_API_KEY", "")
    base_url = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")

    if not api_key:
        print("[LỖI] Chưa tìm thấy BEE_API_KEY trong file .env")
        sys.exit(1)

    print("==========================================================================")
    print("HỆ THỐNG TRÍCH XUẤT TRI THỨC & HỖ TRỢ RA ĐỀ THI TIN HỌC 2018 (CLI MODE)")
    print(f"Khối lớp: {args.grade} | Dải trang: {args.start_page} -> {args.end_page}")
    print("==========================================================================")

    client = AIVisionClient(api_key=api_key, base_url=base_url)
    extractor = MultimodalExtractor(client=client)
    output_dir = os.path.join(PROJECT_DIR, "output")
    exporter = GraphDataExporter(output_dir=output_dir)

    temp_dir = os.path.join(PROJECT_DIR, "temp_cli_images")
    queue = []

    if args.pdf:
        print(f"[1/3] Đang quét và phân đoạn trang PDF: {args.pdf}")
        tasks = PDFPageProcessor.convert_pdf_to_images(
            pdf_path=args.pdf,
            temp_dir=temp_dir,
            start_page=args.start_page,
            end_page=args.end_page
        )
        for img_p, p_num in tasks:
            queue.append({"path": img_p, "source": args.pdf, "page": p_num})
    elif args.image:
        queue.append({"path": args.image, "source": args.image, "page": 1})

    print(f"[2/3] Bắt đầu gọi AI Vision phân tích đa phương thức cho {len(queue)} trang...")
    all_bundles = []
    all_logs = []

    for idx, item in enumerate(queue, 1):
        print(f"  → Đang phân tích trang {item['page']} ({idx}/{len(queue)})...")
        flat_rows, norm_data, log_info = extractor.extract_page(
            image_path=item["path"],
            source_file=item["source"],
            page_num=item["page"],
            grade_hint=args.grade
        )
        all_bundles.append({"flat_rows": flat_rows, "normalized": norm_data})
        all_logs.append(log_info)

    print("[3/3] Đang xuất dữ liệu sang Excel và CSV Neo4j...")
    book_tag = f"sgk_tinhoc_{args.grade.replace(' ', '_').lower()}"
    session_excel, session_dir, master_dir = exporter.export(
        extracted_bundles=all_bundles,
        logs=all_logs,
        base_name=book_tag
    )

    if os.path.exists(temp_dir):
        import shutil
        shutil.rmtree(temp_dir)

    print("\n✅ HOÀN TẤT THÀNH CÔNG!")
    print(f"📊 File Excel phiên cào này: {session_excel}")
    print(f"📁 Thư mục CSV riêng phiên này: {session_dir}")
    print(f"🌐 Thư mục Master Neo4j (Tích lũy & Loại trùng): {master_dir}")


def launch_web_dashboard():
    """Khởi chạy ứng dụng Web Dashboard hiện đại bằng Streamlit."""
    web_script = os.path.join(PROJECT_DIR, "web_dashboard.py")
    print("\n🌐 Đang khởi chạy Hệ thống Web Dashboard trên trình duyệt...")
    print("📍 URL: http://localhost:8501")
    print("👉 Hãy mở trình duyệt nếu không tự động bật.")
    cmd = [sys.executable, "-m", "streamlit", "run", web_script, "--server.headless=false"]
    try:
        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\nĐã đóng hệ thống Web Dashboard.")


def launch_desktop_gui():
    """Khởi chạy giao diện Desktop Tkinter."""
    import tkinter as tk
    from gui.app_window import EduGraphAppGUI
    root = tk.Tk()
    app = EduGraphAppGUI(root)
    root.mainloop()


def main():
    parser = argparse.ArgumentParser(
        description="Hệ thống Quản lý Tri thức, Phân tích Chương trình & Ra Đề thi Tự động Môn Tin học GDPT 2018 (Neo4j)"
    )
    parser.add_argument("--mode", choices=["web", "desktop", "cli"], default="web",
                        help="Chế độ giao diện: 'web' (Web Dashboard trên trình duyệt - Mặc định), 'desktop' (Cửa sổ Tkinter), hoặc 'cli'")
    parser.add_argument("--desktop", action="store_true", help="Khởi chạy giao diện Desktop Tkinter")
    parser.add_argument("--pdf", type=str, help="Đường dẫn tệp PDF SGK Tin học (chạy CLI)")
    parser.add_argument("--image", type=str, help="Đường dẫn tệp ảnh trang SGK (chạy CLI)")
    parser.add_argument("--start_page", type=int, default=None, help="Trang bắt đầu (1-indexed)")
    parser.add_argument("--end_page", type=int, default=None, help="Trang kết thúc (1-indexed)")
    parser.add_argument("--grade", type=str, default="Lớp 9", help="Khối lớp (Ví dụ: 'Lớp 9')")

    args = parser.parse_args()

    if args.pdf or args.image:
        run_cli_mode(args)
    elif args.desktop or args.mode == "desktop":
        launch_desktop_gui()
    else:
        # Mặc định khởi chạy Web Dashboard
        launch_web_dashboard()


if __name__ == "__main__":
    main()
