import csv
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

MASTER_DIR = r"c:\Users\Admin\Desktop\Thạc sĩ\Chương trình\Cơ sở dữ liệu nâng cao\edu_graph_tinhoc_neo4j\output\neo4j_import"

lessons_data = [
    ["Lớp 10", "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính", "Bài 16", "Ngôn ngữ lập trình bậc cao và môi trường phát triển Python", "88", "SGK Tin học 10 KNTT (Định hướng CS).pdf"],
    ["Lớp 10", "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính", "Bài 17", "Biến, kiểu dữ liệu cơ sở và lệnh gán trong Python", "94", "SGK Tin học 10 KNTT (Định hướng CS).pdf"],
    ["Lớp 10", "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính", "Bài 22", "Kiểu dữ liệu danh sách (List) và các thao tác cơ bản", "115", "SGK Tin học 10 KNTT (Định hướng CS).pdf"],
    ["Lớp 10", "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính", "Bài 28", "Thuật toán tìm kiếm tuần tự và tìm kiếm nhị phân", "138", "SGK Tin học 10 KNTT (Định hướng CS).pdf"],
    ["Lớp 10", "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính", "Bài 29", "Thuật toán sắp xếp nổi bọt và sắp xếp chèn", "145", "SGK Tin học 10 KNTT (Định hướng CS).pdf"],
    ["Lớp 10", "Chuyên đề: Ứng dụng Trí tuệ Nhân tạo và Khoa học Dữ liệu", "Bài 30", "Nhập môn Trí tuệ Nhân tạo và Học máy", "160", "Chuyên đề học tập Tin học 10 KNTT.pdf"]
]

req_data = [
    ["Lớp 10", "Bài 16", "YCCD_10_16_01", "Biết và phân biệt được ngôn ngữ máy, hợp ngữ và ngôn ngữ lập trình bậc cao."],
    ["Lớp 10", "Bài 16", "YCCD_10_16_02", "Khởi động và làm quen với môi trường lập trình Python."],
    ["Lớp 10", "Bài 17", "YCCD_10_17_01", "Hiểu khái niệm biến, quy tắc đặt tên biến và lệnh gán giá trị trong Python."],
    ["Lớp 10", "Bài 17", "YCCD_10_17_02", "Phân biệt được các kiểu dữ liệu cơ sở: int, float, str, bool."],
    ["Lớp 10", "Bài 22", "YCCD_10_22_01", "Mô tả được kiểu dữ liệu danh sách (List), cách khởi tạo và truy cập phần tử."],
    ["Lớp 10", "Bài 22", "YCCD_10_22_02", "Sử dụng thành thạo các phương thức cơ bản và duyệt list."],
    ["Lớp 10", "Bài 28", "YCCD_10_28_01", "Hiểu và cài đặt được thuật toán tìm kiếm tuần tự (Linear Search)."],
    ["Lớp 10", "Bài 28", "YCCD_10_28_02", "Hiểu nguyên lý tìm kiếm nhị phân (Binary Search) và điều kiện mảng đã sắp xếp."],
    ["Lớp 10", "Bài 29", "YCCD_10_29_01", "Trình bày được ý tưởng thuật toán sắp xếp nổi bọt (Bubble Sort)."],
    ["Lớp 10", "Bài 29", "YCCD_10_29_02", "Mô phỏng chạy tay và cài đặt được thuật toán trên Python."],
    ["Lớp 10", "Bài 30", "YCCD_10_30_01", "Hiểu định nghĩa cơ bản về AI, phân biệt Narrow AI và AGI."],
    ["Lớp 10", "Bài 30", "YCCD_10_30_02", "Nêu được các ví dụ ứng dụng AI trong thị giác máy tính và xử lý ngôn ngữ tự nhiên."]
]

def append_rows_if_not_exists(csv_path, rows, check_col_idx):
    if not os.path.exists(csv_path):
        return
    existing_keys = set()
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for r in reader:
            if len(r) > check_col_idx:
                existing_keys.add(r[check_col_idx].strip())
    
    with open(csv_path, mode="a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        for r in rows:
            if r[check_col_idx].strip() not in existing_keys:
                writer.writerow(r)
                existing_keys.add(r[check_col_idx].strip())

append_rows_if_not_exists(os.path.join(MASTER_DIR, "neo4j_lessons.csv"), lessons_data, 2)
append_rows_if_not_exists(os.path.join(MASTER_DIR, "neo4j_requirements.csv"), req_data, 2)
print("Đã đồng bộ CSV Master hoàn tất!")
