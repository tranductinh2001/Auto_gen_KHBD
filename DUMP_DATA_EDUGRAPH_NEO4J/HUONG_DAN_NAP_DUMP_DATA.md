# HƯỚNG DẪN KHÔI PHỤC VÀ SỬ DỤNG DUMP DATA CSDL ĐỒ THỊ NEO4J
**Đề tài:** Ứng dụng Cơ sở Dữ liệu Đồ thị Tri thức (Knowledge Graph DBMS) Quản lý và Khai phá Chương trình Giáo dục Phổ thông 2018 Môn Tin học Cấp THCS  
**Học viên thực hiện:** Trần Đức Tình – **MSHV:** 8291581013  
**Chuyên ngành:** Khoa học Máy tính – **Trường:** Đại học Quy Nhơn  
**Môn học:** Cơ sở dữ liệu nâng cao  

---

## 1. DANH MỤC CÁC THÀNH PHẦN TRONG GÓI DUMP DATA

Gói dữ liệu này cung cấp đầy đủ các phương thức để Quý Thầy/Cô hoặc Hội đồng chấm thi có thể dễ dàng nạp lại toàn bộ Cơ sở Dữ liệu Đồ thị Neo4j:

| STT | Tên tệp / Thư mục | Định dạng | Mô tả công dụng |
|:---:|:---|:---:|:---|
| **1** | `edugraph_full_dump.cypher` | `.cypher` | **(Khuyên dùng)** Bản dump Cypher độc lập toàn bộ DDL (Constraints) và DML (Nodes & Relationships). Chạy trực tiếp trên Neo4j Browser hoặc `cypher-shell` mà **không cần copy file CSV**. |
| **2** | `restore_database.py` | `.py` | Script Python tự động nạp 1-click qua giao thức Bolt `bolt://localhost:7687` vào Database `edugraph`. |
| **3** | `import_via_load_csv.cypher` | `.cypher` | Lệnh Cypher nạp qua `LOAD CSV` truyền thống của Neo4j. |
| **4** | `csv_master_data/` | Thư mục CSV | Chứa trọn bộ 7 tệp dữ liệu Master đã làm sạch và chuẩn hóa (Lessons, Requirements, Knowledge, Skills, Visual Assets, Questions, Answers). |
| **5** | `sgk_tinhoc_tong_hop_master.xlsx` | Excel | Dữ liệu tổng hợp toàn diện 4 khối lớp (Lớp 6, 7, 8, 9) phân tách theo từng Sheet nghiệp vụ. |

---

## 2. HƯỚNG DẪN KHÔI PHỤC (CHỌN 1 TRONG 2 CÁCH DƯỚI ĐÂY)

### CÁCH 1: Nạp trực tiếp qua Python 1-Click (Nhanh nhất & Đơn giản nhất)
*Yêu cầu:* Máy đã mở Neo4j Desktop (hoặc Neo4j Server) và bấm **Start**.

1. Mở Terminal / PowerShell tại thư mục này.
2. Chạy lệnh:
   ```bash
   python restore_database.py
   ```
*(Nếu mật khẩu Neo4j khác mặc định `12345678`, có thể truyền tham số: `python restore_database.py --password <mat_khau_cua_ban>`)*.  
Toàn bộ Database `edugraph` sẽ được tạo mới, thiết lập ràng buộc và nạp đầy đủ trong vòng 10 giây kèm bảng thống kê tổng kết.

---

### CÁCH 2: Nạp bằng file Cypher Dump (`edugraph_full_dump.cypher`)
*Thực hiện trên Neo4j Browser:*
1. Mở **Neo4j Browser** (thường tại `http://localhost:7474`).
2. Chọn kết nối vào database `:use edugraph` (hoặc database mặc định `:use neo4j`).
3. Mở file `edugraph_full_dump.cypher`, copy từng khối lệnh và dán vào thanh lệnh của Neo4j Browser rồi ấn **Run** (Ctrl + Enter).

*Hoặc thực hiện qua cypher-shell:*
```bash
cypher-shell -u neo4j -p 12345678 -d edugraph -f edugraph_full_dump.cypher
```

---

## 3. THỐNG KÊ KIỂM CHỨNG DỮ LIỆU SAU KHI NẠP

Sau khi nạp xong, Quý Thầy/Cô có thể chạy các câu truy vấn sau trong **Neo4j Browser** để kiểm tra tính toàn vẹn:

### 3.1. Đếm tổng số Node theo từng thực thể:
```cypher
MATCH (n) 
RETURN labels(n)[0] AS Entity, count(n) AS TotalNodes 
ORDER BY TotalNodes DESC;
```
*Kết quả chuẩn:*
- `AnswerOption` (Phương án lựa chọn): ~778 nodes
- `VisualAsset` (Hình ảnh/Sơ đồ SGK): ~391 nodes
- `Knowledge` (Tri thức cốt lõi): ~382 nodes
- `Skill` (Kỹ năng thực hành): ~341 nodes
- `Question` (Câu hỏi đánh giá): ~323 nodes
- `Requirement` (Yêu cầu cần đạt YCCĐ): ~237 nodes
- `Lesson` (Bài học THCS): ~125 nodes
- `Chapter` (Chủ đề học tập): ~56 nodes
- `Grade` (Khối lớp 6, 7, 8, 9): 4 nodes

### 3.2. Truy vấn cây tri thức của một bài học bất kỳ:
```cypher
MATCH (l:Lesson {id: 'Lớp 9_Bài 1'})-[r]->(target)
RETURN l.title, type(r) AS Relationship, labels(target)[0] AS TargetType, target;
```

---
*Trân trọng cảm ơn Quý Thầy/Cô đã đánh giá đồ án!*
