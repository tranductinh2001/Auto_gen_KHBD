# ĐỒ ÁN CƠ SỞ DỮ LIỆU NÂNG CAO
## ĐỀ TÀI: XÂY DỰNG HỆ THỐNG QUẢN LÝ TRI THỨC, PHÂN TÍCH CHƯƠNG TRÌNH VÀ HỖ TRỢ RA ĐỀ THI TỰ ĐỘNG MÔN TIN HỌC PHỔ THÔNG GDPT 2018 SỬ DỤNG NEO4J

---

## 1. GIỚI THIỆU TỔNG QUAN
Hệ thống là giải pháp hoàn chỉnh từ khâu **thu thập và trích xuất tri thức đa phương thức (Multimodal Knowledge Extraction)** từ Sách giáo khoa (PDF / Hình ảnh) bằng AI Vision API cho đến khâu **tổ chức, lưu trữ đồ thị tri thức trên Neo4j Graph Database** và **hỗ trợ ra đề thi tự động theo ma trận chuẩn sư phạm GDPT 2018**.

### Các tính năng cốt lõi:
1. **Cấu hình phạm vi trang (`start_page` & `end_page`):** Kiểm soát chính xác khoảng trang cần cào (ví dụ: test từ trang 30 đến trang 40) giúp tiết kiệm tài nguyên API và kiểm soát chất lượng dữ liệu.
2. **Trích xuất Đa phương thức (Multimodal Extraction):**
   - **Văn bản:** Tự động phân đoạn theo cấp bậc sư phạm (Khối lớp $\rightarrow$ Chủ đề lớn $\rightarrow$ Bài học cụ thể $\rightarrow$ Yêu cầu cần đạt $\rightarrow$ Kiến thức cốt lõi $\rightarrow$ Kỹ năng thực hành).
   - **Hình ảnh / Sơ đồ / Giao diện (Visual Assets):** AI Vision nhận diện mọi hình minh họa, sơ đồ thuật toán, khối lệnh Scratch, code Python, giao diện Excel... và sinh ra **Detailed Image Description** (phân tích thành phần cấu trúc và ý nghĩa sư phạm).
   - **Đánh giá & Ngân hàng câu hỏi:** Trích xuất câu hỏi luyện tập/vận dụng, phân loại mức độ nhận thức ("Nhận biết" | "Thông hiểu" | "Vận dụng"), các phương án trắc nghiệm (A, B, C, D) và đáp án đúng.
3. **Cơ chế Bảo vệ Kép (Anti-Overwrite & Master Accumulation):**
   - Lưu trữ phiên cào độc lập trong `output/sessions/` (không bao giờ mất dữ liệu cũ).
   - Tự động gộp tích lũy và khử trùng lặp vào thư mục Master Neo4j: `output/neo4j_import/`.
4. **Hỗ trợ Ra đề thi tự động:** Cung cấp sẵn các mẫu truy vấn Cypher để sinh đề thi theo ma trận chuẩn (40% Nhận biết, 30% Thông hiểu, 30% Vận dụng) và truy vết yêu cầu cần đạt.

---

## 2. KIẾN TRÚC ĐỒ THỊ TRI THỨC (NEO4J PROPERTY GRAPH SCHEMA)

### 2.1. Danh mục các Node Labels & Thuộc tính:
- `(:Grade)`: `{grade_level: "Lớp 9"}`
- `(:Chapter)`: `{id: "Lớp 9_Chủ đề 1: Máy tính và cộng đồng", title: "Chủ đề 1: Máy tính và cộng đồng"}`
- `(:Lesson)`: `{id: "Lớp 9_Bài 1", lesson_id: "Bài 1", title: "Thế giới kĩ thuật số", page: 7, source_file: "sgk.pdf"}`
- `(:Requirement)`: `{id: "Lớp 9_Bài 1_YC01", code: "YC01", content: "Nêu được vai trò của thiết bị số..."}`
- `(:Knowledge)`: `{id: "Lớp 9_Bài 1_1. Thiết bị số", name: "1. Thiết bị số", content: "Nội dung lý thuyết..."}`
- `(:VisualAsset)`: `{id: "Lớp 9_Bài 1_Hình 1.1", code: "Hình 1.1", caption: "Ti vi kĩ thuật số", detailed_description: "Ảnh chụp giao diện...", page: 7}`
- `(:Skill)`: `{id: "Lớp 9_Bài 1_Kỹ năng chuột", name: "Kỹ năng chuột", description: "Thao tác điều khiển..."}`
- `(:Question)`: `{id: "Lớp 9_Bài 1_CH01", question_id: "CH01", content: "Nội dung câu hỏi...", difficulty: "Nhận biết"}`
- `(:AnswerOption)`: `{option_key: "A", content: "Phương án A", is_correct: true}`

### 2.2. Các Quan hệ Đồ thị (Graph Relationships):
```
(:Grade)-[:HAS_CHAPTER]->(:Chapter)
(:Chapter)-[:HAS_LESSON]->(:Lesson)
(:Lesson)-[:HAS_REQUIREMENT]->(:Requirement)
(:Lesson)-[:TEACHES_KNOWLEDGE]->(:Knowledge)
(:Lesson)-[:HAS_VISUAL_ASSET]->(:VisualAsset)
(:Lesson)-[:DEVELOPS_SKILL]->(:Skill)
(:Lesson)-[:EVALUATED_BY]->(:Question)
(:Question)-[:HAS_OPTION]->(:AnswerOption)
```

---

## 3. CẤU TRÚC THƯ MỤC DỰ ÁN

```
edu_graph_tinhoc_neo4j/
│
├── core/
│   ├── ai_vision_client.py        # Client gọi AI Vision API với timeout & retry
│   ├── pdf_processor.py           # Quét & cắt trang PDF theo start_page - end_page (200 DPI)
│   └── multimodal_extractor.py    # Prompt chuyên sâu trích xuất sư phạm đa phương thức
│
├── exporters/
│   └── graph_exporter.py          # Xuất Excel đa Sheet, CSV Neo4j và tích lũy Master
├── web_dashboard.py             # Hệ thống Web Dashboard hiện đại bằng Streamlit (4 Phân hệ)
├── run_web_dashboard.bat         # Click đúp để mở ngay Web Dashboard trên trình duyệt
├── run_desktop_app.bat           # Click đúp để mở giao diện Desktop Tkinter
│
├── core/
│   ├── ai_vision_client.py        # Client gọi AI Vision API với timeout & retry
│   ├── pdf_processor.py           # Quét & cắt trang PDF theo start_page - end_page (200 DPI)
│   ├── multimodal_extractor.py    # Prompt chuyên sâu trích xuất sư phạm đa phương thức
│   └── graph_service.py           # Nghiệp vụ Đồ thị: Phân tích chương trình & Sinh đề thi tự động
│
├── exporters/
│   └── graph_exporter.py          # Xuất Excel đa Sheet, CSV Neo4j và tích lũy Master
│
├── gui/
│   └── app_window.py              # Giao diện Desktop Tkinter đa phân hệ (4 Tabs)
│
├── scripts/
│   ├── neo4j_import_script.cypher         # Toàn bộ lệnh LOAD CSV nạp vào Neo4j
│   └── auto_exam_generation_queries.cypher# Truy vấn Cypher hỗ trợ ra đề thi tự động
│
├── output/                        # Thư mục xuất dữ liệu tự động
│   ├── sessions/                  # Lưu trữ từng phiên cào riêng biệt (không bao giờ mất)
│   └── neo4j_import/              # Kho Master nạp Neo4j (Tích lũy & Loại trùng)
│
├── main.py                        # Điểm khởi chạy linh hoạt (Mặc định mở Web Dashboard)
├── requirements.txt               # Danh mục thư viện Python
├── .env                           # Cấu hình API Key và model
└── README.md                      # Báo cáo kỹ thuật chi tiết
```

---

## 4. HƯỚNG DẪN VẬN HÀNH HỆ THỐNG

### Cách 1: Khởi chạy Web Dashboard Hiện Đại (Khuyến nghị cho Đồ án)
Mở terminal tại thư mục `edu_graph_tinhoc_neo4j`:
```bash
python main.py
```
*(Hoặc chạy lệnh `streamlit run web_dashboard.py` hoặc click đúp file `run_web_dashboard.bat`)*  
Hệ thống sẽ tự động bật trình duyệt Web tại địa chỉ `http://localhost:8501` với đầy đủ 4 phân hệ chức năng:
1. **📥 1. Trích xuất & Số hóa SGK (AI Vision):** Cấu hình `start_page` - `end_page`, chọn PDF/ảnh, phân tích đa phương thức và tải Excel/CSV.
2. **🌐 2. Quản lý Tri thức & Neo4j:** Duyệt cây tri thức sư phạm, xem schema Property Graph, kết nối Neo4j DBMS trực tiếp hoặc nạp file Cypher.
3. **📊 3. Phân tích Chương trình GDPT 2018:** Biểu đồ ma trận Bloom (Nhận biết/Thông hiểu/Vận dụng), phân loại học liệu trực quan, mật độ yêu cầu cần đạt.
4. **📝 4. Hỗ trợ Ra Đề Thi Tự Động:** Chọn khối lớp, số lượng câu, ma trận tỷ lệ %, ưu tiên Visual Assets $\rightarrow$ Sinh đề thi hoàn chỉnh kèm đáp án & xuất file text.

### Cách 2: Khởi chạy Giao diện Desktop (Tkinter GUI)
```bash
python main.py --desktop
```
*(Hoặc click đúp file `run_desktop_app.bat`)*

### Cách 3: Chạy Dòng lệnh (CLI Headless)
```bash
python main.py --pdf "../CamScanner 01-10-2026 13.30_1.pdf" --start_page 30 --end_page 40 --grade "Lớp 9"
```

---

## 5. HƯỚNG DẪN NẠP VÀO NEO4J & RA ĐỀ THI TỰ ĐỘNG

1. Mở Neo4j Desktop $\rightarrow$ Chọn Database $\rightarrow$ Mở thư mục `import/`.
2. Copy toàn bộ các file `.csv` trong thư mục `output/neo4j_import/` vào thư mục `import/` của Neo4j.
3. Mở **Neo4j Browser**, dán và chạy nội dung file `scripts/neo4j_import_script.cypher`.
4. Để thực hiện tính năng **Phân tích chương trình & Hỗ trợ ra đề thi tự động**, mở file `scripts/auto_exam_generation_queries.cypher` hoặc dùng trực tiếp Tab 4 trên giao diện:
   - **Kịch bản 1:** Tự động sinh đề thi trắc nghiệm theo ma trận chuẩn Bộ GD&ĐT (40% Nhận biết, 30% Thông hiểu, 30% Vận dụng).
   - **Kịch bản 2:** Ra đề thi thực hành gắn liền với Visual Assets (hình ảnh giao diện phần mềm / sơ đồ thuật toán).
   - **Kịch bản 3:** Kiểm toán ma trận đánh giá (truy vết Yêu cầu cần đạt nào đã có câu hỏi kiểm tra).
   - **Kịch bản 4:** Phân tích độ bao phủ chương trình học (Curriculum Coverage Analysis).
