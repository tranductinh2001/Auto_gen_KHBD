import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import shutil
from dotenv import load_dotenv, set_key

from core.ai_vision_client import AIVisionClient
from core.pdf_processor import PDFPageProcessor
from core.multimodal_extractor import MultimodalExtractor
from exporters.graph_exporter import GraphDataExporter
from core.graph_service import GraphService


class EduGraphAppGUI:
    """
    Giao diện đồ họa Đa phân hệ cho đồ án:
    'Xây dựng hệ thống quản lý tri thức, phân tích chương trình và hỗ trợ ra đề thi tự động môn Tin học phổ thông GDPT 2018 sử dụng Neo4j'
    """

    def __init__(self, root):
        self.root = root
        self.root.title("HỆ THỐNG QUẢN LÝ TRI THỨC & RA ĐỀ THI TIN HỌC GDPT 2018 (NEO4J GRAPH)")
        self.root.geometry("900x750")
        self.root.minsize(850, 700)

        # Đảm bảo cửa sổ nổi lên trên khi khởi chạy
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(self.root.attributes, '-topmost', False)

        project_dir = os.path.dirname(os.path.dirname(__file__))
        env_path = os.path.join(project_dir, ".env")
        load_dotenv(env_path)
        self.api_key = os.getenv("BEE_API_KEY", "")
        self.base_url = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")
        self.env_path = env_path
        self.project_dir = project_dir
        self.master_dir = os.path.join(project_dir, "output", "neo4j_import")

        self.graph_svc = GraphService(master_dir=self.master_dir)
        self.selected_files = []
        self.is_processing = False

        self._build_ui()
        self._refresh_master_stats()

    def _build_ui(self):
        # Header banner
        header_frame = tk.Frame(self.root, bg="#1a365d", height=60)
        header_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            header_frame,
            text="HỆ THỐNG QUẢN LÝ TRI THỨC, PHÂN TÍCH CHƯƠNG TRÌNH & RA ĐỀ THI TIN HỌC 2018 (NEO4J)",
            font=("Segoe UI", 12, "bold"),
            fg="white",
            bg="#1a365d"
        )
        title_lbl.pack(pady=(10, 2))

        sub_lbl = tk.Label(
            header_frame,
            text="Đồ án Cơ sở dữ liệu Nâng cao | Kiến trúc Đồ thị Tri thức Multimodal Property Graph",
            font=("Segoe UI", 9),
            fg="#cbd5e1",
            bg="#1a365d"
        )
        sub_lbl.pack(pady=(0, 10))

        # Notebook tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_extract = ttk.Frame(self.notebook, padding=10)
        self.tab_graph = ttk.Frame(self.notebook, padding=10)
        self.tab_analytics = ttk.Frame(self.notebook, padding=10)
        self.tab_exam = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_extract, text=" 📥 1. Trích xuất Đa phương thức (AI Vision) ")
        self.notebook.add(self.tab_graph, text=" 🌐 2. Quản lý Tri thức & Neo4j ")
        self.notebook.add(self.tab_analytics, text=" 📊 3. Phân tích Chương trình 2018 ")
        self.notebook.add(self.tab_exam, text=" 📝 4. Hỗ trợ Ra Đề Thi Tự Động ")

        self._build_tab_extract()
        self._build_tab_graph()
        self._build_tab_analytics()
        self._build_tab_exam()

    # ---------------- TAB 1: CÀO & TRÍCH XUẤT ----------------
    def _build_tab_extract(self):
        f = self.tab_extract

        # Cấu hình API
        f_api = ttk.LabelFrame(f, text="Cấu hình AI Vision API (Beeknoee / Gemini)", padding=8)
        f_api.pack(fill=tk.X, pady=(0, 8))

        row_api = ttk.Frame(f_api)
        row_api.pack(fill=tk.X)
        ttk.Label(row_api, text="API Key:").pack(side=tk.LEFT, padx=(0, 5))
        self.api_key_var = tk.StringVar(value=self.api_key)
        self.entry_key = ttk.Entry(row_api, textvariable=self.api_key_var, show="*", width=45)
        self.entry_key.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row_api, text="Lưu Cấu Hình", command=self.save_config).pack(side=tk.LEFT, padx=5)

        # Cấu hình phạm vi trang & Khối lớp
        f_cfg = ttk.LabelFrame(f, text="Cấu hình Khối lớp & Phạm vi trang (start_page - end_page)", padding=8)
        f_cfg.pack(fill=tk.X, pady=(0, 8))

        row_c = ttk.Frame(f_cfg)
        row_c.pack(fill=tk.X)

        ttk.Label(row_c, text="Khối lớp:").pack(side=tk.LEFT, padx=(0, 5))
        self.grade_var = tk.StringVar(value="Lớp 9")
        cbo = ttk.Combobox(row_c, textvariable=self.grade_var, values=["Lớp 9", "Lớp 6", "Lớp 7", "Lớp 8", "Lớp 10", "Lớp 11", "Lớp 12"], state="readonly", width=10)
        cbo.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row_c, text="Trang bắt đầu (start_page):").pack(side=tk.LEFT, padx=(0, 5))
        self.start_page_var = tk.StringVar(value="")
        ttk.Entry(row_c, textvariable=self.start_page_var, width=6).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row_c, text="Trang kết thúc (end_page):").pack(side=tk.LEFT, padx=(0, 5))
        self.end_page_var = tk.StringVar(value="")
        ttk.Entry(row_c, textvariable=self.end_page_var, width=6).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(row_c, text="(Để trống = Quét toàn bộ sách)", font=("Segoe UI", 8, "italic"), foreground="#64748b").pack(side=tk.LEFT)

        # Chọn tài liệu
        f_btn = ttk.Frame(f)
        f_btn.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(f_btn, text="📁 Chọn Tệp PDF SGK", command=self.select_pdf).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(f_btn, text="🖼️ Chọn Ảnh Trang SGK", command=self.select_images).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_btn, text="❌ Xóa danh sách", command=self.clear_files).pack(side=tk.LEFT, padx=5)

        # Danh sách tệp
        self.listbox = tk.Listbox(f, height=5)
        self.listbox.pack(fill=tk.X, pady=(0, 8))

        # Nút Run & Tiến trình
        self.btn_run = tk.Button(
            f,
            text="🚀 BẮT ĐẦU TRÍCH XUẤT ĐA PHƯƠNG THỨC & NẠP VÀO MASTER GRAPH",
            bg="#0284c7",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            pady=8,
            command=self.start_processing
        )
        self.btn_run.pack(fill=tk.X, pady=(0, 5))

        self.progress_bar = ttk.Progressbar(f, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        self.lbl_status = ttk.Label(f, text="Trạng thái: Sẵn sàng.", font=("Segoe UI", 9))
        self.lbl_status.pack(anchor=tk.W)

    # ---------------- TAB 2: QUẢN LÝ TRI THỨC & NEO4J ----------------
    def _build_tab_graph(self):
        f = self.tab_graph

        f_stats = ttk.LabelFrame(f, text="Thống kê Đồ thị Tri thức Neo4j (Kho Master)", padding=10)
        f_stats.pack(fill=tk.X, pady=(0, 10))

        self.lbl_graph_stats = ttk.Label(
            f_stats,
            text="Đang tải dữ liệu...",
            font=("Segoe UI", 10),
            justify=tk.LEFT
        )
        self.lbl_graph_stats.pack(anchor=tk.W)

        f_actions = ttk.Frame(f)
        f_actions.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(f_actions, text="🔄 Làm mới số liệu", command=self._refresh_master_stats).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(f_actions, text="📂 Mở thư mục CSV Neo4j", command=self._open_master_folder).pack(side=tk.LEFT, padx=10)
        ttk.Button(f_actions, text="📜 Xem Cypher Script", command=self._open_cypher_script).pack(side=tk.LEFT, padx=10)

        # Trình xem cây tri thức
        f_tree = ttk.LabelFrame(f, text="Danh mục Bài học & Cây Tri thức", padding=10)
        f_tree.pack(fill=tk.BOTH, expand=True)

        self.txt_graph_detail = scrolledtext.ScrolledText(f_tree, wrap=tk.WORD, font=("Consolas", 9))
        self.txt_graph_detail.pack(fill=tk.BOTH, expand=True)

    # ---------------- TAB 3: PHÂN TÍCH CHƯƠNG TRÌNH ----------------
    def _build_tab_analytics(self):
        f = self.tab_analytics

        lbl_desc = ttk.Label(
            f,
            text="Phân tích Chương trình Giáo dục Phổ thông 2018 môn Tin học theo dữ liệu đồ thị:",
            font=("Segoe UI", 10, "bold")
        )
        lbl_desc.pack(anchor=tk.W, pady=(0, 10))

        self.txt_analytics = scrolledtext.ScrolledText(f, wrap=tk.WORD, font=("Consolas", 10))
        self.txt_analytics.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ttk.Button(f, text="🔄 Chạy lại Phân tích", command=self._refresh_analytics).pack(anchor=tk.E)

    # ---------------- TAB 4: RA ĐỀ THI TỰ ĐỘNG ----------------
    def _build_tab_exam(self):
        f = self.tab_exam

        f_matrix = ttk.LabelFrame(f, text="Cấu hình Ma trận Đề thi (Theo chuẩn GDPT 2018)", padding=10)
        f_matrix.pack(fill=tk.X, pady=(0, 10))

        r1 = ttk.Frame(f_matrix)
        r1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(r1, text="Khối lớp:").pack(side=tk.LEFT, padx=(0, 5))
        self.exam_grade_var = tk.StringVar(value="Lớp 9")
        ttk.Combobox(r1, textvariable=self.exam_grade_var, values=["Lớp 9", "Lớp 6", "Lớp 7", "Lớp 8", "Lớp 10", "Tất cả"], state="readonly", width=10).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(r1, text="Tổng số câu:").pack(side=tk.LEFT, padx=(0, 5))
        self.exam_num_var = tk.StringVar(value="10")
        ttk.Entry(r1, textvariable=self.exam_num_var, width=6).pack(side=tk.LEFT, padx=(0, 20))

        self.exam_prio_vis_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r1, text="Ưu tiên câu hỏi có gắn Sơ đồ / Giao diện (Visual Assets)", variable=self.exam_prio_vis_var).pack(side=tk.LEFT)

        r2 = ttk.Frame(f_matrix)
        r2.pack(fill=tk.X)
        ttk.Label(r2, text="Tỷ lệ: Nhận biết = 40% | Thông hiểu = 30% | Vận dụng = 30% (Chuẩn Bộ GD&ĐT)").pack(side=tk.LEFT)

        btn_gen = tk.Button(
            f,
            text="🎲 SINH ĐỀ THI TỰ ĐỘNG THEO MA TRẬN TỪ GRAPH DATABASE",
            bg="#16a34a",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            pady=6,
            command=self._generate_exam
        )
        btn_gen.pack(fill=tk.X, pady=(0, 10))

        self.txt_exam = scrolledtext.ScrolledText(f, wrap=tk.WORD, font=("Segoe UI", 10))
        self.txt_exam.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        f_down = ttk.Frame(f)
        f_down.pack(fill=tk.X)
        ttk.Button(f_down, text="💾 Lưu đề thi thành file Text", command=self._save_exam_file).pack(side=tk.RIGHT)

    # ---------------- LOGIC METHODS ----------------
    def save_config(self):
        new_key = self.api_key_var.get().strip()
        set_key(self.env_path, "BEE_API_KEY", new_key)
        self.api_key = new_key
        messagebox.showinfo("Thông báo", "Đã lưu cấu hình API Key vào .env!")

    def select_pdf(self):
        p = filedialog.askopenfilename(title="Chọn tệp PDF SGK", filetypes=[("PDF files", "*.pdf")])
        if p:
            self.selected_files = [p]
            self.listbox.delete(0, tk.END)
            self.listbox.insert(tk.END, f"[PDF] {p}")

    def select_images(self):
        imgs = filedialog.askopenfilenames(title="Chọn ảnh SGK", filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if imgs:
            self.selected_files = list(imgs)
            self.listbox.delete(0, tk.END)
            for img in self.selected_files:
                self.listbox.insert(tk.END, f"[ẢNH] {img}")

    def clear_files(self):
        self.selected_files = []
        self.listbox.delete(0, tk.END)

    def _refresh_master_stats(self):
        stats = self.graph_svc.get_graph_summary()
        text = (
            f"• Số Khối lớp: {stats['total_grades']}  |  Số Chủ đề/Chương: {stats['total_chapters']}  |  Số Bài học: {stats['total_lessons']}\n"
            f"• Số Yêu cầu cần đạt (AIMS_AT): {stats['total_requirements']}  |  Số Khái niệm cốt lõi (CONTAINS): {stats['total_knowledges']}\n"
            f"• Số Kỹ năng (DEVELOPS): {stats['total_skills']}  |  Số Tài nguyên trực quan/Sơ đồ (ILLUSTRATED_BY): {stats['total_visual_assets']}\n"
            f"• Ngân hàng Câu hỏi (EVALUATED_BY): {stats['total_questions']} câu  |  Số Phương án trắc nghiệm: {stats['total_options']}"
        )
        self.lbl_graph_stats.config(text=text)

        # Cập nhật chi tiết bài học
        catalog = self.graph_svc.get_lessons_catalog()
        self.txt_graph_detail.delete("1.0", tk.END)
        for c in catalog:
            self.txt_graph_detail.insert(tk.END, f"======================================================================\n")
            self.txt_graph_detail.insert(tk.END, f"📖 [{c['grade_level']}] {c['lesson_id']}: {c['lesson_title']} (Trang: {c['page_num']})\n")
            self.txt_graph_detail.insert(tk.END, f"   Chương: {c['chapter_title']}\n")
            self.txt_graph_detail.insert(tk.END, f"   🎯 Yêu cầu cần đạt: {len(c['requirements'])} mục | 🧠 Kiến thức cốt lõi: {len(c['knowledges'])} mục\n")
            self.txt_graph_detail.insert(tk.END, f"   🖼️ Sơ đồ/Hình ảnh trực quan: {len(c['visual_assets'])} mục | ❓ Câu hỏi ôn tập: {len(c['questions'])} câu\n\n")

        self._refresh_analytics()

    def _refresh_analytics(self):
        an = self.graph_svc.get_curriculum_analytics()
        self.txt_analytics.delete("1.0", tk.END)
        self.txt_analytics.insert(tk.END, "=== 1. CƠ CẤU MỨC ĐỘ NHẬN THỨC CÂU HỎI (BLOOM) ===\n")
        for k, v in an["difficulty_distribution"].items():
            self.txt_analytics.insert(tk.END, f"  • {k:15}: {v} câu\n")

        self.txt_analytics.insert(tk.END, "\n=== 2. PHÂN BỐ HỌC LIỆU TRỰC QUAN MÔN TIN HỌC ===\n")
        for k, v in an["visual_asset_distribution"].items():
            self.txt_analytics.insert(tk.END, f"  • {k:25}: {v} đối tượng\n")

        self.txt_analytics.insert(tk.END, "\n=== 3. MẬT ĐỘ YÊU CẦU CẦN ĐẠT THEO CHỦ ĐỀ ===\n")
        for ch, count in an["chapter_requirements"].items():
            self.txt_analytics.insert(tk.END, f"  • {ch}: {count} YCCD\n")

    def _open_master_folder(self):
        if os.path.exists(self.master_dir):
            os.startfile(self.master_dir)
        else:
            messagebox.showinfo("Thông báo", "Thư mục Master chưa được tạo.")

    def _open_cypher_script(self):
        c_path = os.path.join(self.master_dir, "neo4j_import_script.cypher")
        if os.path.exists(c_path):
            os.startfile(c_path)
        else:
            messagebox.showinfo("Thông báo", "Chưa có file neo4j_import_script.cypher.")

    def _generate_exam(self):
        grade = self.exam_grade_var.get()
        try:
            num = int(self.exam_num_var.get())
        except ValueError:
            num = 10
        prio = self.exam_prio_vis_var.get()

        res = self.graph_svc.generate_exam(
            grade_level=grade,
            total_questions=num,
            nb_ratio=0.4,
            th_ratio=0.3,
            vd_ratio=0.3,
            prioritize_visuals=prio
        )

        qs = res.get("questions", [])
        mat = res.get("matrix", {})

        self.txt_exam.delete("1.0", tk.END)
        if not qs:
            self.txt_exam.insert(tk.END, "Không có câu hỏi phù hợp trong cơ sở dữ liệu đồ thị.")
            return

        self.txt_exam.insert(tk.END, f"ĐỀ THI MÔN TIN HỌC GDPT 2018 - {grade.upper()}\n")
        self.txt_exam.insert(tk.END, f"Ma trận đề thi: {mat.get('Nhận biết', 0)} Nhận biết | {mat.get('Thông hiểu', 0)} Thông hiểu | {mat.get('Vận dụng', 0)} Vận dụng\n")
        self.txt_exam.insert(tk.END, "="*70 + "\n\n")

        for idx, q in enumerate(qs, 1):
            self.txt_exam.insert(tk.END, f"Câu {idx} [{q['difficulty']}] ({q['lesson_id']}): {q['content']}\n")
            if q.get("visuals") and len(q["visuals"]) > 0:
                vis = q["visuals"][0]
                self.txt_exam.insert(tk.END, f"   [Sơ đồ/Hình ảnh {vis.get('visual_asset_code', '')}: {vis.get('caption', '')}]\n")
                self.txt_exam.insert(tk.END, f"   *Mô tả: {vis.get('detailed_description', '')}*\n")

            if q.get("options"):
                for opt in q["options"]:
                    is_corr = " (ĐÁP ÁN ĐÚNG)" if str(opt.get("is_correct", "")).lower() == "true" else ""
                    self.txt_exam.insert(tk.END, f"   {opt.get('option_key', '')}. {opt.get('content', '')}{is_corr}\n")
            self.txt_exam.insert(tk.END, "\n")

    def _save_exam_file(self):
        content = self.txt_exam.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Cảnh báo", "Chưa có nội dung đề thi để lưu.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Thành công", f"Đã lưu đề thi vào:\n{p}")

    def start_processing(self):
        if not self.selected_files:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp PDF hoặc ảnh SGK trước!")
            return
        if not self.api_key_var.get().strip():
            messagebox.showwarning("Cảnh báo", "Vui lòng cấu hình API Key!")
            return

        self.btn_run.config(state=tk.DISABLED)
        self.is_processing = True
        threading.Thread(target=self._process_worker, daemon=True).start()

    def _process_worker(self):
        try:
            client = AIVisionClient(api_key=self.api_key_var.get().strip(), base_url=self.base_url)
            extractor = MultimodalExtractor(client=client)
            exporter = GraphDataExporter(output_dir=os.path.join(self.project_dir, "output"))

            start_p = int(self.start_page_var.get()) if self.start_page_var.get().strip().isdigit() else None
            end_p = int(self.end_page_var.get()) if self.end_page_var.get().strip().isdigit() else None
            grade = self.grade_var.get()

            temp_dir = os.path.join(self.project_dir, "temp_gui_images")
            tasks = []

            for f_path in self.selected_files:
                if f_path.lower().endswith(".pdf"):
                    t_list = PDFPageProcessor.convert_pdf_to_images(
                        pdf_path=f_path,
                        temp_dir=temp_dir,
                        start_page=start_p,
                        end_page=end_p
                    )
                    for ip, pn in t_list:
                        tasks.append({"path": ip, "source": f_path, "page": pn})
                else:
                    tasks.append({"path": f_path, "source": f_path, "page": 1})

            total = len(tasks)
            bundles = []
            logs = []

            for idx, item in enumerate(tasks, 1):
                self.lbl_status.config(text=f"Trạng thái: Đang phân tích trang {item['page']} ({idx}/{total})...")
                self.progress_bar["value"] = (idx / total) * 100
                self.root.update_idletasks()

                flat_rows, norm_data, log_info = extractor.extract_page(
                    image_path=item["path"],
                    source_file=item["source"],
                    page_num=item["page"],
                    grade_hint=grade
                )
                bundles.append({"flat_rows": flat_rows, "normalized": norm_data})
                logs.append(log_info)

            book_tag = f"sgk_tinhoc_{grade.replace(' ', '_').lower()}"
            s_excel, s_dir, m_dir = exporter.export(bundles, logs, base_name=book_tag)

            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

            self.lbl_status.config(text="Trạng thái: Hoàn tất thành công!")
            self._refresh_master_stats()
            messagebox.showinfo("Thành công", f"Đã trích xuất và tích lũy vào Master Graph thành công!\n\nFile Excel: {os.path.basename(s_excel)}")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")
            self.lbl_status.config(text="Trạng thái: Đã dừng do lỗi.")
        finally:
            self.btn_run.config(state=tk.NORMAL)
            self.progress_bar["value"] = 0
            self.is_processing = False
