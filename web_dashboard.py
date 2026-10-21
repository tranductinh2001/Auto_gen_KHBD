import os
import sys
import shutil
import tempfile
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

# Thiết lập thư mục gốc
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from core.ai_vision_client import AIVisionClient
from core.pdf_processor import PDFPageProcessor
from core.multimodal_extractor import MultimodalExtractor
from exporters.graph_exporter import GraphDataExporter
from core.graph_service import GraphService
from core.exam_builder import ExamMatrixEngine, ExamDocxExporter, CURRICULUM_DATA
from core.ai_exam_generator import AIExamGenerator
from core.curriculum_planner import CurriculumPlanner
from core.exam_exporters import ExamAuxiliaryExporter
from core.study_guide_builder import StudyGuideBuilder, StudyGuideDocxExporter
from core.khbd_ai_generator import KHBDAIGenerator
from core.khbd_docx_exporter import KHBDDocxExporter

def to_clean_slug(text: str) -> str:
    import unicodedata
    if not text:
        return "lop_9"
    nfkd = unicodedata.normalize('NFKD', str(text))
    ascii_str = ''.join([c for c in nfkd if not unicodedata.combining(c)])
    clean = ascii_str.replace('đ', 'd').replace('Đ', 'D').replace(' ', '_').lower()
    return ''.join([c if c.isalnum() or c in ['_', '-'] else '' for c in clean])


# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Hệ thống Tri thức & Ra đề thi Tin học GDPT 2018 (Neo4j)",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------- NÂNG CẤP GIAO DIỆN (PREMIUM MODERN DESIGN SYSTEM) -----------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700;800&family=Outfit:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Nâng cấp Header & Hero Title */
.main-hero-box {
    background: linear-gradient(135deg, rgba(248, 250, 252, 0.9) 0%, rgba(238, 242, 255, 0.95) 100%);
    border: 1px solid rgba(99, 102, 241, 0.25);
    border-radius: 18px;
    padding: 24px 30px;
    margin-bottom: 22px;
    box-shadow: 0 4px 20px -2px rgba(99, 102, 241, 0.08);
}
.hero-pill-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 14px;
    font-size: 0.78rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    color: #4338ca;
    background: rgba(99, 102, 241, 0.12);
    border: 1px solid rgba(99, 102, 241, 0.3);
    border-radius: 20px;
    margin-bottom: 10px;
}
.hero-main-title {
    font-family: 'Outfit', sans-serif;
    font-size: 2.2rem;
    font-weight: 800;
    background: linear-gradient(135deg, #0f172a 0%, #3730a3 55%, #6366f1 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    line-height: 1.25;
    margin-bottom: 8px;
}
.hero-sub-text {
    color: #475569;
    font-size: 1.05rem;
    font-weight: 400;
    line-height: 1.5;
}

/* Nâng cấp các Card chỉ số KPI / Metrics */
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 16px 20px;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02);
    transition: all 0.25s ease;
}
div[data-testid="stMetric"]:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 24px rgba(99, 102, 241, 0.14);
    border-color: #c7d2fe;
}
div[data-testid="stMetricLabel"] {
    color: #64748b;
    font-size: 0.88rem;
    font-weight: 600;
}
div[data-testid="stMetricValue"] {
    color: #0f172a;
    font-family: 'Outfit', sans-serif;
    font-size: 1.85rem;
    font-weight: 700;
}

/* Nâng cấp Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background-color: #f1f5f9;
    padding: 6px;
    border-radius: 14px;
    border: 1px solid #e2e8f0;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    padding: 10px 18px;
    font-weight: 600;
    font-size: 0.95rem;
    color: #475569;
    border: none;
    transition: all 0.2s ease;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%) !important;
    color: #ffffff !important;
    box-shadow: 0 4px 14px rgba(79, 70, 229, 0.32);
}

/* Nút bấm (Buttons) */
div.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%);
    border: none;
    color: white;
    font-weight: 600;
    padding: 12px 24px;
    border-radius: 12px;
    box-shadow: 0 4px 14px rgba(79, 70, 229, 0.35);
    transition: all 0.25s ease;
}
div.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(79, 70, 229, 0.45);
    background: linear-gradient(135deg, #4338ca 0%, #3730a3 100%);
}
div.stDownloadButton > button {
    border-radius: 12px;
    font-weight: 600;
    transition: all 0.25s ease;
}

/* Expanders */
div[data-testid="stExpander"] {
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    background: #ffffff;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02);
    margin-bottom: 16px;
}
div[data-testid="stExpander"] > details > summary {
    font-weight: 600;
    border-radius: 14px;
    padding: 14px 18px;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #f8fafc;
    border-right: 1px solid #e2e8f0;
}
</style>
""", unsafe_allow_html=True)

# Load biến môi trường
ENV_PATH = os.path.join(CURRENT_DIR, ".env")
load_dotenv(ENV_PATH, override=True)
BEE_API_KEY = os.getenv("BEE_API_KEY", "")
BEE_BASE_URL = os.getenv("BEE_BASE_URL", "https://platform.beeknoee.com/api/v1")
OUTPUT_DIR = os.path.join(CURRENT_DIR, "output")
MASTER_DIR = os.path.join(OUTPUT_DIR, "neo4j_import")

# Khởi tạo Graph Service trong Session State để lưu phiên kết nối
if "graph_svc" not in st.session_state:
    st.session_state["graph_svc"] = GraphService(master_dir=MASTER_DIR)
graph_svc = st.session_state["graph_svc"]

if "neo4j_connected" not in st.session_state:
    st.session_state["neo4j_connected"] = False

# ----------------- SIDEBAR: CẤU HÌNH HỆ THỐNG -----------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/network.png", width=64)
    st.title("HỆ THỐNG NEO4J TIN HỌC")
    st.markdown("**Đồ án CSDL Nâng cao**  \n*Quản lý tri thức & Ra đề thi tự động GDPT 2018*")
    st.divider()

    st.subheader("⚙️ Cấu hình Beeknoee AI API")
    api_key_input = st.text_input("BEE_API_KEY", value=BEE_API_KEY, type="password")
    base_url_input = st.text_input("BEE_BASE_URL", value=BEE_BASE_URL)
    current_model = os.getenv("VISION_MODEL", "google/gemini-3.1-flash-lite")
    if "2.5" in current_model or not current_model:
        current_model = "google/gemini-3.1-flash-lite"
    model_input = st.text_input("VISION_MODEL", value=current_model)
    # Tự động chuẩn hóa model cũ sang google/gemini-3.1-flash-lite
    if "2.5" in model_input:
        model_input = "google/gemini-3.1-flash-lite"

    col_btn_api1, col_btn_api2 = st.columns(2)
    with col_btn_api1:
        if st.button("💾 Lưu Cấu hình", use_container_width=True):
            from dotenv import set_key
            set_key(ENV_PATH, "BEE_API_KEY", api_key_input)
            set_key(ENV_PATH, "BEE_BASE_URL", base_url_input)
            set_key(ENV_PATH, "VISION_MODEL", model_input)
            os.environ["BEE_API_KEY"] = api_key_input
            os.environ["BEE_BASE_URL"] = base_url_input
            os.environ["VISION_MODEL"] = model_input
            st.success("Đã cập nhật file .env!")
    with col_btn_api2:
        if st.button("🔍 Kiểm tra API", use_container_width=True):
            import requests as rq
            headers_t = {"Authorization": f"Bearer {api_key_input}", "Content-Type": "application/json"}
            payload_t = {
                "model": model_input,
                "messages": [{"role": "user", "content": "Ping"}],
                "max_tokens": 5
            }
            try:
                r_test = rq.post(f"{base_url_input.rstrip('/')}/chat/completions", headers=headers_t, json=payload_t, timeout=12)
                if r_test.status_code == 200:
                    st.success("🟢 Kết nối Beeknoee AI thành công!")
                else:
                    st.error(f"🔴 Lỗi ({r_test.status_code}): {r_test.text[:120]}")
            except Exception as ex_t:
                st.error(f"🔴 Không thể kết nối: {ex_t}")

    st.divider()
    st.subheader("🛡️ Chế độ Vận Hành & Nguồn Dữ Liệu")
    app_mode = st.radio(
        "Chọn chế độ:",
        [
            "📁 Chế độ Tự Động / Ngoại Tuyến (Xem Ngay Đầy Đủ Dữ Liệu CSV Master)",
            "🔒 Chế độ Kiểm Tra Neo4j Live (Bắt buộc xác thực mật khẩu)"
        ],
        index=0,
        help="Chế độ Tự Động hiển thị ngay cây tri thức và toàn bộ dữ liệu. Khi cần chấm đồ án Live Neo4j, hãy chọn chế độ thứ 2 và nhập mật khẩu 12345678."
    )

    st.divider()
    st.subheader("🔌 Kết nối Neo4j Live DBMS")
    st.info("🔑 **Mật khẩu Neo4j DBMS:** `12345678`\n*(Nhập sai mật khẩu hệ thống sẽ từ chối và KHÔNG đổ dữ liệu)*")
    
    neo_uri = st.text_input("Neo4j Bolt URI", value="bolt://localhost:7687")
    neo_db = st.text_input("Database Name", value="edugraph", help="Tên database sạch 100% không chứa dữ liệu mẫu khác")
    neo_user = st.text_input("Username", value="neo4j")
    neo_pass = st.text_input("Password", value="", type="password", placeholder="Nhập 12345678 để kết nối Live")

    col_btn_conn, col_btn_dis = st.columns(2)
    with col_btn_conn:
        if st.button("🔌 Kết Nối Live", type="primary", use_container_width=True):
            try:
                if graph_svc.connect_neo4j(neo_uri, neo_user, neo_pass, database=neo_db):
                    st.session_state["neo4j_connected"] = True
                    st.session_state["neo_conn_uri"] = neo_uri
                    st.session_state["neo_conn_user"] = neo_user
                    st.session_state["neo_conn_db"] = neo_db
                    st.success(f"✅ Xác thực thành công! Đã kết nối Neo4j Live Database '{neo_db}'.")
                    st.rerun()
            except Exception as e:
                st.session_state["neo4j_connected"] = False
                err_msg = str(e)
                if "10061" in err_msg or "actively refused" in err_msg or "ServiceUnavailable" in err_msg:
                    st.error(
                        "⚠️ **KHÔNG THỂ KẾT NỐI ĐẾN MÁY CHỦ NEO4J!**\n\n"
                        "- **Nguyên nhân:** Dịch vụ Neo4j DBMS chưa được khởi động (chưa Start), hoặc cổng 7687 chưa sẵn sàng.\n"
                        "- **Cách khắc phục:**\n"
                        "  1. Mở **Neo4j Desktop**, tìm đến Database của bạn và bấm nút **Start** (hiển thị màu xanh lá).\n"
                        "  2. Đợi vài giây rồi nhấn lại nút **🔌 Kết Nối Live**.\n"
                        "  3. Hoặc bạn có thể chuyển sang **'📁 Chế độ Ngoại tuyến (Offline Demo - File CSV)'** ở sidebar bên trên để xem dữ liệu ngay mà không cần bật Neo4j."
                    )
                elif "AuthError" in err_msg or "authentication" in err_msg.lower() or "unauthorized" in err_msg.lower():
                    st.error(f"❌ **XÁC THỰC THẤT BẠI:** Tên đăng nhập hoặc mật khẩu không chính xác!\nChi tiết: {err_msg}")
                else:
                    st.error(f"❌ **LỖI KẾT NỐI NEO4J:**\n{err_msg}")

    with col_btn_dis:
        if st.button("Ngắt kết nối", use_container_width=True):
            graph_svc.close_neo4j()
            st.session_state["neo4j_connected"] = False
            st.info("Đã ngắt kết nối.")
            st.rerun()

    st.caption("Developed for Neo4j Advanced Database Project | GDPT 2018")

# ----------------- MAIN TITLE (PREMIUM HERO BANNER) -----------------
st.markdown("""
<div class="main-hero-box">
    <div class="hero-pill-badge">🎓 GDPT 2018 • NEO4J GRAPH DATABASE • AI VISION ENGINE</div>
    <div class="hero-main-title">Hệ Thống Quản Lý Tri Thức & Ra Đề Thi Tự Động</div>
    <div class="hero-sub-text">Môn Tin học Phổ thông GDPT 2018 — Phân tích chương trình, bản đặc tả kỹ thuật và sinh đề thi chuẩn Bộ GD&ĐT</div>
</div>
""", unsafe_allow_html=True)

# Kiểm tra trạng thái Live Neo4j thực tế
is_live = st.session_state.get("neo4j_connected", False) and graph_svc.is_connected()
strict_mode = ("Kiểm Tra Neo4j Live" in app_mode)

if strict_mode:
    if is_live:
        live_stats = graph_svc.get_live_stats()
        n_count = live_stats.get("total_nodes", 0)
        r_count = live_stats.get("total_relationships", 0)
        cur_db = getattr(graph_svc, "database", "edugraph")
        st.success(
            f"🟢 **ĐÃ XÁC THỰC MẬT KHẨU & KẾT NỐI NEO4J LIVE DBMS THÀNH CÔNG** (`bolt://localhost:7687` | Database: `{cur_db}` | User: `neo4j`)\n\n"
            f"Dữ liệu bên dưới được **truy vấn 100% trực tiếp từ CSDL Neo4j** đang chạy trên máy: "
            f"**{n_count:,} Nodes** và **{r_count:,} Relationships**."
        )
    else:
        st.warning(
            "🔒 **CHẾ ĐỘ KIỂM TRA LIVE: ĐANG CHỜ KẾT NỐI VỚI NEO4J DBMS**\n\n"
            "- Bạn đang bật chế độ kiểm tra trực tiếp từ Neo4j Live Database.\n"
            "- **Để kết nối:** Mở **Neo4j Desktop** bấm nút **Start DBMS**, sau đó nhập mật khẩu **`12345678`** tại sidebar bên trái và bấm **'🔌 Kết Nối Live'**.\n"
            "- *(Mẹo: Để xem và kiểm tra đề thi/phân tích ngay mà không cần bật Neo4j, hãy chọn chế độ **'📁 Chế độ Tự Động / Ngoại Tuyến'** ở menu bên trái)*."
        )
else:
    st.info(
        "🟢 **CHẾ ĐỘ TỰ ĐỘNG (HOẠT ĐỘNG HOÀN TOÀN TỰ DO):** Hệ thống đang hiển thị toàn bộ tri thức, biểu đồ phân tích Bloom và công cụ biên soạn đề thi từ kho dữ liệu tích lũy Master."
    )

# Thống kê nhanh: Chỉ hiển thị khi ĐÃ KẾT NỐI (ở strict mode) hoặc khi ở Offline Demo
if strict_mode and not is_live:
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("📚 Bài học (Lessons)", "🔒 Bị khóa")
    kpi2.metric("🎯 YC cần đạt", "🔒 Bị khóa")
    kpi3.metric("🧠 Khái niệm cốt lõi", "🔒 Bị khóa")
    kpi4.metric("🖼️ Trực quan", "🔒 Bị khóa")
    kpi5.metric("❓ Ngân hàng Câu hỏi", "🔒 Bị khóa")
else:
    summary = graph_svc.get_graph_summary()
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("📚 Bài học (Lessons)", summary["total_lessons"])
    kpi2.metric("🎯 YC cần đạt (Requirements)", summary["total_requirements"])
    kpi3.metric("🧠 Khái niệm cốt lõi", summary["total_knowledges"])
    kpi4.metric("🖼️ Trực quan (Visual Assets)", summary["total_visual_assets"])
    kpi5.metric("❓ Ngân hàng Câu hỏi", summary["total_questions"])

st.markdown("---")

# ----------------- 5 TABS CHỨC NĂNG CHÍNH -----------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 1. Trích xuất & Số hóa SGK (AI Vision)",
    "🌐 2. Quản lý Tri thức & Neo4j Graph DB",
    "📊 3. Phân tích Chương trình GDPT 2018",
    "📝 4. Hỗ trợ Ra Đề Thi Tự Động",
    "📑 5. Soạn KHBD Chuẩn 5512 & NLS (TT 02/2025)"
])

# =========================================================
# TAB 1: TRÍCH XUẤT ĐA PHƯƠNG THỨC TỪ SGK
# =========================================================
with tab1:
    st.header("Trích xuất Đa phương thức từ Sách giáo khoa bằng AI Vision")
    st.info(
        "💡 **Tính năng:** Tự động nhận diện cấu trúc sư phạm Tin học 2018 (Chủ đề, Bài học, Yêu cầu cần đạt, "
        "Kiến thức cốt lõi, Kỹ năng, Hình ảnh/Sơ đồ thuật toán/Giao diện kèm mô tả chi tiết, Câu hỏi & Đáp án)."
    )

    col_cfg1, col_cfg2, col_cfg3 = st.columns([1.5, 1, 1])
    with col_cfg1:
        grade_choice = st.selectbox("Khối lớp (Grade Level)", ["Lớp 9", "Lớp 6", "Lớp 7", "Lớp 8", "Lớp 10", "Lớp 11", "Lớp 12"])
    with col_cfg2:
        start_p = st.number_input("Trang bắt đầu (start_page)", min_value=1, value=30, step=1,
                                  help="Giới hạn trang bắt đầu quét từ file PDF để test từng phần.")
    with col_cfg3:
        end_p = st.number_input("Trang kết thúc (end_page)", min_value=1, value=40, step=1,
                                help="Giới hạn trang kết thúc quét từ file PDF.")

    st.markdown("---")
    file_type = st.radio("Chọn nguồn dữ liệu đầu vào:", ["Tệp PDF Sách giáo khoa", "Hình ảnh trang SGK"], horizontal=True)

    uploaded_pdf = None
    uploaded_images = None

    if file_type == "Tệp PDF Sách giáo khoa":
        uploaded_pdf = st.file_uploader("Tải lên file PDF SGK (Ví dụ: SGK Tin học 9)", type=["pdf"])
    else:
        uploaded_images = st.file_uploader("Tải lên một hoặc nhiều ảnh trang SGK", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    btn_extract = st.button("🚀 Bắt đầu Quét & Trích xuất Tri thức", type="primary", use_container_width=True)

    if btn_extract:
        if not api_key_input:
            st.error("Vui lòng nhập BEE_API_KEY ở cột bên trái!")
        elif not uploaded_pdf and not uploaded_images:
            st.warning("Vui lòng tải lên file PDF hoặc hình ảnh trang SGK!")
        else:
            with st.spinner("Đang chuẩn bị môi trường xử lý..."):
                client = AIVisionClient(api_key=api_key_input, base_url=base_url_input, model=model_input)
                extractor = MultimodalExtractor(client=client)
                exporter = GraphDataExporter(output_dir=OUTPUT_DIR)

                temp_work_dir = tempfile.mkdtemp()
                tasks_queue = []

                if uploaded_pdf:
                    temp_pdf_path = os.path.join(temp_work_dir, uploaded_pdf.name)
                    with open(temp_pdf_path, "wb") as f:
                        f.write(uploaded_pdf.getvalue())

                    st.write(f"📄 Đang xử lý PDF: **{uploaded_pdf.name}** (Trang {start_p} đến {end_p})...")
                    img_tasks = PDFPageProcessor.convert_pdf_to_images(
                        pdf_path=temp_pdf_path,
                        temp_dir=temp_work_dir,
                        start_page=start_p,
                        end_page=end_p
                    )
                    for ip, pn in img_tasks:
                        tasks_queue.append({"path": ip, "source": uploaded_pdf.name, "page": pn})
                elif uploaded_images:
                    for idx, img_file in enumerate(uploaded_images, 1):
                        img_path = os.path.join(temp_work_dir, img_file.name)
                        with open(img_path, "wb") as f:
                            f.write(img_file.getvalue())
                        tasks_queue.append({"path": img_path, "source": img_file.name, "page": idx})

                st.success(f"Đã chuẩn bị {len(tasks_queue)} trang cần phân tích bởi AI Vision!")

                # Tiến trình chạy
                progress_bar = st.progress(0)
                status_text = st.empty()
                all_bundles = []
                all_logs = []

                for i, item in enumerate(tasks_queue):
                    status_text.text(f"Đang gọi AI Vision phân tích trang {item['page']} ({i+1}/{len(tasks_queue)})...")
                    flat_rows, norm_data, log_info = extractor.extract_page(
                        image_path=item["path"],
                        source_file=item["source"],
                        page_num=item["page"],
                        grade_hint=grade_choice
                    )
                    all_bundles.append({"flat_rows": flat_rows, "normalized": norm_data})
                    all_logs.append(log_info)
                    progress_bar.progress((i + 1) / len(tasks_queue))

                status_text.text("Đang xuất dữ liệu và đồng bộ vào kho Master Neo4j...")
                book_tag = f"sgk_tinhoc_{grade_choice.replace(' ', '_').lower()}"
                session_excel, session_dir, master_dir = exporter.export(
                    extracted_bundles=all_bundles,
                    logs=all_logs,
                    base_name=book_tag
                )

                # Dọn dẹp temp
                shutil.rmtree(temp_work_dir, ignore_errors=True)

                st.balloons()
                st.success("🎉 TRÍCH XUẤT THÀNH CÔNG VÀ ĐÃ TÍCH LŨY VÀO MASTER GRAPH DATABASE!")

                # Hiển thị kết quả tóm tắt
                col_res1, col_res2 = st.columns(2)
                with col_res1:
                    st.write(f"📁 **Thư mục phiên cào:** `{session_dir}`")
                    st.write(f"🌐 **Thư mục Master Neo4j:** `{master_dir}`")
                with col_res2:
                    if os.path.exists(session_excel):
                        with open(session_excel, "rb") as f:
                            st.download_button(
                                label="📥 Tải file Excel Báo cáo Phiên này",
                                data=f.read(),
                                file_name=os.path.basename(session_excel),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

# =========================================================
# ĐỊNH NGHĨA CÁC HÀM RENDER DỮ LIỆU
# =========================================================

def render_tab2_content():
    """Hiển thị toàn bộ dữ liệu quản lý tri thức và đồ thị Neo4j."""
    sub_t0, sub_t1, sub_t2, sub_t3, sub_t4 = st.tabs([
        "📋 Xem Dữ Liệu Đồ Thị Live (Data Explorer)",
        "🌳 Cây Tri thức Bài học",
        "🕸️ Trực quan hóa Đồ thị (Graph Visualization)",
        "🔍 Cypher Console & Duyệt Đồ thị (Graph Traversal)",
        "📊 Dữ liệu CSV & Script Nạp Neo4j"
    ])

    # --- SUB-TAB 0: XEM DỮ LIỆU ĐỒ THỊ TRỰC TIẾP TỪ NEO4J (DATA EXPLORER) ---
    with sub_t0:
        st.subheader("📋 Trình Khám Phá & Hiển Thị Dữ Liệu Thực Tế Từ Cơ Sở Dữ Liệu Neo4j")
        st.markdown(
            "Tính năng này cho phép bạn **xem toàn bộ dữ liệu** của từng thực thể sư phạm (Nodes & Properties) "
            "được lưu trữ trực tiếp trong CSDL Đồ thị Neo4j hoặc kho Master CSV."
        )

        if is_live:
            st.success("🟢 **Nguồn dữ liệu:** Đang truy vấn TRỰC TIẾP từ Neo4j Database Server qua giao thức Bolt (`bolt://localhost:7687`).")
            l_stats = graph_svc.get_live_stats()
            col_ns1, col_ns2 = st.columns(2)
            with col_ns1:
                with st.expander("📊 Thống kê Số lượng Nodes theo Nhãn (Live Neo4j)", expanded=False):
                    df_labels = pd.DataFrame(
                        [{"Nhãn Node (Label)": k, "Số lượng": v} for k, v in l_stats.get("labels", {}).items() if k in ["Grade", "Chapter", "Lesson", "Requirement", "Knowledge", "Skill", "VisualAsset", "Question", "AnswerOption"]]
                    )
                    st.dataframe(df_labels, use_container_width=True, hide_index=True)
            with col_ns2:
                with st.expander("🔗 Thống kê Mối quan hệ (Relationships) trong Live Neo4j", expanded=False):
                    df_rels = pd.DataFrame(
                        [{"Mối quan hệ (Type)": k, "Số lượng": v} for k, v in l_stats.get("relationships", {}).items()]
                    )
                    st.dataframe(df_rels, use_container_width=True, hide_index=True)
        else:
            st.info("🟡 **Nguồn dữ liệu:** Đang đọc từ Bộ nhớ đệm Master CSV (Chế độ Ngoại tuyến).")

        entity_map = {
            "📚 Bài học (Lesson)": {"label": "Lesson", "csv": "neo4j_lessons.csv"},
            "🎯 Yêu cầu cần đạt chuẩn 2018 (Requirement)": {"label": "Requirement", "csv": "neo4j_requirements.csv"},
            "🧠 Khái niệm & Kiến thức cốt lõi (Knowledge)": {"label": "Knowledge", "csv": "neo4j_knowledges.csv"},
            "🖼️ Hình ảnh & Sơ đồ trực quan (VisualAsset)": {"label": "VisualAsset", "csv": "neo4j_visual_assets.csv"},
            "❓ Ngân hàng Câu hỏi trắc nghiệm & tự luận (Question)": {"label": "Question", "csv": "neo4j_questions.csv"},
            "🔘 Phương án trắc nghiệm A-B-C-D (AnswerOption)": {"label": "AnswerOption", "csv": "neo4j_answer_options.csv"},
            "⚡ Kỹ năng thao tác & lập trình (Skill)": {"label": "Skill", "csv": "neo4j_skills.csv"},
            "📑 Chủ đề / Chương SGK (Chapter)": {"label": "Chapter", "csv": "neo4j_lessons.csv"}
        }

        col_ent_sel, col_lim = st.columns([3, 1])
        with col_ent_sel:
            selected_entity = st.selectbox("Chọn bảng thực thể cần xem dữ liệu:", list(entity_map.keys()))
        with col_lim:
            limit_rows = st.number_input("Giới hạn số dòng hiển thị:", min_value=10, max_value=2000, value=200, step=50)

        ent_info = entity_map[selected_entity]
        df_display = pd.DataFrame()

        if is_live:
            with st.spinner(f"Đang truy vấn Nodes {ent_info['label']} từ Neo4j Live..."):
                raw_nodes = graph_svc.get_live_nodes(label=ent_info["label"], limit=limit_rows)
                if raw_nodes:
                    df_display = pd.DataFrame(raw_nodes)
                    if "_element_id" in df_display.columns:
                        cols = [c for c in df_display.columns if c != "_element_id"] + ["_element_id"]
                        df_display = df_display[cols]
        else:
            csv_path = os.path.join(MASTER_DIR, ent_info["csv"])
            if os.path.exists(csv_path):
                df_display = pd.read_csv(csv_path, encoding="utf-8-sig")
                if len(df_display) > limit_rows:
                    df_display = df_display.head(limit_rows)

        if not df_display.empty:
            search_kw = st.text_input(f"🔍 Tìm kiếm nhanh trong bảng '{selected_entity}':", "")
            if search_kw:
                mask = df_display.astype(str).apply(lambda row: row.str.contains(search_kw, case=False, na=False)).any(axis=1)
                df_filtered = df_display[mask]
            else:
                df_filtered = df_display

            st.markdown(f"**Tổng số bản ghi hiển thị:** `{len(df_filtered)}` / `{len(df_display)}` dòng")
            st.dataframe(df_filtered, use_container_width=True)

            csv_data = df_filtered.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                label=f"📥 Xuất dữ liệu {ent_info['label']} ra file CSV",
                data=csv_data,
                file_name=f"neo4j_live_{ent_info['label'].lower()}.csv",
                mime="text/csv"
            )
        else:
            st.warning(f"Hiện chưa có dữ liệu cho thực thể '{selected_entity}'.")

        st.markdown("---")

    # --- SUB-TAB 1: CÂY TRI THỨC SƯ PHẠM ---
    with sub_t1:
        st.subheader("Duyệt cây tri thức theo Bài học môn Tin học")
        lessons_list = graph_svc.get_lessons_catalog()

        if not lessons_list:
            st.warning("Chưa có dữ liệu bài học nào. Vui lòng kết nối Neo4j hoặc chạy trích xuất!")
        else:
            lesson_options = [f"[{l['grade_level']}] {l['lesson_id']}: {l['lesson_title']}" for l in lessons_list]
            selected_idx = st.selectbox("Chọn bài học cần xem cấu trúc tri thức:", range(len(lesson_options)), format_func=lambda x: lesson_options[x], key="sb_lesson_tree")
            sel_l = lessons_list[selected_idx]

            st.markdown(f"### 📖 {sel_l['lesson_id']}: {sel_l['lesson_title']}")
            st.caption(f"Khối lớp: **{sel_l['grade_level']}** | Chương/Chủ đề: **{sel_l['chapter_title']}** | Trang SGK: **{sel_l['page_num']}**")

            c_tree1, c_tree2 = st.columns(2)
            with c_tree1:
                st.markdown("#### 🎯 Yêu cầu cần đạt (Requirements)")
                if sel_l["requirements"]:
                    for r in sel_l["requirements"]:
                        st.markdown(f"- **[{r.get('requirement_code', 'YCCD')}]:** {r.get('requirement_content', '')}")
                else:
                    st.write("*(Chưa có bản ghi)*")

                st.markdown("#### 🧠 Kiến thức cốt lõi (Knowledge Nodes)")
                if sel_l["knowledges"]:
                    for k in sel_l["knowledges"]:
                        with st.expander(f"🔹 {k.get('knowledge_name', '')}"):
                            st.write(k.get("knowledge_content", ""))
                else:
                    st.write("*(Chưa có bản ghi)*")

            with c_tree2:
                st.markdown("#### 🖼️ Hình ảnh & Sơ đồ trực quan (Visual Assets)")
                if sel_l["visual_assets"]:
                    for v in sel_l["visual_assets"]:
                        with st.expander(f"📷 {v.get('visual_asset_code', '')}: {v.get('caption', '')}"):
                            st.write(f"**Mô tả chi tiết của AI Vision:**")
                            st.info(v.get("detailed_description", ""))
                else:
                    st.write("*(Chưa có bản ghi)*")

                st.markdown("#### ❓ Câu hỏi ôn tập & Luyện tập")
                if sel_l["questions"]:
                    for q in sel_l["questions"]:
                        st.markdown(f"- **{q.get('question_id', '')}** `[{q.get('difficulty', '')}]`: {q.get('content', '')}")
                else:
                    st.write("*(Chưa có bản ghi)*")

    # --- SUB-TAB 2: TRỰC QUAN HÓA ĐỒ THỊ TƯƠNG TÁC (PYVIS) ---
    with sub_t2:
        st.subheader("Trực quan hóa Lược đồ Đồ thị & Mạng lưới Tri thức (Interactive Graph)")
        st.info("💡 **Tính năng đồ họa:** Tương tác trực tiếp (kéo thả node, phóng to/thu nhỏ, rê chuột để xem chi tiết thuộc tính).")

        viz_mode = st.radio(
            "Chọn chế độ hiển thị đồ thị:",
            ["Lược đồ cấu trúc tổng thể (Meta-Schema như arrows.app)", "Mạng lưới tri thức cụ thể từng bài học (Lesson Subgraph)"],
            horizontal=True
        )

        import streamlit.components.v1 as components

        if viz_mode == "Lược đồ cấu trúc tổng thể (Meta-Schema như arrows.app)":
            st.markdown("##### 📌 Lược đồ Đồ thị Tri thức Chuẩn GDPT 2018 (`CALL db.schema.visualization();`)")
            html_schema = graph_svc.generate_pyvis_schema_html()
            components.html(html_schema, height=550, scrolling=True)
            st.caption("Các đỉnh (Nodes) đại diện cho thực thể sư phạm; Các cạnh (Edges) định nghĩa quan hệ logic.")
        else:
            st.markdown("##### 🌐 Mạng lưới Tri thức Tương tác Chi tiết theo Bài học")
            lessons_list = graph_svc.get_lessons_catalog()
            if lessons_list:
                sel_lid = st.selectbox(
                    "Chọn bài học cần trực quan hóa đồ thị liên kết:",
                    [l["lesson_id"] for l in lessons_list],
                    format_func=lambda x: f"{x} - " + next((l["lesson_title"] for l in lessons_list if l["lesson_id"] == x), "")
                )
                html_sub = graph_svc.generate_pyvis_lesson_subgraph_html(lesson_id=sel_lid)
                components.html(html_sub, height=550, scrolling=True)

    # --- SUB-TAB 3: CYPHER GRAPH TRAVERSAL CONSOLE ---
    with sub_t3:
        st.subheader("Bảng Điều Khiển Truy Vấn Đồ Thị Bằng Cypher (Graph Traversal)")
        st.markdown("Duyệt đồ thị qua các mối quan hệ mà **không cần các câu lệnh JOIN phức tạp**:")

        query_templates = {
            "1. Duyệt toàn bộ cây tri thức của Bài 1":
                "MATCH (l:Lesson {lesson_id: 'Bài 1'})-[r]->(x)\nRETURN l.lesson_title AS Lesson, type(r) AS Relationship, labels(x)[0] AS TargetType, coalesce(x.content, x.name, x.caption, x.title) AS TargetDetail\nLIMIT 20;",
            "2. Tìm các bài học có gắn Sơ đồ thuật toán & Câu hỏi Vận dụng":
                "MATCH (l:Lesson)-[:HAS_VISUAL_ASSET]->(v:VisualAsset),\n      (l)-[:TESTS_KNOWLEDGE]->(q:Question {difficulty: 'Vận dụng'})\nRETURN l.lesson_id AS BaiHoc, l.lesson_title AS TieuDe, v.caption AS SoDo, q.content AS CauHoi\nLIMIT 15;",
            "3. Thống kê số lượng Chuẩn đầu ra (Requirement) theo từng Chủ đề":
                "MATCH (c:Chapter)-[:HAS_LESSON]->(l:Lesson)-[:MEETS_REQUIREMENT]->(r:Requirement)\nRETURN c.title AS ChuDe, count(DISTINCT l) AS SoBaiHoc, count(r) AS TongYCCD\nORDER BY TongYCCD DESC;",
            "4. Duyệt Ngân hàng Đề thi kèm các phương án lựa chọn (Graph Traversal)":
                "MATCH (l:Lesson)-[:TESTS_KNOWLEDGE]->(q:Question)-[:HAS_OPTION]->(opt:AnswerOption)\nRETURN l.lesson_id AS BaiHoc, q.content AS CauHoi, q.difficulty AS MucDo, collect(opt.option_key + '. ' + opt.content) AS PhuongAn\nLIMIT 10;",
            "5. Kiểm tra danh mục các Nhãn (Labels) trong Neo4j":
                "CALL db.labels() YIELD label\nRETURN label AS NodeLabel;"
        }

        sel_tpl_name = st.selectbox("Chọn mẫu truy vấn Cypher chuẩn sư phạm:", list(query_templates.keys()))
        default_cypher = query_templates[sel_tpl_name]
        cypher_input = st.text_area("Câu lệnh Cypher:", value=default_cypher, height=130)

        if st.button("🚀 Thực Thi Cypher", type="primary"):
            import time
            t_start = time.time()
            try:
                results = graph_svc.execute_cypher(cypher_input)
                dur = round((time.time() - t_start) * 1000, 2)
                st.success(f"✅ Truy vấn thành công! Trả về {len(results)} bản ghi (Thời gian thực thi: {dur} ms)")
                if results:
                    df_res = pd.DataFrame(results)
                    st.dataframe(df_res, use_container_width=True)
                else:
                    st.info("Không có dữ liệu trả về cho điều kiện truy vấn này.")
            except Exception as e:
                st.error(f"❌ Lỗi thực thi Cypher: {str(e)}")

    # --- SUB-TAB 4: DỮ LIỆU CSV & CYPHER SCRIPT NẠP ---
    with sub_t4:
        st.subheader("Dữ Liệu Chuẩn Hóa CSV & Tập Lệnh LOAD CSV")
        col_csv1, col_csv2 = st.columns(2)
        with col_csv1:
            csv_choice = st.selectbox(
                "Xem trước bảng dữ liệu CSV đã trích xuất:",
                ["neo4j_lessons.csv", "neo4j_requirements.csv", "neo4j_knowledges.csv", "neo4j_visual_assets.csv", "neo4j_questions.csv", "neo4j_answer_options.csv"]
            )
            csv_p = os.path.join(MASTER_DIR, csv_choice)
            if os.path.exists(csv_p):
                df_view = pd.read_csv(csv_p, encoding="utf-8-sig")
                st.dataframe(df_view, use_container_width=True)

        with col_csv2:
            st.markdown("##### Tập lệnh Cypher Script Nạp Dữ Liệu vào Neo4j")
            cypher_path = os.path.join(MASTER_DIR, "neo4j_import_script.cypher")
            if os.path.exists(cypher_path):
                with open(cypher_path, "r", encoding="utf-8") as f:
                    c_content = f.read()
                st.code(c_content[:1000] + "\n\n// ... (và các lệnh khác)", language="cypher")
                st.download_button(
                    label="📥 Tải tệp neo4j_import_script.cypher",
                    data=c_content,
                    file_name="neo4j_import_script.cypher",
                    mime="text/plain",
                    use_container_width=True
                )


def render_tab3_content():
    """Hiển thị phân tích chương trình GDPT 2018."""
    analytics = graph_svc.get_curriculum_analytics()
    col_an1, col_an2 = st.columns(2)

    with col_an1:
        st.subheader("1. Cơ cấu Mức độ Nhận thức Ngân hàng Đề thi")
        diff_data = analytics["difficulty_distribution"]
        df_diff = pd.DataFrame(list(diff_data.items()), columns=["Mức độ", "Số lượng"])
        st.bar_chart(df_diff.set_index("Mức độ"))
        st.caption("Tỷ lệ phân bố câu hỏi theo Bloom: Nhận biết - Thông hiểu - Vận dụng.")

    with col_an2:
        st.subheader("2. Phân bố Học liệu Trực quan (Visual Assets)")
        vis_data = analytics["visual_asset_distribution"]
        df_vis = pd.DataFrame(list(vis_data.items()), columns=["Phân loại", "Số lượng"])
        st.bar_chart(df_vis.set_index("Phân loại"))
        st.caption("Mật độ Sơ đồ thuật toán, Giao diện phần mềm ứng dụng và Sơ đồ tư duy.")

    st.subheader("3. Mật độ Yêu cầu cần đạt theo Chương / Chủ đề")
    ch_data = analytics["chapter_requirements"]
    if ch_data:
        df_ch = pd.DataFrame(list(ch_data.items()), columns=["Chương / Chủ đề", "Số YCCD"]).sort_values(by="Số YCCD", ascending=False)
        st.dataframe(df_ch, use_container_width=True)
    else:
        st.info("Chưa có dữ liệu chương.")


def render_tab4_content():
    """Hỗ trợ ra đề thi tự động chuẩn ma trận và bản đặc tả GDPT 2018 (Xuất file DOCX chuẩn mực 100%)."""
    st.markdown("### 🎓 Công Cụ Biên Soạn Đề Thi Theo Khung Ma Trận & Bản Đặc Tả Chuẩn Bộ GD&ĐT")
    st.info(
        "💡 **Cơ chế hoạt động:** Hệ thống tích hợp trực tiếp **Beeknoee AI** để phân loại chuẩn xác mức độ nhận thức "
        "(Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao), bám sát chuẩn Yêu cầu cần đạt môn Tin học GDPT 2018, "
        "tự động lập Khung ma trận (Landscape), Bản đặc tả (Landscape), sinh 02 Mã đề hoán vị TH01 & TH02 kèm bảng điền đáp án và dòng kẻ làm bài, "
        "xuất trọn gói thành file Word (.docx) chuyên nghiệp không thiếu sót bất kỳ phần nào."
    )

    ai_gen_svc = AIExamGenerator(api_key=api_key_input, base_url=base_url_input, model=model_input)

    # Lựa chọn động cơ khảo thí
    st.markdown("##### 🚀 Lựa chọn Động cơ Biên soạn Đề thi:")
    col_eng1, col_eng2 = st.columns([2.5, 1.2])
    with col_eng1:
        gen_engine_choice = st.radio(
            "Phương thức tạo đề:",
            [
                "🤖 Beeknoee AI Engine (Khuyến nghị: Phân loại sâu ma trận nhận thức & ngữ cảnh hoá câu hỏi chuẩn GDPT 2018)",
                "📁 Local Offline Engine (Trích xuất nhanh từ ngân hàng câu hỏi định dạng sẵn)"
            ],
            index=0,
            horizontal=False
        )
    with col_eng2:
        if ai_gen_svc.is_available():
            st.success(f"🟢 **Beeknoee AI: Đã kết nối**\n- Base URL: `{base_url_input}`\n- Model: `{ai_gen_svc.model}`")
        else:
            st.warning("⚠️ Chưa nhập API Key Beeknoee. Hãy điền BEE_API_KEY ở cột trái để mở khóa AI.")

    with st.expander("⚙️ Thiết lập Thông tin Hành chính & Đơn vị Khảo thí", expanded=True):
        col_adm1, col_adm2 = st.columns(2)
        with col_adm1:
            ex_gov = st.text_input("Đơn vị chủ quản / Phòng GD&ĐT:", value="UBND XÃ KON CHIÊNG")
            ex_school = st.text_input("Trường học:", value="TRƯỜNG TH&THCS ĐĂK TRÔI")
            ex_grade = st.selectbox("Khối lớp ra đề:", ["Lớp 9", "Lớp 6", "Lớp 7", "Lớp 8"])
            ex_exam = st.selectbox("Kỳ kiểm tra:", [
                "KIỂM TRA GIỮA HỌC KỲ I",
                "KIỂM TRA CUỐI HỌC KỲ I",
                "KIỂM TRA GIỮA HỌC KỲ II",
                "KIỂM TRA CUỐI HỌC KỲ II",
                "KIỂM TRA ĐỊNH KỲ THƯỜNG XUYÊN"
            ])
        with col_adm2:
            ex_year = st.text_input("Năm học:", value="2026 - 2027")
            ex_time = st.number_input("Thời gian làm bài (phút):", min_value=15, max_value=120, value=45, step=5)
            ex_teacher = st.text_input("Giáo viên ra đề:", value="Trần Đức Tình")
            ex_head = st.text_input("Tổ trưởng chuyên môn thẩm định:", value="Hoàng Thị Kim Tuyết")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                ex_matrix_date = st.text_input("Ngày làm ma trận & đặc tả:", value="20/10/2026")
            with col_d2:
                ex_exam_date = st.text_input("Ngày ra đề & đáp án:", value="21/10/2026")

    # ----------------- PHÂN HỆ QUẢN LÝ KẾ HOẠCH GIÁO DỤC (KHGD) CẢ NĂM -----------------
    st.markdown("##### 📅 Kế Hoạch Giáo Dục (KHGD) / Phân Phối Chương Trình Cả Năm & Phạm Vi Ra Đề")
    with st.expander("📋 Quản lý KHGD Cả Năm (35 tuần) & Lọc Bài Học Đã Dạy Tránh Ra Nhầm Bài Chưa Học", expanded=True):
        st.info(
            "💡 **Ý nghĩa sư phạm:** Giáo viên có thể tải lên Kế hoạch giáo dục của nhà trường hoặc sử dụng KHGD 35 tuần chuẩn của Bộ GD&ĐT. "
            "Hệ thống sẽ **khóa hoàn toàn** các bài học chưa dạy tới, chỉ sinh câu hỏi trắc nghiệm, đúng/sai và tự luận nằm trong các bài học đã học!"
        )

        col_khgd_src, col_khgd_dl = st.columns([2.5, 1.5])
        with col_khgd_src:
            khgd_source_mode = st.radio(
                "Nguồn Kế hoạch giáo dục:",
                [
                    "📖 Sử dụng KHGD 35 tuần chuẩn GDPT 2018 (SGK Kết nối tri thức)",
                    "📤 Tải lên file KHGD riêng của trường (.docx, .xlsx, .pdf, .csv, .json)"
                ],
                index=0,
                horizontal=True
            )
        with col_khgd_dl:
            sample_khgd_bytes = CurriculumPlanner.export_sample_excel(grade=ex_grade)
            st.download_button(
                label=f"📥 Tải KHGD Mẫu ({ex_grade})",
                data=sample_khgd_bytes,
                file_name=f"KHGD_Mau_TinHoc_{to_clean_slug(ex_grade)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help=f"Tải về file Excel mẫu 35 tuần chuẩn của {ex_grade} để giáo viên chỉnh sửa theo thực tế nhà trường rồi tải lên lại."
            )

        active_full_plan = []
        if khgd_source_mode.startswith("📤"):
            uploaded_khgd = st.file_uploader(
                "Chọn file KHGD của trường (hỗ trợ .docx, .xlsx, .pdf, .csv, .json - có thể chứa gộp cả Tin 6, 7, 8, 9):",
                type=["docx", "doc", "xlsx", "xls", "pdf", "csv", "json", "txt"]
            )
            if uploaded_khgd is not None:
                try:
                    # Bóc tách thông minh đa định dạng và đa khối lớp
                    school_plans_dict = CurriculumPlanner.parse_multimodal_school_plan(
                        uploaded_khgd.getvalue(),
                        uploaded_khgd.name
                    )
                    st.session_state["uploaded_school_khgd"] = school_plans_dict
                    
                    detected_summary = []
                    for g_key in ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"]:
                        g_items = school_plans_dict.get(g_key, [])
                        if g_items:
                            detected_summary.append(f"**{g_key}**: {len(g_items)} bài")
                    
                    if detected_summary:
                        st.success(f"🤖 **AI Parser:** Đã nhận diện & phân tách KHGD nhà trường thành công: {' | '.join(detected_summary)}!")
                    else:
                        st.info("✅ Đã nạp dữ liệu KHGD từ file tải lên.")

                    # Lấy KHGD ứng với khối lớp đang chọn
                    target_plan = school_plans_dict.get(ex_grade, [])
                    if target_plan:
                        active_full_plan = target_plan
                    else:
                        # Thử lấy lớp đầu tiên có dữ liệu hoặc fallback về chuẩn
                        first_avail = next((v for v in school_plans_dict.values() if v), [])
                        if first_avail:
                            st.warning(f"⚠️ File tải lên không có mục riêng cho **{ex_grade}**. Đang áp dụng mục bài học tìm thấy trong file.")
                            active_full_plan = first_avail
                        else:
                            st.warning(f"⚠️ Chưa trích xuất được bài học cho **{ex_grade}**. Đang sử dụng KHGD chuẩn thay thế.")
                            active_full_plan = CurriculumPlanner.get_standard_plan(ex_grade)
                except Exception as ex_kh:
                    st.error(f"❌ Lỗi phân tích file KHGD: {ex_kh}. Đang dùng KHGD chuẩn thay thế.")
                    active_full_plan = CurriculumPlanner.get_standard_plan(ex_grade)
            else:
                st.caption("💡 Giáo viên có thể tải lên file Kế hoạch dạy học của Tổ chuyên môn (Word/Excel/PDF). Hệ thống sẽ tự động bóc tách bài học và Yêu cầu cần đạt.")
                active_full_plan = CurriculumPlanner.get_standard_plan(ex_grade)
        else:
            active_full_plan = CurriculumPlanner.get_standard_plan(ex_grade)

        # Thiết lập mốc tuần tối đa đã học theo kỳ thi
        default_max_week = 9
        if "GIỮA HỌC KỲ I" in ex_exam.upper():
            default_max_week = 9
        elif "CUỐI HỌC KỲ I" in ex_exam.upper():
            default_max_week = 18
        elif "GIỮA HỌC KỲ II" in ex_exam.upper():
            default_max_week = 27
        elif "CUỐI HỌC KỲ II" in ex_exam.upper():
            default_max_week = 35

        col_w1, col_w2 = st.columns([1.5, 2.5])
        with col_w1:
            selected_max_week = st.slider(
                "Đã giảng dạy đến tuần thứ mấy?",
                min_value=1,
                max_value=35,
                value=int(default_max_week),
                step=1,
                help="Chỉ ra đề vào các bài học thuộc từ Tuần 1 đến Tuần này."
            )
        with col_w2:
            st.markdown(f"**Thời điểm kiểm tra:** `{ex_exam}`  \n**Phạm vi áp dụng:** Từ **Tuần 1** đến **Tuần {selected_max_week}**.")

        # Lọc danh sách bài học đã dạy
        filtered_plan_items = CurriculumPlanner.filter_by_exam_period(active_full_plan, ex_exam, max_week=selected_max_week)
        unique_units_map = {}
        for it in filtered_plan_items:
            u_name = it.get("unit", "")
            if u_name and u_name not in unique_units_map:
                unique_units_map[u_name] = it

        all_candidate_units = list(unique_units_map.keys())
        selected_units_for_exam = st.multiselect(
            "📌 Chọn các bài học đưa vào đề thi (mặc định đã chọn toàn bộ các bài đã dạy):",
            options=all_candidate_units,
            default=all_candidate_units,
            help="Bỏ chọn nếu có bài nào học sinh chưa kịp học hoặc không thi."
        )

        st.success(
            f"🛡️ **CƠ CHẾ BẢO VỆ PHẠM VI:** Đã khóa chặt đề thi trong **{len(selected_units_for_exam)} bài học đã chọn**. "
            f"Các bài từ Tuần {selected_max_week + 1} trở đi và các bài chưa chọn **TUYỆT ĐỐI KHÔNG XUẤT HIỆN TRONG ĐỀ THI**!"
        )

        with st.expander("👁️ Xem bảng chi tiết KHGD các bài học trong phạm vi ra đề", expanded=False):
            raw_list = [unique_units_map[u] for u in selected_units_for_exam if u in unique_units_map]
            if raw_list:
                df_khgd_view = pd.DataFrame(raw_list)
                rename_cols = {
                    "week": "Tuần",
                    "lesson_index": "Tiết",
                    "semester": "Học kỳ",
                    "chapter": "Chủ đề / Chương",
                    "unit": "Tên bài học",
                    "content": "Yêu cầu cần đạt (YCCĐ)",
                    "assessment": "Thời điểm kiểm tra"
                }
                df_khgd_view = df_khgd_view.rename(columns=rename_cols)
                st.dataframe(df_khgd_view, use_container_width=True, hide_index=True)

    st.markdown("##### 📊 Thiết lập Tỷ lệ Ma trận Mức độ Nhận thức (%)")
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        ratio_nb = st.slider("Nhận biết (%)", 0, 100, 30, step=5, help="Mức độ nhận biết kiến thức cốt lõi")
    with col_m2:
        ratio_th = st.slider("Thông hiểu (%)", 0, 100, 40, step=5, help="Mức độ giải thích, phân biệt, suy luận")
    with col_m3:
        ratio_vd = st.slider("Vận dụng (%)", 0, 100, 30, step=5, help="Mức độ vận dụng kiến thức giải quyết vấn đề")
    with col_m4:
        ratio_vdc = st.slider("Vận dụng cao (%)", 0, 50, 0, step=5, help="Mức độ giải quyết vấn đề sáng tạo, liên môn")

    sum_ratio = ratio_nb + ratio_th + ratio_vd + ratio_vdc
    if sum_ratio != 100:
        st.warning(f"⚠️ Tổng tỷ lệ hiện tại là {sum_ratio}%. Khi sinh đề hệ thống sẽ tự động quy chuẩn (normalize) về đúng 100%.")

    st.markdown("##### 📚 Thiết lập Tỷ lệ Phân bổ theo Chủ đề SGK (%)")
    topics_list = CURRICULUM_DATA.get(ex_grade, {}).get("topics", [])
    topic_ratios_input = {}
    if topics_list:
        n_top = len(topics_list)
        default_pct = int(100 / n_top)
        cols_top = st.columns(min(n_top, 4))
        for t_idx, top in enumerate(topics_list):
            c_target = cols_top[t_idx % len(cols_top)]
            ch_title = top["chapter"]
            val_init = default_pct if t_idx < n_top - 1 else (100 - default_pct * (n_top - 1))
            with c_target:
                topic_ratios_input[ch_title] = st.slider(
                    f"{ch_title} (%)",
                    min_value=0,
                    max_value=100,
                    value=int(val_init),
                    step=5,
                    key=f"topic_ratio_{ex_grade}_{t_idx}"
                )

    custom_yccd_payload = []
    with st.expander("📋 Kế Hoạch Bài Dạy & Cấu Hình Yêu Cầu Cần Đạt (YCCĐ) Chuẩn Sư Phạm", expanded=False):
        st.markdown(
            "💡 **YCCĐ bám sát KHGD:** Giáo viên có thể điều chỉnh yêu cầu cần đạt cho các bài học đã chọn. "
            "Hệ thống sẽ ép AI hoặc Local Engine bám sát chính xác các nội dung này."
        )
        use_custom_yccd = st.checkbox("🎯 Bắt buộc bám sát chính xác các Yêu cầu cần đạt bên dưới khi ra đề", value=True)

        # Lấy YCCĐ từ các bài đã chọn trong KHGD (hỗ trợ toàn bộ các bài học trong phạm vi ra đề)
        units_to_show = selected_units_for_exam[:8] if len(selected_units_for_exam) > 8 else selected_units_for_exam
        for idx_y, u_name in enumerate(units_to_show):
            it_info = unique_units_map.get(u_name, {})
            c_y1, c_y2 = st.columns([1, 2.5])
            with c_y1:
                ch_val = st.text_input(f"Chương/Chủ đề {idx_y+1}:", value=it_info.get("chapter", "Chủ đề SGK"), key=f"yccd_ch_{ex_grade}_{idx_y}")
                un_val = st.text_input(f"Tên bài {idx_y+1}:", value=u_name, key=f"yccd_un_{ex_grade}_{idx_y}")
            with c_y2:
                default_content = it_info.get("content", "Nhận biết kiến thức cốt lõi, thông hiểu và vận dụng thực tiễn.")
                req_val = st.text_area(f"YCCĐ - Bài {idx_y+1}:", value=default_content, height=80, key=f"yccd_req_{ex_grade}_{idx_y}")

            if use_custom_yccd and req_val.strip():
                custom_yccd_payload.append({
                    "chapter": ch_val.strip(),
                    "unit": un_val.strip(),
                    "requirement": req_val.strip(),
                    "req_nb": it_info.get("req_nb", ""),
                    "req_th": it_info.get("req_th", ""),
                    "req_vd": it_info.get("req_vd", "")
                })

    with st.expander("📝 Cấu trúc Đề thi & Định dạng Đề mẫu", expanded=False):
        c_fmt1, c_fmt2, c_fmt3 = st.columns(3)
        c_fmt1.markdown("**Phần I: TNKQ Nhiều lựa chọn**\n- 12 câu hỏi (4 phương án A, B, C, D)\n- Tổng: 3,0 điểm (0,25đ / câu)")
        c_fmt2.markdown("**Phần II: TNKQ Đúng / Sai**\n- 04 câu tình huống (mỗi câu 4 lệnh hỏi a, b, c, d)\n- Tổng: 4,0 điểm (0,25đ / lệnh hỏi)")
        c_fmt3.markdown("**Phần III: Tự luận**\n- 03 câu hỏi thực tiễn / vận dụng\n- Tổng: 3,0 điểm (1,0đ / câu)")
        st.caption("✔️ Điểm toàn bài: **10,0 điểm** | Trắc nghiệm: **7,0 điểm (70%)** - Tự luận: **3,0 điểm (30%)** theo đúng Thông tư và đề mẫu.")

    btn_build_exam = st.button("🎲 TẠO MA TRẬN & BIÊN SOẠN HỒ SƠ ĐỀ THI CHUẨN ĐẶC TẢ", type="primary", use_container_width=True)

    if btn_build_exam or "current_dossier" in st.session_state:
        if btn_build_exam:
            use_ai = gen_engine_choice.startswith("🤖")
            dossier = None

            if use_ai:
                if not ai_gen_svc.is_available():
                    st.warning("⚠️ Không tìm thấy API Key Beeknoee. Hệ thống sẽ tự động chuyển sang chế độ Local Offline Engine.")
                    use_ai = False
                else:
                    with st.spinner("🤖 Đang kết nối Beeknoee AI để phân tích ma trận nhận thức và biên soạn câu hỏi bám sát bài đã học..."):
                        try:
                            dossier = ai_gen_svc.generate_pedagogical_exam(
                                grade_level=ex_grade,
                                exam_title=ex_exam,
                                school_name=ex_school,
                                governing_body=ex_gov,
                                school_year=ex_year,
                                teacher_name=ex_teacher,
                                head_name=ex_head,
                                pct_nb=float(ratio_nb),
                                pct_th=float(ratio_th),
                                pct_vd=float(ratio_vd),
                                pct_vdc=float(ratio_vdc),
                                topic_ratios=topic_ratios_input,
                                custom_requirements=custom_yccd_payload if custom_yccd_payload else None,
                                allowed_units=selected_units_for_exam
                            )
                            st.session_state["used_engine"] = f"Beeknoee AI Engine ({ai_gen_svc.model})"
                        except Exception as e:
                            st.warning(f"⚠️ {str(e)}\n\n👉 **Hệ thống đã tự động kích hoạt Local Offline Engine dự phòng chuẩn GDPT 2018.**")
                            use_ai = False

            if not use_ai or dossier is None:
                with st.spinner("Đang sử dụng Local Offline Engine biên soạn hồ sơ đề thi bám sát bài đã học..."):
                    engine = ExamMatrixEngine(
                        grade_level=ex_grade,
                        exam_title=ex_exam,
                        school_name=ex_school,
                        governing_body=ex_gov,
                        school_year=ex_year,
                        teacher_name=ex_teacher,
                        head_name=ex_head,
                        pct_nb=float(ratio_nb),
                        pct_th=float(ratio_th),
                        pct_vd=float(ratio_vd),
                        pct_vdc=float(ratio_vdc),
                        num_mcq=12,
                        num_tf=4,
                        num_essay=3,
                        duration_minutes=ex_time,
                        topic_ratios=topic_ratios_input,
                        allowed_units=selected_units_for_exam,
                        custom_requirements=custom_yccd_payload if custom_yccd_payload else None
                    )
                    dossier = engine.generate_exam_dossier()
                    st.session_state["used_engine"] = "Local Offline Engine (Bám sát KHGD)"

            if dossier and "meta" in dossier:
                dossier["meta"]["matrix_date"] = ex_matrix_date
                dossier["meta"]["exam_date"] = ex_exam_date
            st.session_state["current_dossier"] = dossier

            # Xuất file Word (.docx) Toàn bộ Hồ sơ khảo thí (gồm Ma trận + Đặc tả + Đề thi)
            out_dir = os.path.join(CURRENT_DIR, "output")
            os.makedirs(out_dir, exist_ok=True)
            safe_g = to_clean_slug(ex_grade)
            doc_path = os.path.join(out_dir, f"Ho_So_De_Thi_Chuan_{safe_g}.docx")
            ExamDocxExporter.export(dossier, doc_path)
            st.session_state["current_docx_path"] = doc_path

            # Xuất riêng file Word (.docx) Đề thi chính thức (TH01, TH02 + Đáp án) để in ấn phát học sinh
            exam_only_path = os.path.join(out_dir, f"De_Thi_Chinh_Thuc_{safe_g}_TH01_TH02.docx")
            ExamDocxExporter.export_exam_only(dossier, exam_only_path)
            st.session_state["current_exam_only_path"] = exam_only_path

            # Tự động biên soạn và xuất Đề Cương Ôn Tập (Study Guide) đồng bộ
            study_guide = StudyGuideBuilder.build_study_guide(
                dossier=dossier,
                extra_mcq_count=12,
                extra_tf_count=4,
                extra_essay_count=3,
                shuffle=True
            )
            guide_doc_path = os.path.join(out_dir, f"De_Cuong_On_Tap_{safe_g}.docx")
            StudyGuideDocxExporter.export(study_guide, guide_doc_path)
            st.session_state["current_study_guide"] = study_guide
            st.session_state["current_guide_docx_path"] = guide_doc_path

        dossier = st.session_state.get("current_dossier")
        doc_path = st.session_state.get("current_docx_path")
        used_eng = st.session_state.get("used_engine", "Local Offline Engine")

        if dossier:
            # Tự động bù Đề cương ôn tập nếu session cũ đã có dossier
            study_guide = st.session_state.get("current_study_guide")
            guide_doc_path = st.session_state.get("current_guide_docx_path")
            if not study_guide:
                safe_g = to_clean_slug(dossier["meta"]["grade"])
                study_guide = StudyGuideBuilder.build_study_guide(dossier)
                guide_doc_path = os.path.join(CURRENT_DIR, "output", f"De_Cuong_On_Tap_{safe_g}.docx")
                StudyGuideDocxExporter.export(study_guide, guide_doc_path)
                st.session_state["current_study_guide"] = study_guide
                st.session_state["current_guide_docx_path"] = guide_doc_path

            st.success(f"🎉 ĐÃ BIÊN SOẠN THÀNH CÔNG HỒ SƠ ĐỀ THI & ĐỀ CƯƠNG ÔN TẬP CHUẨN ĐẶC TẢ GDPT 2018! (Nguồn: **{used_eng}**)")

            # KPI Tóm tắt
            kpi_e1, kpi_e2, kpi_e3, kpi_e4, kpi_e5 = st.columns(5)
            kpi_e1.metric("Khối lớp", dossier["meta"]["grade"])
            kpi_e2.metric("Ma trận nhận thức", f"{dossier['meta']['pct_nb']}% - {dossier['meta']['pct_th']}% - {dossier['meta']['pct_vd']}%")
            kpi_e3.metric("Số mã đề thi", "02 Mã đề (TH01, TH02)")
            kpi_e4.metric("Câu hỏi đề thi", "19 câu (12 TN + 4 Đ/S + 3 TL)")
            sg_cnt = study_guide.get("counts", {}) if study_guide else {}
            kpi_e5.metric("Câu hỏi ôn tập", f"{sg_cnt.get('total_questions', 38)} câu (Đã xáo trộn)")

            # ----------------- TRUNG TÂM TẢI XUỐNG HỒ SƠ (DOWNLOAD HUB) -----------------
            st.markdown("### 📥 TRUNG TÂM TẢI XUỐNG VĂN BẢN WORD & HỒ SƠ KHẢO THÍ")
            st.info("💡 **Hệ thống cung cấp các tệp Word (.DOCX) chuẩn thể thức Bộ GD&ĐT** để in ấn, phát học sinh hoặc nộp Ban Giám Hiệu. Chọn đúng tệp bạn cần dưới đây:")

            slug_grade = to_clean_slug(dossier.get("meta", {}).get("grade", "lop_9"))

            # --- Chuẩn bị dữ liệu bytes cho các tệp Word ---
            # 1. File Word Đề thi chính thức (TH01, TH02 + Đáp án)
            exam_only_name = f"De_Thi_Chinh_Thuc_{slug_grade}_TH01_TH02.docx"
            exam_only_path = st.session_state.get("current_exam_only_path")
            exam_only_bytes = None
            if exam_only_path and os.path.exists(exam_only_path):
                try:
                    with open(exam_only_path, "rb") as f_ex:
                        exam_only_bytes = f_ex.read()
                except Exception:
                    exam_only_bytes = None
            if not exam_only_bytes and dossier:
                exam_only_path = os.path.join(CURRENT_DIR, "output", exam_only_name)
                ExamDocxExporter.export_exam_only(dossier, exam_only_path)
                with open(exam_only_path, "rb") as f_ex:
                    exam_only_bytes = f_ex.read()

            # 2. File Word Toàn bộ Hồ sơ khảo thí (Ma trận ngang + Đặc tả ngang + Đề thi)
            full_docx_name = f"Ho_So_De_Thi_Chuan_{slug_grade}.docx"
            full_docx_bytes = None
            if doc_path and os.path.exists(doc_path):
                try:
                    with open(doc_path, "rb") as f_doc:
                        full_docx_bytes = f_doc.read()
                except Exception:
                    full_docx_bytes = None
            if not full_docx_bytes and dossier:
                doc_path = os.path.join(CURRENT_DIR, "output", full_docx_name)
                ExamDocxExporter.export(dossier, doc_path)
                with open(doc_path, "rb") as f_doc:
                    full_docx_bytes = f_doc.read()

            # 3. File Word Đề cương ôn tập
            guide_docx_name = f"De_Cuong_On_Tap_{slug_grade}.docx"
            guide_docx_bytes = None
            if guide_doc_path and os.path.exists(guide_doc_path):
                try:
                    with open(guide_doc_path, "rb") as f_sg:
                        guide_docx_bytes = f_sg.read()
                except Exception:
                    guide_docx_bytes = None
            if not guide_docx_bytes and study_guide:
                guide_doc_path = os.path.join(CURRENT_DIR, "output", guide_docx_name)
                StudyGuideDocxExporter.export(study_guide, guide_doc_path)
                with open(guide_doc_path, "rb") as f_sg:
                    guide_docx_bytes = f_sg.read()

            # =========================================================================
            # NHÓM 1: CÁC TỆP WORD (.DOCX) CHUẨN IN ẤN (ƯU TIÊN HÀNG ĐẦU, NỔI BẬT NHẤT)
            # =========================================================================
            st.markdown("##### 📄 NHÓM 1: VĂN BẢN WORD (.DOCX) CHUẨN IN ẤN & PHOTO")
            w_col1, w_col2, w_col3 = st.columns([1.3, 1.3, 1.3])

            with w_col1:
                st.download_button(
                    label="📄 TẢI ĐỀ THI & ĐÁP ÁN (.DOCX)\n(In & Photo Phát Học Sinh Làm Bài)",
                    data=exam_only_bytes if exam_only_bytes else b"",
                    file_name=exam_only_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                    help="Tệp Word khổ dọc gồm Đề thi TH01 + Đề thi TH02 có ô phách, đường cắt kéo 100% không lệch và Hướng dẫn chấm biểu điểm chi tiết."
                )

            with w_col2:
                st.download_button(
                    label="📑 TẢI TOÀN BỘ HỒ SƠ (.DOCX)\n(Gồm Ma Trận, Đặc Tả & Đề Thi Nộp BGH)",
                    data=full_docx_bytes if full_docx_bytes else b"",
                    file_name=full_docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    help="Tệp Word đầy đủ gồm Khung Ma trận ngang + Bản Đặc tả ngang + Đề thi TH01/TH02 và Đáp án (dùng nộp Ban Giám Hiệu duyệt)."
                )

            with w_col3:
                st.download_button(
                    label="📚 TẢI ĐỀ CƯƠNG ÔN TẬP (.DOCX)\n(Phát Học Sinh Ôn Tập Có Dòng Kẻ)",
                    data=guide_docx_bytes if guide_docx_bytes else b"",
                    file_name=guide_docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    help="Tệp Word chuẩn thể thức gồm Khung thông tin học sinh, Hướng dẫn ôn tập, Câu hỏi xáo trộn có dòng kẻ làm bài và Đáp án chi tiết ở trang cuối."
                )

            # --- CÁC NÚT THAO TÁC TRỰC TIẾP TRÊN MÁY TÍNH WINDOWS ---
            action_c1, action_c2 = st.columns([1, 1])
            with action_c1:
                if st.button("🚀 MỞ TRỰC TIẾP FILE WORD BẰNG MICROSOFT WORD", use_container_width=True, type="secondary"):
                    try:
                        target_to_open = exam_only_path if (exam_only_path and os.path.exists(exam_only_path)) else doc_path
                        if target_to_open and os.path.exists(target_to_open):
                            os.startfile(target_to_open)
                            st.toast(f"✅ Đang mở file Word: {os.path.basename(target_to_open)}!")
                    except Exception as e_open:
                        st.error(f"Không thể mở file tự động: {e_open}")

            with action_c2:
                if st.button("📂 MỞ THƯ MỤC OUTPUT CHỨA TẤT CẢ FILE WORD", use_container_width=True):
                    try:
                        out_fld = os.path.join(CURRENT_DIR, "output")
                        if os.path.exists(out_fld):
                            os.startfile(out_fld)
                            st.toast("✅ Đã mở thư mục output trong File Explorer!")
                    except Exception as e_fld:
                        st.error(f"Không thể mở thư mục: {e_fld}")

            # =========================================================================
            # NHÓM 2: TỆP PHỤ TRỢ (DỮ LIỆU SỐ EXCEL & VĂN BẢN THÔ - KHÔNG DÙNG ĐỂ IN WORD)
            # =========================================================================
            st.markdown("##### 🛠️ NHÓM 2: ĐỊNH DẠNG PHỤ TRỢ (EXCEL & TEXT THÔ - KHÔNG PHẢI FILE WORD)")
            sub_col1, sub_col2, sub_col3, sub_col4 = st.columns(4)

            with sub_col1:
                excel_matrix_bytes = ExamAuxiliaryExporter.export_excel_matrix_spec(dossier)
                st.download_button(
                    label="📊 Xuất Ma Trận Excel (.XLSX)",
                    data=excel_matrix_bytes,
                    file_name=f"Ma_Tran_Dac_Ta_{slug_grade}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="File Excel 2 Sheet gồm Khung Ma Trận và Bản Đặc Tả kỹ thuật để nộp Ban Giám Hiệu."
                )

            with sub_col2:
                txt_exam_content = ExamAuxiliaryExporter.export_text_format(dossier, exam_code="th01")
                st.download_button(
                    label="📝 Text Thô Copy Azota (.TXT)",
                    data=txt_exam_content.encode("utf-8"),
                    file_name=f"De_Thi_Text_{slug_grade}_TH01.txt",
                    mime="text/plain",
                    use_container_width=True,
                    help="Định dạng văn bản thô giúp copy nhanh vào phần mềm tạo đề thi trực tuyến Azota, Shub, Google Forms."
                )

            with sub_col3:
                if study_guide:
                    sg_txt_content = StudyGuideDocxExporter.export_text(study_guide)
                    st.download_button(
                        label="📖 Text Đề Cương Gửi Zalo (.TXT)",
                        data=sg_txt_content.encode("utf-8"),
                        file_name=f"De_Cuong_On_Tap_{dossier['meta']['grade'].replace(' ', '_')}.txt",
                        mime="text/plain",
                        use_container_width=True,
                        help="Định dạng văn bản phân đoạn rõ ràng để thầy cô gửi nhanh vào Zalo nhóm phụ huynh/học sinh hoặc đăng lên Azota, OLM."
                    )
                else:
                    st.empty()

            with sub_col4:
                st.download_button(
                    label="📅 Tải KHGD Đã Lọc (.XLSX)",
                    data=sample_khgd_bytes,
                    file_name=f"KHGD_Pham_Vi_Kiem_Tra_{slug_grade}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Kế hoạch giáo dục chứa phạm vi các bài học được sử dụng trong kỳ kiểm tra này."
                )

            st.caption("✔️ **Để in đề thi phát cho học sinh:** Bấm nút đỏ **'📄 TẢI ĐỀ THI & ĐÁP ÁN (.DOCX)'** để mở bằng Microsoft Word và in trực tiếp.")

            st.markdown("---")
            st.subheader("👁️ Xem Trước Chi Tiết Hồ Sơ Khảo Thí & Đề Cương Ôn Tập")

            subtab1, subtab2, subtab3, subtab4, subtab5 = st.tabs([
                "📊 1. Khung Ma Trận Đề Kiểm Tra",
                "📑 2. Bản Đặc Tả Kỹ Thuật",
                "📝 3. Đề Thi Chính Thức (TH01 & TH02)",
                "🔑 4. Hướng Dẫn Chấm & Biểu Điểm",
                "📚 5. Đề Cương Ôn Tập (Study Guide)"
            ])

            # SUBTAB 1: KHUNG MA TRẬN
            with subtab1:
                st.markdown(f"#### KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KỲ - {dossier['meta']['grade'].upper()}")
                st.caption(f"Tỉ lệ phân bổ: **{dossier['meta']['pct_nb']}% Nhận biết - {dossier['meta']['pct_th']}% Thông hiểu - {dossier['meta']['pct_vd']}% Vận dụng** | Điểm tối đa: 10,0 điểm")
                mat_rows = []
                for u in dossier["matrix_rows"]["units"]:
                    mat_rows.append({
                        "TT": u["tt"],
                        "Chương / Chủ đề": u["chapter"],
                        "Nội dung / Đơn vị kiến thức": u["unit"],
                        "MCQ (Biết)": u["mcq_nb"],
                        "MCQ (Hiểu)": u["mcq_th"],
                        "Đúng/Sai (Biết)": u["tf_nb"],
                        "Đúng/Sai (Hiểu)": u["tf_th"],
                        "Đúng/Sai (Vận dụng)": u["tf_vd"],
                        "Tự luận (Hiểu)": u["es_th"],
                        "Tự luận (Vận dụng)": u["es_vd"],
                        "Tổng điểm / Tỷ lệ": u["total_score"]
                    })
                df_mat = pd.DataFrame(mat_rows)
                st.dataframe(df_mat, use_container_width=True, hide_index=True)

                st.markdown("##### 📌 Tổng kết phân bổ theo hình thức:")
                tot = dossier["matrix_rows"]["totals"]
                c_t1, c_t2, c_t3, c_t4 = st.columns(4)
                c_t1.info(f"**Phần I: TN Nhiều lựa chọn**\n\n{tot['total_mcq']}")
                c_t2.info(f"**Phần II: TN Đúng - Sai**\n\n{tot['total_tf']}")
                c_t3.info(f"**Phần III: Tự luận**\n\n{tot['total_es']}")
                c_t4.success(f"**Tổng điểm toàn bài**\n\n{tot['sum_score']}\n\n*({tot['ratio_str']})*")

            # SUBTAB 2: BẢN ĐẶC TẢ
            with subtab2:
                st.markdown(f"#### BẢNG ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KỲ - {dossier['meta']['grade'].upper()}")
                st.caption("Chi tiết Yêu cầu cần đạt chuẩn GDPT 2018 theo từng cấp độ nhận thức")
                for s in dossier["spec_rows"]:
                    with st.expander(f"📍 [{s['tt']}] {s['chapter']} - {s['unit']}", expanded=True):
                        st.markdown(f"**Yêu cầu cần đạt:**\n```\n{s['requirement']}\n```")
                        col_sp1, col_sp2, col_sp3 = st.columns(3)
                        col_sp1.markdown(f"• **TNKQ 4 lựa chọn:** {s['mcq_nb']} Biết | {s['mcq_th']} Hiểu")
                        col_sp2.markdown(f"• **TNKQ Đúng/Sai:** {s['tf_nb'] or 0} Biết | {s['tf_th'] or 0} Hiểu | {s['tf_vd'] or 0} Vận dụng")
                        col_sp3.markdown(f"• **Tự luận:** {s['es_th'] or 0} Hiểu | {s['es_vd'] or 0} Vận dụng")

            # SUBTAB 3: ĐỀ THI CHÍNH THỨC
            with subtab3:
                exam_choice = st.radio("Chọn mã đề xem trước:", ["MÃ ĐỀ TH01 (Đề gốc)", "MÃ ĐỀ TH02 (Đề hoán vị đảo câu & đáp án)"], horizontal=True)
                chosen_exam = dossier["th01"] if "TH01" in exam_choice else dossier["th02"]

                st.markdown(f"### 📋 ĐỀ KIỂM TRA GIỮA HỌC KỲ I - MÃ ĐỀ: {chosen_exam['code']}")
                st.caption(f"{dossier['meta']['school_name']} | Năm học {dossier['meta']['school_year']} | Thời gian làm bài: {dossier['meta']['duration']} phút")

                st.markdown("#### I. PHẦN TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 điểm)")
                for idx, q in enumerate(chosen_exam["mcqs"], 1):
                    st.markdown(f"**Câu {idx}.** {q['content']}")
                    o = q["options"]
                    c_o1, c_o2 = st.columns(2)
                    c_o1.markdown(f"**A.** {o['A']}  \n**C.** {o['C']}")
                    c_o2.markdown(f"**B.** {o['B']}  \n**D.** {o['D']}")

                st.markdown("---")
                st.markdown("#### II. PHẦN TRẮC NGHIỆM ĐÚNG / SAI (4,0 điểm)")
                for idx, tf_it in enumerate(chosen_exam["tfs"], 1):
                    st.markdown(f"**Câu {idx}.** *“{tf_it['topic']}”*")
                    for stm in tf_it["statements"]:
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**{stm['label']})** {stm['content']}")

                st.markdown("---")
                st.markdown("#### III. PHẦN TỰ LUẬN (3,0 điểm)")
                for idx, es_it in enumerate(chosen_exam["essays"], 1):
                    st.markdown(f"**Câu {idx} ({es_it['score']} điểm):** {es_it['content']}")

            # SUBTAB 4: ĐÁP ÁN & BIỂU ĐIỂM
            with subtab4:
                st.markdown("### 🔑 BẢNG ĐÁP ÁN VÀ THANG ĐIỂM CHI TIẾT")
                code_ans = st.selectbox("Chọn mã đề đối chiếu:", ["Mã đề TH01", "Mã đề TH02"])
                cur_ex = dossier["th01"] if "TH01" in code_ans else dossier["th02"]

                st.markdown(f"#### 1. Đáp án Phần I (Nhiều lựa chọn) - Mã đề {cur_ex['code']}")
                ans_cols = [f"Câu {i}" for i in range(1, len(cur_ex["mcqs"]) + 1)]
                ans_vals = [q["correct"] for q in cur_ex["mcqs"]]
                df_ans1 = pd.DataFrame([ans_vals], columns=ans_cols, index=["Đáp án"])
                st.table(df_ans1)

                st.markdown(f"#### 2. Đáp án Phần II (Đúng / Sai) - Mã đề {cur_ex['code']}")
                tf_rows = []
                for idx, tf_q in enumerate(cur_ex["tfs"], 1):
                    for stm in tf_q["statements"]:
                        tf_rows.append({
                            "Câu": f"Câu {idx}",
                            "Lệnh hỏi": stm["label"].upper(),
                            "Đáp án": stm["correct"],
                            "Giải thích / Căn cứ": stm.get("reason", "")
                        })
                st.dataframe(pd.DataFrame(tf_rows), use_container_width=True, hide_index=True)

                st.markdown(f"#### 3. Thang điểm chi tiết Phần III (Tự luận) - Mã đề {cur_ex['code']}")
                for idx, es_q in enumerate(cur_ex["essays"], 1):
                    with st.expander(f"Câu {idx} ({es_q['score']} điểm) - Hướng dẫn chấm", expanded=True):
                        st.markdown(f"**Câu hỏi:** {es_q['content']}")
                        rub_rows = []
                        for r in es_q.get("rubric", []):
                            rub_rows.append({"Yêu cầu cần đạt / Nội dung trả lời": r["criterion"], "Điểm": f"{r['score']}đ"})
                        st.table(pd.DataFrame(rub_rows))

            # SUBTAB 5: ĐỀ CƯƠNG ÔN TẬP ĐỒNG BỘ
            with subtab5:
                if study_guide:
                    st.markdown(f"### 📚 ĐỀ CƯƠNG ÔN TẬP MÔN TIN HỌC {study_guide.get('grade', '').upper()}")
                    st.info(
                        "💡 **Đặc tả sư phạm & Tính năng tự động:**\n\n"
                        "1. **Bao hàm 100% câu hỏi đề thi:** Tất cả các câu hỏi có trong Đề kiểm tra chính thức (TH01/TH02) đều nằm trọn vẹn trong Đề cương này.\n"
                        "2. **Bổ sung câu hỏi rèn luyện:** Tự động lấy thêm các câu hỏi mở rộng cùng chủ đề/bài học từ ngân hàng câu hỏi để học sinh rèn luyện phong phú.\n"
                        "3. **Xáo trộn ngẫu nhiên:** Đã đảo thứ tự tất cả câu hỏi và các phương án A, B, C, D để tránh việc học sinh học vẹt thứ tự.\n"
                        "4. **Chuẩn thể thức:** Đầy đủ thông tin trường lớp, phần ghi họ tên học sinh, có dòng kẻ làm bài tự luận và bảng đáp án chi tiết ở trang cuối."
                    )

                    sg_c = study_guide.get("counts", {})
                    sg_k1, sg_k2, sg_k3, sg_k4 = st.columns(4)
                    sg_k1.metric("Tổng câu hỏi ôn tập", f"{sg_c.get('total_questions', 0)} câu")
                    sg_k2.metric("Trắc nghiệm 4 lựa chọn", f"{sg_c.get('total_mcq', 0)} câu")
                    sg_k3.metric("Trắc nghiệm Đúng / Sai", f"{sg_c.get('total_tf', 0)} câu chùm")
                    sg_k4.metric("Tự luận & Vận dụng", f"{sg_c.get('total_essay', 0)} bài tập")

                    st.markdown("---")
                    st.markdown(f"#### I. PHẦN CÂU HỎI TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN ({len(study_guide.get('mcqs', []))} câu)")
                    st.caption("Em hãy khoanh tròn vào chữ cái đứng trước phương án trả lời đúng nhất:")

                    for q in study_guide.get("mcqs", []):
                        st.markdown(f"**Câu {q.get('guide_num', 1)}.** {q.get('content', '')}")
                        o = q.get("options", {})
                        c1, c2 = st.columns(2)
                        c1.markdown(f"**A.** {o.get('A', '')}  \n**C.** {o.get('C', '')}")
                        c2.markdown(f"**B.** {o.get('B', '')}  \n**D.** {o.get('D', '')}")

                    st.markdown("---")
                    st.markdown(f"#### II. PHẦN CÂU HỎI TRẮC NGHIỆM ĐÚNG / SAI ({len(study_guide.get('tfs', []))} câu)")
                    st.caption("Trong mỗi ý a), b), c), d) ở từng câu, học sinh chọn Đúng hoặc Sai:")

                    for tf in study_guide.get("tfs", []):
                        st.markdown(f"**Câu {tf.get('guide_num', 1)}.** *“{tf.get('topic', '')}”*")
                        for stm in tf.get("statements", []):
                            st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**{stm.get('label', '')})** {stm.get('content', '')}")

                    st.markdown("---")
                    st.markdown(f"#### III. PHẦN CÂU HỎI TỰ LUẬN VÀ BÀI TẬP VẬN DỤNG THỰC HÀNH ({len(study_guide.get('essays', []))} bài)")
                    for es in study_guide.get("essays", []):
                        st.markdown(f"**Bài tập {es.get('guide_num', 1)}:** {es.get('content', '')}")

                    st.markdown("---")
                    with st.expander("🔑 BẢNG ĐÁP ÁN & HƯỚNG DẪN GIẢI CHI TIẾT ĐỀ CƯƠNG ÔN TẬP (Dành cho Giáo viên & Tự học)", expanded=False):
                        st.markdown("##### 1. Đáp án Trắc nghiệm nhiều lựa chọn:")
                        sg_mcqs = study_guide.get("mcqs", [])
                        cols_per_row = 10
                        mcq_chunks = [sg_mcqs[i:i + cols_per_row] for i in range(0, len(sg_mcqs), cols_per_row)]
                        for chunk in mcq_chunks:
                            hdr = [f"Câu {q.get('guide_num')}" for q in chunk]
                            val = [q.get("correct") for q in chunk]
                            st.table(pd.DataFrame([val], columns=hdr, index=["Đáp án"]))

                        st.markdown("##### 2. Đáp án Trắc nghiệm Đúng / Sai:")
                        sg_tf_rows = []
                        for tf in study_guide.get("tfs", []):
                            for stm in tf.get("statements", []):
                                sg_tf_rows.append({
                                    "Câu ôn tập": f"Câu {tf.get('guide_num')}",
                                    "Lệnh hỏi": stm.get("label", "").upper(),
                                    "Đáp án": stm.get("correct", ""),
                                    "Giải thích / Căn cứ": stm.get("reason", "")
                                })
                        st.dataframe(pd.DataFrame(sg_tf_rows), use_container_width=True, hide_index=True)

                        st.markdown("##### 3. Gợi ý trả lời & Thang điểm Tự luận:")
                        for es in study_guide.get("essays", []):
                            st.markdown(f"**Bài tập {es.get('guide_num')}:** {es.get('content')}")
                            rub_data = []
                            for r in es.get("rubric", []):
                                rub_data.append({"Yêu cầu cần đạt / Nội dung trả lời": r.get("criterion", ""), "Biểu điểm gợi ý": f"{r.get('score', 0)}đ"})
                            st.table(pd.DataFrame(rub_data))
                else:
                    st.info("Vui lòng nhấn nút 'Biên Soạn & Xuất Bộ Đề Thi' ở trên để tự động khởi tạo Đề Cương Ôn Tập tương ứng.")


# =========================================================
# GẮN CÁC HÀM VÀO TỪNG TAB VÀ ÁP DỤNG CƠ CHẾ KHÓA DỮ LIỆU
# =========================================================

# TAB 2
with tab2:
    st.header("Quản lý Cơ sở Dữ liệu Đồ thị Neo4j & Trực Quan Hóa Tri Thức Môn Tin Học")
    if strict_mode and not is_live:
        st.error("🔒 **DỮ LIỆU ĐỒ THỊ NEO4J ĐANG BỊ KHÓA!**")
        st.warning(
            "Hệ thống đang ở chế độ **Bắt buộc kết nối Neo4j Live** nhưng bạn chưa đăng nhập hoặc đã nhập SAI mật khẩu!\n\n"
            "- Để đảm bảo an toàn và tính chân thực, hệ thống **TUYỆT ĐỐI KHÔNG ĐỔ BẤT KỲ DỮ LIỆU NÀO** ra màn hình.\n"
            "- **Hướng dẫn mở khóa:** Nhập đúng Mật khẩu **`12345678`** tại thanh bên trái và bấm nút **'🔌 Kết Nối Live'**.\n"
            "- *(Nếu bạn muốn xem dữ liệu tĩnh từ file CSV, hãy chọn 'Chế độ Ngoại tuyến' ở menu bên trái)*."
        )
    else:
        render_tab2_content()

# TAB 3
with tab3:
    st.header("Phân Tích Chương Trình & Phân Bố Tri Thức GDPT 2018")
    if strict_mode and not is_live:
        st.error("🔒 **DỮ LIỆU PHÂN TÍCH ĐANG BỊ KHÓA!**")
        st.warning("Vui lòng kết nối thành công tới Neo4j Live để xem biểu đồ và số liệu phân tích chương trình.")
    else:
        render_tab3_content()

# TAB 4
with tab4:
    st.header("Hỗ Trợ Ra Đề Thi Tự Động Từ Cơ Sở Dữ Liệu Đồ Thị Neo4j")
    if strict_mode and not is_live:
        st.error("🔒 **TÍNH NĂNG SINH ĐỀ THI ĐANG BỊ KHÓA!**")
        st.warning("Hệ thống sinh đề thi trực tiếp từ các câu hỏi trong Neo4j. Vui lòng kết nối Neo4j Live để mở khóa tính năng này.")
    else:
        render_tab4_content()

# =========================================================
# TAB 5: SOẠN KHBD CHUẨN 5512 & NĂNG LỰC SỐ (TT 02/2025)
# =========================================================
def render_tab5_content():
    import zipfile
    st.markdown("""
    <div style="background: linear-gradient(135deg, rgba(238, 242, 255, 0.95) 0%, rgba(224, 231, 255, 0.9) 100%); border: 1px solid #c7d2fe; border-radius: 14px; padding: 18px 24px; margin-bottom: 22px; box-shadow: 0 4px 15px rgba(99, 102, 241, 0.08);">
        <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
            <span style="background: #4338ca; color: white; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: 700;">CHÍNH THỨC</span>
            <span style="color: #4338ca; font-weight: 700; font-size: 0.95rem;">CÔNG VĂN 5512/BGDĐT-GDTrH • THÔNG TƯ 02/2025/TT-BGDĐT</span>
        </div>
        <h3 style="color: #1e1b4b; margin: 0 0 8px 0; font-family: 'Outfit', sans-serif;">Biên Soạn Kế Hoạch Bài Dạy (KHBD) Chuẩn & Tích Hợp Năng Lực Số (NLS)</h3>
        <p style="color: #475569; margin: 0; font-size: 0.95rem; line-height: 1.5;">
            Trợ lý AI chuyên sâu kết nối trực tiếp với <b>Beekone AI Engine</b> (mô hình tối ưu <code>gemini-3.1-flash-lite</code>), 
            tự động đồng bộ với <b>Phân phối chương trình (PPCT)</b>, hỗ trợ soạn thảo <b>đơn lẻ</b> hoặc <b>hàng loạt (Batch)</b> 
            với định dạng Word (.docx) chuẩn 100%: <b>chữ đỏ NLS</b>, <b>Table 0 chống rớt dòng tên GV</b> và <b>khối chữ ký chống tách trang</b>.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Khung cấu hình hành chính
    with st.expander("⚙️ Cấu Hình Thông Tin Hành Chính & Phê Duyệt", expanded=True):
        col_c1, col_c2, col_c3 = st.columns([1.2, 1.2, 1.2])
        with col_c1:
            st.markdown("##### 👨‍🏫 Thông Tin Giáo Viên")
            t_name = st.text_input("Họ và tên giáo viên:", value="Trần Đức Tình", key="tab5_t_name")
            t_dept = st.text_input("Tổ chuyên môn:", value="Tổ Tự Nhiên", key="tab5_t_dept")
            t_school = st.text_input("Đơn vị trường:", value="TRƯỜNG TH & THCS ĐĂK TRÔI", key="tab5_t_school")
        with col_c2:
            st.markdown("##### ✍️ Khối Ký Phê Duyệt")
            s_cm = st.text_input("Duyệt của Chuyên môn:", value="Lưu Văn Chí", key="tab5_s_cm")
            s_to = st.text_input("Duyệt của tổ chuyên môn:", value="Hoàng Thị Kim Tuyết", key="tab5_s_to")
            t_subj = st.selectbox("Môn học & Khối lớp:", ["Tin học 6", "Tin học 7", "Tin học 8", "Tin học 9"], index=0, key="tab5_subj")
        with col_c3:
            st.markdown("##### 📁 Nguồn Phân Phối Chương Trình")
            ppct_default = "_Lop_6A.xlsx"
            ppct_path = st.text_input("Đường dẫn file PPCT Excel:", value=ppct_default, key="tab5_ppct_path")
            uploaded_file = st.file_uploader("Hoặc tải file Excel PPCT mới (.xlsx):", type=["xlsx"], key="tab5_upload_ppct")

    # Xác định đường dẫn file Excel thực tế
    active_ppct_path = ppct_path
    if uploaded_file is not None:
        os.makedirs("output", exist_ok=True)
        temp_excel = os.path.join("output", uploaded_file.name)
        with open(temp_excel, "wb") as f:
            f.write(uploaded_file.getbuffer())
        active_ppct_path = temp_excel

    # Tải danh sách bài dạy
    ai_gen = KHBDAIGenerator()
    try:
        lessons = ai_gen.load_ppct_excel(active_ppct_path)
    except Exception as e:
        st.error(f"❌ Không thể đọc file Phân phối chương trình tại '{active_ppct_path}': {e}")
        return

    # Thống kê nhanh PPCT
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📖 Tổng Tiết PPCT", len(lessons))
    k2.metric("📅 Khối Lớp", t_subj)
    k3.metric("🤖 AI Engine", ai_gen.model.split("/")[-1])
    k4.metric("📁 Thư mục xuất Word", "output/khbd_lop_6a/")

    subtab1, subtab2, subtab3 = st.tabs([
        "🎯 1. Soạn 1 Tiết Cụ Thể (Xem Trước & Tải Về)",
        "⚡ 2. Sinh Hàng Loạt Theo Tuần / Học Kỳ (Batch Generation)",
        "📂 3. Kho Giáo Án Đã Xuất (.docx)"
    ])

    # ----------------- SUBTAB 1: SOẠN 1 TIẾT CỤ THỂ -----------------
    with subtab1:
        st.markdown("#### 🎯 Soạn Thảo Chi Tiết Từng Tiết Học")
        lesson_options = [
            f"Tuần {l['week']} • Tiết {l['period']}: {l['lesson_title']} ({l['date_range']})"
            for l in lessons
        ]
        sel_idx = st.selectbox("Chọn tiết học cần soạn:", range(len(lesson_options)), format_func=lambda i: lesson_options[i])
        cur_lesson = lessons[sel_idx]

        col_b1, col_b2 = st.columns([1.5, 3])
        with col_b1:
            btn_gen_single = st.button("🚀 Soạn Kế Hoạch Bài Dạy Bằng Beekone AI", type="primary", use_container_width=True)

        if btn_gen_single:
            with st.spinner(f"Đang biên soạn bài '{cur_lesson['lesson_title']}' (Tuần {cur_lesson['week']} - Tiết {cur_lesson['period']})..."):
                try:
                    data = ai_gen.generate_khbd_json(
                        lesson_info=cur_lesson,
                        teacher_name=t_name,
                        department=t_dept,
                        signer_cm=s_cm,
                        signer_to=s_to,
                        subject="Tin học",
                        grade="6",
                        school_name=t_school
                    )
                    out_dir = "output/khbd_lop_6a"
                    os.makedirs(out_dir, exist_ok=True)
                    out_docx = os.path.join(out_dir, cur_lesson["output_filename"])
                    KHBDDocxExporter.export_khbd_to_docx(data, out_docx)
                    st.session_state["last_khbd_data"] = data
                    st.session_state["last_khbd_docx"] = out_docx
                    st.success(f"✅ Đã soạn và xuất file Word thành công: **{cur_lesson['output_filename']}**!")
                except Exception as e:
                    st.error(f"Lỗi khi soạn bài: {e}")

        # Xem trước nội dung bài vừa sinh
        if "last_khbd_data" in st.session_state and os.path.exists(st.session_state.get("last_khbd_docx", "")):
            data = st.session_state["last_khbd_data"]
            docx_path = st.session_state["last_khbd_docx"]

            st.markdown("---")
            c_d1, c_d2 = st.columns([2, 1])
            with c_d1:
                st.subheader(f"📄 {data.get('title', 'KẾ HOẠCH BÀI DẠY')}")
                st.caption(data.get("sub_title", ""))
            with c_d2:
                with open(docx_path, "rb") as f:
                    st.download_button(
                        label=f"⬇️ Tải File Word ({os.path.basename(docx_path)})",
                        data=f,
                        file_name=os.path.basename(docx_path),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True
                    )

            # Preview cấu trúc 5 phần
            with st.expander("👁️ Xem chi tiết Mục tiêu & Năng lực số (NLS)", expanded=True):
                obj = data.get("objectives", {})
                st.markdown("**1. Về kiến thức:**")
                for k in obj.get("knowledge", []):
                    st.markdown(f"- {k}")

                st.markdown("**2. Về năng lực:**")
                comp = obj.get("competencies", {})
                for g in comp.get("general", []):
                    st.markdown(f"- *Năng lực chung:* {g}")
                for s in comp.get("specific", []):
                    st.markdown(f"- *Năng lực đặc thù:* {s}")

                st.markdown("<p style='color: #dc2626; font-weight: bold; margin-bottom: 4px;'>2.3. Năng lực số (NLS) tích hợp theo Thông tư 02/2025/TT-BGDĐT:</p>", unsafe_allow_html=True)
                for d in comp.get("digital", []):
                    code = d.get("code", "") if isinstance(d, dict) else ""
                    desc = d.get("desc", str(d)) if isinstance(d, dict) else str(d)
                    st.markdown(f"<p style='color: #dc2626; margin: 2px 0 2px 20px;'>• <b>[{code}]:</b> {desc}</p>", unsafe_allow_html=True)

            with st.expander("👁️ Xem chi tiết 4 Hoạt động Dạy học (Tiến trình 5512)", expanded=False):
                for act in data.get("activities", []):
                    st.markdown(f"##### 📌 {act.get('name')} ({act.get('time', '')})")
                    st.markdown(f"- **Mục tiêu:** {act.get('goal', '')}")
                    st.markdown(f"- **Nội dung:** {act.get('content', '')}")
                    st.markdown(f"- **Sản phẩm:** {act.get('product', '')}")
                    step_data = []
                    for stp in act.get("steps", []):
                        step_data.append({
                            "Các bước": stp.get("step_name", ""),
                            "Hoạt động của GV": stp.get("teacher_action", ""),
                            "Hoạt động của HS": stp.get("student_action", "")
                        })
                    st.table(pd.DataFrame(step_data))

            with st.expander("👁️ Bảng tổng hợp Năng lực số (NLS) tích hợp", expanded=True):
                nls_t = data.get("nls_summary_table", [])
                st.table(pd.DataFrame([
                    {"TT": item.get("stt"), "Nội dung hoạt động số của học sinh": item.get("activity"), "Mã NLS đạt được": item.get("code")}
                    for item in nls_t
                ]))

    # ----------------- SUBTAB 2: SINH HÀNG LOẠT (BATCH) -----------------
    with subtab2:
        st.markdown("#### ⚡ Sinh Hàng Loạt Kế Hoạch Bài Dạy (Batch Generation)")
        st.info("💡 **Tiết kiệm thời gian tối đa:** Tự động tạo giáo án cho nhiều tuần liên tiếp chỉ với 1 cú nhấp chuột. AI Beekone xử lý siêu tốc ~7-9 giây mỗi bài.")

        col_b_cfg1, col_b_cfg2 = st.columns([1.5, 2])
        with col_b_cfg1:
            range_mode = st.radio("Chọn phạm vi sinh:", [
                "Tuỳ chọn khoảng tuần",
                "Học kỳ 1 (Tuần 1 đến 18)",
                "Học kỳ 2 (Tuần 19 đến 35)",
                "Toàn bộ năm học (Tuần 1 đến 35)"
            ])

            if range_mode == "Tuỳ chọn khoảng tuần":
                c_w1, c_w2 = st.columns(2)
                f_w = c_w1.number_input("Từ tuần:", min_value=1, max_value=35, value=1, step=1)
                t_w = c_w2.number_input("Đến tuần:", min_value=1, max_value=35, value=5, step=1)
            elif range_mode == "Học kỳ 1 (Tuần 1 đến 18)":
                f_w, t_w = 1, 18
            elif range_mode == "Học kỳ 2 (Tuần 19 đến 35)":
                f_w, t_w = 19, 35
            else:
                f_w, t_w = 1, 35

        # Lọc danh sách bài học cần sinh
        batch_lessons = []
        for l in lessons:
            try:
                w_num = int(l["week"])
                if f_w <= w_num <= t_w:
                    batch_lessons.append(l)
            except ValueError:
                pass

        with col_b_cfg2:
            st.markdown(f"**Danh sách tiết sẽ biên soạn ({len(batch_lessons)} tiết):**")
            df_preview = pd.DataFrame([{
                "Tuần": l["week"],
                "Tiết": l["period"],
                "Tên bài học": l["lesson_title"],
                "Ngày dạy": l["date_range"]
            } for l in batch_lessons])
            st.dataframe(df_preview, use_container_width=True, height=180, hide_index=True)

        btn_start_batch = st.button("🚀 BẮT ĐẦU SINH HÀNG LOẠT VỚI BEEKONE AI", type="primary", use_container_width=True)

        if btn_start_batch:
            if not batch_lessons:
                st.warning("Không có bài học nào nằm trong khoảng tuần đã chọn.")
            else:
                prog_bar = st.progress(0)
                status_txt = st.empty()
                out_dir = "output/khbd_lop_6a"
                os.makedirs(out_dir, exist_ok=True)
                generated_batch_files = []

                for b_i, b_l in enumerate(batch_lessons):
                    status_txt.markdown(f"⏳ **[{b_i+1}/{len(batch_lessons)}] Đang biên soạn bài:** Tuần {b_l['week']} - Tiết {b_l['period']}: *{b_l['lesson_title']}*...")
                    try:
                        b_data = ai_gen.generate_khbd_json(
                            lesson_info=b_l,
                            teacher_name=t_name,
                            department=t_dept,
                            signer_cm=s_cm,
                            signer_to=s_to,
                            subject="Tin học",
                            grade="6",
                            school_name=t_school
                        )
                        b_out_docx = os.path.join(out_dir, b_l["output_filename"])
                        KHBDDocxExporter.export_khbd_to_docx(b_data, b_out_docx)
                        generated_batch_files.append(b_out_docx)
                    except Exception as err:
                        st.error(f"Lỗi tại bài '{b_l['lesson_title']}': {err}")

                    prog_bar.progress((b_i + 1) / len(batch_lessons))

                status_txt.success(f"🎉 **ĐÃ HOÀN TẤT SINH {len(generated_batch_files)}/{len(batch_lessons)} GIÁO ÁN THÀNH CÔNG!**")

                # Đóng gói ZIP
                zip_filename = f"KHBD_Lop_6A_Tuan_{f_w:02d}_den_{t_w:02d}.zip"
                zip_filepath = os.path.join("output", zip_filename)
                with zipfile.ZipFile(zip_filepath, "w", zipfile.ZIP_DEFLATED) as zf:
                    for f_p in generated_batch_files:
                        zf.write(f_p, os.path.basename(f_p))

                st.session_state["latest_batch_zip"] = zip_filepath

        if "latest_batch_zip" in st.session_state and os.path.exists(st.session_state["latest_batch_zip"]):
            zip_p = st.session_state["latest_batch_zip"]
            with open(zip_p, "rb") as zf:
                st.download_button(
                    label=f"📦 TẢI TRỌN BỘ FILE NÉN ZIP ({os.path.basename(zip_p)} - {os.path.getsize(zip_p)/1024:.1f} KB)",
                    data=zf,
                    file_name=os.path.basename(zip_p),
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

    # ----------------- SUBTAB 3: KHO GIÁO ÁN ĐÃ XUẤT -----------------
    with subtab3:
        st.markdown("#### 📂 Danh Mục Giáo Án Word Đã Xuất Trong Thư Mục `output/khbd_lop_6a/`")
        out_dir = "output/khbd_lop_6a"
        if os.path.exists(out_dir):
            all_files = [f for f in os.listdir(out_dir) if f.endswith(".docx")]
            all_files.sort()
            if all_files:
                st.success(f"Hiện có **{len(all_files)} file giáo án Word (.docx)** sẵn sàng sử dụng:")
                col_f1, col_f2 = st.columns([3, 1])
                for idx_f, f_name in enumerate(all_files, start=1):
                    full_p = os.path.join(out_dir, f_name)
                    sz_kb = os.path.getsize(full_p) / 1024
                    c_a, c_b = st.columns([3, 1])
                    c_a.markdown(f"**{idx_f}.** 📄 `{f_name}` *({sz_kb:.1f} KB)*")
                    with open(full_p, "rb") as f_down:
                        c_b.download_button(
                            label="⬇️ Tải",
                            data=f_down,
                            file_name=f_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"down_btn_{idx_f}"
                        )
            else:
                st.info("Chưa có file giáo án nào trong thư mục. Hãy chọn tab 1 hoặc tab 2 để bắt đầu sinh.")
        else:
            st.info("Thư mục `output/khbd_lop_6a/` chưa được tạo.")

# GẮN HÀM TAB 5
with tab5:
    render_tab5_content()
