import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import os, sys

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter
from scratch.test_user_mcqs import user_mcqs

def est_mcq_pt(q):
    content = q.get('content', '')
    stem_len = len(content) + 12
    stem_lines = max(1, (stem_len + 84) // 85)
    stem_pt = stem_lines * 11.55 + 2.0

    opts = q.get('options', {})
    opt_keys = ['A', 'B', 'C', 'D']
    max_opt_len = max((len(str(opts.get(k, ''))) for k in opt_keys), default=0)

    if max_opt_len <= 14:
        opt_pt = 11.55 + 1.2
    elif max_opt_len <= 56:
        opt_pt = 2 * 11.55 + 2 * 1.2
    else:
        opt_lines = sum(max(1, (len(str(opts.get(k, ''))) + 80) // 80) for k in opt_keys)
        opt_pt = opt_lines * 11.55 + 4 * 1.2

    return stem_pt + opt_pt

def est_tf_pt(tf_item):
    top_lines = max(1, (len(tf_item.get('topic', '')) + 15) // 85)
    top_pt = top_lines * 11.55 + 2.0
    stmts_pt = 0.0
    for stmt in tf_item.get('statements', []):
        s_lines = max(1, (len(stmt.get('content', '')) + 8) // 85)
        stmts_pt += s_lines * 11.55 + 1.0
    return top_pt + stmts_pt

def find_optimal_page1_count(mcqs, tfs):
    """
    Tìm số câu đưa lên Trang 1 sao cho:
    1. Trang 1 chứa tối đa các câu vừa vặn mà không bị tràn (ngân sách 535 pt).
    2. Các câu còn lại + Phần II trên Trang 2 cũng vừa vặn mà không bị tràn (ngân sách 740 pt).
    3. Giảm thiểu khoảng trống thừa ở cuối Trang 1 và Trang 2.
    """
    total_mcqs = len(mcqs)
    # Chiều cao cố định trên Trang 2: X-box (93) + Tiêu đề Phần II (16) + Bảng 12 cột (65) + 4 câu Đúng/Sai
    tf_total_pt = sum(est_tf_pt(tf) for tf in tfs)
    fixed_p2_pt = 93.0 + 16.0 + 65.0 + tf_total_pt

    best_k = 7
    best_balance_diff = 9999.0

    for k in range(min(total_mcqs, 10), 4, -1):
        p1_mcqs_pt = sum(est_mcq_pt(q) for q in mcqs[:k])
        p2_mcqs_pt = sum(est_mcq_pt(q) for q in mcqs[k:])
        p2_total_pt = fixed_p2_pt + p2_mcqs_pt

        # Không bao giờ được tràn trang (A4 printable = 756 pt)
        p1_full_pt = 208.0 + p1_mcqs_pt
        if p1_full_pt <= 735.0 and p2_total_pt <= 735.0:
            diff = abs(p1_full_pt - p2_total_pt)
            print(f"  Candidate k={k}: P1={p1_full_pt:.1f} pt, P2={p2_total_pt:.1f} pt (diff={diff:.1f})")
            if diff < best_balance_diff:
                best_balance_diff = diff
                best_k = k

    return best_k

# Test with User short exam
engine = ExamMatrixEngine(grade_level='Lớp 9')
dossier = engine.generate_exam_dossier()
tfs_std = dossier['th01']['tfs']

# User short tfs from inspect
user_tfs = [
    {"topic": "Sử dụng thiết bị số trong đời sống", "statements": [
        {"label": "a", "content": "Thiết bị số giúp con người kết nối nhanh hơn."},
        {"label": "b", "content": "Chỉ người lớn mới cần sử dụng thiết bị số."},
        {"label": "c", "content": "Thiết bị số thông minh có khả năng tự động hóa."},
        {"label": "d", "content": "Sử dụng thiết bị số không gây tác động đến môi trường."}
    ]},
    {"topic": "Đánh giá thông tin trên mạng", "statements": [
        {"label": "a", "content": "Mọi thông tin trên Internet đều chính xác."},
        {"label": "b", "content": "Cần kiểm tra tác giả của bài viết."},
        {"label": "c", "content": "Tin giả thường có mục đích lừa đảo."},
        {"label": "d", "content": "Không nên chia sẻ thông tin khi chưa rõ nguồn gốc."}
    ]},
    {"topic": "Pháp luật và Internet", "statements": [
        {"label": "a", "content": "Luật an ninh mạng bảo vệ người dùng."},
        {"label": "b", "content": "Tự ý lấy ảnh người khác đăng lên mạng là đúng."},
        {"label": "c", "content": "Bản quyền số bảo vệ quyền lợi tác giả."},
        {"label": "d", "content": "Internet là môi trường không có luật lệ."}
    ]},
    {"topic": "Phần mềm mô phỏng", "statements": [
        {"label": "a", "content": "Mô phỏng giúp tiết kiệm chi phí thực nghiệm."},
        {"label": "b", "content": "Phần mềm mô phỏng luôn cho kết quả sai."},
        {"label": "c", "content": "Mô phỏng giúp quan sát các hiện tượng khó."},
        {"label": "d", "content": "Chỉ nhà khoa học mới được dùng phần mềm mô phỏng."}
    ]}
]

print("=== User Short Exam ===")
k_user_th01 = find_optimal_page1_count(user_mcqs[:12], user_tfs)
print(f"Optimal k for User TH01: {k_user_th01}")

print("\n=== Standard Detailed Exam ===")
k_std_th01 = find_optimal_page1_count(dossier['th01']['mcqs'], tfs_std)
print(f"Optimal k for Standard TH01: {k_std_th01}")
k_std_th02 = find_optimal_page1_count(dossier['th02']['mcqs'], tfs_std)
print(f"Optimal k for Standard TH02: {k_std_th02}")
