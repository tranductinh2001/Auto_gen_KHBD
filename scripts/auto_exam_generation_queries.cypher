// ==============================================================================
// KHO TRUY VẤN NÂNG CAO CHO ĐỒ ÁN CƠ SỞ DỮ LIỆU NÂNG CAO
// Đề tài: "Xây dựng hệ thống quản lý tri thức, phân tích chương trình và hỗ trợ ra đề thi tự động môn Tin học phổ thông GDPT 2018 sử dụng Neo4j"
// ==============================================================================

// ------------------------------------------------------------------------------
// KỊCH BẢN 1: SINH ĐỀ THI TRẮC NGHIỆM TỰ ĐỘNG THEO MA TRẬN CHUẨN SƯ PHẠM
// Yêu cầu: Đề thi gồm 10 câu (4 Nhận biết, 3 Thông hiểu, 3 Vận dụng) từ Khối Lớp 9
// ------------------------------------------------------------------------------
MATCH (l:Lesson {id: 'Lớp 9_' + l.lesson_id})-[:EVALUATED_BY]->(q:Question)
WITH q, q.difficulty AS diff, rand() AS r
ORDER BY r
WITH diff, collect(q) AS questions_by_diff
WITH 
    [q IN questions_by_diff WHERE diff = 'Nhận biết'][0..4] AS nhan_biet,
    [q IN questions_by_diff WHERE diff = 'Thông hiểu'][0..3] AS thong_hieu,
    [q IN questions_by_diff WHERE diff = 'Vận dụng'][0..3] AS van_dung
WITH nhan_biet + thong_hieu + van_dung AS exam_questions
UNWIND exam_questions AS q
MATCH (q)-[:HAS_OPTION]->(opt:AnswerOption)
RETURN 
    q.question_id AS MaCauHoi,
    q.difficulty AS MucDoNhanThuc,
    q.content AS NoiDungCauHoi,
    collect({key: opt.option_key, content: opt.content, is_correct: opt.is_correct}) AS CacDapAn
ORDER BY 
    CASE q.difficulty 
        WHEN 'Nhận biết' THEN 1 
        WHEN 'Thông hiểu' THEN 2 
        WHEN 'Vận dụng' THEN 3 
    END, q.question_id;

// ------------------------------------------------------------------------------
// KỊCH BẢN 2: RA ĐỀ THI TRỰC QUAN GẮN VỚI HÌNH ẢNH / GIAO DIỆN PHẦN MỀM / THUẬT TOÁN
// Tìm bài học có chứa Visual Assets và câu hỏi liên quan để ra đề thi thực hành
// ------------------------------------------------------------------------------
MATCH (l:Lesson)-[:HAS_VISUAL_ASSET]->(v:VisualAsset)
MATCH (l)-[:EVALUATED_BY]->(q:Question)
RETURN 
    l.title AS TenBaiHoc,
    v.code AS MaHinhAnh,
    v.caption AS ChieuThichHinh,
    v.detailed_description AS MoTaGiaoDienThuatToan,
    q.question_id AS CauHoiLienQuan,
    q.difficulty AS DoKho,
    q.content AS NoiDungCauHoi;

// ------------------------------------------------------------------------------
// KỊCH BẢN 3: TRUY VẾT MA TRẬN ĐÁNH GIÁ (REQUIREMENT-BASED ASSESSMENT AUDIT)
// Kiểm tra xem Yêu cầu cần đạt nào đã có câu hỏi kiểm tra, yêu cầu nào chưa có
// ------------------------------------------------------------------------------
MATCH (l:Lesson)-[:HAS_REQUIREMENT]->(r:Requirement)
OPTIONAL MATCH (l)-[:EVALUATED_BY]->(q:Question)
RETURN 
    l.title AS BaiHoc,
    r.code AS MaYeuCau,
    r.content AS YeuCauCanDat,
    count(q) AS SoLuongCauHoiKiemTra,
    collect(q.question_id) AS DanhSachCauHoi;

// ------------------------------------------------------------------------------
// KỊCH BẢN 4: PHÂN TÍCH ĐỘ BAO PHỦ CHƯƠNG TRÌNH (CURRICULUM COVERAGE ANALYSIS)
// Thống kê số lượng Kiến thức, Kỹ năng, Hình ảnh và Câu hỏi theo từng Chủ đề lớn
// ------------------------------------------------------------------------------
MATCH (g:Grade)-[:HAS_CHAPTER]->(c:Chapter)-[:HAS_LESSON]->(l:Lesson)
OPTIONAL MATCH (l)-[:TEACHES_KNOWLEDGE]->(k:Knowledge)
OPTIONAL MATCH (l)-[:DEVELOPS_SKILL]->(s:Skill)
OPTIONAL MATCH (l)-[:HAS_VISUAL_ASSET]->(v:VisualAsset)
OPTIONAL MATCH (l)-[:EVALUATED_BY]->(q:Question)
RETURN 
    g.grade_level AS KhoiLop,
    c.title AS ChuDeLon,
    count(DISTINCT l) AS TongSoBaiHoc,
    count(DISTINCT k) AS TongSoKienThuc,
    count(DISTINCT s) AS TongSoKyNang,
    count(DISTINCT v) AS TongSoVisualAssets,
    count(DISTINCT q) AS TongSoCauHoi
ORDER BY KhoiLop, ChuDeLon;
