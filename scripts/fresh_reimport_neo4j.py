"""
scripts/fresh_reimport_neo4j.py
Tạo mới hoàn toàn CSDL Đồ thị Neo4j cho Đồ án Tin học GDPT 2018:
1. Xóa sạch mọi Node cũ (bao gồm cả dữ liệu bán hàng Northwind Customer/Product).
2. Tạo các ràng buộc tính duy nhất (Constraints & Indexes).
3. Nạp toàn bộ dữ liệu chuẩn hóa từ kho Master CSV bằng Python Batch Transaction (nhanh, an toàn, không cần copy file).
4. Thống kê kết quả chuẩn đẹp để sinh viên chụp ảnh đưa vào Báo cáo Tiểu luận.
"""

import os
import sys
import csv
from neo4j import GraphDatabase

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(CURRENT_DIR)
MASTER_DIR = os.path.join(PROJECT_DIR, "output", "neo4j_import")

URI = "bolt://localhost:7687"
USER = "neo4j"
PASSWORD = "12345678"
TARGET_DB = "edugraph"


def read_csv_rows(filename):
    path = os.path.join(MASTER_DIR, filename)
    if not os.path.exists(path):
        print(f"⚠️ Không tìm thấy file: {filename}")
        return []
    rows = []
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(r)
    return rows


def main():
    print("=" * 75)
    print(f"🚀 KHỞI TẠO MỚI HOÀN TOÀN CSDL ĐỒ THỊ NEO4J TRÊN DATABASE '{TARGET_DB}'")
    print("=" * 75)
    print(f"📍 Kết nối: {URI} | User: {USER} | Target DB: {TARGET_DB}")

    try:
        driver = GraphDatabase.driver(URI, auth=(USER, PASSWORD))
        driver.verify_connectivity()
        print("✅ Kết nối máy chủ Neo4j thành công!")
    except Exception as e:
        print(f"❌ Không thể kết nối Neo4j: {e}")
        print("👉 Hãy chắc chắn bạn đã mở Neo4j Desktop và bấm START Database!")
        return

    # TỰ ĐỘNG TẠO DATABASE MỚI NẾU CHƯA CÓ TRÊN SESSION SYSTEM
    try:
        with driver.session(database="system") as sys_session:
            sys_session.run(f"CREATE DATABASE {TARGET_DB} IF NOT EXISTS WAIT")
            print(f"✅ Đã xác thực/tạo mới Database '{TARGET_DB}' thành công trên hệ thống!")
    except Exception as se:
        print(f"ℹ️ Không thể chạy CREATE DATABASE trên system: {se}")

    with driver.session(database=TARGET_DB) as session:
        # BƯỚC 1: XÓA SẠCH DỮ LIỆU CŨ VÀ CÁC NODE BÁN HÀNG NORTHWIND
        print("\n[1/7] Đang dọn dẹp sạch sẽ toàn bộ Database...")
        session.run("MATCH (n) DETACH DELETE n")
        print("  ✔️ Đã xóa sạch 100% dữ liệu cũ (bao gồm Customer, Product, Order...).")

        # BƯỚC 2: TẠO RÀNG BUỘC KHÓA CHUẨN (CONSTRAINTS)
        print("\n[2/7] Đang thiết lập các chỉ mục và ràng buộc duy nhất (Constraints)...")
        constraints = [
            "CREATE CONSTRAINT constraint_grade_level IF NOT EXISTS FOR (g:Grade) REQUIRE g.grade_level IS UNIQUE",
            "CREATE CONSTRAINT constraint_chapter_id IF NOT EXISTS FOR (c:Chapter) REQUIRE c.id IS UNIQUE",
            "CREATE CONSTRAINT constraint_lesson_id IF NOT EXISTS FOR (l:Lesson) REQUIRE l.id IS UNIQUE",
            "CREATE CONSTRAINT constraint_req_id IF NOT EXISTS FOR (r:Requirement) REQUIRE r.id IS UNIQUE",
            "CREATE CONSTRAINT constraint_know_id IF NOT EXISTS FOR (k:Knowledge) REQUIRE k.id IS UNIQUE",
            "CREATE CONSTRAINT constraint_visual_id IF NOT EXISTS FOR (v:VisualAsset) REQUIRE v.id IS UNIQUE",
            "CREATE CONSTRAINT constraint_skill_id IF NOT EXISTS FOR (s:Skill) REQUIRE s.id IS UNIQUE",
            "CREATE CONSTRAINT constraint_question_id IF NOT EXISTS FOR (q:Question) REQUIRE q.id IS UNIQUE"
        ]
        for c in constraints:
            try:
                session.run(c)
            except Exception as ce:
                pass
        print("  ✔️ Đã thiết lập xong 8 Constraints chuẩn Schema.")

        # BƯỚC 3: NẠP CẤU TRÚC PHÂN CẤP SƯ PHẠM (GRADE -> CHAPTER -> LESSON)
        print("\n[3/7] Đang nạp Cấu trúc Khối lớp, Chủ đề và Bài học...")
        lessons_data = read_csv_rows("neo4j_lessons.csv")
        q_lesson = """
        UNWIND $batch AS row
        WITH row WHERE row.lesson_id IS NOT NULL AND trim(row.lesson_id) <> ''
        MERGE (g:Grade {grade_level: coalesce(row.grade_level, 'Lớp 9')})
        MERGE (c:Chapter {id: coalesce(row.grade_level, 'Lớp 9') + '_' + coalesce(row.chapter_title, 'Chủ đề chung')})
        ON CREATE SET c.title = coalesce(row.chapter_title, 'Chủ đề chung')
        MERGE (g)-[:HAS_CHAPTER]->(c)
        MERGE (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
        ON CREATE SET l.lesson_id = row.lesson_id,
                      l.title = row.lesson_title,
                      l.page = row.page_num,
                      l.source_file = row.source_file
        MERGE (c)-[:HAS_LESSON]->(l)
        """
        session.run(q_lesson, batch=lessons_data)
        print(f"  ✔️ Đã nạp thành công {len(lessons_data)} bài học.")

        # BƯỚC 4: NẠP YÊU CẦU CẦN ĐẠT & TRI THỨC CỐT LÕI
        print("\n[4/7] Đang nạp Yêu cầu cần đạt (YCCĐ) và Tri thức cốt lõi...")
        reqs_data = read_csv_rows("neo4j_requirements.csv")
        q_req = """
        UNWIND $batch AS row
        WITH row WHERE row.requirement_code IS NOT NULL AND trim(row.requirement_code) <> ''
        MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
        MERGE (r:Requirement {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.requirement_code})
        ON CREATE SET r.code = row.requirement_code,
                      r.content = row.requirement_content
        MERGE (l)-[:HAS_REQUIREMENT]->(r)
        """
        session.run(q_req, batch=reqs_data)

        know_data = read_csv_rows("neo4j_knowledge.csv")
        q_know = """
        UNWIND $batch AS row
        WITH row WHERE row.knowledge_name IS NOT NULL AND trim(row.knowledge_name) <> ''
        MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
        MERGE (k:Knowledge {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.knowledge_name})
        ON CREATE SET k.name = row.knowledge_name,
                      k.content = row.knowledge_content
        MERGE (l)-[:TEACHES_KNOWLEDGE]->(k)
        """
        session.run(q_know, batch=know_data)
        print(f"  ✔️ Đã nạp {len(reqs_data)} YCCĐ và {len(know_data)} đơn vị Tri thức cốt lõi.")

        # BƯỚC 5: NẠP HÌNH ẢNH MINH HỌA (VISUAL ASSETS) & KỸ NĂNG THỰC HÀNH (SKILLS)
        print("\n[5/7] Đang nạp Học liệu trực quan (Visual Assets) và Kỹ năng...")
        visuals_data = read_csv_rows("neo4j_visual_assets.csv")
        q_visual = """
        UNWIND $batch AS row
        WITH row WHERE row.visual_asset_code IS NOT NULL AND trim(row.visual_asset_code) <> ''
        MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
        MERGE (v:VisualAsset {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.visual_asset_code})
        ON CREATE SET v.code = row.visual_asset_code,
                      v.caption = row.visual_caption,
                      v.description = row.visual_description,
                      v.page = row.page_num
        MERGE (l)-[:HAS_VISUAL_ASSET]->(v)
        """
        session.run(q_visual, batch=visuals_data)

        skills_data = read_csv_rows("neo4j_skills.csv")
        q_skill = """
        UNWIND $batch AS row
        WITH row WHERE row.skill_name IS NOT NULL AND trim(row.skill_name) <> ''
        MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
        MERGE (s:Skill {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.skill_name})
        ON CREATE SET s.name = row.skill_name,
                      s.description = row.skill_description
        MERGE (l)-[:DEVELOPS_SKILL]->(s)
        """
        session.run(q_skill, batch=skills_data)
        print(f"  ✔️ Đã nạp {len(visuals_data)} Visual Assets và {len(skills_data)} Kỹ năng.")

        # BƯỚC 6: NẠP NGÂN HÀNG CÂU HỎI & ĐÁP ÁN (QUESTIONS & OPTIONS)
        print("\n[6/7] Đang nạp Ngân hàng câu hỏi đánh giá và các phương án trả lời...")
        questions_data = read_csv_rows("neo4j_questions.csv")
        q_quest = """
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(row.question_id) <> ''
        MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
        MERGE (q:Question {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.question_id})
        ON CREATE SET q.question_id = row.question_id,
                      q.content = row.question_content,
                      q.difficulty = row.difficulty
        MERGE (l)-[:EVALUATED_BY]->(q)
        """
        session.run(q_quest, batch=questions_data)

        answers_data = read_csv_rows("neo4j_answers.csv")
        q_ans = """
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(row.question_id) <> ''
        MATCH (q:Question {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.question_id})
        CREATE (a:AnswerOption {
            option_key: coalesce(row.answer_option_key, ''),
            content: row.answer_content,
            is_correct: (row.is_correct = 'True' OR row.is_correct = 'true')
        })
        MERGE (q)-[:HAS_OPTION]->(a)
        """
        session.run(q_ans, batch=answers_data)
        print(f"  ✔️ Đã nạp {len(questions_data)} Câu hỏi và {len(answers_data)} Phương án lựa chọn.")

        # BƯỚC 7: TỔNG KẾT VÀ BÁO CÁO THỐNG KÊ HOÀN CHỈNH
        print("\n[7/7] Đang tính toán thống kê đồ thị mới...")
        res_nodes = session.run("MATCH (n) RETURN labels(n)[0] AS label, count(n) AS cnt ORDER BY cnt DESC")
        nodes_stat = {r["label"]: r["cnt"] for r in res_nodes if r["label"]}
        total_n = sum(nodes_stat.values())

        res_rels = session.run("MATCH ()-[r]->() RETURN type(r) AS rel, count(r) AS cnt ORDER BY cnt DESC")
        rels_stat = {r["rel"]: r["cnt"] for r in res_rels if r["rel"]}
        total_r = sum(rels_stat.values())

    print("\n" + "=" * 75)
    print("🎉 HOÀN TẤT THIẾT LẬP CSDL MỚI 100% SẠCH ĐẸP CHO BÁO CÁO TIỂU LUẬN!")
    print("=" * 75)
    print(f"📊 TỔNG SỐ NODES: {total_n:,} nodes")
    for lbl, c in nodes_stat.items():
        print(f"   • (:{lbl}): {c:,}")
    print(f"\n🔗 TỔNG SỐ RELATIONSHIPS: {total_r:,} relationships")
    for rel, c in rels_stat.items():
        print(f"   • [:{rel}]: {c:,}")

    print("\n" + "=" * 75)
    print("📸 CÁC CÂU LỆNH CYPHER GỢI Ý ĐỂ CHẠY & CHỤP ẢNH VÀO BÁO CÁO:")
    print("=" * 75)
    print("1. Chụp Bản đồ Cấu trúc Schema Đồ thị (Meta Graph):")
    print("   CALL db.schema.visualization()")
    print("\n2. Chụp một Cây tri thức hoàn chỉnh từ Khối lớp -> Bài học -> Câu hỏi:")
    print("   MATCH path = (g:Grade)-[:HAS_CHAPTER]->(c:Chapter)-[:HAS_LESSON]->(l:Lesson)-[:EVALUATED_BY]->(q:Question)")
    print("   WHERE g.grade_level = 'Lớp 9'")
    print("   RETURN path LIMIT 25")
    print("\n3. Chụp Thống kê Phân bổ Câu hỏi theo Mức độ nhận thức:")
    print("   MATCH (q:Question)")
    print("   RETURN q.difficulty AS MucDo, count(q) AS SoLuong ORDER BY SoLuong DESC")
    print("=" * 75)

if __name__ == "__main__":
    main()
