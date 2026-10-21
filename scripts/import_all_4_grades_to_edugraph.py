"""
scripts/import_all_4_grades_to_edugraph.py
Nạp trọn bộ dữ liệu 4 khối lớp THCS (Lớp 6, Lớp 7, Lớp 8, Lớp 9) vào Database 'edugraph':
1. Đọc 4 file Excel: sgk_tinhoc_lớp_6.xlsx, sgk_tinhoc_lớp_7.xlsx, sgk_tinhoc_lớp_8.xlsx, sgk_tinhoc_lớp_9.xlsx.
2. Đồng bộ hóa và ghi lại các file CSV Master trong output/neo4j_import/.
3. Xóa sạch và nạp mới toàn bộ 4 khối lớp vào Database Neo4j 'edugraph'.
4. Báo cáo thống kê chi tiết từng khối lớp phục vụ chụp ảnh báo cáo tiểu luận.
"""

import os
import sys
import pandas as pd
from neo4j import GraphDatabase

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")
MASTER_DIR = os.path.join(OUTPUT_DIR, "neo4j_import")
os.makedirs(MASTER_DIR, exist_ok=True)

URI = "bolt://localhost:7687"
USER = "neo4j"
PASSWORD = "12345678"
TARGET_DB = "edugraph"

GRADE_FILES = [
    ("Lớp 6", os.path.join(OUTPUT_DIR, "sgk_tinhoc_lớp_6.xlsx")),
    ("Lớp 7", os.path.join(OUTPUT_DIR, "sgk_tinhoc_lớp_7.xlsx")),
    ("Lớp 8", os.path.join(OUTPUT_DIR, "sgk_tinhoc_lớp_8.xlsx")),
    ("Lớp 9", os.path.join(OUTPUT_DIR, "sgk_tinhoc_lớp_9.xlsx")),
]


def load_all_grades_data():
    all_lessons = []
    all_reqs = []
    all_know = []
    all_skills = []
    all_visuals = []
    all_questions = []
    all_answers = []

    for grade_label, file_path in GRADE_FILES:
        if not os.path.exists(file_path):
            print(f"⚠️ Không tìm thấy file: {file_path}")
            continue
        print(f"📖 Đang đọc dữ liệu {grade_label} từ: {os.path.basename(file_path)}...")
        
        # 1. BaiHoc
        try:
            df_bh = pd.read_excel(file_path, sheet_name="BaiHoc")
            df_bh["grade_level"] = grade_label
            all_lessons.append(df_bh)
        except Exception as e:
            print(f"  Lỗi đọc BaiHoc: {e}")

        # 2. YeuCauCanDat
        try:
            df_yc = pd.read_excel(file_path, sheet_name="YeuCauCanDat")
            df_yc["grade_level"] = grade_label
            all_reqs.append(df_yc)
        except Exception as e:
            print(f"  Lỗi đọc YeuCauCanDat: {e}")

        # 3. KienThuc
        try:
            df_kt = pd.read_excel(file_path, sheet_name="KienThuc")
            df_kt["grade_level"] = grade_label
            all_know.append(df_kt)
        except Exception as e:
            print(f"  Lỗi đọc KienThuc: {e}")

        # 4. KyNang
        try:
            df_kn = pd.read_excel(file_path, sheet_name="KyNang")
            df_kn["grade_level"] = grade_label
            all_skills.append(df_kn)
        except Exception as e:
            print(f"  Lỗi đọc KyNang: {e}")

        # 5. HinhAnh_SoDo_GiaoDien
        try:
            df_vis = pd.read_excel(file_path, sheet_name="HinhAnh_SoDo_GiaoDien")
            df_vis["grade_level"] = grade_label
            all_visuals.append(df_vis)
        except Exception as e:
            print(f"  Lỗi đọc HinhAnh_SoDo_GiaoDien: {e}")

        # 6. CauHoi_DapAn
        try:
            df_ch = pd.read_excel(file_path, sheet_name="CauHoi_DapAn")
            df_ch["grade_level"] = grade_label
            all_answers.append(df_ch)

            # Tách riêng câu hỏi duy nhất
            cols_q = ["grade_level", "lesson_id", "question_id", "question_content", "difficulty"]
            valid_cols_q = [c for c in cols_q if c in df_ch.columns]
            df_q_unique = df_ch[valid_cols_q].drop_duplicates(subset=["grade_level", "lesson_id", "question_id"])
            all_questions.append(df_q_unique)
        except Exception as e:
            print(f"  Lỗi đọc CauHoi_DapAn: {e}")

    # Ghép và khử trùng lặp
    df_merged_lessons = pd.concat(all_lessons, ignore_index=True).drop_duplicates(subset=["grade_level", "lesson_id"]).fillna("")
    df_merged_reqs = pd.concat(all_reqs, ignore_index=True).drop_duplicates(subset=["grade_level", "lesson_id", "requirement_code"]).fillna("")
    df_merged_know = pd.concat(all_know, ignore_index=True).drop_duplicates(subset=["grade_level", "lesson_id", "knowledge_name"]).fillna("")
    df_merged_skills = pd.concat(all_skills, ignore_index=True).drop_duplicates(subset=["grade_level", "lesson_id", "skill_name"]).fillna("")
    df_merged_visuals = pd.concat(all_visuals, ignore_index=True).drop_duplicates(subset=["grade_level", "lesson_id", "visual_asset_code"]).fillna("")
    df_merged_questions = pd.concat(all_questions, ignore_index=True).drop_duplicates(subset=["grade_level", "lesson_id", "question_id"]).fillna("")
    df_merged_answers = pd.concat(all_answers, ignore_index=True).fillna("")

    # Xuất ra thư mục Master CSV
    print("\n💾 Đang ghi đồng bộ vào kho Master CSV (output/neo4j_import/)...")
    df_merged_lessons.to_csv(os.path.join(MASTER_DIR, "neo4j_lessons.csv"), index=False, encoding="utf-8-sig")
    df_merged_reqs.to_csv(os.path.join(MASTER_DIR, "neo4j_requirements.csv"), index=False, encoding="utf-8-sig")
    df_merged_know.to_csv(os.path.join(MASTER_DIR, "neo4j_knowledge.csv"), index=False, encoding="utf-8-sig")
    df_merged_skills.to_csv(os.path.join(MASTER_DIR, "neo4j_skills.csv"), index=False, encoding="utf-8-sig")
    df_merged_visuals.to_csv(os.path.join(MASTER_DIR, "neo4j_visual_assets.csv"), index=False, encoding="utf-8-sig")
    df_merged_questions.to_csv(os.path.join(MASTER_DIR, "neo4j_questions.csv"), index=False, encoding="utf-8-sig")
    df_merged_answers.to_csv(os.path.join(MASTER_DIR, "neo4j_answers.csv"), index=False, encoding="utf-8-sig")
    print("  ✔️ Đã cập nhật xong toàn bộ 7 file Master CSV với đầy đủ 4 khối lớp THCS!")

    return {
        "lessons": df_merged_lessons.to_dict(orient="records"),
        "reqs": df_merged_reqs.to_dict(orient="records"),
        "know": df_merged_know.to_dict(orient="records"),
        "skills": df_merged_skills.to_dict(orient="records"),
        "visuals": df_merged_visuals.to_dict(orient="records"),
        "questions": df_merged_questions.to_dict(orient="records"),
        "answers": df_merged_answers.to_dict(orient="records"),
    }


def import_to_neo4j(data_dict):
    print("\n" + "=" * 75)
    print(f"🚀 TIẾN HÀNH NẠP ĐẦY ĐỦ 4 KHỐI LỚP THCS VÀO DATABASE '{TARGET_DB}'")
    print("=" * 75)

    try:
        driver = GraphDatabase.driver(URI, auth=(USER, PASSWORD))
        driver.verify_connectivity()
    except Exception as e:
        print(f"❌ Không thể kết nối Neo4j: {e}")
        return

    # Đảm bảo database edugraph tồn tại
    try:
        with driver.session(database="system") as sys_s:
            sys_s.run(f"CREATE DATABASE {TARGET_DB} IF NOT EXISTS WAIT")
    except Exception:
        pass

    with driver.session(database=TARGET_DB) as session:
        # Xóa sạch dữ liệu cũ
        print("[1/7] Dọn dẹp Database 'edugraph'...")
        session.run("MATCH (n) DETACH DELETE n")

        # Tạo Constraints
        print("[2/7] Tạo Constraints & Indexes...")
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
            try: session.run(c)
            except Exception: pass

        # Nạp Lessons & Hierarchy (Grade -> Chapter -> Lesson)
        print(f"[3/7] Nạp {len(data_dict['lessons'])} Bài học cho cả 4 khối lớp...")
        q_lesson = """
        UNWIND $batch AS row
        WITH row WHERE row.lesson_id IS NOT NULL AND trim(toString(row.lesson_id)) <> ''
        MERGE (g:Grade {grade_level: toString(row.grade_level)})
        MERGE (c:Chapter {id: toString(row.grade_level) + '_' + toString(row.chapter_title)})
        ON CREATE SET c.title = toString(row.chapter_title)
        MERGE (g)-[:HAS_CHAPTER]->(c)
        MERGE (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        ON CREATE SET l.lesson_id = toString(row.lesson_id),
                      l.title = toString(row.lesson_title),
                      l.page = toString(row.page_num),
                      l.source_file = toString(row.source_file)
        MERGE (c)-[:HAS_LESSON]->(l)
        """
        session.run(q_lesson, batch=data_dict["lessons"])

        # Nạp Requirements & Knowledge
        print(f"[4/7] Nạp {len(data_dict['reqs'])} YCCĐ và {len(data_dict['know'])} Tri thức cốt lõi...")
        q_req = """
        UNWIND $batch AS row
        WITH row WHERE row.requirement_code IS NOT NULL AND trim(toString(row.requirement_code)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (r:Requirement {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.requirement_code)})
        ON CREATE SET r.code = toString(row.requirement_code),
                      r.content = toString(row.requirement_content)
        MERGE (l)-[:HAS_REQUIREMENT]->(r)
        """
        session.run(q_req, batch=data_dict["reqs"])

        q_know = """
        UNWIND $batch AS row
        WITH row WHERE row.knowledge_name IS NOT NULL AND trim(toString(row.knowledge_name)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (k:Knowledge {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.knowledge_name)})
        ON CREATE SET k.name = toString(row.knowledge_name),
                      k.content = toString(row.knowledge_content)
        MERGE (l)-[:TEACHES_KNOWLEDGE]->(k)
        """
        session.run(q_know, batch=data_dict["know"])

        # Nạp Visual Assets & Skills
        print(f"[5/7] Nạp {len(data_dict['visuals'])} Visual Assets và {len(data_dict['skills'])} Kỹ năng...")
        q_vis = """
        UNWIND $batch AS row
        WITH row WHERE row.visual_asset_code IS NOT NULL AND trim(toString(row.visual_asset_code)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (v:VisualAsset {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.visual_asset_code)})
        ON CREATE SET v.code = toString(row.visual_asset_code),
                      v.caption = toString(row.visual_caption),
                      v.description = toString(row.visual_description),
                      v.page = toString(row.page_num)
        MERGE (l)-[:HAS_VISUAL_ASSET]->(v)
        """
        session.run(q_vis, batch=data_dict["visuals"])

        q_skill = """
        UNWIND $batch AS row
        WITH row WHERE row.skill_name IS NOT NULL AND trim(toString(row.skill_name)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (s:Skill {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.skill_name)})
        ON CREATE SET s.name = toString(row.skill_name),
                      s.description = toString(row.skill_description)
        MERGE (l)-[:DEVELOPS_SKILL]->(s)
        """
        session.run(q_skill, batch=data_dict["skills"])

        # Nạp Questions & Answer Options
        print(f"[6/7] Nạp {len(data_dict['questions'])} Câu hỏi và {len(data_dict['answers'])} Phương án lựa chọn...")
        q_quest = """
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(toString(row.question_id)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})
        ON CREATE SET q.question_id = toString(row.question_id),
                      q.content = toString(row.question_content),
                      q.difficulty = toString(row.difficulty)
        MERGE (l)-[:EVALUATED_BY]->(q)
        """
        session.run(q_quest, batch=data_dict["questions"])

        q_ans = """
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(toString(row.question_id)) <> ''
        MATCH (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})
        CREATE (a:AnswerOption {
            option_key: toString(coalesce(row.answer_option_key, '')),
            content: toString(row.answer_content),
            is_correct: (toString(row.is_correct) = 'True' OR toString(row.is_correct) = 'true' OR toString(row.is_correct) = '1')
        })
        MERGE (q)-[:HAS_OPTION]->(a)
        """
        session.run(q_ans, batch=data_dict["answers"])

        # Thống kê tổng kết
        print("\n[7/7] Đang lấy thống kê phân bổ theo 4 khối lớp...")
        res_grades = session.run("MATCH (g:Grade)-[:HAS_CHAPTER]->(c)-[:HAS_LESSON]->(l) RETURN g.grade_level AS grade, count(DISTINCT l) AS total_lessons, count(DISTINCT c) AS total_chapters ORDER BY grade").data()
        
        res_nodes = session.run("MATCH (n) RETURN labels(n)[0] AS label, count(n) AS cnt ORDER BY cnt DESC")
        nodes_stat = {r["label"]: r["cnt"] for r in res_nodes if r["label"]}
        total_n = sum(nodes_stat.values())

        res_rels = session.run("MATCH ()-[r]->() RETURN type(r) AS rel, count(r) AS cnt ORDER BY cnt DESC")
        rels_stat = {r["rel"]: r["cnt"] for r in res_rels if r["rel"]}
        total_r = sum(rels_stat.values())

    print("\n" + "=" * 75)
    print("🎉 HOÀN TẤT NẠP TOÀN DIỆN CẢ 4 KHỐI LỚP THCS VÀO DATABASE 'edugraph'!")
    print("=" * 75)
    print("📚 PHÂN BỔ THEO TỪNG KHỐI LỚP:")
    for row in res_grades:
        print(f"   • {row['grade']}: {row['total_chapters']} Chủ đề | {row['total_lessons']} Bài học")

    print(f"\n📊 TỔNG CỘNG NODES TRÊN TOÀN HỆ THỐNG: {total_n:,} nodes")
    for lbl, c in nodes_stat.items():
        print(f"   • (:{lbl}): {c:,}")
    print(f"\n🔗 TỔNG CỘNG RELATIONSHIPS: {total_r:,} relationships")
    for rel, c in rels_stat.items():
        print(f"   • [:{rel}]: {c:,}")
    print("=" * 75)


def main():
    data = load_all_grades_data()
    import_to_neo4j(data)


if __name__ == "__main__":
    main()
