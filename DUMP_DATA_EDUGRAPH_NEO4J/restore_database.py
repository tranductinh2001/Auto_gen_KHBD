# -*- coding: utf-8 -*-
"""
restore_database.py - KHÔI PHỤC CSDL ĐỒ THỊ NEO4J 'edugraph' TỰ ĐỘNG 1-CLICK
Đồ án: Cơ sở dữ liệu nâng cao - Học viên: Trần Đức Tình (MSHV: 8291581013)
Trường Đại học Quy Nhơn

Cách chạy:
    python restore_database.py
Hoặc tùy biến tham số:
    python restore_database.py --uri bolt://localhost:7687 --user neo4j --password 12345678 --db edugraph
"""

import os
import sys
import argparse
import pandas as pd
from neo4j import GraphDatabase

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_DIR = os.path.join(CURRENT_DIR, "csv_master_data")

def main():
    parser = argparse.ArgumentParser(description="Khôi phục CSDL Đồ thị Neo4j 'edugraph'")
    parser.add_argument("--uri", default="bolt://localhost:7687", help="Neo4j Bolt URI")
    parser.add_argument("--user", default="neo4j", help="Tên đăng nhập")
    parser.add_argument("--password", default="12345678", help="Mật khẩu")
    parser.add_argument("--db", default="edugraph", help="Tên Database")
    args = parser.parse_args()

    print("=" * 75)
    print(f"🚀 TIẾN HÀNH KHÔI PHỤC DATABASE '{args.db}' TRÊN NEO4J")
    print(f"📍 Kết nối: {args.uri} | User: {args.user}")
    print("=" * 75)

    try:
        driver = GraphDatabase.driver(args.uri, auth=(args.user, args.password))
        driver.verify_connectivity()
        print("✅ Kết nối máy chủ Neo4j thành công!")
    except Exception as e:
        print(f"❌ Không thể kết nối Neo4j: {e}")
        print("👉 Vui lòng mở Neo4j Desktop và bấm Start DBMS, sau đó thử lại.")
        sys.exit(1)

    # Đảm bảo database tồn tại
    try:
        with driver.session(database="system") as sys_s:
            sys_s.run(f"CREATE DATABASE {args.db} IF NOT EXISTS WAIT")
            print(f"✅ Đã đảm bảo Database '{args.db}' sẵn sàng hoạt động.")
    except Exception:
        pass

    with driver.session(database=args.db) as session:
        print("\n[1/6] Dọn dẹp dữ liệu cũ...")
        session.run("MATCH (n) DETACH DELETE n")

        print("[2/6] Tạo Constraints & Indexes...")
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

        # Nạp Lessons
        df_lessons = pd.read_csv(os.path.join(CSV_DIR, "neo4j_lessons.csv")).fillna("")
        print(f"[3/6] Nạp {len(df_lessons)} Bài học & phân cấp sư phạm (Grade -> Chapter -> Lesson)...")
        session.run("""
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
        """, batch=df_lessons.to_dict(orient="records"))

        # Nạp YCCĐ & Tri thức
        df_reqs = pd.read_csv(os.path.join(CSV_DIR, "neo4j_requirements.csv")).fillna("")
        df_know = pd.read_csv(os.path.join(CSV_DIR, "neo4j_knowledge.csv")).fillna("")
        print(f"[4/6] Nạp {len(df_reqs)} Yêu cầu cần đạt và {len(df_know)} Tri thức cốt lõi...")
        session.run("""
        UNWIND $batch AS row
        WITH row WHERE row.requirement_code IS NOT NULL AND trim(toString(row.requirement_code)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (r:Requirement {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.requirement_code)})
        ON CREATE SET r.code = toString(row.requirement_code),
                      r.content = toString(row.requirement_content)
        MERGE (l)-[:HAS_REQUIREMENT]->(r)
        """, batch=df_reqs.to_dict(orient="records"))

        session.run("""
        UNWIND $batch AS row
        WITH row WHERE row.knowledge_name IS NOT NULL AND trim(toString(row.knowledge_name)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (k:Knowledge {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.knowledge_name)})
        ON CREATE SET k.name = toString(row.knowledge_name),
                      k.content = toString(row.knowledge_content)
        MERGE (l)-[:TEACHES_KNOWLEDGE]->(k)
        """, batch=df_know.to_dict(orient="records"))

        # Nạp Visual Assets & Skills
        df_vis = pd.read_csv(os.path.join(CSV_DIR, "neo4j_visual_assets.csv")).fillna("")
        df_skills = pd.read_csv(os.path.join(CSV_DIR, "neo4j_skills.csv")).fillna("")
        print(f"[5/6] Nạp {len(df_vis)} Visual Assets và {len(df_skills)} Kỹ năng...")
        session.run("""
        UNWIND $batch AS row
        WITH row WHERE row.visual_asset_code IS NOT NULL AND trim(toString(row.visual_asset_code)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (v:VisualAsset {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.visual_asset_code)})
        ON CREATE SET v.code = toString(row.visual_asset_code),
                      v.caption = toString(row.visual_caption),
                      v.description = toString(row.visual_description),
                      v.page = toString(row.page_num)
        MERGE (l)-[:HAS_VISUAL_ASSET]->(v)
        """, batch=df_vis.to_dict(orient="records"))

        session.run("""
        UNWIND $batch AS row
        WITH row WHERE row.skill_name IS NOT NULL AND trim(toString(row.skill_name)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (s:Skill {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.skill_name)})
        ON CREATE SET s.name = toString(row.skill_name),
                      s.description = toString(row.skill_description)
        MERGE (l)-[:DEVELOPS_SKILL]->(s)
        """, batch=df_skills.to_dict(orient="records"))

        # Nạp Questions & Answers
        df_questions = pd.read_csv(os.path.join(CSV_DIR, "neo4j_questions.csv")).fillna("")
        df_answers = pd.read_csv(os.path.join(CSV_DIR, "neo4j_answers.csv")).fillna("")
        print(f"[6/6] Nạp {len(df_questions)} Câu hỏi và {len(df_answers)} Phương án lựa chọn...")
        session.run("""
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(toString(row.question_id)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})
        ON CREATE SET q.question_id = toString(row.question_id),
                      q.content = toString(row.question_content),
                      q.difficulty = toString(row.difficulty)
        MERGE (l)-[:EVALUATED_BY]->(q)
        """, batch=df_questions.to_dict(orient="records"))

        session.run("""
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(toString(row.question_id)) <> ''
        MATCH (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})
        CREATE (a:AnswerOption {
            option_key: toString(coalesce(row.answer_option_key, '')),
            content: toString(row.answer_content),
            is_correct: (toString(row.is_correct) = 'True' OR toString(row.is_correct) = 'true' OR toString(row.is_correct) = '1')
        })
        MERGE (q)-[:HAS_OPTION]->(a)
        """, batch=df_answers.to_dict(orient="records"))

        # Tổng kết thống kê
        res_nodes = session.run("MATCH (n) RETURN labels(n)[0] AS label, count(n) AS cnt ORDER BY cnt DESC")
        nodes_stat = {r["label"]: r["cnt"] for r in res_nodes if r["label"]}
        total_n = sum(nodes_stat.values())

        res_rels = session.run("MATCH ()-[r]->() RETURN type(r) AS rel, count(r) AS cnt ORDER BY cnt DESC")
        rels_stat = {r["rel"]: r["cnt"] for r in res_rels if r["rel"]}
        total_r = sum(rels_stat.values())

    print("\n" + "=" * 75)
    print(f"🎉 KHÔI PHỤC THÀNH CÔNG DATABASE '{args.db}' TRÊN NEO4J!")
    print("=" * 75)
    print(f"📊 Tổng số Đỉnh (Nodes): {total_n:,}")
    for lbl, cnt in nodes_stat.items():
        print(f"   • (:{lbl}): {cnt:,}")
    print(f"🔗 Tổng số Quan hệ (Relationships): {total_r:,}")
    for rel, cnt in rels_stat.items():
        print(f"   • [:{rel}]: {cnt:,}")
    print("=" * 75)

if __name__ == "__main__":
    main()
