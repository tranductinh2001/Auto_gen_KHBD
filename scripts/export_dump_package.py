import os
import sys
import json
import shutil
import zipfile
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")
MASTER_CSV_DIR = os.path.join(OUTPUT_DIR, "neo4j_import")
DUMP_DIR = os.path.join(PROJECT_DIR, "DUMP_DATA_EDUGRAPH_NEO4J")

os.makedirs(DUMP_DIR, exist_ok=True)
CSV_DEST_DIR = os.path.join(DUMP_DIR, "csv_master_data")
os.makedirs(CSV_DEST_DIR, exist_ok=True)

print("=" * 75)
print("📦 BẮT ĐẦU TẠO GÓI DUMP DATA CHO ĐỀ ÁN CSDL NÂNG CAO - NEO4J")
print("=" * 75)

# 1. Copy các file CSV Master
csv_files = [
    "neo4j_lessons.csv",
    "neo4j_requirements.csv",
    "neo4j_knowledge.csv",
    "neo4j_skills.csv",
    "neo4j_visual_assets.csv",
    "neo4j_questions.csv",
    "neo4j_answers.csv"
]

print("\n[1/5] Sao chép các file CSV Master...")
for f in csv_files:
    src = os.path.join(MASTER_CSV_DIR, f)
    dst = os.path.join(CSV_DEST_DIR, f)
    if os.path.exists(src):
        shutil.copy2(src, dst)
        df = pd.read_csv(src)
        print(f"  ✓ {f:<26}: {len(df):>5} dòng dữ liệu")
    else:
        print(f"  ⚠️ Thiếu file: {f}")

# Copy file Excel master nếu có
excel_master = os.path.join(OUTPUT_DIR, "sgk_tinhoc_tong_hop_master.xlsx")
if os.path.exists(excel_master):
    shutil.copy2(excel_master, os.path.join(DUMP_DIR, "sgk_tinhoc_tong_hop_master.xlsx"))
    print("  ✓ Đã copy file Excel tổng hợp: sgk_tinhoc_tong_hop_master.xlsx")

# 2. Sinh file edugraph_full_dump.cypher độc lập (Self-Contained Cypher Script)
print("\n[2/5] Đang tạo file Dump Cypher độc lập (edugraph_full_dump.cypher)...")
cypher_dump_path = os.path.join(DUMP_DIR, "edugraph_full_dump.cypher")

def escape_cypher_str(val):
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return s.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"').replace("\n", "\\n").replace("\r", "")

with open(cypher_dump_path, "w", encoding="utf-8") as f_out:
    f_out.write("// ==============================================================================\n")
    f_out.write("// NEO4J FULL DATABASE DUMP - EDŮGRAPH (TIN HỌC GDPT 2018)\n")
    f_out.write("// Đồ án: Cơ sở dữ liệu nâng cao - Học viên: Trần Đức Tình (MSHV: 8291581013)\n")
    f_out.write("// Trường Đại học Quy Nhơn - Khóa Thạc sĩ KHMT\n")
    f_out.write("// File DDL & DML độc lập (Self-contained), có thể thực thi trực tiếp trên Neo4j Browser,\n")
    f_out.write("// cypher-shell hoặc Neo4j AuraDB mà KHÔNG cần đặt file CSV vào thư mục import.\n")
    f_out.write("// ==============================================================================\n\n")

    # DDL: Constraints
    f_out.write("// --- 1. TẠO RÀNG BUỘC TOÀN VẸN (CONSTRAINTS & INDEXES) ---\n")
    constraints = [
        "CREATE CONSTRAINT constraint_grade_level IF NOT EXISTS FOR (g:Grade) REQUIRE g.grade_level IS UNIQUE;",
        "CREATE CONSTRAINT constraint_chapter_id IF NOT EXISTS FOR (c:Chapter) REQUIRE c.id IS UNIQUE;",
        "CREATE CONSTRAINT constraint_lesson_id IF NOT EXISTS FOR (l:Lesson) REQUIRE l.id IS UNIQUE;",
        "CREATE CONSTRAINT constraint_req_id IF NOT EXISTS FOR (r:Requirement) REQUIRE r.id IS UNIQUE;",
        "CREATE CONSTRAINT constraint_know_id IF NOT EXISTS FOR (k:Knowledge) REQUIRE k.id IS UNIQUE;",
        "CREATE CONSTRAINT constraint_visual_id IF NOT EXISTS FOR (v:VisualAsset) REQUIRE v.id IS UNIQUE;",
        "CREATE CONSTRAINT constraint_skill_id IF NOT EXISTS FOR (s:Skill) REQUIRE s.id IS UNIQUE;",
        "CREATE CONSTRAINT constraint_question_id IF NOT EXISTS FOR (q:Question) REQUIRE q.id IS UNIQUE;"
    ]
    for c in constraints:
        f_out.write(c + "\n")
    f_out.write("\n")

    # DML 1: Lessons & Hierarchy
    df_lessons = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_lessons.csv")).fillna("")
    f_out.write(f"// --- 2. NẠP CẤU TRÚC PHÂN CẤP SƯ PHẠM ({len(df_lessons)} BÀI HỌC) ---\n")
    batch_lessons = []
    for _, r in df_lessons.iterrows():
        batch_lessons.append({
            "grade_level": str(r.get("grade_level", "")),
            "chapter_title": str(r.get("chapter_title", "")),
            "lesson_id": str(r.get("lesson_id", "")),
            "lesson_title": str(r.get("lesson_title", "")),
            "page_num": str(r.get("page_num", "")),
            "source_file": str(r.get("source_file", ""))
        })
    f_out.write("UNWIND " + json.dumps(batch_lessons, ensure_ascii=False) + " AS row\n")
    f_out.write("MERGE (g:Grade {grade_level: toString(row.grade_level)})\n")
    f_out.write("MERGE (c:Chapter {id: toString(row.grade_level) + '_' + toString(row.chapter_title)})\n")
    f_out.write("ON CREATE SET c.title = toString(row.chapter_title)\n")
    f_out.write("MERGE (g)-[:HAS_CHAPTER]->(c)\n")
    f_out.write("MERGE (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})\n")
    f_out.write("ON CREATE SET l.lesson_id = toString(row.lesson_id),\n")
    f_out.write("              l.title = toString(row.lesson_title),\n")
    f_out.write("              l.page = toString(row.page_num),\n")
    f_out.write("              l.source_file = toString(row.source_file)\n")
    f_out.write("MERGE (c)-[:HAS_LESSON]->(l);\n\n")

    # DML 2: Requirements
    df_reqs = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_requirements.csv")).fillna("")
    f_out.write(f"// --- 3. NẠP YÊU CẦU CẦN ĐẠT ({len(df_reqs)} YCCĐ) ---\n")
    batch_reqs = []
    for _, r in df_reqs.iterrows():
        batch_reqs.append({
            "grade_level": str(r.get("grade_level", "")),
            "lesson_id": str(r.get("lesson_id", "")),
            "requirement_code": str(r.get("requirement_code", "")),
            "requirement_content": str(r.get("requirement_content", ""))
        })
    f_out.write("UNWIND " + json.dumps(batch_reqs, ensure_ascii=False) + " AS row\n")
    f_out.write("MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})\n")
    f_out.write("MERGE (r:Requirement {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.requirement_code)})\n")
    f_out.write("ON CREATE SET r.code = toString(row.requirement_code),\n")
    f_out.write("              r.content = toString(row.requirement_content)\n")
    f_out.write("MERGE (l)-[:HAS_REQUIREMENT]->(r);\n\n")

    # DML 3: Knowledge
    df_know = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_knowledge.csv")).fillna("")
    f_out.write(f"// --- 4. NẠP NỘI DUNG KIẾN THỨC CỐT LÕI ({len(df_know)} TRI THỨC) ---\n")
    batch_know = []
    for _, r in df_know.iterrows():
        batch_know.append({
            "grade_level": str(r.get("grade_level", "")),
            "lesson_id": str(r.get("lesson_id", "")),
            "knowledge_name": str(r.get("knowledge_name", "")),
            "knowledge_content": str(r.get("knowledge_content", ""))
        })
    f_out.write("UNWIND " + json.dumps(batch_know, ensure_ascii=False) + " AS row\n")
    f_out.write("MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})\n")
    f_out.write("MERGE (k:Knowledge {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.knowledge_name)})\n")
    f_out.write("ON CREATE SET k.name = toString(row.knowledge_name),\n")
    f_out.write("              k.content = toString(row.knowledge_content)\n")
    f_out.write("MERGE (l)-[:TEACHES_KNOWLEDGE]->(k);\n\n")

    # DML 4: Visual Assets
    df_vis = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_visual_assets.csv")).fillna("")
    f_out.write(f"// --- 5. NẠP HÌNH ẢNH, SƠ ĐỒ VÀ GIAO DIỆN ({len(df_vis)} VISUAL ASSETS) ---\n")
    batch_vis = []
    for _, r in df_vis.iterrows():
        batch_vis.append({
            "grade_level": str(r.get("grade_level", "")),
            "lesson_id": str(r.get("lesson_id", "")),
            "visual_asset_code": str(r.get("visual_asset_code", "")),
            "visual_caption": str(r.get("visual_caption", "")),
            "visual_description": str(r.get("visual_description", "")),
            "page_num": str(r.get("page_num", ""))
        })
    f_out.write("UNWIND " + json.dumps(batch_vis, ensure_ascii=False) + " AS row\n")
    f_out.write("MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})\n")
    f_out.write("MERGE (v:VisualAsset {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.visual_asset_code)})\n")
    f_out.write("ON CREATE SET v.code = toString(row.visual_asset_code),\n")
    f_out.write("              v.caption = toString(row.visual_caption),\n")
    f_out.write("              v.description = toString(row.visual_description),\n")
    f_out.write("              v.page = toString(row.page_num)\n")
    f_out.write("MERGE (l)-[:HAS_VISUAL_ASSET]->(v);\n\n")

    # DML 5: Skills
    df_skills = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_skills.csv")).fillna("")
    f_out.write(f"// --- 6. NẠP KỸ NĂNG THỰC HÀNH ({len(df_skills)} KỸ NĂNG) ---\n")
    batch_skills = []
    for _, r in df_skills.iterrows():
        batch_skills.append({
            "grade_level": str(r.get("grade_level", "")),
            "lesson_id": str(r.get("lesson_id", "")),
            "skill_name": str(r.get("skill_name", "")),
            "skill_description": str(r.get("skill_description", ""))
        })
    f_out.write("UNWIND " + json.dumps(batch_skills, ensure_ascii=False) + " AS row\n")
    f_out.write("MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})\n")
    f_out.write("MERGE (s:Skill {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.skill_name)})\n")
    f_out.write("ON CREATE SET s.name = toString(row.skill_name),\n")
    f_out.write("              s.description = toString(row.skill_description)\n")
    f_out.write("MERGE (l)-[:DEVELOPS_SKILL]->(s);\n\n")

    # DML 6: Questions
    df_questions = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_questions.csv")).fillna("")
    f_out.write(f"// --- 7. NẠP CÂU HỎI ĐÁNH GIÁ ({len(df_questions)} CÂU HỎI) ---\n")
    batch_questions = []
    for _, r in df_questions.iterrows():
        batch_questions.append({
            "grade_level": str(r.get("grade_level", "")),
            "lesson_id": str(r.get("lesson_id", "")),
            "question_id": str(r.get("question_id", "")),
            "question_content": str(r.get("question_content", "")),
            "difficulty": str(r.get("difficulty", ""))
        })
    f_out.write("UNWIND " + json.dumps(batch_questions, ensure_ascii=False) + " AS row\n")
    f_out.write("MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})\n")
    f_out.write("MERGE (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})\n")
    f_out.write("ON CREATE SET q.question_id = toString(row.question_id),\n")
    f_out.write("              q.content = toString(row.question_content),\n")
    f_out.write("              q.difficulty = toString(row.difficulty)\n")
    f_out.write("MERGE (l)-[:EVALUATED_BY]->(q);\n\n")

    # DML 7: Answers
    df_answers = pd.read_csv(os.path.join(CSV_DEST_DIR, "neo4j_answers.csv")).fillna("")
    f_out.write(f"// --- 8. NẠP PHƯƠNG ÁN LỰA CHỌN ({len(df_answers)} ĐÁP ÁN) ---\n")
    # Batch answers in groups of 200 for safe execution
    batch_size = 200
    for chunk_idx in range(0, len(df_answers), batch_size):
        chunk = df_answers.iloc[chunk_idx:chunk_idx + batch_size]
        batch_ans = []
        for _, r in chunk.iterrows():
            is_cor = str(r.get("is_correct", "")).strip().lower() in ["true", "1"]
            batch_ans.append({
                "grade_level": str(r.get("grade_level", "")),
                "lesson_id": str(r.get("lesson_id", "")),
                "question_id": str(r.get("question_id", "")),
                "option_key": str(r.get("answer_option_key", "")),
                "content": str(r.get("answer_content", "")),
                "is_correct": is_cor
            })
        f_out.write(f"// Batch {chunk_idx // batch_size + 1} ({len(batch_ans)} items)\n")
        f_out.write("UNWIND " + json.dumps(batch_ans, ensure_ascii=False) + " AS row\n")
        f_out.write("MATCH (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})\n")
        f_out.write("CREATE (a:AnswerOption {\n")
        f_out.write("    option_key: toString(row.option_key),\n")
        f_out.write("    content: toString(row.content),\n")
        f_out.write("    is_correct: row.is_correct\n")
        f_out.write("})\n")
        f_out.write("MERGE (q)-[:HAS_OPTION]->(a);\n\n")

print(f"  ✓ Đã sinh thành công: {os.path.basename(cypher_dump_path)} ({os.path.getsize(cypher_dump_path):,} bytes)")

# 3. Tạo script LOAD CSV truyền thống
print("\n[3/5] Tạo file Cypher dạng LOAD CSV (import_via_load_csv.cypher)...")
load_csv_path = os.path.join(DUMP_DIR, "import_via_load_csv.cypher")
shutil.copy2(os.path.join(MASTER_CSV_DIR, "neo4j_import_script.cypher"), load_csv_path)
print("  ✓ Đã lưu import_via_load_csv.cypher")

# 4. Tạo script Python Restore 1-Click
print("\n[4/5] Tạo script Python Restore 1-Click (restore_database.py)...")
restore_py_content = """# -*- coding: utf-8 -*-
\"\"\"
restore_database.py - KHÔI PHỤC CSDL ĐỒ THỊ NEO4J 'edugraph' TỰ ĐỘNG 1-CLICK
Đồ án: Cơ sở dữ liệu nâng cao - Học viên: Trần Đức Tình (MSHV: 8291581013)
Trường Đại học Quy Nhơn

Cách chạy:
    python restore_database.py
Hoặc tùy biến tham số:
    python restore_database.py --uri bolt://localhost:7687 --user neo4j --password 12345678 --db edugraph
\"\"\"

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
        print("\\n[1/6] Dọn dẹp dữ liệu cũ...")
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
        session.run(\"\"\"
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
        \"\"\", batch=df_lessons.to_dict(orient="records"))

        # Nạp YCCĐ & Tri thức
        df_reqs = pd.read_csv(os.path.join(CSV_DIR, "neo4j_requirements.csv")).fillna("")
        df_know = pd.read_csv(os.path.join(CSV_DIR, "neo4j_knowledge.csv")).fillna("")
        print(f"[4/6] Nạp {len(df_reqs)} Yêu cầu cần đạt và {len(df_know)} Tri thức cốt lõi...")
        session.run(\"\"\"
        UNWIND $batch AS row
        WITH row WHERE row.requirement_code IS NOT NULL AND trim(toString(row.requirement_code)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (r:Requirement {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.requirement_code)})
        ON CREATE SET r.code = toString(row.requirement_code),
                      r.content = toString(row.requirement_content)
        MERGE (l)-[:HAS_REQUIREMENT]->(r)
        \"\"\", batch=df_reqs.to_dict(orient="records"))

        session.run(\"\"\"
        UNWIND $batch AS row
        WITH row WHERE row.knowledge_name IS NOT NULL AND trim(toString(row.knowledge_name)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (k:Knowledge {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.knowledge_name)})
        ON CREATE SET k.name = toString(row.knowledge_name),
                      k.content = toString(row.knowledge_content)
        MERGE (l)-[:TEACHES_KNOWLEDGE]->(k)
        \"\"\", batch=df_know.to_dict(orient="records"))

        # Nạp Visual Assets & Skills
        df_vis = pd.read_csv(os.path.join(CSV_DIR, "neo4j_visual_assets.csv")).fillna("")
        df_skills = pd.read_csv(os.path.join(CSV_DIR, "neo4j_skills.csv")).fillna("")
        print(f"[5/6] Nạp {len(df_vis)} Visual Assets và {len(df_skills)} Kỹ năng...")
        session.run(\"\"\"
        UNWIND $batch AS row
        WITH row WHERE row.visual_asset_code IS NOT NULL AND trim(toString(row.visual_asset_code)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (v:VisualAsset {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.visual_asset_code)})
        ON CREATE SET v.code = toString(row.visual_asset_code),
                      v.caption = toString(row.visual_caption),
                      v.description = toString(row.visual_description),
                      v.page = toString(row.page_num)
        MERGE (l)-[:HAS_VISUAL_ASSET]->(v)
        \"\"\", batch=df_vis.to_dict(orient="records"))

        session.run(\"\"\"
        UNWIND $batch AS row
        WITH row WHERE row.skill_name IS NOT NULL AND trim(toString(row.skill_name)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (s:Skill {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.skill_name)})
        ON CREATE SET s.name = toString(row.skill_name),
                      s.description = toString(row.skill_description)
        MERGE (l)-[:DEVELOPS_SKILL]->(s)
        \"\"\", batch=df_skills.to_dict(orient="records"))

        # Nạp Questions & Answers
        df_questions = pd.read_csv(os.path.join(CSV_DIR, "neo4j_questions.csv")).fillna("")
        df_answers = pd.read_csv(os.path.join(CSV_DIR, "neo4j_answers.csv")).fillna("")
        print(f"[6/6] Nạp {len(df_questions)} Câu hỏi và {len(df_answers)} Phương án lựa chọn...")
        session.run(\"\"\"
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(toString(row.question_id)) <> ''
        MATCH (l:Lesson {id: toString(row.grade_level) + '_' + toString(row.lesson_id)})
        MERGE (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})
        ON CREATE SET q.question_id = toString(row.question_id),
                      q.content = toString(row.question_content),
                      q.difficulty = toString(row.difficulty)
        MERGE (l)-[:EVALUATED_BY]->(q)
        \"\"\", batch=df_questions.to_dict(orient="records"))

        session.run(\"\"\"
        UNWIND $batch AS row
        WITH row WHERE row.question_id IS NOT NULL AND trim(toString(row.question_id)) <> ''
        MATCH (q:Question {id: toString(row.grade_level) + '_' + toString(row.lesson_id) + '_' + toString(row.question_id)})
        CREATE (a:AnswerOption {
            option_key: toString(coalesce(row.answer_option_key, '')),
            content: toString(row.answer_content),
            is_correct: (toString(row.is_correct) = 'True' OR toString(row.is_correct) = 'true' OR toString(row.is_correct) = '1')
        })
        MERGE (q)-[:HAS_OPTION]->(a)
        \"\"\", batch=df_answers.to_dict(orient="records"))

        # Tổng kết thống kê
        res_nodes = session.run("MATCH (n) RETURN labels(n)[0] AS label, count(n) AS cnt ORDER BY cnt DESC")
        nodes_stat = {r["label"]: r["cnt"] for r in res_nodes if r["label"]}
        total_n = sum(nodes_stat.values())

        res_rels = session.run("MATCH ()-[r]->() RETURN type(r) AS rel, count(r) AS cnt ORDER BY cnt DESC")
        rels_stat = {r["rel"]: r["cnt"] for r in res_rels if r["rel"]}
        total_r = sum(rels_stat.values())

    print("\\n" + "=" * 75)
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
"""

with open(os.path.join(DUMP_DIR, "restore_database.py"), "w", encoding="utf-8") as f_py:
    f_py.write(restore_py_content)
print("  ✓ Đã lưu restore_database.py")

# 5. Tạo tài liệu hướng dẫn khôi phục chi tiết
print("\n[5/5] Soạn thảo Hướng dẫn Khôi phục dữ liệu (HUONG_DAN_NAP_DUMP_DATA.md)...")
readme_content = """# HƯỚNG DẪN KHÔI PHỤC VÀ SỬ DỤNG DUMP DATA CSDL ĐỒ THỊ NEO4J
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
"""

with open(os.path.join(DUMP_DIR, "HUONG_DAN_NAP_DUMP_DATA.md"), "w", encoding="utf-8") as f_rd:
    f_rd.write(readme_content)
print("  ✓ Đã soạn thảo HUONG_DAN_NAP_DUMP_DATA.md")

# 6. Đóng gói ZIP nén hoàn chỉnh
zip_filename = os.path.join(PROJECT_DIR, "DUMP_DATA_CSDL_NANG_CAO_EDUGRAPH.zip")
print(f"\n📦 Đang nén toàn bộ thành file ZIP nộp kèm: {os.path.basename(zip_filename)}...")
with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
    for root, dirs, files in os.walk(DUMP_DIR):
        for file in files:
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, PROJECT_DIR)
            zipf.write(full_path, arcname=rel_path)

print(f"🎉 HOÀN TẤT TẠO FILE DUMP! Kích thước file ZIP: {os.path.getsize(zip_filename):,} bytes")
