// ==============================================================================
// NEO4J IMPORT SCRIPT: CƠ SỞ DỮ LIỆU ĐỒ THỊ MÔN TIN HỌC PHỔ THÔNG GDPT 2018
// Hướng dẫn: Copy các file CSV vào thư mục import của Neo4j và thực thi các khối lệnh dưới đây.
// ==============================================================================

// 1. TẠO RÀNG BUỘC KHÓA DUY NHẤT (CONSTRAINTS) ĐỂ TĂNG HIỆU SUẤT TRUY VẤN
CREATE CONSTRAINT constraint_grade_level IF NOT EXISTS FOR (g:Grade) REQUIRE g.grade_level IS UNIQUE;
CREATE CONSTRAINT constraint_chapter_id IF NOT EXISTS FOR (c:Chapter) REQUIRE c.id IS UNIQUE;
CREATE CONSTRAINT constraint_lesson_id IF NOT EXISTS FOR (l:Lesson) REQUIRE l.id IS UNIQUE;
CREATE CONSTRAINT constraint_req_id IF NOT EXISTS FOR (r:Requirement) REQUIRE r.id IS UNIQUE;
CREATE CONSTRAINT constraint_know_id IF NOT EXISTS FOR (k:Knowledge) REQUIRE k.id IS UNIQUE;
CREATE CONSTRAINT constraint_visual_id IF NOT EXISTS FOR (v:VisualAsset) REQUIRE v.id IS UNIQUE;
CREATE CONSTRAINT constraint_skill_id IF NOT EXISTS FOR (s:Skill) REQUIRE s.id IS UNIQUE;
CREATE CONSTRAINT constraint_question_id IF NOT EXISTS FOR (q:Question) REQUIRE q.id IS UNIQUE;

// 2. NẠP CẤU TRÚC PHÂN CẤP SƯ PHẠM (GRADE -> CHAPTER -> LESSON)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_lessons.csv' AS row
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
MERGE (c)-[:HAS_LESSON]->(l);

// 3. NẠP YÊU CẦU CẦN ĐẠT (REQUIREMENTS)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_requirements.csv' AS row
WITH row WHERE row.requirement_code IS NOT NULL AND trim(row.requirement_code) <> ''
MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
MERGE (r:Requirement {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.requirement_code})
ON CREATE SET r.code = row.requirement_code,
              r.content = row.requirement_content
MERGE (l)-[:HAS_REQUIREMENT]->(r);

// 4. NẠP NỘI DUNG KIẾN THỨC CỐT LÕI (KNOWLEDGE)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_knowledge.csv' AS row
WITH row WHERE row.knowledge_name IS NOT NULL AND trim(row.knowledge_name) <> ''
MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
MERGE (k:Knowledge {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.knowledge_name})
ON CREATE SET k.name = row.knowledge_name,
              k.content = row.knowledge_content
MERGE (l)-[:TEACHES_KNOWLEDGE]->(k);

// 5. NẠP HÌNH ẢNH, SƠ ĐỒ VÀ GIAO DIỆN PHẦN MỀM (VISUAL ASSETS)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_visual_assets.csv' AS row
WITH row WHERE row.visual_asset_code IS NOT NULL AND trim(row.visual_asset_code) <> ''
MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
MERGE (v:VisualAsset {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.visual_asset_code})
ON CREATE SET v.code = row.visual_asset_code,
              v.caption = row.visual_caption,
              v.description = row.visual_description,
              v.page = row.page_num
MERGE (l)-[:HAS_VISUAL_ASSET]->(v);

// 6. NẠP KỸ NĂNG THỰC HÀNH (SKILLS)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_skills.csv' AS row
WITH row WHERE row.skill_name IS NOT NULL AND trim(row.skill_name) <> ''
MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
MERGE (s:Skill {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.skill_name})
ON CREATE SET s.name = row.skill_name,
              s.description = row.skill_description
MERGE (l)-[:DEVELOPS_SKILL]->(s);

// 7. NẠP CÂU HỎI LUYỆN TẬP VÀ VẬN DỤNG (QUESTIONS)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_questions.csv' AS row
WITH row WHERE row.question_id IS NOT NULL AND trim(row.question_id) <> ''
MATCH (l:Lesson {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id})
MERGE (q:Question {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.question_id})
ON CREATE SET q.question_id = row.question_id,
              q.content = row.question_content,
              q.difficulty = row.difficulty
MERGE (l)-[:EVALUATED_BY]->(q);

// 8. NẠP PHƯƠNG ÁN LỰA CHỌN VÀ ĐÁP ÁN (ANSWER OPTIONS)
LOAD CSV WITH HEADERS FROM 'file:///neo4j_answers.csv' AS row
WITH row WHERE row.question_id IS NOT NULL AND trim(row.question_id) <> ''
MATCH (q:Question {id: coalesce(row.grade_level, 'Lớp 9') + '_' + row.lesson_id + '_' + row.question_id})
CREATE (a:AnswerOption {
    option_key: coalesce(row.answer_option_key, ''),
    content: row.answer_content,
    is_correct: (row.is_correct = 'True' OR row.is_correct = 'true')
})
MERGE (q)-[:HAS_OPTION]->(a);
