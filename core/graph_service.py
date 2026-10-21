import os
import csv
import random
from typing import List, Dict, Any, Optional

try:
    from neo4j import GraphDatabase
    NEO4J_AVAILABLE = True
except ImportError:
    NEO4J_AVAILABLE = False


class GraphService:
    """
    Dịch vụ quản lý cơ sở dữ liệu đồ thị tri thức môn Tin học GDPT 2018.
    Hỗ trợ cả chế độ:
    1. Trực tiếp từ kho dữ liệu Master CSV (Local Engine)
    2. Kết nối trực tiếp qua Bolt Protocol đến Neo4j Database Server
    """

    def __init__(self, master_dir: str):
        self.master_dir = master_dir
        self.driver = None
        self.database = "edugraph"

    def connect_neo4j(self, uri: str, user: str, password: str, database: str = "edugraph") -> bool:
        """Kết nối đến Neo4j DBMS qua giao thức Bolt trên database chỉ định."""
        if not NEO4J_AVAILABLE:
            raise RuntimeError("Thư viện 'neo4j' chưa được cài đặt.")
        try:
            self.driver = GraphDatabase.driver(uri, auth=(user, password))
            self.driver.verify_connectivity()
            self.database = database
            return True
        except Exception as e:
            self.driver = None
            raise RuntimeError(f"Không thể kết nối Neo4j: {str(e)}")

    def is_connected(self) -> bool:
        if not self.driver:
            return False
        try:
            self.driver.verify_connectivity()
            return True
        except Exception:
            return False

    def get_live_stats(self) -> Dict[str, Any]:
        """Lấy thống kê trực tiếp từ Neo4j DBMS đang kết nối."""
        if not self.is_connected():
            return {}
        try:
            with self.driver.session(database=self.database) as session:
                res_nodes = session.run("MATCH (n) RETURN labels(n)[0] AS label, count(n) AS cnt ORDER BY cnt DESC")
                label_counts = {rec["label"]: rec["cnt"] for rec in res_nodes if rec["label"]}
                res_rels = session.run("MATCH ()-[r]->() RETURN type(r) AS rel_type, count(r) AS cnt ORDER BY cnt DESC")
                rel_counts = {rec["rel_type"]: rec["cnt"] for rec in res_rels if rec["rel_type"]}
                total_nodes = sum(label_counts.values())
                total_rels = sum(rel_counts.values())
                return {
                    "total_nodes": total_nodes,
                    "total_relationships": total_rels,
                    "labels": label_counts,
                    "relationships": rel_counts
                }
        except Exception as e:
            print(f"Lỗi lấy live stats: {e}")
            return {}

    def get_live_nodes(self, label: str, limit: int = 500) -> List[Dict[str, Any]]:
        """Lấy dữ liệu thuộc tính của các Node theo nhãn trực tiếp từ Neo4j."""
        if not self.is_connected():
            return []
        try:
            with self.driver.session(database=self.database) as session:
                query = f"MATCH (n:{label}) RETURN properties(n) AS props, elementId(n) AS element_id LIMIT $limit"
                res = session.run(query, limit=limit)
                rows = []
                for rec in res:
                    props = dict(rec["props"])
                    props["_element_id"] = rec["element_id"]
                    rows.append(props)
                return rows
        except Exception as e:
            print(f"Lỗi lấy live nodes {label}: {e}")
            return []

    def _read_csv(self, file_name: str) -> List[Dict[str, str]]:
        file_path = os.path.join(self.master_dir, file_name)
        if not os.path.exists(file_path):
            return []
        rows = []
        try:
            with open(file_path, mode="r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    rows.append(r)
        except Exception as e:
            print(f"Lỗi đọc file {file_name}: {e}")
        return rows

    def get_graph_summary(self) -> Dict[str, Any]:
        """Thống kê tổng quan số lượng các thực thể trong Đồ thị tri thức."""
        if self.is_connected():
            try:
                stats = self.get_live_stats()
                labels = stats.get("labels", {})
                return {
                    "total_grades": labels.get("Grade", 0),
                    "total_chapters": labels.get("Chapter", 0),
                    "total_lessons": labels.get("Lesson", 0),
                    "total_requirements": labels.get("Requirement", 0),
                    "total_knowledges": labels.get("Knowledge", 0),
                    "total_skills": labels.get("Skill", 0),
                    "total_visual_assets": labels.get("VisualAsset", 0),
                    "total_questions": labels.get("Question", 0),
                    "total_options": labels.get("AnswerOption", 0),
                    "grades_list": ["Lớp 9", "Lớp 6"],
                    "chapters_list": [],
                    "is_live_neo4j": True
                }
            except Exception as e:
                print(f"Lỗi lấy summary từ Neo4j Live: {e}")

        # Fallback đọc CSV khi ở chế độ offline
        lessons = self._read_csv("neo4j_lessons.csv")
        reqs = self._read_csv("neo4j_requirements.csv")
        knowledges = self._read_csv("neo4j_knowledges.csv")
        skills = self._read_csv("neo4j_skills.csv")
        visuals = self._read_csv("neo4j_visual_assets.csv")
        questions = self._read_csv("neo4j_questions.csv")
        options = self._read_csv("neo4j_answer_options.csv")

        grades = set(r.get("grade_level", "") for r in lessons if r.get("grade_level"))
        chapters = set(r.get("chapter_title", "") for r in lessons if r.get("chapter_title"))

        return {
            "total_grades": len(grades),
            "total_chapters": len(chapters),
            "total_lessons": len(lessons),
            "total_requirements": len(reqs),
            "total_knowledges": len(knowledges),
            "total_skills": len(skills),
            "total_visual_assets": len(visuals),
            "total_questions": len(questions),
            "total_options": len(options),
            "grades_list": sorted(list(grades)),
            "chapters_list": sorted(list(chapters)),
            "is_live_neo4j": False
        }

    def get_lessons_catalog(self, grade_filter: Optional[str] = None) -> List[Dict[str, Any]]:
        """Lấy danh mục các bài học kèm chi tiết cây tri thức sư phạm."""
        if self.is_connected():
            try:
                with self.driver.session(database=self.database) as session:
                    cypher = """
                    MATCH (c:Chapter)-[:HAS_LESSON]->(l:Lesson)
                    OPTIONAL MATCH (l)-[:MEETS_REQUIREMENT]->(r:Requirement)
                    OPTIONAL MATCH (l)-[:COVERS_KNOWLEDGE]->(k:Knowledge)
                    OPTIONAL MATCH (l)-[:DEVELOPS_SKILL]->(s:Skill)
                    OPTIONAL MATCH (l)-[:HAS_VISUAL_ASSET]->(v:VisualAsset)
                    OPTIONAL MATCH (l)-[:TESTS_KNOWLEDGE]->(q:Question)
                    RETURN l.grade_level AS grade_level, c.title AS chapter_title, 
                           l.lesson_id AS lesson_id, l.lesson_title AS lesson_title, l.page_num AS page_num,
                           collect(DISTINCT properties(r)) AS requirements,
                           collect(DISTINCT properties(k)) AS knowledges,
                           collect(DISTINCT properties(s)) AS skills,
                           collect(DISTINCT properties(v)) AS visual_assets,
                           collect(DISTINCT properties(q)) AS questions
                    ORDER BY l.lesson_id
                    """
                    records = session.run(cypher)
                    catalog = []
                    for rec in records:
                        g = rec["grade_level"] or "Lớp 9"
                        if grade_filter and grade_filter != "Tất cả" and g != grade_filter:
                            continue
                        catalog.append({
                            "grade_level": g,
                            "chapter_title": rec["chapter_title"] or "",
                            "lesson_id": rec["lesson_id"] or "",
                            "lesson_title": rec["lesson_title"] or "",
                            "page_num": rec["page_num"] or "",
                            "requirements": [x for x in rec["requirements"] if x and x.get("requirement_content")],
                            "knowledges": [x for x in rec["knowledges"] if x and x.get("knowledge_name")],
                            "skills": [x for x in rec["skills"] if x and x.get("skill_name")],
                            "visual_assets": [x for x in rec["visual_assets"] if x and x.get("caption")],
                            "questions": [x for x in rec["questions"] if x and x.get("content")]
                        })
                    if catalog:
                        return catalog
            except Exception as e:
                print(f"Lỗi lấy catalog từ Neo4j Live: {e}")

        # Fallback đọc CSV
        lessons = self._read_csv("neo4j_lessons.csv")
        reqs = self._read_csv("neo4j_requirements.csv")
        knowledges = self._read_csv("neo4j_knowledges.csv")
        skills = self._read_csv("neo4j_skills.csv")
        visuals = self._read_csv("neo4j_visual_assets.csv")
        questions = self._read_csv("neo4j_questions.csv")

        catalog = []
        for l in lessons:
            g = l.get("grade_level", "")
            lid = l.get("lesson_id", "")
            if grade_filter and grade_filter != "Tất cả" and g != grade_filter:
                continue

            l_reqs = [r for r in reqs if r.get("grade_level") == g and r.get("lesson_id") == lid]
            l_knows = [k for k in knowledges if k.get("grade_level") == g and k.get("lesson_id") == lid]
            l_skills = [s for s in skills if s.get("grade_level") == g and s.get("lesson_id") == lid]
            l_visuals = [v for v in visuals if v.get("grade_level") == g and v.get("lesson_id") == lid]
            l_qs = [q for q in questions if q.get("grade_level") == g and q.get("lesson_id") == lid]

            catalog.append({
                "grade_level": g,
                "chapter_title": l.get("chapter_title", ""),
                "lesson_id": lid,
                "lesson_title": l.get("lesson_title", ""),
                "page_num": l.get("page_num", ""),
                "requirements": l_reqs,
                "knowledges": l_knows,
                "skills": l_skills,
                "visual_assets": l_visuals,
                "questions": l_qs
            })
        return catalog

    def get_curriculum_analytics(self) -> Dict[str, Any]:
        """Phân tích chương trình Tin học GDPT 2018 theo ma trận sư phạm."""
        lessons = self._read_csv("neo4j_lessons.csv")
        reqs = self._read_csv("neo4j_requirements.csv")
        visuals = self._read_csv("neo4j_visual_assets.csv")
        questions = self._read_csv("neo4j_questions.csv")

        diff_counts = {"Nhận biết": 0, "Thông hiểu": 0, "Vận dụng": 0, "Khác": 0}
        for q in questions:
            diff = q.get("difficulty", "Nhận biết").strip()
            if diff in diff_counts:
                diff_counts[diff] += 1
            else:
                diff_counts["Khác"] += 1

        # Phân bố yêu cầu cần đạt theo từng chương
        chapter_reqs = {}
        for r in reqs:
            # Tra cứu chapter từ lesson
            lid = r.get("lesson_id", "")
            g = r.get("grade_level", "")
            ch = "Chưa phân loại"
            for l in lessons:
                if l.get("lesson_id") == lid and l.get("grade_level") == g:
                    ch = l.get("chapter_title", "Chưa phân loại")
                    break
            chapter_reqs[ch] = chapter_reqs.get(ch, 0) + 1

        # Phân bố Visual Assets theo loại
        visual_types = {"Sơ đồ thuật toán/Tư duy": 0, "Giao diện phần mềm": 0, "Hình ảnh/Minh họa": 0}
        for v in visuals:
            desc = (v.get("caption", "") + " " + v.get("detailed_description", "")).lower()
            if "thuật toán" in desc or "khối" in desc or "sơ đồ" in desc:
                visual_types["Sơ đồ thuật toán/Tư duy"] += 1
            elif "giao diện" in desc or "cửa sổ" in desc or "hộp thoại" in desc or "excel" in desc or "phần mềm" in desc:
                visual_types["Giao diện phần mềm"] += 1
            else:
                visual_types["Hình ảnh/Minh họa"] += 1

        return {
            "difficulty_distribution": diff_counts,
            "chapter_requirements": chapter_reqs,
            "visual_asset_distribution": visual_types
        }

    def generate_exam(
        self,
        grade_level: str,
        lesson_ids: Optional[List[str]] = None,
        total_questions: int = 10,
        nb_ratio: float = 0.4,
        th_ratio: float = 0.3,
        vd_ratio: float = 0.3,
        prioritize_visuals: bool = False
    ) -> Dict[str, Any]:
        """
        Thuật toán sinh đề thi tự động từ Đồ thị tri thức môn Tin học GDPT 2018
        tuân thủ ma trận phân loại nhận thức (Nhận biết - Thông hiểu - Vận dụng).
        """
        all_questions = self._read_csv("neo4j_questions.csv")
        all_options = self._read_csv("neo4j_answer_options.csv")
        all_visuals = self._read_csv("neo4j_visual_assets.csv")

        # Gom nhóm options theo question
        options_map = {}
        for opt in all_options:
            q_key = f"{opt.get('grade_level', '')}_{opt.get('lesson_id', '')}_{opt.get('question_id', '')}"
            if q_key not in options_map:
                options_map[q_key] = []
            options_map[q_key].append(opt)

        # Gom nhóm visuals theo lesson
        visuals_map = {}
        for v in all_visuals:
            l_key = f"{v.get('grade_level', '')}_{v.get('lesson_id', '')}"
            if l_key not in visuals_map:
                visuals_map[l_key] = []
            visuals_map[l_key].append(v)

        # Lọc câu hỏi theo lớp và bài học
        pool = []
        for q in all_questions:
            g = q.get("grade_level", "")
            lid = q.get("lesson_id", "")
            if grade_level != "Tất cả" and g != grade_level:
                continue
            if lesson_ids and len(lesson_ids) > 0 and lid not in lesson_ids:
                continue

            q_key = f"{g}_{lid}_{q.get('question_id', '')}"
            q_opts = options_map.get(q_key, [])
            l_key = f"{g}_{lid}"
            q_visuals = visuals_map.get(l_key, [])

            q_obj = {
                "grade_level": g,
                "lesson_id": lid,
                "question_id": q.get("question_id", ""),
                "content": q.get("content", ""),
                "difficulty": q.get("difficulty", "Nhận biết").strip(),
                "page_num": q.get("page_num", ""),
                "options": sorted(q_opts, key=lambda x: x.get("option_key", "")),
                "visuals": q_visuals
            }
            pool.append(q_obj)

        if not pool:
            return {"questions": [], "matrix": {}, "message": "Không tìm thấy câu hỏi phù hợp với tiêu chí lọc."}

        # Phân loại pool theo mức độ
        nb_pool = [q for q in pool if q["difficulty"] == "Nhận biết"]
        th_pool = [q for q in pool if q["difficulty"] == "Thông hiểu"]
        vd_pool = [q for q in pool if q["difficulty"] == "Vận dụng"]
        other_pool = [q for q in pool if q["difficulty"] not in ["Nhận biết", "Thông hiểu", "Vận dụng"]]

        # Tính chỉ tiêu số lượng câu hỏi theo ma trận
        target_nb = int(round(total_questions * nb_ratio))
        target_th = int(round(total_questions * th_ratio))
        target_vd = total_questions - (target_nb + target_th)

        def pick_items(source_list: List[Dict], count: int) -> List[Dict]:
            if not source_list or count <= 0:
                return []
            if prioritize_visuals:
                # Ưu tiên câu hỏi thuộc bài có visual asset
                with_vis = [x for x in source_list if len(x.get("visuals", [])) > 0]
                without_vis = [x for x in source_list if len(x.get("visuals", [])) == 0]
                random.shuffle(with_vis)
                random.shuffle(without_vis)
                combined = with_vis + without_vis
                return combined[:min(count, len(combined))]
            else:
                shuffled = list(source_list)
                random.shuffle(shuffled)
                return shuffled[:min(count, len(shuffled))]

        selected_nb = pick_items(nb_pool, target_nb)
        selected_th = pick_items(th_pool, target_th)
        selected_vd = pick_items(vd_pool, target_vd)

        selected_all = selected_nb + selected_th + selected_vd

        # Bù thêm nếu thiếu do số lượng ở một mức độ không đủ
        if len(selected_all) < total_questions:
            remaining_needed = total_questions - len(selected_all)
            selected_ids = set(f"{q['grade_level']}_{q['lesson_id']}_{q['question_id']}" for q in selected_all)
            remain_pool = [
                q for q in (nb_pool + th_pool + vd_pool + other_pool)
                if f"{q['grade_level']}_{q['lesson_id']}_{q['question_id']}" not in selected_ids
            ]
            fill_items = pick_items(remain_pool, remaining_needed)
            selected_all.extend(fill_items)

        actual_matrix = {
            "Nhận biết": sum(1 for q in selected_all if q["difficulty"] == "Nhận biết"),
            "Thông hiểu": sum(1 for q in selected_all if q["difficulty"] == "Thông hiểu"),
            "Vận dụng": sum(1 for q in selected_all if q["difficulty"] == "Vận dụng"),
            "Khác": sum(1 for q in selected_all if q["difficulty"] not in ["Nhận biết", "Thông hiểu", "Vận dụng"]),
            "Tổng số": len(selected_all)
        }

        return {
            "questions": selected_all,
            "matrix": actual_matrix,
            "total_requested": total_questions,
            "grade_level": grade_level
        }

    def execute_cypher(self, query: str, uri: str = "bolt://localhost:7687", user: str = "neo4j", password: str = "") -> List[Dict[str, Any]]:
        """Thực thi câu lệnh Cypher trực tiếp trên Neo4j DBMS và trả về kết quả dạng dict."""
        if not NEO4J_AVAILABLE:
            raise RuntimeError("Thư viện 'neo4j' chưa được cài đặt.")
        
        # Sử dụng driver hiện có hoặc tạo driver tạm thời
        drv = self.driver
        temp_drv = False
        if drv is None:
            try:
                drv = GraphDatabase.driver(uri, auth=(user, password) if password else None)
                temp_drv = True
            except Exception as e:
                raise RuntimeError(f"Không thể kết nối Neo4j: {str(e)}")

        try:
            with drv.session() as session:
                result = session.run(query)
                records = []
                for record in result:
                    r_dict = {}
                    for key in record.keys():
                        val = record[key]
                        # Chuyển đổi Node/Relationship của Neo4j thành dict thân thiện
                        if hasattr(val, "items"):
                            r_dict[key] = dict(val.items())
                        elif hasattr(val, "element_id"):
                            r_dict[key] = f"Node({val.labels}): {dict(val)}"
                        else:
                            r_dict[key] = val
                    records.append(r_dict)
                return records
        finally:
            if temp_drv and drv:
                drv.close()

    def generate_pyvis_schema_html(self) -> str:
        """
        Sinh mã HTML hiển thị Lược đồ Đồ thị Tri thức (Meta-Schema Visualization)
        tương tự arrows.app hoặc CALL db.schema.visualization();
        """
        try:
            from pyvis.network import Network
        except ImportError:
            return "<div style='color:red;'>Thư viện pyvis chưa được cài đặt.</div>"

        net = Network(height="520px", width="100%", bgcolor="#f8fafc", font_color="#1e293b", directed=True)
        net.barnes_hut(gravity=-3500, central_gravity=0.3, spring_length=150, spring_strength=0.05)

        # Định nghĩa các Node trong Lược đồ (Meta-Schema)
        nodes_info = [
            ("Grade", "Grade\n(Khối lớp)", "#3b82f6", 32, "Node đại diện cho Khối lớp GDPT 2018 (Lớp 6, 7, 8, 9, 10, 11, 12)"),
            ("Chapter", "Chapter\n(Chủ đề/Chương)", "#8b5cf6", 28, "Chủ đề lớn hoặc Chương trong SGK môn Tin học"),
            ("Lesson", "Lesson\n(Bài học)", "#0ea5e9", 30, "Đơn vị bài học chính trong chương trình GDPT 2018"),
            ("Requirement", "Requirement\n(Yêu cầu cần đạt)", "#10b981", 24, "Mục tiêu chuẩn đầu ra 'Sau bài học này em sẽ...'"),
            ("Knowledge", "Knowledge\n(Kiến thức cốt lõi)", "#f59e0b", 24, "Khái niệm, định nghĩa và nội dung lý thuyết chính"),
            ("Skill", "Skill\n(Kỹ năng thực hành)", "#ec4899", 22, "Kỹ năng thao tác máy tính, lập trình Scratch/Python, bảng tính"),
            ("VisualAsset", "VisualAsset\n(Hình ảnh/Sơ đồ trực quan)", "#06b6d4", 26, "Sơ đồ thuật toán, giao diện phần mềm, tranh minh họa kèm mô tả AI"),
            ("Question", "Question\n(Ngân hàng câu hỏi)", "#ef4444", 26, "Câu hỏi ôn tập, luyện tập phân theo thang Bloom (NB-TH-VD)"),
            ("AnswerOption", "AnswerOption\n(Phương án trắc nghiệm)", "#64748b", 20, "Các lựa chọn A, B, C, D kèm nhãn đáp án đúng/sai")
        ]

        for n_id, n_label, n_color, n_size, n_title in nodes_info:
            net.add_node(
                n_id,
                label=n_label,
                color=n_color,
                size=n_size,
                title=n_title,
                font={"size": 14, "color": "#0f172a", "face": "Segoe UI", "strokeWidth": 2, "strokeColor": "#ffffff"},
                shape="dot",
                shadow=True
            )

        # Định nghĩa các Relationship trong Lược đồ
        edges_info = [
            ("Grade", "Chapter", ":HAS_CHAPTER"),
            ("Chapter", "Lesson", ":HAS_LESSON"),
            ("Lesson", "Requirement", ":MEETS_REQUIREMENT"),
            ("Lesson", "Knowledge", ":COVERS_KNOWLEDGE"),
            ("Lesson", "Skill", ":DEVELOPS_SKILL"),
            ("Lesson", "VisualAsset", ":HAS_VISUAL_ASSET"),
            ("Lesson", "Question", ":TESTS_KNOWLEDGE"),
            ("Question", "AnswerOption", ":HAS_OPTION")
        ]

        for u, v, rel_label in edges_info:
            net.add_edge(
                u,
                v,
                label=rel_label,
                color="#64748b",
                arrows="to",
                font={"size": 11, "align": "middle", "color": "#475569"},
                smooth={"type": "curvedCW", "roundness": 0.15}
            )

        return net.generate_html()

    def generate_pyvis_lesson_subgraph_html(self, lesson_id: str, grade_level: str = "Lớp 9") -> str:
        """
        Sinh mã HTML hiển thị Mạng lưới Tri thức cụ thể (Interactive Lesson Subgraph)
        cho một Bài học cụ thể.
        """
        try:
            from pyvis.network import Network
        except ImportError:
            return "<div style='color:red;'>Thư viện pyvis chưa được cài đặt.</div>"

        catalog = self.get_lessons_catalog(grade_filter=grade_level)
        target_lesson = None
        for l in catalog:
            if l.get("lesson_id") == lesson_id:
                target_lesson = l
                break

        if not target_lesson:
            return "<div style='padding:20px; color:#64748b;'>Không tìm thấy dữ liệu bài học này.</div>"

        net = Network(height="520px", width="100%", bgcolor="#f8fafc", font_color="#1e293b", directed=True)
        net.barnes_hut(gravity=-2500, central_gravity=0.4, spring_length=120)

        # Nút trung tâm Bài học
        l_node_id = f"L_{target_lesson['lesson_id']}"
        net.add_node(
            l_node_id,
            label=f"{target_lesson['lesson_id']}\n{target_lesson['lesson_title'][:25]}...",
            color="#0ea5e9",
            size=35,
            title=f"Bài học: {target_lesson['lesson_title']} (Trang {target_lesson.get('page_num', '')})",
            font={"size": 14, "color": "#ffffff", "face": "Segoe UI"},
            shape="dot",
            shadow=True
        )

        # Nút Yêu cầu cần đạt
        for idx, r in enumerate(target_lesson.get("requirements", [])[:5], 1):
            r_id = f"REQ_{idx}"
            r_text = r.get("requirement_content", "")
            net.add_node(
                r_id,
                label=f"[{r.get('requirement_code', 'YCCD')}]\n{r_text[:20]}...",
                color="#10b981",
                size=22,
                title=f"Yêu cầu cần đạt: {r_text}",
                shape="dot"
            )
            net.add_edge(l_node_id, r_id, label=":MEETS_REQUIREMENT", color="#10b981", arrows="to")

        # Nút Kiến thức cốt lõi
        for idx, k in enumerate(target_lesson.get("knowledges", [])[:6], 1):
            k_id = f"KNOW_{idx}"
            k_name = k.get("knowledge_name", "")
            net.add_node(
                k_id,
                label=f"{k_name[:20]}...",
                color="#f59e0b",
                size=22,
                title=f"Kiến thức: {k_name}\n{k.get('knowledge_content', '')[:100]}...",
                shape="dot"
            )
            net.add_edge(l_node_id, k_id, label=":COVERS_KNOWLEDGE", color="#f59e0b", arrows="to")

        # Nút Visual Asset
        for idx, v in enumerate(target_lesson.get("visual_assets", [])[:4], 1):
            v_id = f"VIS_{idx}"
            v_cap = v.get("caption", "")
            net.add_node(
                v_id,
                label=f"🖼️ {v.get('visual_asset_code', 'Hình')}\n{v_cap[:18]}...",
                color="#06b6d4",
                size=24,
                title=f"Visual Asset: {v_cap}\nAI Mô tả: {v.get('detailed_description', '')[:120]}...",
                shape="dot"
            )
            net.add_edge(l_node_id, v_id, label=":HAS_VISUAL_ASSET", color="#06b6d4", arrows="to")

        # Nút Câu hỏi ôn tập
        for idx, q in enumerate(target_lesson.get("questions", [])[:4], 1):
            q_id = f"Q_{idx}"
            q_content = q.get("content", "")
            net.add_node(
                q_id,
                label=f"❓ {q.get('question_id', 'CH')}\n[{q.get('difficulty', 'NB')}]",
                color="#ef4444",
                size=24,
                title=f"Câu hỏi: {q_content}",
                shape="dot"
            )
            net.add_edge(l_node_id, q_id, label=":TESTS_KNOWLEDGE", color="#ef4444", arrows="to")

        return net.generate_html()
