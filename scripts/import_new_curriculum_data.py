import os
import sys
from neo4j import GraphDatabase

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

def add_new_curriculum_data():
    uri = "bolt://localhost:7687"
    user = "neo4j"
    password = "12345678"
    
    print("[1/3] Đang kết nối tới Neo4j DBMS...")
    driver = GraphDatabase.driver(uri, auth=(user, password))
    
    # Dữ liệu mới: Khối lớp 10 - Khoa học máy tính & Chuyên đề AI
    new_data = {
        "grade": "Lớp 10",
        "chapters": [
            {
                "id": "Lớp 10_Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính",
                "title": "Chủ đề F: Giải quyết vấn đề với sự trợ giúp của máy tính (Lập trình Python)",
                "lessons": [
                    {
                        "lesson_id": "Bài 16",
                        "lesson_title": "Ngôn ngữ lập trình bậc cao và môi trường phát triển Python",
                        "page_num": 88,
                        "source_file": "SGK Tin học 10 KNTT (Định hướng CS).pdf",
                        "requirements": [
                            {"code": "YCCD_10_16_01", "content": "Biết và phân biệt được ngôn ngữ máy, hợp ngữ và ngôn ngữ lập trình bậc cao."},
                            {"code": "YCCD_10_16_02", "content": "Khởi động và làm quen với môi trường lập trình Python (chế độ gõ lệnh trực tiếp và chế độ soạn thảo tệp script)."}
                        ],
                        "knowledges": [
                            {"name": "Ngôn ngữ lập trình bậc cao", "content": "Ngôn ngữ gần gũi với ngôn ngữ tự nhiên của con người, có tính độc lập phần cứng cao và cần trình biên dịch/thông dịch để thực thi."},
                            {"name": "Môi trường IDLE Python", "content": "Môi trường tích hợp cung cấp giao diện tương tác trực tiếp (Shell) và trình soạn thảo file .py để viết mã nguồn."}
                        ],
                        "skills": [
                            {"name": "Kỹ năng cài đặt và khởi động IDLE Python", "description": "Mở IDLE, chạy thử lệnh print() trực tiếp trên dấu nhắc lệnh >>> và lưu file script .py."},
                            {"name": "Kỹ năng khắc phục lỗi cú pháp cơ bản", "description": "Nhận diện thông báo SyntaxError khi gõ thiếu ngoặc hoặc sai từ khóa."}
                        ],
                        "visuals": [
                            {"code": "HINH_10_16_01", "type": "Ảnh chụp giao diện", "caption": "Giao diện Python Shell và chế độ soạn thảo Editor", "desc": "Cửa sổ dòng lệnh Python 3.11 hiển thị prompt >>> và cửa sổ soạn thảo script không tên Untitled."},
                            {"code": "HINH_10_16_02", "type": "Sơ đồ khối", "caption": "Quy trình thông dịch mã nguồn Python sang Bytecode", "desc": "Lưu đồ mô tả mã nguồn .py chuyển qua Python Compiler tạo ra file .pyc bytecode rồi nạp vào máy ảo PVM."}
                        ],
                        "questions": [
                            {
                                "id": "CH_10_16_01",
                                "content": "Phát biểu nào sau đây là ĐÚNG nhất về ngôn ngữ lập trình Python?",
                                "difficulty": "Nhận biết",
                                "options": [
                                    {"key": "A", "content": "Python là ngôn ngữ lập trình bậc cao, thông dịch, cú pháp ngắn gọn và đa nền tảng.", "is_correct": True},
                                    {"key": "B", "content": "Python là hợp ngữ chỉ chạy được trên hệ điều hành Windows.", "is_correct": False},
                                    {"key": "C", "content": "Python là ngôn ngữ máy mà vi xử lý có thể hiểu trực tiếp không cần dịch.", "is_correct": False},
                                    {"key": "D", "content": "Python bắt buộc phải kết thúc câu lệnh bằng dấu chấm phẩy (;).", "is_correct": False}
                                ]
                            }
                        ]
                    },
                    {
                        "lesson_id": "Bài 17",
                        "lesson_title": "Biến, kiểu dữ liệu cơ sở và lệnh gán trong Python",
                        "page_num": 94,
                        "source_file": "SGK Tin học 10 KNTT (Định hướng CS).pdf",
                        "requirements": [
                            {"code": "YCCD_10_17_01", "content": "Hiểu khái niệm biến, quy tắc đặt tên biến và lệnh gán giá trị trong Python."},
                            {"code": "YCCD_10_17_02", "content": "Phân biệt được các kiểu dữ liệu cơ sở: int (số nguyên), float (số thực), str (chuỗi ký tự), bool (logic)."}
                        ],
                        "knowledges": [
                            {"name": "Biến và Định danh trong Python", "content": "Vùng nhớ được đặt tên để lưu trữ giá trị. Tên biến bắt đầu bằng chữ cái hoặc gạch dưới, không chứa ký tự đặc biệt, phân biệt hoa thường."},
                            {"name": "Kiểu dữ liệu động (Dynamic Typing)", "content": "Python tự động suy luận kiểu dữ liệu của biến tại thời điểm gán giá trị mà không cần khai báo kiểu trước."}
                        ],
                        "skills": [
                            {"name": "Kỹ năng khai báo biến và gán giá trị", "description": "Sử dụng toán tử gán = để khởi tạo và cập nhật giá trị cho biến."},
                            {"name": "Kỹ năng ép kiểu dữ liệu", "description": "Sử dụng các hàm int(), float(), str() để chuyển đổi linh hoạt giữa các kiểu."}
                        ],
                        "visuals": [
                            {"code": "HINH_10_17_01", "type": "Sơ đồ bộ nhớ", "caption": "Cơ chế quản lý tham chiếu vùng nhớ của biến trong Python", "desc": "Sơ đồ ô nhớ minh họa biến x trỏ tới đối tượng số nguyên 10 trong Heap memory."}
                        ],
                        "questions": [
                            {
                                "id": "CH_10_17_01",
                                "content": "Trong Python, tên biến nào sau đây là HỢP LỆ theo chuẩn quy tắc định danh?",
                                "difficulty": "Thông hiểu",
                                "options": [
                                    {"key": "A", "content": "tong_diem_2026", "is_correct": True},
                                    {"key": "B", "content": "2026_tongdiem", "is_correct": False},
                                    {"key": "C", "content": "tong-diem", "is_correct": False},
                                    {"key": "D", "content": "def", "is_correct": False}
                                ]
                            }
                        ]
                    },
                    {
                        "lesson_id": "Bài 22",
                        "lesson_title": "Kiểu dữ liệu danh sách (List) và các thao tác cơ bản",
                        "page_num": 115,
                        "source_file": "SGK Tin học 10 KNTT (Định hướng CS).pdf",
                        "requirements": [
                            {"code": "YCCD_10_22_01", "content": "Mô tả được kiểu dữ liệu danh sách (List), cách khởi tạo và truy cập phần tử qua chỉ số (Index)."},
                            {"code": "YCCD_10_22_02", "content": "Sử dụng thành thạo các phương thức cơ bản: append(), remove(), insert(), pop() và duyệt list bằng vòng lặp."}
                        ],
                        "knowledges": [
                            {"name": "Cấu trúc dữ liệu List trong Python", "content": "Tập hợp các phần tử có thứ tự, có thể thay đổi (mutable), cho phép chứa các phần tử có kiểu dữ liệu khác nhau."},
                            {"name": "Cơ chế đánh chỉ số âm (Negative Indexing)", "content": "Chỉ số dương tính từ 0 từ trái sang phải; chỉ số âm tính từ -1 từ phần tử cuối cùng trở về trước."}
                        ],
                        "skills": [
                            {"name": "Kỹ năng thao tác List slicing", "description": "Trích xuất danh sách con theo cú pháp a[start:stop:step]."},
                            {"name": "Kỹ năng duyệt phần tử danh sách", "description": "Duyệt qua danh sách bằng vòng lặp for item in a hoặc for i in range(len(a))."}
                        ],
                        "visuals": [
                            {"code": "HINH_10_22_01", "type": "Sơ đồ chỉ số", "caption": "Đánh chỉ số hai chiều xuôi và ngược trên List", "desc": "Bảng mô tả các ô nhớ [10, 20, 30] với chỉ số xuôi 0, 1, 2 và chỉ số ngược -3, -2, -1."}
                        ],
                        "questions": [
                            {
                                "id": "CH_10_22_01",
                                "content": "Cho mảng a = [5, 10, 15, 20]. Giá trị của biểu thức a[-1] + a[1] là bao nhiêu?",
                                "difficulty": "Vận dụng",
                                "options": [
                                    {"key": "A", "content": "30 (vì 20 + 10 = 30)", "is_correct": True},
                                    {"key": "B", "content": "25", "is_correct": False},
                                    {"key": "C", "content": "15", "is_correct": False},
                                    {"key": "D", "content": "Báo lỗi IndexError", "is_correct": False}
                                ]
                            }
                        ]
                    },
                    {
                        "lesson_id": "Bài 28",
                        "lesson_title": "Thuật toán tìm kiếm tuần tự và tìm kiếm nhị phân",
                        "page_num": 138,
                        "source_file": "SGK Tin học 10 KNTT (Định hướng CS).pdf",
                        "requirements": [
                            {"code": "YCCD_10_28_01", "content": "Hiểu và cài đặt được thuật toán tìm kiếm tuần tự (Linear Search) trên danh sách."},
                            {"code": "YCCD_10_28_02", "content": "Hiểu nguyên lý chia để trị của thuật toán tìm kiếm nhị phân (Binary Search) và điều kiện mảng phải được sắp xếp."}
                        ],
                        "knowledges": [
                            {"name": "Thuật toán tìm kiếm tuần tự (Linear Search)", "content": "So sánh lần lượt từng phần tử từ đầu đến cuối danh sách với khóa cần tìm. Độ phức tạp thời gian O(n)."},
                            {"name": "Thuật toán tìm kiếm nhị phân (Binary Search)", "content": "Tại mỗi bước so sánh khóa với phần tử ở giữa, loại bỏ nửa danh sách không chứa khóa. Độ phức tạp O(log n)."}
                        ],
                        "skills": [
                            {"name": "Kỹ năng cài đặt thuật toán nhị phân", "description": "Viết mã Python với 2 con trỏ left, right và vòng lặp while left <= right: mid = (left+right)//2."},
                            {"name": "Kỹ năng đánh giá độ phức tạp thuật toán", "description": "So sánh số phép tính giữa O(n) và O(log n) khi kích thước dữ liệu tăng lên."}
                        ],
                        "visuals": [
                            {"code": "HINH_10_28_01", "type": "Lưu đồ thuật toán", "caption": "Lưu đồ thuật toán tìm kiếm nhị phân (Binary Search Flowchart)", "desc": "Sơ đồ khối thể hiện nhánh kiểm tra khóa k == a[mid], k < a[mid] thì thu hẹp right = mid - 1."}
                        ],
                        "questions": [
                            {
                                "id": "CH_10_28_01",
                                "content": "Điều kiện TIÊN QUYẾT bắt buộc để áp dụng thuật toán tìm kiếm nhị phân là gì?",
                                "difficulty": "Thông hiểu",
                                "options": [
                                    {"key": "A", "content": "Danh sách dữ liệu phải được sắp xếp theo thứ tự (tăng dần hoặc giảm dần).", "is_correct": True},
                                    {"key": "B", "content": "Danh sách phải có số lượng phần tử là một số chẵn.", "is_correct": False},
                                    {"key": "C", "content": "Tất cả phần tử trong danh sách bắt buộc phải là số dương.", "is_correct": False},
                                    {"key": "D", "content": "Danh sách không được chứa quá 1000 phần tử.", "is_correct": False}
                                ]
                            }
                        ]
                    },
                    {
                        "lesson_id": "Bài 29",
                        "lesson_title": "Thuật toán sắp xếp nổi bọt và sắp xếp chèn",
                        "page_num": 145,
                        "source_file": "SGK Tin học 10 KNTT (Định hướng CS).pdf",
                        "requirements": [
                            {"code": "YCCD_10_29_01", "content": "Trình bày được ý tưởng thuật toán sắp xếp nổi bọt (Bubble Sort) thông qua so sánh và đổi chỗ hai phần tử kề nhau."},
                            {"code": "YCCD_10_29_02", "content": "Mô phỏng từng bước chạy tay (Dry-run) và cài đặt được thuật toán trên Python."}
                        ],
                        "knowledges": [
                            {"name": "Thuật toán sắp xếp nổi bọt (Bubble Sort)", "content": "Duyệt nhiều lượt, mỗi lượt so sánh 2 phần tử liền kề và hoán đổi nếu sai thứ tự, đưa phần tử lớn nhất nổi dần về cuối mảng."},
                            {"name": "Kỹ thuật hoán đổi giá trị trong Python", "content": "Cú pháp đồng thời thanh lịch a[i], a[j] = a[j], a[i] không cần dùng biến phụ trung gian."}
                        ],
                        "skills": [
                            {"name": "Kỹ năng Dry-run thuật toán sắp xếp", "description": "Lập bảng theo dõi trạng thái các phần tử của mảng qua từng vòng lặp i và j."},
                            {"name": "Kỹ năng tối ưu hóa cờ hiệu (Flag Optimization)", "description": "Dùng biến cờ swapped để dừng sớm nếu mảng đã được sắp xếp hoàn toàn."}
                        ],
                        "visuals": [
                            {"code": "HINH_10_29_01", "type": "Sơ đồ trực quan hóa", "caption": "Các bước hoán đổi phần tử trong thuật toán Bubble Sort", "desc": "Minh họa 4 bước của lượt 1 đưa số 9 từ vị trí đầu mảng về vị trí cuối mảng."}
                        ],
                        "questions": [
                            {
                                "id": "CH_10_29_01",
                                "content": "Trong thuật toán sắp xếp nổi bọt với mảng có n phần tử, số lần so sánh tối đa ở trường hợp xấu nhất là bao nhiêu?",
                                "difficulty": "Vận dụng",
                                "options": [
                                    {"key": "A", "content": "n * (n - 1) / 2 phép so sánh (độ phức tạp O(n^2)).", "is_correct": True},
                                    {"key": "B", "content": "Chỉ đúng n phép so sánh.", "is_correct": False},
                                    {"key": "C", "content": "log2(n) phép so sánh.", "is_correct": False},
                                    {"key": "D", "content": "n * log2(n) phép so sánh.", "is_correct": False}
                                ]
                            }
                        ]
                    }
                ]
            },
            {
                "id": "Lớp 10_Chuyên đề: Ứng dụng Trí tuệ Nhân tạo và Khoa học Dữ liệu",
                "title": "Chuyên đề: Ứng dụng Trí tuệ Nhân tạo và Khoa học Dữ liệu trong GDPT 2018",
                "lessons": [
                    {
                        "lesson_id": "Bài 30",
                        "lesson_title": "Nhập môn Trí tuệ Nhân tạo và Học máy",
                        "page_num": 160,
                        "source_file": "Chuyên đề học tập Tin học 10 KNTT.pdf",
                        "requirements": [
                            {"code": "YCCD_10_30_01", "content": "Hiểu định nghĩa cơ bản về AI, phân biệt được Trí tuệ nhân tạo hẹp (Narrow AI) và Trí tuệ nhân tạo tổng quát (AGI)."},
                            {"code": "YCCD_10_30_02", "content": "Nêu được các ví dụ ứng dụng AI trong nhận dạng giọng nói, thị giác máy tính và xử lý ngôn ngữ tự nhiên."}
                        ],
                        "knowledges": [
                            {"name": "Trí tuệ nhân tạo (Artificial Intelligence - AI)", "content": "Ngành khoa học máy tính nghiên cứu cách thức tạo ra máy móc có khả năng mô phỏng các năng lực trí tuệ của con người như học tập, suy luận và giải quyết vấn đề."},
                            {"name": "Học máy (Machine Learning)", "content": "Phân ngành của AI cho phép hệ thống tự động học và cải thiện hiệu năng từ dữ liệu kinh nghiệm mà không cần lập trình quy tắc tường minh."}
                        ],
                        "skills": [
                            {"name": "Kỹ năng đánh giá tác động đạo đức của AI", "description": "Phân tích vấn đề thiên kiến dữ liệu (Bias), quyền riêng tư dữ liệu cá nhân và an toàn thông tin khi sử dụng AI."}
                        ],
                        "visuals": [
                            {"code": "HINH_10_30_01", "type": "Sơ đồ Venn", "caption": "Mối quan hệ phân tầng giữa AI, Machine Learning và Deep Learning", "desc": "Sơ đồ vòng tròn đồng tâm thể hiện Deep Learning nằm trong Machine Learning và Machine Learning nằm trong AI."}
                        ],
                        "questions": [
                            {
                                "id": "CH_10_30_01",
                                "content": "Đặc trưng cốt lõi nào phân biệt Học máy (Machine Learning) với lập trình giải thuật truyền thống?",
                                "difficulty": "Thông hiểu",
                                "options": [
                                    {"key": "A", "content": "Học máy tự động tìm ra quy luật/mô hình dựa trên dữ liệu đầu vào và đáp án mẫu, thay vì con người phải tự viết luật thủ công.", "is_correct": True},
                                    {"key": "B", "content": "Học máy không cần máy tính để thực thi.", "is_correct": False},
                                    {"key": "C", "content": "Học máy chỉ chạy được với các số nguyên dương.", "is_correct": False},
                                    {"key": "D", "content": "Học máy luôn cho độ chính xác 100% không bao giờ có sai số.", "is_correct": False}
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    }
    
    # Quan hệ tri thức tiên quyết & liên kết ngữ nghĩa
    prerequisites = [
        ("Biến và Định danh trong Python", "Cấu trúc dữ liệu List trong Python"),
        ("Cấu trúc dữ liệu List trong Python", "Thuật toán tìm kiếm tuần tự (Linear Search)"),
        ("Thuật toán sắp xếp nổi bọt (Bubble Sort)", "Thuật toán tìm kiếm nhị phân (Binary Search)"),
        ("Thuật toán tìm kiếm nhị phân (Binary Search)", "Học máy (Machine Learning)")
    ]

    print("[2/3] Đang thực thi nạp dữ liệu vào Neo4j Live...")
    with driver.session() as session:
        # Nạp Grade
        session.run("MERGE (g:Grade {grade_level: $grade})", grade=new_data["grade"])
        
        for ch in new_data["chapters"]:
            session.run("""
                MATCH (g:Grade {grade_level: $grade})
                MERGE (c:Chapter {id: $ch_id})
                ON CREATE SET c.title = $ch_title, c.grade_level = $grade
                MERGE (g)-[:HAS_CHAPTER]->(c)
            """, grade=new_data["grade"], ch_id=ch["id"], ch_title=ch["title"])
            
            for les in ch["lessons"]:
                les_node_id = f"{new_data['grade']}_{les['lesson_id']}"
                session.run("""
                    MATCH (c:Chapter {id: $ch_id})
                    MERGE (l:Lesson {id: $les_id})
                    ON CREATE SET l.lesson_id = $lid, l.title = $title, l.page = $page, l.source_file = $src, l.grade_level = $grade
                    MERGE (c)-[:HAS_LESSON]->(l)
                """, ch_id=ch["id"], les_id=les_node_id, lid=les["lesson_id"], title=les["lesson_title"], page=les["page_num"], src=les["source_file"], grade=new_data["grade"])
                
                # Requirements
                for req in les.get("requirements", []):
                    req_id = f"{les_node_id}_{req['code']}"
                    session.run("""
                        MATCH (l:Lesson {id: $les_id})
                        MERGE (r:Requirement {id: $req_id})
                        ON CREATE SET r.code = $code, r.content = $content, r.grade_level = $grade
                        MERGE (l)-[:MEETS_REQUIREMENT]->(r)
                        MERGE (l)-[:HAS_REQUIREMENT]->(r)
                    """, les_id=les_node_id, req_id=req_id, code=req["code"], content=req["content"], grade=new_data["grade"])
                
                # Knowledge
                for kn in les.get("knowledges", []):
                    kn_id = f"{les_node_id}_{kn['name']}"
                    session.run("""
                        MATCH (l:Lesson {id: $les_id})
                        MERGE (k:Knowledge {id: $kn_id})
                        ON CREATE SET k.name = $name, k.content = $content, k.grade_level = $grade
                        MERGE (l)-[:COVERS_KNOWLEDGE]->(k)
                        MERGE (l)-[:TEACHES_KNOWLEDGE]->(k)
                    """, les_id=les_node_id, kn_id=kn_id, name=kn["name"], content=kn["content"], grade=new_data["grade"])
                
                # Skills
                for sk in les.get("skills", []):
                    sk_id = f"{les_node_id}_{sk['name']}"
                    session.run("""
                        MATCH (l:Lesson {id: $les_id})
                        MERGE (s:Skill {id: $sk_id})
                        ON CREATE SET s.name = $name, s.description = $desc, s.grade_level = $grade
                        MERGE (l)-[:DEVELOPS_SKILL]->(s)
                    """, les_id=les_node_id, sk_id=sk_id, name=sk["name"], desc=sk["description"], grade=new_data["grade"])
                
                # Visual Assets
                for vis in les.get("visuals", []):
                    vis_id = f"{les_node_id}_{vis['code']}"
                    session.run("""
                        MATCH (l:Lesson {id: $les_id})
                        MERGE (v:VisualAsset {id: $vis_id})
                        ON CREATE SET v.code = $code, v.caption = $caption, v.description = $desc, v.detailed_description = $desc, v.type = $type, v.page = $page, v.grade_level = $grade
                        MERGE (l)-[:HAS_VISUAL_ASSET]->(v)
                    """, les_id=les_node_id, vis_id=vis_id, code=vis["code"], caption=vis["caption"], desc=vis["desc"], type=vis["type"], page=les["page_num"], grade=new_data["grade"])
                
                # Questions & Options
                for q in les.get("questions", []):
                    q_id = f"{les_node_id}_{q['id']}"
                    session.run("""
                        MATCH (l:Lesson {id: $les_id})
                        MERGE (que:Question {id: $q_id})
                        ON CREATE SET que.question_id = $qid, que.content = $content, que.difficulty = $diff, que.grade_level = $grade
                        MERGE (l)-[:TESTS_KNOWLEDGE]->(que)
                        MERGE (l)-[:EVALUATED_BY]->(que)
                    """, les_id=les_node_id, q_id=q_id, qid=q["id"], content=q["content"], diff=q["difficulty"], grade=new_data["grade"])
                    
                    for opt in q.get("options", []):
                        session.run("""
                            MATCH (que:Question {id: $q_id})
                            CREATE (a:AnswerOption {option_key: $key, content: $content, is_correct: $is_cor, question_id: $q_id})
                            MERGE (que)-[:HAS_OPTION]->(a)
                        """, q_id=q_id, key=opt["key"], content=opt["content"], is_cor=opt["is_correct"])

        # Nạp quan hệ tiên quyết PREREQUISITE_FOR
        print("-> Nạp các quan hệ tri thức tiên quyết (:PREREQUISITE_FOR)...")
        for src_name, tgt_name in prerequisites:
            session.run("""
                MATCH (k1:Knowledge {name: $src})
                MATCH (k2:Knowledge {name: $tgt})
                MERGE (k1)-[:PREREQUISITE_FOR]->(k2)
            """, src=src_name, tgt=tgt_name)

    print("[3/3] Kiểm tra lại tổng số Nodes và Relationships sau khi nạp...")
    with driver.session() as session:
        t_nodes = session.run("MATCH (n) RETURN count(n) AS c").single()["c"]
        t_rels = session.run("MATCH ()-[r]->() RETURN count(r) AS c").single()["c"]
        g10_nodes = session.run("MATCH (n {grade_level: 'Lớp 10'}) RETURN count(n) AS c").single()["c"]
        g10_rels = session.run("MATCH (n {grade_level: 'Lớp 10'})-[r]->() RETURN count(r) AS c").single()["c"]
        print(f"==> TỔNG SỐ NODES TRONG TOÀN BỘ CSDL: {t_nodes}")
        print(f"==> TỔNG SỐ RELATIONSHIPS: {t_rels}")
        print(f"==> DỮ LIỆU MỚI (LỚP 10): {g10_nodes} nodes, {g10_rels} relationships đã nạp thành công!")

    driver.close()

if __name__ == "__main__":
    add_new_curriculum_data()
