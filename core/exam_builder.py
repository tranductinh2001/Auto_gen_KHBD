"""
core/exam_builder.py
Module chuyên sâu: Xây dựng Ma trận, Bản đặc tả và Sinh bộ hồ sơ Đề thi chuẩn GDPT 2018
Định dạng file Word (.docx) chuyên nghiệp, đầy đủ 100% các phần theo mẫu của Sở/Phòng GD&ĐT.
"""

import os
import sys
import copy
import random
from typing import List, Dict, Any, Tuple, Optional

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


# ==============================================================================
# 1. NGÂN HÀNG DỮ LIỆU ĐẶC TẢ VÀ CÂU HỎI CHUẨN GDPT 2018 (LỚP 6, 7, 8, 9)
# ==============================================================================

CURRICULUM_DATA: Dict[str, Dict[str, Any]] = {
    "Lớp 6": {
        "topics": [
            {
                "tt": "1",
                "chapter": "Chủ đề 1. Máy tính và cộng đồng",
                "units": [
                    {
                        "unit_name": "1. Thông tin và dữ liệu",
                        "req_nb": "Trong các tình huống cụ thể có sẵn:\n- Phân biệt được thông tin với vật mang tin.\n- Nhận biết được sự khác nhau giữa thông tin và dữ liệu.\n- Nêu được các bước cơ bản trong xử lí thông tin.",
                        "req_th": "- Nêu được ví dụ minh hoạ về mối quan hệ giữa thông tin và dữ liệu.\n- Nêu được ví dụ minh hoạ tầm quan trọng của thông tin.",
                        "req_vd": "- Giải thích được máy tính và các thiết bị số là công cụ hiệu quả để thu thập, lưu trữ, xử lí và truyền thông tin.\n- Phân loại được những công việc trong cuộc sống theo các hoạt động của quá trình xử lý thông tin.",
                        "req_vdc": "Đề xuất được phương án tổ chức và xử lý thông tin số hoá cho bài toán thực tiễn của trường học."
                    },
                    {
                        "unit_name": "2. Biểu diễn thông tin và lưu trữ trong máy tính",
                        "req_nb": "- Biết được bit là đơn vị nhỏ nhất trong lưu trữ thông tin.\n- Nêu được tên và độ lớn xấp xỉ của các đơn vị đo: Byte, KB, MB, GB, TB.",
                        "req_th": "- Giải thích được vì sao thông tin trong máy tính được biểu diễn bằng dãy bit (0 và 1).",
                        "req_vd": "- Tính toán và xác định được khả năng lưu trữ của các thiết bị nhớ thông dụng (USB, đĩa cứng, thẻ nhớ, bộ nhớ đám mây).",
                        "req_vdc": "Ước lượng tối ưu dung lượng lưu trữ cho dữ liệu đa phương tiện trong dự án học tập."
                    }
                ]
            },
            {
                "tt": "2",
                "chapter": "Chủ đề 2. Mạng máy tính và Internet",
                "units": [
                    {
                        "unit_name": "1. Giới thiệu về mạng máy tính và Internet",
                        "req_nb": "- Nêu được khái niệm và lợi ích cơ bản của mạng máy tính.\n- Kể tên các thành phần chủ yếu của mạng máy tính (thiết bị đầu cuối, thiết bị kết nối, phần mềm mạng).",
                        "req_th": "- Phân biệt được đặc điểm giữa mạng có dây và mạng không dây.\n- Nêu được các dịch vụ tiêu biểu trên Internet.",
                        "req_vd": "- Nêu ví dụ cụ thể về trường hợp mạng không dây tiện dụng hơn mạng có dây trong đời sống và học tập.",
                        "req_vdc": "Đánh giá lựa chọn giải pháp kết nối mạng an toàn, tiết kiệm chi phí cho phòng tin học."
                    }
                ]
            }
        ],
        "mcq_bank": [
            {
                "id": "6_MCQ_01",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Nhận biết",
                "content": "Phương án nào sau đây KHÔNG nêu đúng ví dụ về vật mang tin?",
                "options": {
                    "A": "Quyển sách giáo khoa Tin học 6",
                    "B": "Tấm biển báo giao thông trên đường",
                    "C": "Suy nghĩ xuất hiện trong đầu của em",
                    "D": "Thẻ nhớ và ổ đĩa USB"
                },
                "correct": "C"
            },
            {
                "id": "6_MCQ_02",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Nhận biết",
                "content": "Phương án nào sau đây chỉ ra đúng thứ tự các bước cơ bản trong quá trình xử lý thông tin của con người?",
                "options": {
                    "A": "Thu nhận thông tin → Xử lý thông tin → Lưu trữ thông tin → Truyền (trao đổi) thông tin",
                    "B": "Xử lý thông tin → Thu nhận thông tin → Lưu trữ thông tin",
                    "C": "Thu nhận thông tin → Trao đổi thông tin → Xử lý thông tin",
                    "D": "Lưu trữ thông tin → Xử lý thông tin → Thu nhận thông tin"
                },
                "correct": "A"
            },
            {
                "id": "6_MCQ_03",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Nhận biết",
                "content": "Các số, văn bản, hình ảnh, âm thanh ghi lại sự vật hiện tượng được gọi là gì?",
                "options": {
                    "A": "Vật mang tin",
                    "B": "Dữ liệu",
                    "C": "Thông tin",
                    "D": "Tri thức"
                },
                "correct": "B"
            },
            {
                "id": "6_MCQ_04",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Nhận biết",
                "content": "Máy tính điện tử gồm có bao nhiêu khối thành phần chính để thực hiện quá trình xử lý thông tin?",
                "options": {
                    "A": "2 thành phần",
                    "B": "3 thành phần",
                    "C": "4 thành phần (Vào, Ra, Bộ nhớ, Bộ xử lý)",
                    "D": "5 thành phần"
                },
                "correct": "C"
            },
            {
                "id": "6_MCQ_05",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Thông hiểu",
                "content": "Thao tác ghi nhớ nội dung bài học vào sổ tay của học sinh tương ứng với hoạt động nào trong quá trình xử lý thông tin?",
                "options": {
                    "A": "Thu nhận thông tin",
                    "B": "Lưu trữ thông tin",
                    "C": "Xử lý thông tin",
                    "D": "Trao đổi thông tin"
                },
                "correct": "B"
            },
            {
                "id": "6_MCQ_06",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Thông hiểu",
                "content": "Bảng chỉ dẫn phân làn đường gắn tại ngã tư là:",
                "options": {
                    "A": "Vật mang tin chứa thông tin giao thông",
                    "B": "Dữ liệu dạng âm thanh",
                    "C": "Bộ nhớ của máy tính",
                    "D": "Phần mềm hệ thống"
                },
                "correct": "A"
            },
            {
                "id": "6_MCQ_07",
                "unit": "2. Biểu diễn thông tin và lưu trữ trong máy tính",
                "level": "Nhận biết",
                "content": "1 Megabyte (MB) xấp xỉ bằng bao nhiêu Byte?",
                "options": {
                    "A": "Xấp xỉ 1 triệu byte (khoảng 1.000.000 byte)",
                    "B": "Xấp xỉ 1 nghìn byte",
                    "C": "Xấp xỉ 1 tỉ byte",
                    "D": "Xấp xỉ 100 byte"
                },
                "correct": "A"
            },
            {
                "id": "6_MCQ_08",
                "unit": "2. Biểu diễn thông tin và lưu trữ trong máy tính",
                "level": "Nhận biết",
                "content": "Đơn vị đo dung lượng thông tin nào là nhỏ nhất trong các đơn vị dưới đây?",
                "options": {
                    "A": "Byte",
                    "B": "Bit",
                    "C": "Kilobyte (KB)",
                    "D": "Megabyte (MB)"
                },
                "correct": "B"
            },
            {
                "id": "6_MCQ_09",
                "unit": "2. Biểu diễn thông tin và lưu trữ trong máy tính",
                "level": "Thông hiểu",
                "content": "Phát biểu nào sau đây KHÔNG chính xác về biểu diễn thông tin trong máy tính?",
                "options": {
                    "A": "Máy tính có thể xử lý trực tiếp âm thanh và hình ảnh nguyên bản ngoài đời thực mà không cần mã hóa",
                    "B": "Mọi dữ liệu trong máy tính đều được biểu diễn dưới dạng các dãy bit 0 và 1",
                    "C": "Dãy bit là tập hợp chỉ gồm các kí hiệu 0 và 1",
                    "D": "Một byte gồm có 8 bit ghép lại"
                },
                "correct": "A"
            },
            {
                "id": "6_MCQ_10",
                "unit": "1. Giới thiệu về mạng máy tính và Internet",
                "level": "Nhận biết",
                "content": "Thành phần nào dưới đây thuộc nhóm thiết bị kết nối mạng?",
                "options": {
                    "A": "Bộ định tuyến (Router) và Switch",
                    "B": "Màn hình và chuột",
                    "C": "Loa và tai nghe",
                    "D": "Máy in màu"
                },
                "correct": "A"
            },
            {
                "id": "6_MCQ_11",
                "unit": "1. Giới thiệu về mạng máy tính và Internet",
                "level": "Vận dụng",
                "content": "Trong trường hợp nào dưới đây mạng không dây (Wi-Fi) tiện dụng và ưu thế hơn mạng có dây?",
                "options": {
                    "A": "Khi người dùng sử dụng điện thoại thông minh, laptop di chuyển liên tục trong khuôn viên trường",
                    "B": "Khi cần tốc độ truyền dữ liệu khổng lồ tuyệt đối không bị nhiễu sóng trong phòng máy chủ",
                    "C": "Khi lắp đặt hệ thống máy tính bàn cố định 20 năm không thay đổi vị trí",
                    "D": "Khi không có nguồn điện dự phòng"
                },
                "correct": "A"
            },
            {
                "id": "6_MCQ_12",
                "unit": "1. Giới thiệu về mạng máy tính và Internet",
                "level": "Thông hiểu",
                "content": "Phát biểu nào sau đây nêu SAI đặc điểm của mạng máy tính?",
                "options": {
                    "A": "Mạng máy tính chỉ có thể kết nối các máy tính đặt chung trong một phòng",
                    "B": "Mạng máy tính giúp người dùng chia sẻ tệp tin, dữ liệu và máy in dùng chung",
                    "C": "Internet là mạng kết nối hàng triệu máy tính trên toàn thế giới",
                    "D": "Mạng không dây sử dụng sóng điện từ để truyền thông tin"
                },
                "correct": "A"
            }
        ],
        "tf_bank": [
            {
                "id": "6_TF_01",
                "unit": "1. Thông tin và dữ liệu",
                "topic": "Hoa nhìn thấy giấy khen của bạn Nam dán trên bảng tin lớp và thấy nó rất đẹp.",
                "statements": [
                    {"label": "a", "content": "Tờ giấy khen là vật mang tin.", "correct": "Đ", "reason": "Giấy khen chứa chữ và hình ảnh được in ra."},
                    {"label": "b", "content": "Mắt Hoa tiếp nhận hình ảnh giấy khen là quá trình xử lý thông tin.", "correct": "S", "reason": "Đây là hoạt động thu nhận thông tin qua giác quan."},
                    {"label": "c", "content": "Dòng chữ ghi 'Học sinh xuất sắc' trên giấy khen là dữ liệu văn bản.", "correct": "Đ", "reason": "Văn bản được ghi lại là dữ liệu."},
                    {"label": "d", "content": "Hoa khen bạn Nam học giỏi là hoạt động trao đổi (truyền) thông tin.", "correct": "Đ", "reason": "Hoa bày tỏ và chia sẻ ý kiến với mọi người."}
                ]
            },
            {
                "id": "6_TF_02",
                "unit": "2. Biểu diễn thông tin và lưu trữ trong máy tính",
                "topic": "Khi nghiên cứu về nguyên lý biểu diễn và lưu trữ dữ liệu trong máy tính điện tử:",
                "statements": [
                    {"label": "a", "content": "Bit là đơn vị nhỏ nhất để đo dung lượng lưu trữ thông tin trong máy tính.", "correct": "Đ", "reason": "Bit chỉ nhận 1 trong 2 giá trị 0 hoặc 1."},
                    {"label": "b", "content": "1 Kilobyte (KB) có độ lớn chính xác bằng 100 Byte.", "correct": "S", "reason": "1 KB = 1024 Byte (xấp xỉ 1.000 Byte)."},
                    {"label": "c", "content": "Một tệp văn bản có dung lượng 2 MB sẽ có kích thước xấp xỉ 2 triệu Byte.", "correct": "Đ", "reason": "1 MB xấp xỉ 1 triệu byte."},
                    {"label": "d", "content": "Các thiết bị nhớ như thẻ nhớ, đĩa cứng chỉ lưu trữ được chữ viết, không thể lưu trữ âm thanh hay video.", "correct": "S", "reason": "Thiết bị nhớ số lưu trữ mọi định dạng đa phương tiện đã được số hóa."}
                ]
            },
            {
                "id": "6_TF_03",
                "unit": "1. Giới thiệu về mạng máy tính và Internet",
                "topic": "Xét về cấu trúc và vai trò của mạng máy tính trong cuộc sống số hiện nay:",
                "statements": [
                    {"label": "a", "content": "Mạng máy tính gồm ba thành phần cơ bản: các thiết bị đầu cuối, các thiết bị kết nối và phần mềm mạng.", "correct": "Đ", "reason": "Theo SGK Tin học 6 GDPT 2018."},
                    {"label": "b", "content": "Switch và Router là các thiết bị đầu cuối giống như máy in và điện thoại.", "correct": "S", "reason": "Switch và Router là thiết bị kết nối mạng."},
                    {"label": "c", "content": "Mạng không dây truyền tín hiệu bằng sóng vô tuyến thay vì dây cáp quang hoặc cáp xoắn đôi.", "correct": "Đ", "reason": "Sóng điện từ/vô tuyến truyền tải trong không gian."},
                    {"label": "d", "content": "Khi tham gia mạng máy tính, người dùng không cần quan tâm đến an toàn thông tin hay mật khẩu.", "correct": "S", "reason": "Bảo vệ thông tin tài khoản là nguyên tắc sống còn."}
                ]
            },
            {
                "id": "6_TF_04",
                "unit": "1. Thông tin và dữ liệu",
                "topic": "Trong các hoạt động thông tin thường ngày của học sinh phổ thông:",
                "statements": [
                    {"label": "a", "content": "Tiếng trống trường vang lên báo hiệu giờ ra chơi là thông tin dạng âm thanh.", "correct": "Đ", "reason": "Học sinh nghe thấy tiếng trống và hiểu được thông điệp."},
                    {"label": "b", "content": "Việc đọc sách để tích lũy kiến thức là quá trình thu nhận và xử lý thông tin.", "correct": "Đ", "reason": "Bộ não con người tiếp nhận và phân tích nội dung."},
                    {"label": "c", "content": "Thông tin và dữ liệu luôn có ý nghĩa và giá trị giống hệt nhau đối với mọi người trong mọi hoàn cảnh.", "correct": "S", "reason": "Dữ liệu chỉ trở thành thông tin có ích khi được con người hiểu và sử dụng đúng bối cảnh."},
                    {"label": "d", "content": "Máy tính điện tử có thể hoàn toàn thay thế mọi cảm xúc và trực giác sáng tạo của con người.", "correct": "S", "reason": "Máy tính chỉ là công cụ hỗ trợ xử lý theo chỉ dẫn."}
                ]
            }
        ],
        "essay_bank": [
            {
                "id": "6_ES_01",
                "unit": "1. Giới thiệu về mạng máy tính và Internet",
                "level": "Thông hiểu",
                "score": 1.0,
                "content": "Em hãy kể tên 4 dịch vụ tiêu biểu và phổ biến nhất của mạng máy tính và Internet giúp ích cho học sinh trong học tập.",
                "rubric": [
                    {"criterion": "Nêu được dịch vụ tìm kiếm thông tin (Google Search, Bing...)", "score": 0.25},
                    {"criterion": "Nêu được dịch vụ thư điện tử (Email) hoặc trao đổi tin nhắn học tập", "score": 0.25},
                    {"criterion": "Nêu được dịch vụ học trực tuyến và lưu trữ tài liệu (Google Classroom, Drive...)", "score": 0.25},
                    {"criterion": "Nêu được dịch vụ giải trí lành mạnh hoặc chia sẻ video bài giảng (YouTube Edu...)", "score": 0.25}
                ]
            },
            {
                "id": "6_ES_02",
                "unit": "1. Thông tin và dữ liệu",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Hãy phân tích hoạt động thông tin của con người trong tình huống: 'Khi nghe thấy tiếng còi xe cứu thương từ phía sau, người lái xe lập tức quan sát gương chiếu hậu và điều khiển xe tấp vào lề đường bên phải để nhường đường'.",
                "rubric": [
                    {"criterion": "Chỉ ra khâu thu nhận thông tin: Tai nghe tiếng còi, mắt quan sát gương", "score": 0.25},
                    {"criterion": "Chỉ ra khâu xử lý thông tin: Nhận biết xe ưu tiên khẩn cấp và cần phải nhường đường", "score": 0.25},
                    {"criterion": "Chỉ ra khâu lưu trữ/tra cứu tri thức: Vận dụng luật giao thông đường bộ đã học", "score": 0.25},
                    {"criterion": "Chỉ ra khâu hành động/truyền thông tin: Bật xi-nhan và tấp xe sang phải an toàn", "score": 0.25}
                ]
            },
            {
                "id": "6_ES_03",
                "unit": "2. Biểu diễn thông tin và lưu trữ trong máy tính",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Giả sử một bộ phim hoạt hình ngắn có dung lượng khoảng 500 MB. Hỏi một thẻ nhớ dung lượng 16 GB có thể lưu trữ tối đa được khoảng bao nhiêu bộ phim hoạt hình như vậy? (Biết 1 GB xấp xỉ bằng 1000 MB).",
                "rubric": [
                    {"criterion": "Đổi đơn vị đúng: 16 GB = 16 x 1000 MB = 16.000 MB", "score": 0.5},
                    {"criterion": "Thực hiện phép chia: 16.000 / 500 = 32", "score": 0.25},
                    {"criterion": "Kết luận rõ ràng: Thẻ nhớ có thể chứa tối đa khoảng 32 bộ phim", "score": 0.25}
                ]
            }
        ]
    },

    "Lớp 7": {
        "topics": [
            {
                "tt": "1",
                "chapter": "Chủ đề 1. Máy tính và cộng đồng (5 tiết)",
                "units": [
                    {
                        "unit_name": "1.1. Sơ lược về các thành phần của máy tính",
                        "req_nb": "Biết và nhận ra được các thiết bị vào ra trong mô hình thiết bị máy tính, tính đa dạng và hình dạng của các thiết bị. Biết được chức năng của một số thiết bị vào ra trong thu thập, lưu trữ, xử lí và truyền thông tin.",
                        "req_th": "Nêu được ví dụ cụ thể về những thao tác không đúng cách sẽ gây ra lỗi cho các thiết bị và hệ thống xử lí thông tin.",
                        "req_vd": "Thực hiện đúng các thao tác với các thiết bị thông dụng của máy tính.",
                        "req_vdc": "Khắc phục được các sự cố kết nối cơ bản của thiết bị ngoại vi với thân máy."
                    },
                    {
                        "unit_name": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                        "req_nb": "Biết được tệp chương trình cũng là dữ liệu, có thể được lưu trữ trong máy tính. Nêu được tên một số phần mềm ứng dụng đã sử dụng. Nêu được một số biện pháp để bảo vệ máy tính cá nhân, tài khoản và dữ liệu cá nhân.",
                        "req_th": "Giải thích được chức năng điều khiển của hệ điều hành, qua đó phân biệt được hệ điều hành với phần mềm ứng dụng. Phân biệt được loại tệp chương trình và tệp dữ liệu thông qua phần mở rộng.",
                        "req_vd": "Thao tác thành thạo với tệp và thư mục (tạo, đổi tên, sao chép, di chuyển, xóa trên cây thư mục).",
                        "req_vdc": "Tổ chức cấu trúc thư mục khoa học phục vụ quản lý dự án học tập nhóm."
                    }
                ]
            },
            {
                "tt": "2",
                "chapter": "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin (2 tiết)",
                "units": [
                    {
                        "unit_name": "2.1. Mạng xã hội và một số kênh trao đổi thông tin thông dụng",
                        "req_nb": "Nhận biết một số website là mạng xã hội. Nêu được một số chức năng cơ bản của mạng xã hội: kết nối, giao lưu, chia sẻ, thảo luận và trao đổi thông tin.",
                        "req_th": "Nêu được ví dụ cụ thể về hậu quả của việc sử dụng thông tin vào mục đích sai trái trên không gian mạng.",
                        "req_vd": "Sử dụng được một số chức năng cơ bản của một mạng xã hội để giao lưu và chia sẻ thông tin học tập an toàn.",
                        "req_vdc": "Xây dựng quy tắc ứng xử nhóm trực tuyến đảm bảo văn minh và bảo mật."
                    }
                ]
            },
            {
                "tt": "3",
                "chapter": "Chủ đề 3. Đạo đức, pháp luật và văn hóa trong môi trường số (1 tiết)",
                "units": [
                    {
                        "unit_name": "3.1. Văn hoá ứng xử qua phương tiện truyền thông số",
                        "req_nb": "Biết được cách ứng xử hợp lí khi gặp trên mạng hoặc các kênh truyền thông tin số những thông tin có nội dung xấu, thông tin không phù hợp lứa tuổi.",
                        "req_th": "Nêu được một số ví dụ truy cập không hợp lệ vào các nguồn thông tin và kênh truyền thông tin.",
                        "req_vd": "Thực hiện được giao tiếp qua mạng (trực tuyến hay không trực tuyến) theo đúng quy tắc và bằng ngôn ngữ lịch sự, thể hiện ứng xử có văn hoá.",
                        "req_vdc": "Tuyên truyền phòng chống bắt nạt qua mạng và bảo vệ bản quyền số."
                    }
                ]
            }
        ],
        "mcq_bank": [
            {
                "id": "7_MCQ_01",
                "unit": "1.1. Sơ lược về các thành phần của máy tính",
                "level": "Nhận biết",
                "content": "Thiết bị vào (Input device) có chức năng gì trong các phương án sau đây?",
                "options": {
                    "A": "Hiển thị kết quả xử lý ra màn hình",
                    "B": "Tiếp nhận dữ liệu từ thế giới bên ngoài và chuyển vào máy tính",
                    "C": "Lưu trữ dữ liệu lâu dài ngay cả khi mất điện",
                    "D": "In dữ liệu ra giấy"
                },
                "correct": "B"
            },
            {
                "id": "7_MCQ_02",
                "unit": "1.1. Sơ lược về các thành phần của máy tính",
                "level": "Nhận biết",
                "content": "Thiết bị nào sau đây vừa có thể đưa thông tin vào, vừa có thể đưa thông tin ra?",
                "options": {
                    "A": "Màn hình cảm ứng (Touch screen)",
                    "B": "Bàn phím cơ",
                    "C": "Chuột quang",
                    "D": "Máy quét (Scanner)"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_03",
                "unit": "1.1. Sơ lược về các thành phần của máy tính",
                "level": "Thông hiểu",
                "content": "Thao tác nào dưới đây KHÔNG nên làm khi sử dụng máy tính để tránh gây lỗi phần cứng hoặc mất dữ liệu?",
                "options": {
                    "A": "Rút phích cắm điện nguồn trực tiếp khi máy tính đang hoạt động",
                    "B": "Tắt máy bằng lệnh Shut down trong menu Start",
                    "C": "Sử dụng khăn mềm khô để lau bụi bám trên vỏ máy",
                    "D": "Đóng các ứng dụng trước khi tắt máy"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_04",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "level": "Nhận biết",
                "content": "Tệp có phần mở rộng .exe (Executable) thuộc loại tệp nào sau đây?",
                "options": {
                    "A": "Tệp âm thanh",
                    "B": "Tệp chương trình có thể thực thi",
                    "C": "Tệp hình ảnh",
                    "D": "Tệp tài liệu văn bản Word"
                },
                "correct": "B"
            },
            {
                "id": "7_MCQ_05",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "level": "Nhận biết",
                "content": "Trong các phần mềm dưới đây, phần mềm ứng dụng nào có chức năng giúp người dùng vẽ hình và xử lý ảnh đơn giản?",
                "options": {
                    "A": "MS Paint",
                    "B": "Calculator",
                    "C": "Notepad",
                    "D": "File Explorer"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_06",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "level": "Thông hiểu",
                "content": "Em hãy cho biết phần mềm nào sau đây đóng vai trò nền tảng và quan trọng nhất trên máy tính?",
                "options": {
                    "A": "Hệ điều hành (Windows, Linux, macOS)",
                    "B": "Phần mềm nghe nhạc",
                    "C": "Trình duyệt Web",
                    "D": "Game điện tử"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_07",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "level": "Vận dụng",
                "content": "Cách đặt mật khẩu tài khoản nào dưới đây được xem là mật khẩu mạnh và an toàn?",
                "options": {
                    "A": "12345678",
                    "B": "nguyenvana",
                    "C": "Thao#2026@Tin7",
                    "D": "password"
                },
                "correct": "C"
            },
            {
                "id": "7_MCQ_08",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "level": "Thông hiểu",
                "content": "Dựa vào phần mở rộng của các tệp tin dưới đây, em hãy chỉ ra đâu là tệp video?",
                "options": {
                    "A": "baihoc.docx",
                    "B": "hinhnen.png",
                    "C": "clip_lop7a.mp4",
                    "D": "bansoan.txt"
                },
                "correct": "C"
            },
            {
                "id": "7_MCQ_09",
                "unit": "2.1. Mạng xã hội và một số kênh trao đổi thông tin thông dụng",
                "level": "Nhận biết",
                "content": "Trong các địa chỉ website sau đây, địa chỉ nào là nền tảng mạng xã hội thông dụng?",
                "options": {
                    "A": "https://facebook.com",
                    "B": "https://moet.gov.vn",
                    "C": "https://wikipedia.org",
                    "D": "https://dantri.com.vn"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_10",
                "unit": "2.1. Mạng xã hội và một số kênh trao đổi thông tin thông dụng",
                "level": "Nhận biết",
                "content": "Mạng xã hội KHÔNG có chức năng nào sau đây?",
                "options": {
                    "A": "Tự động sửa chữa phần cứng máy tính bị hỏng khi kết nối mạng",
                    "B": "Kết nối và giao lưu với bạn bè, thầy cô",
                    "C": "Chia sẻ thông tin, hình ảnh và video học tập",
                    "D": "Tạo nhóm thảo luận chuyên môn"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_11",
                "unit": "3.1. Văn hoá ứng xử qua phương tiện truyền thông số",
                "level": "Thông hiểu",
                "content": "Cách ứng xử nào sau đây là hợp lý nhất khi em vô tình truy cập vào một trang web có nội dung xấu, độc hại?",
                "options": {
                    "A": "Đóng ngay trang web đó và thông báo cho cha mẹ hoặc thầy cô giáo",
                    "B": "Chia sẻ đường liên kết lên mạng xã hội để cảnh báo mọi người",
                    "C": "Tải các hình ảnh về máy tính để nghiên cứu",
                    "D": "Bình luận khiêu khích dưới bài viết"
                },
                "correct": "A"
            },
            {
                "id": "7_MCQ_12",
                "unit": "3.1. Văn hoá ứng xử qua phương tiện truyền thông số",
                "level": "Vận dụng",
                "content": "Hành vi nào dưới đây thể hiện văn hóa ứng xử đúng mực trên không gian mạng?",
                "options": {
                    "A": "Sử dụng ngôn ngữ lịch sự, tôn trọng quan điểm của người khác và kiểm chứng trước khi chia sẻ",
                    "B": "Dùng tài khoản ẩn danh để thóa mạ bạn cùng trường",
                    "C": "Sao chép nguyên văn bài viết của người khác mà không ghi nguồn tác giả",
                    "D": "Phát tán thông tin chưa được kiểm chứng để câu like"
                },
                "correct": "A"
            }
        ],
        "tf_bank": [
            {
                "id": "7_TF_01",
                "unit": "1.1. Sơ lược về các thành phần của máy tính",
                "topic": "Khi sử dụng máy tính cá nhân để học tập và làm việc, xét các thao tác sau:",
                "statements": [
                    {"label": "a", "content": "Bàn phím và chuột là các thiết bị vào, có chức năng đưa thông tin vào máy tính.", "correct": "Đ", "reason": "Chuột và bàn phím là input devices tiêu biểu."},
                    {"label": "b", "content": "Máy in và loa là các thiết bị ra, có chức năng đưa thông tin từ máy tính ra ngoài.", "correct": "Đ", "reason": "Máy in xuất văn bản ra giấy, loa xuất âm thanh."},
                    {"label": "c", "content": "Rút cáp USB đột ngột khi đang chép tệp lớn sẽ không thể làm hỏng tệp hay thiết bị.", "correct": "S", "reason": "Có thể gây hỏng bảng chỉ mục file hệ thống (FAT/NTFS) hoặc mất dữ liệu."},
                    {"label": "d", "content": "Vệ sinh bàn phím bằng nước ướt đẫm trực tiếp khi máy tính đang bật nguồn là an toàn.", "correct": "S", "reason": "Rất nguy hiểm, gây chập cháy mạch điện tử."}
                ]
            },
            {
                "id": "7_TF_02",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "topic": "Khi tìm hiểu về hệ điều hành và quản lý tệp trên hệ điều hành Windows:",
                "statements": [
                    {"label": "a", "content": "Hệ điều hành đóng vai trò trung gian điều khiển phần cứng và cung cấp môi trường chạy ứng dụng.", "correct": "Đ", "reason": "Chức năng cốt lõi của OS."},
                    {"label": "b", "content": "Trong cùng một thư mục, có thể tồn tại hai tệp có tên và phần mở rộng giống hệt nhau.", "correct": "S", "reason": "Hệ điều hành không cho phép trùng cả tên và phần mở rộng trong cùng thư mục."},
                    {"label": "c", "content": "Phần mở rộng của tệp (ví dụ .docx, .xlsx, .pptx) giúp hệ điều hành nhận biết ứng dụng phù hợp để mở tệp.", "correct": "Đ", "reason": "Phần mở rộng xác định loại định dạng tệp tin."},
                    {"label": "d", "content": "Người dùng không cần cài đặt hệ điều hành mà vẫn có thể chạy trực tiếp các phần mềm soạn thảo văn bản.", "correct": "S", "reason": "Phải có hệ điều hành mới cài và chạy được ứng dụng."}
                ]
            },
            {
                "id": "7_TF_03",
                "unit": "2.1. Mạng xã hội và một số kênh trao đổi thông tin thông dụng",
                "topic": "Để tham gia và sử dụng mạng xã hội một cách an toàn, lành mạnh:",
                "statements": [
                    {"label": "a", "content": "Mạng xã hội cho phép người dùng giao lưu, kết nối và học tập trực tuyến.", "correct": "Đ", "reason": "Mạng xã hội là công cụ kết nối số mạnh mẽ."},
                    {"label": "b", "content": "Chia sẻ công khai số điện thoại, địa chỉ nhà riêng và mật khẩu lên trang cá nhân là hành vi an toàn.", "correct": "S", "reason": "Tiềm ẩn nguy cơ lộ lọt dữ liệu và bị kẻ xấu lợi dụng lừa đảo."},
                    {"label": "c", "content": "Có thể dùng mạng xã hội để tổ chức học nhóm và chia sẻ tài liệu ôn thi.", "correct": "Đ", "reason": "Đây là ứng dụng tích cực trong giáo dục."},
                    {"label": "d", "content": "Tất cả các tin tức được đăng tải trên mạng xã hội đều chính xác 100% không bao giờ có tin giả.", "correct": "S", "reason": "Mạng xã hội có rất nhiều tin giả, cần phải kiểm chứng."}
                ]
            },
            {
                "id": "7_TF_04",
                "unit": "3.1. Văn hoá ứng xử qua phương tiện truyền thông số",
                "topic": "Về văn hóa ứng xử và bảo vệ quyền tác giả trong môi trường số:",
                "statements": [
                    {"label": "a", "content": "Khi trích dẫn tranh ảnh hoặc bài viết của người khác, cần ghi rõ nguồn và tác giả.", "correct": "Đ", "reason": "Thể hiện sự tôn trọng quyền tác giả và luật sở hữu trí tuệ."},
                    {"label": "b", "content": "Tham gia vào các hội nhóm kêu gọi bắt nạt, xúc phạm danh dự của bạn bè là hành vi vi phạm đạo đức và pháp luật.", "correct": "Đ", "reason": "Hành vi bạo lực mạng bị pháp luật nghiêm cấm."},
                    {"label": "c", "content": "Gửi tin nhắn rác hoặc liên kết chứa mã độc cho bạn bè chỉ là trò đùa vô hại không cần xin lỗi.", "correct": "S", "reason": "Đây là hành vi thiếu văn hóa và có thể gây hại nghiêm trọng cho thiết bị."},
                    {"label": "d", "content": "Khi phát hiện tài khoản cá nhân bị xâm nhập trái phép, cần đổi ngay mật khẩu và kích hoạt bảo mật 2 lớp.", "correct": "Đ", "reason": "Biện pháp xử lý ứng phó sự cố kịp thời."}
                ]
            }
        ],
        "essay_bank": [
            {
                "id": "7_ES_01",
                "unit": "1.1. Sơ lược về các thành phần của máy tính",
                "level": "Thông hiểu",
                "score": 1.0,
                "content": "a) Chuột máy tính và bàn phím thuộc nhóm thiết bị nào trong mô hình kiến trúc máy tính?\nb) Giả sử khi cắm chuột vào máy tính, con trỏ chuột không di chuyển được. Em hãy nêu 2 nguyên nhân khả dĩ và cách kiểm tra khắc phục.",
                "rubric": [
                    {"criterion": "Nêu đúng: Chuột và bàn phím thuộc nhóm thiết bị vào (Input devices)", "score": 0.5},
                    {"criterion": "Nêu nguyên nhân 1: Cổng cắm USB bị lỏng hoặc bám bụi -> Kiểm tra cắm lại cổng khác", "score": 0.25},
                    {"criterion": "Nêu nguyên nhân 2: Chuột không dây hết pin hoặc thiếu driver/máy tính bị treo", "score": 0.25}
                ]
            },
            {
                "id": "7_ES_02",
                "unit": "1.2. Khái niệm hệ điều hành và phần mềm ứng dụng",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Cho cây thư mục ổ đĩa D:\\ như sau:\nD:\\\n├── HOC_TAP\\\n│   ├── TIN_HOC\\\n│   │   └── BaiTap.docx\n│   └── TOAN_HOC\\\n└── GIAI_TRI\\\n    └── NhacVui.mp3\n\na) Viết đường dẫn tuyệt đối (Full Path) đến tệp 'BaiTap.docx'.\nb) Nêu các bước thao tác để di chuyển tệp 'BaiTap.docx' sang thư mục 'TOAN_HOC'.",
                "rubric": [
                    {"criterion": "Viết đúng đường dẫn: D:\\HOC_TAP\\TIN_HOC\\BaiTap.docx", "score": 0.5},
                    {"criterion": "Nêu đúng thao tác: Chọn tệp BaiTap.docx -> Nhấn Cut (Ctrl+X)", "score": 0.25},
                    {"criterion": "Mở thư mục TOAN_HOC -> Nhấn Paste (Ctrl+V)", "score": 0.25}
                ]
            },
            {
                "id": "7_ES_03",
                "unit": "3.1. Văn hoá ứng xử qua phương tiện truyền thông số",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Em hãy nêu 4 biện pháp thiết thực giúp học sinh trung học cơ sở bảo vệ tài khoản mạng xã hội và dữ liệu cá nhân của mình khi sử dụng Internet.",
                "rubric": [
                    {"criterion": "Đặt mật khẩu mạnh (gồm chữ hoa, chữ thường, số, ký tự đặc biệt) và đổi định kỳ", "score": 0.25},
                    {"criterion": "Không cung cấp mật khẩu, mã OTP hoặc thông tin nhạy cảm cho bất kỳ ai", "score": 0.25},
                    {"criterion": "Đăng xuất tài khoản khi sử dụng máy tính công cộng hoặc thiết bị của người khác", "score": 0.25},
                    {"criterion": "Bật tính năng xác thực hai yếu tố (2FA) và cẩn trọng khi bấm vào các liên kết lạ", "score": 0.25}
                ]
            }
        ]
    },

    "Lớp 8": {
        "topics": [
            {
                "tt": "1",
                "chapter": "Chủ đề 1. Máy tính và cộng đồng",
                "units": [
                    {
                        "unit_name": "1. Lịch sử phát triển máy tính điện tử",
                        "req_nb": "Nêu được các mốc phát triển chính của năm thế hệ máy tính điện tử. Nhận biết được xu hướng thu nhỏ kích thước, tăng tốc độ và tích hợp trí tuệ nhân tạo.",
                        "req_th": "Giải thích được vai trò của vi xử lý trong việc tạo nên cuộc cách mạng máy tính cá nhân.",
                        "req_vd": "So sánh được đặc trưng kỹ thuật giữa các thế hệ máy tính.",
                        "req_vdc": "Dự đoán xu thế phát triển máy tính lượng tử và điện toán biên."
                    }
                ]
            },
            {
                "tt": "2",
                "chapter": "Chủ đề 4. Ứng dụng tin học (Bảng tính điện tử)",
                "units": [
                    {
                        "unit_name": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                        "req_nb": "Nhận biết được địa chỉ tương đối, địa chỉ tuyệt đối và địa chỉ hỗn hợp trong phần mềm bảng tính. Biết cú pháp các hàm cơ bản (SUM, AVERAGE, MAX, MIN, COUNT, IF).",
                        "req_th": "Hiểu và giải thích được sự thay đổi của địa chỉ ô khi sao chép công thức chứa địa chỉ tương đối hoặc tuyệt đối.",
                        "req_vd": "Vận dụng được các hàm tính toán và địa chỉ phù hợp để giải quyết bảng tính quản lý điểm số hoặc chi tiêu gia đình.",
                        "req_vdc": "Thiết kế bảng tính tự động xếp loại học lực kết hợp hàm IF lồng nhau và định dạng có điều kiện."
                    }
                ]
            },
            {
                "tt": "3",
                "chapter": "Chủ đề 3. Đạo đức, pháp luật và văn hóa trong môi trường số",
                "units": [
                    {
                        "unit_name": "1. Bản quyền phần mềm và sở hữu trí tuệ số",
                        "req_nb": "Nêu được khái niệm bản quyền, phần mềm thương mại, phần mềm nguồn mở và phần mềm miễn phí.",
                        "req_th": "Phân biệt được hành vi tôn trọng bản quyền với các hành vi vi phạm quyền tác giả trong môi trường số.",
                        "req_vd": "Lựa chọn và sử dụng phần mềm có bản quyền hoặc phần mềm nguồn mở hợp pháp cho học tập.",
                        "req_vdc": "Đề xuất quy trình đăng ký giấy phép Creative Commons cho sản phẩm dự án học sinh."
                    }
                ]
            }
        ],
        "mcq_bank": [
            {
                "id": "8_MCQ_01",
                "unit": "1. Lịch sử phát triển máy tính điện tử",
                "level": "Nhận biết",
                "content": "Thế hệ máy tính điện tử đầu tiên (những năm 1940-1950) sử dụng linh kiện điện tử cốt lõi nào?",
                "options": {
                    "A": "Đèn điện tử chân không",
                    "B": "Bóng bán dẫn (Transistor)",
                    "C": "Mạch tích hợp (IC)",
                    "D": "Bộ vi xử lý siêu lớn (VLSI)"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_02",
                "unit": "1. Lịch sử phát triển máy tính điện tử",
                "level": "Nhận biết",
                "content": "Máy tính cá nhân (PC) bắt đầu bùng nổ mạnh mẽ từ thế hệ máy tính thứ mấy nhờ sự ra đời của bộ vi xử lý?",
                "options": {
                    "A": "Thế hệ thứ nhất",
                    "B": "Thế hệ thứ hai",
                    "C": "Thế hệ thứ ba",
                    "D": "Thế hệ thứ tư"
                },
                "correct": "D"
            },
            {
                "id": "8_MCQ_03",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Nhận biết",
                "content": "Ký hiệu nào dưới đây được sử dụng để cố định dòng hoặc cột tạo địa chỉ tuyệt đối trong phần mềm bảng tính Excel?",
                "options": {
                    "A": "Ký hiệu $",
                    "B": "Ký hiệu &",
                    "C": "Ký hiệu #",
                    "D": "Ký hiệu @"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_04",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Nhận biết",
                "content": "Trong Excel, hàm nào sau đây được dùng để tính trung bình cộng của một khối ô số?",
                "options": {
                    "A": "SUM",
                    "B": "AVERAGE",
                    "C": "MAX",
                    "D": "COUNT"
                },
                "correct": "B"
            },
            {
                "id": "8_MCQ_05",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Thông hiểu",
                "content": "Giả sử ô C1 có công thức '=A1+B1'. Khi sao chép công thức từ ô C1 xuống ô C2, công thức tại ô C2 sẽ tự động chuyển thành:",
                "options": {
                    "A": "=A2+B2",
                    "B": "=A1+B1",
                    "C": "=$A$1+$B$1",
                    "D": "=A1+B2"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_06",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Thông hiểu",
                "content": "Địa chỉ ô nào sau đây là địa chỉ tuyệt đối hoàn toàn?",
                "options": {
                    "A": "$C$5",
                    "B": "C$5",
                    "C": "$C5",
                    "D": "C5"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_07",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Vận dụng",
                "content": "Công thức '=IF(D5>=5, \"Đạt\", \"Chưa đạt\")' tại ô E5 sẽ trả về kết quả gì nếu giá trị tại ô D5 là 7.5?",
                "options": {
                    "A": "Đạt",
                    "B": "Chưa đạt",
                    "C": "7.5",
                    "D": "Báo lỗi #VALUE!"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_08",
                "unit": "1. Bản quyền phần mềm và sở hữu trí tuệ số",
                "level": "Nhận biết",
                "content": "Phần mềm nguồn mở (Open Source Software) có đặc điểm nổi bật nào sau đây?",
                "options": {
                    "A": "Mã nguồn được công khai, người dùng có quyền tự do nghiên cứu, sửa đổi và phân phối",
                    "B": "Tuyệt đối không ai được phép xem mã lập trình",
                    "C": "Bắt buộc phải trả phí bản quyền đắt tiền mỗi tháng",
                    "D": "Chỉ hoạt động được khi không có mạng Internet"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_09",
                "unit": "1. Bản quyền phần mềm và sở hữu trí tuệ số",
                "level": "Thông hiểu",
                "content": "Hành vi nào dưới đây được xem là vi phạm bản quyền phần mềm?",
                "options": {
                    "A": "Tải phần mềm thương mại có tính phí từ trang web bẻ khóa (crack) để dùng lậu",
                    "B": "Sử dụng phần mềm mã nguồn mở Linux theo đúng giấy phép GPL",
                    "C": "Mua bản quyền gói phần mềm chính hãng từ nhà sản xuất",
                    "D": "Dùng phiên bản dùng thử miễn phí do nhà sản xuất phát hành"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_10",
                "unit": "1. Lịch sử phát triển máy tính điện tử",
                "level": "Thông hiểu",
                "content": "Đặc điểm chung trong xu hướng phát triển từ thế hệ máy tính đầu tiên đến nay là:",
                "options": {
                    "A": "Kích thước ngày càng thu nhỏ, tiêu thụ ít điện năng hơn nhưng tốc độ tính toán tăng gấp hàng triệu lần",
                    "B": "Kích thước ngày càng to lớn chiếm trọn nhiều tòa nhà",
                    "C": "Giá thành ngày càng đắt đỏ và chỉ các chính phủ mới mua được",
                    "D": "Tốc độ tính toán chậm hơn nhưng bộ nhớ lớn hơn"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_11",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Vận dụng",
                "content": "Tại ô F2 cần tính Thuế = Doanh thu * Thuế suất (Thuế suất cố định tại ô B1 là 10%). Công thức chuẩn nhất để khi sao chép xuống các dòng dưới không bị sai là:",
                "options": {
                    "A": "=E2*$B$1",
                    "B": "=E2*B1",
                    "C": "=$E$2*B1",
                    "D": "=E2*B$1$"
                },
                "correct": "A"
            },
            {
                "id": "8_MCQ_12",
                "unit": "1. Bản quyền phần mềm và sở hữu trí tuệ số",
                "level": "Vận dụng",
                "content": "Khi thực hiện dự án thiết kế poster tuyên truyền cho trường, phương án hợp pháp và an toàn nhất về mặt bản quyền là:",
                "options": {
                    "A": "Sử dụng các hình ảnh miễn phí có giấy phép Creative Commons (CC) hoặc hình ảnh tự chụp",
                    "B": "Tải ảnh có đóng dấu mờ (watermark) của các nhiếp ảnh gia nổi tiếng và xóa mờ dấu đi",
                    "C": "Sao chép nguyên thiết kế của người khác rồi đề tên mình",
                    "D": "Chụp lén tài liệu mật của người khác đưa lên poster"
                },
                "correct": "A"
            }
        ],
        "tf_bank": [
            {
                "id": "8_TF_01",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "topic": "Khi thao tác với công thức và các loại địa chỉ ô trong bảng tính Excel:",
                "statements": [
                    {"label": "a", "content": "Địa chỉ tương đối sẽ tự động thay đổi tọa độ cột hoặc dòng tương ứng khi công thức được sao chép sang vị trí khác.", "correct": "Đ", "reason": "Bản chất của địa chỉ tương đối."},
                    {"label": "b", "content": "Địa chỉ $A$1 là địa chỉ tuyệt đối, tọa độ cột A và dòng 1 được giữ nguyên cố định khi sao chép công thức.", "correct": "Đ", "reason": "Có dấu $ trước cả tên cột và chỉ số dòng."},
                    {"label": "c", "content": "Địa chỉ A$1 là địa chỉ tuyệt đối hoàn toàn, không thể thay đổi cột khi kéo ngang.", "correct": "S", "reason": "Đây là địa chỉ hỗn hợp, chỉ cố định dòng 1, cột A vẫn thay đổi khi kéo ngang."},
                    {"label": "d", "content": "Có thể nhấn phím F4 để chuyển đổi nhanh giữa các kiểu địa chỉ ô khi đang nhập công thức.", "correct": "Đ", "reason": "Phím tắt chuẩn trong Excel."}
                ]
            },
            {
                "id": "8_TF_02",
                "unit": "1. Bản quyền phần mềm và sở hữu trí tuệ số",
                "topic": "Về các khía cạnh pháp lý của việc sở hữu, sử dụng và chia sẻ phần mềm:",
                "statements": [
                    {"label": "a", "content": "Phần mềm thương mại (Commercial software) yêu cầu người dùng phải mua bản quyền mới được phép sử dụng hợp pháp.", "correct": "Đ", "reason": "Theo Luật Sở hữu trí tuệ."},
                    {"label": "b", "content": "Mọi phần mềm miễn phí (Freeware) đều cho phép người dùng xem và sửa đổi mã nguồn theo ý muốn.", "correct": "S", "reason": "Freeware chỉ miễn phí sử dụng, không bắt buộc cung cấp mã nguồn."},
                    {"label": "c", "content": "Phần mềm nguồn mở (Open source) khuyến khích sự cộng tác phát triển và chia sẻ tri thức trong cộng đồng công nghệ.", "correct": "Đ", "reason": "Triết lý của FOSS (Free and Open Source Software)."},
                    {"label": "d", "content": "Việc bẻ khóa và chia sẻ công cụ bẻ khóa phần mềm thương mại lên mạng là hành vi vi phạm pháp luật nghiêm trọng.", "correct": "Đ", "reason": "Vi phạm Luật An ninh mạng và Luật Sở hữu trí tuệ."}
                ]
            },
            {
                "id": "8_TF_03",
                "unit": "1. Lịch sử phát triển máy tính điện tử",
                "topic": "Về tiến trình phát triển và các thế hệ máy tính điện tử:",
                "statements": [
                    {"label": "a", "content": "Máy tính ENIAC là một trong những chiếc máy tính điện tử đa năng đầu tiên trên thế giới.", "correct": "Đ", "reason": "Ra đời năm 1946 tại Mỹ."},
                    {"label": "b", "content": "Sự ra đời của bóng bán dẫn (Transistor) đã giúp máy tính thế hệ thứ hai nhỏ gọn hơn nhiều so với thế hệ dùng đèn điện tử chân không.", "correct": "Đ", "reason": "Bóng bán dẫn thay thế đèn chân không cồng kềnh."},
                    {"label": "c", "content": "Thế hệ máy tính thứ năm hướng tới việc tích hợp trí tuệ nhân tạo (AI), xử lý song song và tương tác ngôn ngữ tự nhiên.", "correct": "Đ", "reason": "Định hướng phát triển hiện nay."},
                    {"label": "d", "content": "Các máy tính thế hệ đầu tiên đã có thể bỏ túi và hoạt động bằng pin năng lượng mặt trời.", "correct": "S", "reason": "Máy tính thế hệ 1 nặng hàng chục tấn và chiếm diện tích cả căn phòng lớn."}
                ]
            },
            {
                "id": "8_TF_04",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "topic": "Khi áp dụng các hàm tính toán trong bảng tính để xử lý dữ liệu:",
                "statements": [
                    {"label": "a", "content": "Hàm =SUM(A1:A5) sẽ tính tổng tất cả các ô từ A1 đến A5.", "correct": "Đ", "reason": "Cú pháp chuẩn của hàm SUM."},
                    {"label": "b", "content": "Hàm =MAX(B1:B10) dùng để tìm giá trị nhỏ nhất trong danh sách.", "correct": "S", "reason": "MAX tìm giá trị lớn nhất, MIN mới tìm giá trị nhỏ nhất."},
                    {"label": "c", "content": "Khi ô tính hiển thị ký hiệu '#####', có nghĩa là độ rộng của cột chưa đủ để hiển thị số.", "correct": "Đ", "reason": "Cần kéo rộng cột ra để xem đầy đủ số."},
                    {"label": "d", "content": "Công thức bắt buộc phải bắt đầu bằng dấu bằng '='.", "correct": "Đ", "reason": "Quy tắc bắt buộc trong bảng tính."}
                ]
            }
        ],
        "essay_bank": [
            {
                "id": "8_ES_01",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Cho bảng tính điểm trung bình học kỳ của học sinh gồm các cột: Cột C (Điểm Toán), Cột D (Điểm Ngữ văn), Cột E (Điểm Tin học). Tại cột F (Điểm Trung Bình), em hãy viết công thức dùng hàm AVERAGE để tính điểm trung bình của học sinh ở dòng 4.",
                "rubric": [
                    {"criterion": "Viết đúng công thức chuẩn: =AVERAGE(C4:E4)", "score": 0.5},
                    {"criterion": "Hoặc công thức tương đương: =AVERAGE(C4, D4, E4)", "score": 0.25},
                    {"criterion": "Giải thích được ý nghĩa: Hàm tự động cộng 3 môn và chia đều cho 3", "score": 0.25}
                ]
            },
            {
                "id": "8_ES_02",
                "unit": "1. Xử lý dữ liệu với địa chỉ ô và hàm tính toán",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Phân biệt sự khác nhau giữa địa chỉ tương đối (ví dụ B2) và địa chỉ tuyệt đối (ví dụ $B$2) khi sao chép công thức sang vị trí khác trong bảng tính Excel.",
                "rubric": [
                    {"criterion": "Nêu rõ địa chỉ tương đối: Khi sao chép sang vị trí khác, tọa độ cột và dòng tự động thay đổi theo vị trí tương đối mới", "score": 0.5},
                    {"criterion": "Nêu rõ địa chỉ tuyệt đối: Luôn giữ nguyên tọa độ cột và dòng không đổi dù sao chép đến bất kỳ ô nào", "score": 0.5}
                ]
            },
            {
                "id": "8_ES_03",
                "unit": "1. Bản quyền phần mềm và sở hữu trí tuệ số",
                "level": "Thông hiểu",
                "score": 1.0,
                "content": "Em hãy nêu 2 lợi ích quan trọng khi sử dụng phần mềm có bản quyền và 2 tác hại nguy hiểm của việc cài đặt phần mềm bẻ khóa (crack) từ Internet.",
                "rubric": [
                    {"criterion": "2 lợi ích phần mềm bản quyền: Ổn định, an toàn, được cập nhật vá lỗi và hỗ trợ kỹ thuật từ nhà phát triển", "score": 0.5},
                    {"criterion": "2 tác hại phần mềm crack: Nguy cơ nhiễm virus/mã độc tống tiền, đánh cắp mật khẩu tài khoản và vi phạm pháp luật", "score": 0.5}
                ]
            }
        ]
    },

    "Lớp 9": {
        "topics": [
            {
                "tt": "1",
                "chapter": "Chủ đề 1. Máy tính và cộng đồng",
                "units": [
                    {
                        "unit_name": "1. Thế giới kỹ thuật số",
                        "req_nb": "Nêu được khả năng của máy tính và chỉ ra được một số ứng dụng thực tế của nó trong khoa học kĩ thuật và đời sống.",
                        "req_th": "Nhận biết được sự có mặt của các thiết bị có gắn bộ xử lí thông tin ở khắp nơi (trong gia đình, trường học, bệnh viện, nhà máy). Giải thích được tác động của công nghệ thông tin lên giáo dục và xã hội thông qua các ví dụ cụ thể.",
                        "req_vd": "Đánh giá được vai trò của thiết bị thông minh đối với việc nâng cao năng suất lao động.",
                        "req_vdc": "Đề xuất giải pháp ứng dụng thiết bị IoT để tối ưu hóa quản lý năng lượng trong lớp học thông minh."
                    }
                ]
            },
            {
                "tt": "2",
                "chapter": "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin",
                "units": [
                    {
                        "unit_name": "1. Thông tin trong giải quyết vấn đề",
                        "req_nb": "Nêu được các tiêu chí cơ bản để đánh giá chất lượng thông tin: tính mới, tính chính xác, tính đầy đủ, tính sử dụng được.",
                        "req_th": "Giải thích được sự cần thiết phải quan tâm đến chất lượng thông tin khi tìm kiếm, tiếp nhận và trao đổi thông tin trong thực tế.",
                        "req_vd": "Tìm kiếm, thu thập và chọn lọc được nguồn thông tin đáng tin cậy để giải quyết một nhiệm vụ học tập cụ thể.",
                        "req_vdc": "Xây dựng ma trận đánh giá phản biện cho các nguồn dữ liệu đa chiều trên Internet."
                    },
                    {
                        "unit_name": "2. Thực hành: Đánh giá chất lượng thông tin",
                        "req_nb": "Biết cách kiểm tra tác giả, cơ quan chủ quản, thời điểm xuất bản của trang thông tin điện tử.",
                        "req_th": "Nhận diện được dấu hiệu của các thông tin giật gân, thiếu căn cứ hoặc giả mạo.",
                        "req_vd": "Thực hiện được quy trình kiểm chứng chéo thông tin từ ít nhất 2 nguồn chính thống.",
                        "req_vdc": "Phát hiện và báo cáo các hành vi thao túng thông tin trên diễn đàn trực tuyến."
                    }
                ]
            },
            {
                "tt": "3",
                "chapter": "Chủ đề 3. Đạo đức, pháp luật và văn hóa trong môi trường số",
                "units": [
                    {
                        "unit_name": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                        "req_nb": "Nêu được một số nội dung liên quan đến Luật Công nghệ thông tin, Luật An ninh mạng, các khía cạnh pháp lý của việc sở hữu, sử dụng và trao đổi thông tin số.",
                        "req_th": "Trình bày được tác động tiêu cực của công nghệ số đối với con người. Nêu được các hành vi vi phạm pháp luật, trái đạo đức, thiếu văn hóa khi hoạt động trên môi trường số.",
                        "req_vd": "Thể hiện được hành vi công dân có trách nhiệm: tôn trọng bản quyền, bảo vệ dữ liệu cá nhân, ứng xử văn minh trên mạng.",
                        "req_vdc": "Xây dựng chiến dịch nâng cao nhận thức cộng đồng về an toàn số học đường."
                    }
                ]
            }
        ],
        "mcq_bank": [
            {
                "id": "9_MCQ_01",
                "unit": "1. Thế giới kỹ thuật số",
                "level": "Nhận biết",
                "content": "Máy tính là công cụ hỗ trợ đắc lực giúp con người thực hiện các hoạt động nào sau đây?",
                "options": {
                    "A": "Lưu trữ, xử lý và trao đổi thông tin tự động, chính xác",
                    "B": "Chỉ dùng để ghi chép văn bản thay thế sổ tay",
                    "C": "Chỉ dùng để chơi game và giải trí",
                    "D": "Thay thế hoàn toàn cơ thể sinh học của con người"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_02",
                "unit": "1. Thế giới kỹ thuật số",
                "level": "Nhận biết",
                "content": "Thiết bị nào sau đây thuộc nhóm phần cứng của máy tính điện tử?",
                "options": {
                    "A": "Hệ điều hành Windows",
                    "B": "Phần mềm soạn thảo văn bản Word",
                    "C": "Bộ nhớ trong RAM và bộ xử lý CPU",
                    "D": "Trình duyệt Google Chrome"
                },
                "correct": "C"
            },
            {
                "id": "9_MCQ_03",
                "unit": "1. Thông tin trong giải quyết vấn đề",
                "level": "Thông hiểu",
                "content": "Trong môi trường số, hoạt động 'kiểm chứng thông tin' có ý nghĩa bản chất là gì?",
                "options": {
                    "A": "Đọc tin tức thật nhanh và lướt qua tiêu đề",
                    "B": "Kiểm tra độ chính xác, nguồn gốc xuất xứ và độ tin cậy của thông tin trước khi sử dụng",
                    "C": "Chia sẻ ngay bài viết lên mạng xã hội để tăng tương tác",
                    "D": "Sao chép nguyên văn bài viết sang trang cá nhân"
                },
                "correct": "B"
            },
            {
                "id": "9_MCQ_04",
                "unit": "1. Thế giới kỹ thuật số",
                "level": "Nhận biết",
                "content": "Thiết bị thông minh (Smart device) là thiết bị có đặc điểm nổi bật nào dưới đây?",
                "options": {
                    "A": "Có gắn bộ vi xử lý, có khả năng kết nối mạng và thực hiện các tác vụ tự động, tương tác linh hoạt",
                    "B": "Là thiết bị gia dụng cơ học cổ điển không dùng điện",
                    "C": "Chỉ hoạt động được khi có người bấm nút điều khiển liên tục",
                    "D": "Tuyệt đối không thể kết nối mạng không dây"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_05",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "level": "Nhận biết",
                "content": "Trong môi trường số, loại thông tin nào dưới đây bắt buộc phải được bảo mật nghiêm ngặt?",
                "options": {
                    "A": "Tên trường học và lớp học của trường công lập",
                    "B": "Mật khẩu tài khoản ngân hàng, mã OTP và số định danh cá nhân (CCCD)",
                    "C": "Bảng thời khóa biểu học tập của cả lớp",
                    "D": "Nội dung bài thơ trong sách giáo khoa"
                },
                "correct": "B"
            },
            {
                "id": "9_MCQ_06",
                "unit": "1. Thế giới kỹ thuật số",
                "level": "Thông hiểu",
                "content": "Tác động tích cực nổi bật của công nghệ thông tin đối với giáo dục và xã hội là:",
                "options": {
                    "A": "Giúp người học tiếp cận kho tri thức toàn cầu mọi lúc mọi nơi, học tập cá nhân hóa và làm việc từ xa hiệu quả",
                    "B": "Làm cho học sinh không cần phải tư duy hay rèn luyện kỹ năng nào nữa",
                    "C": "Khiến con người hoàn toàn không cần tiếp xúc hay giao tiếp trực tiếp ngoài đời",
                    "D": "Khiến tri thức loài người không còn giá trị"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_07",
                "unit": "1. Thông tin trong giải quyết vấn đề",
                "level": "Thông hiểu",
                "content": "Một trong những tiêu chí quan trọng hàng đầu để đánh giá một nguồn thông tin trên Internet là đáng tin cậy:",
                "options": {
                    "A": "Có tiêu đề giật gân, ly kỳ và nhiều biểu tượng cảm xúc kích động",
                    "B": "Xuất phát từ cơ quan báo chí chính thống, tổ chức khoa học uy tín, có tác giả và ngày xuất bản rõ ràng",
                    "C": "Được lan truyền trên các tài khoản nặc danh trên mạng xã hội",
                    "D": "Có hình ảnh cắt ghép bắt mắt không rõ xuất xứ"
                },
                "correct": "B"
            },
            {
                "id": "9_MCQ_08",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "level": "Vận dụng",
                "content": "Khi chia sẻ một bài viết nghiên cứu hoặc tác phẩm của người khác lên Internet, hành động đúng chuẩn văn hóa số là:",
                "options": {
                    "A": "Ghi rõ họ tên tác giả, nguồn gốc bài viết và chỉ chia sẻ khi được phép hoặc tuân thủ giấy phép sử dụng",
                    "B": "Xóa tên tác giả cũ và thay bằng tên của mình",
                    "C": "Cắt xén nội dung làm thay đổi bản chất thông điệp của tác giả gốc",
                    "D": "Bán lại bài viết đó để kiếm tiền mà không xin phép"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_09",
                "unit": "1. Thông tin trong giải quyết vấn đề",
                "level": "Thông hiểu",
                "content": "Việc đánh giá chất lượng thông tin giúp ích gì cho con người khi ra quyết định?",
                "options": {
                    "A": "Giúp hạn chế rủi ro, đưa ra quyết định chính xác, giải quyết vấn đề hiệu quả và tránh bị lừa đảo",
                    "B": "Làm mất thời gian vô ích mà không mang lại lợi ích gì",
                    "C": "Khiến cho thông tin trở nên phức tạp và khó hiểu hơn",
                    "D": "Giúp người dùng không cần phải đọc thông tin nữa"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_10",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "level": "Nhận biết",
                "content": "Đạo đức số (Digital ethics) được thể hiện rõ nét qua hành vi nào dưới đây?",
                "options": {
                    "A": "Tôn trọng quyền riêng tư của người khác, ứng xử văn minh và không phát tán nội dung độc hại trên mạng",
                    "B": "Đọc lén tin nhắn riêng tư của bạn bè khi bạn quên khóa máy",
                    "C": "Chia sẻ thông tin bôi nhọ người khác khi có mâu thuẫn cá nhân",
                    "D": "Sử dụng hình ảnh của người khác để chế ảnh giễu cợt"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_11",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "level": "Thông hiểu",
                "content": "Văn hoá trong môi trường số KHÔNG bao gồm hành vi nào sau đây?",
                "options": {
                    "A": "Vi phạm bản quyền tác giả và phát tán phần mềm lậu",
                    "B": "Tuân thủ pháp luật về an ninh mạng và công nghệ thông tin",
                    "C": "Dùng ngôn ngữ lịch sự, tôn trọng sự khác biệt",
                    "D": "Bảo vệ thông tin cá nhân của bản thân và người khác"
                },
                "correct": "A"
            },
            {
                "id": "9_MCQ_12",
                "unit": "1. Thông tin trong giải quyết vấn đề",
                "level": "Vận dụng",
                "content": "Khi nhận được tin nhắn từ một tài khoản lạ thông báo 'Bạn đã trúng thưởng 500 triệu đồng, hãy nhấp vào liên kết sau và nộp 2 triệu phí nhận thưởng', phản ứng chuẩn mực nhất là:",
                "options": {
                    "A": "Nhận diện đây là dấu hiệu lừa đảo trực tuyến điển hình; tuyệt đối không nhấn vào liên kết, chặn tài khoản và cảnh báo người thân",
                    "B": "Lập tức làm theo hướng dẫn chuyển 2 triệu đồng để nhận thưởng nhanh chóng",
                    "C": "Nhập ngay thông tin số tài khoản và mật khẩu ngân hàng vào liên kết đó",
                    "D": "Chuyển tiếp tin nhắn cho tất cả bạn bè để cùng nhận thưởng"
                },
                "correct": "A"
            }
        ],
        "tf_bank": [
            {
                "id": "9_TF_01",
                "unit": "1. Thế giới kỹ thuật số",
                "topic": "Khi xem xét về thiết bị thông minh và công nghệ số trong cuộc sống hiện đại:",
                "statements": [
                    {"label": "a", "content": "Thiết bị thông minh có thể được điều khiển từ xa và giám sát trạng thái thông qua kết nối Internet.", "correct": "Đ", "reason": "Đặc trưng của thiết bị kết nối IoT."},
                    {"label": "b", "content": "Điện thoại thông minh (Smartphone) là một ví dụ điển hình về thiết bị thông minh đa năng.", "correct": "Đ", "reason": "Tích hợp CPU, cảm biến, hệ điều hành và kho ứng dụng."},
                    {"label": "c", "content": "Thiết bị thông minh chỉ hoạt động đơn độc, hoàn toàn không thể trao đổi dữ liệu với các thiết bị khác.", "correct": "S", "reason": "Thiết bị thông minh kết nối mạng để đồng bộ dữ liệu liên tục."},
                    {"label": "d", "content": "Sự xuất hiện của các dây chuyền robot tự động trong nhà máy giúp tăng năng suất và giảm thiểu nguy hiểm cho con người.", "correct": "Đ", "reason": "Tác động tích cực của tự động hóa công nghiệp."}
                ]
            },
            {
                "id": "9_TF_02",
                "unit": "1. Thông tin trong giải quyết vấn đề",
                "topic": "Về chất lượng và độ tin cậy của thông tin trên không gian mạng:",
                "statements": [
                    {"label": "a", "content": "Mọi thông tin có nhiều lượt thích (like) và chia sẻ (share) trên mạng xã hội đều là sự thật 100% không cần kiểm chứng.", "correct": "S", "reason": "Lượt tương tác có thể bị thao túng hoặc tin giả lan truyền nhanh do yếu tố giật gân."},
                    {"label": "b", "content": "Tính mới của thông tin là tiêu chí quan trọng khi xem xét dữ liệu dự báo thời tiết hoặc giá cả thị trường.", "correct": "Đ", "reason": "Thông tin lỗi thời sẽ làm quyết định sai lệch."},
                    {"label": "c", "content": "Khi tiếp nhận một thông tin quan trọng, nên đối chiếu chéo với các nguồn tin chính thống từ các cơ quan nhà nước uy tín.", "correct": "Đ", "reason": "Phương pháp kiểm chứng chéo khoa học."},
                    {"label": "d", "content": "Một thông tin chỉ có giá trị khi nó đầy đủ, chính xác, cập nhật và có thể sử dụng được cho mục đích cụ thể.", "correct": "Đ", "reason": "Các tiêu chí chất lượng thông tin theo SGK Tin 9."}
                ]
            },
            {
                "id": "9_TF_03",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "topic": "Khi sử dụng Internet và các dịch vụ trực tuyến để học tập, làm việc:",
                "statements": [
                    {"label": "a", "content": "Không nên chia sẻ mật khẩu tài khoản hoặc mã xác thực OTP cho bất kỳ người nào khác.", "correct": "Đ", "reason": "Nguyên tắc bảo vệ tài khoản sống còn."},
                    {"label": "b", "content": "Có thể mở tự do bất kỳ liên kết hoặc tệp đính kèm nào được gửi từ các email lạ không rõ danh tính.", "correct": "S", "reason": "Rất dễ bị dính mã độc đánh cắp tài khoản (Phishing/Malware)."},
                    {"label": "c", "content": "Việc đặt mật khẩu phức tạp gồm cả chữ hoa, chữ thường, số, ký tự đặc biệt giúp nâng cao độ an toàn dữ liệu.", "correct": "Đ", "reason": "Mật khẩu mạnh chống bị tấn công dò quét."},
                    {"label": "d", "content": "Sử dụng phần mềm phòng chống mã độc và thường xuyên cập nhật hệ điều hành giúp bảo vệ máy tính an toàn.", "correct": "Đ", "reason": "Vá các lỗ hổng bảo mật hệ thống."}
                ]
            },
            {
                "id": "9_TF_04",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "topic": "Về vai trò và trách nhiệm của công dân số (Digital Citizen) trong thời đại mới:",
                "statements": [
                    {"label": "a", "content": "Công dân số có trách nhiệm bảo vệ thông tin cá nhân của mình và tôn trọng quyền riêng tư của người khác.", "correct": "Đ", "reason": "Tiêu chuẩn cơ bản của công dân số văn minh."},
                    {"label": "b", "content": "Tự do ngôn luận trên mạng đồng nghĩa với việc được phép bịa đặt, xúc phạm danh dự của bất kỳ ai mà không bị xử lý.", "correct": "S", "reason": "Hành vi vu khống, xúc phạm trên mạng bị xử phạt theo Luật An ninh mạng và Bộ luật Hình sự."},
                    {"label": "c", "content": "Chủ động phát hiện, báo cáo các nội dung độc hại, lừa đảo trực tuyến cho cơ quan chức năng là hành vi có trách nhiệm.", "correct": "Đ", "reason": "Góp phần làm trong sạch môi trường số."},
                    {"label": "d", "content": "Ứng dụng học tập trực tuyến giúp học sinh tự chủ kế hoạch học tập, tiết kiệm thời gian và tài nguyên di chuyển.", "correct": "Đ", "reason": "Lợi ích nổi bật của chuyển đổi số giáo dục."}
                ]
            }
        ],
        "essay_bank": [
            {
                "id": "9_ES_01",
                "unit": "1. Thế giới kỹ thuật số",
                "level": "Thông hiểu",
                "score": 1.0,
                "content": "Em hãy cho biết vì sao thông tin lại đóng vai trò đặc biệt quan trọng đối với con người trong hoạt động học tập, lao động sản xuất và đời sống xã hội?",
                "rubric": [
                    {"criterion": "Giúp con người hiểu biết thế giới tự nhiên và xã hội, mở mang tri thức", "score": 0.25},
                    {"criterion": "Là căn cứ sống còn để con người đưa ra các quyết định hành động đúng đắn và kịp thời", "score": 0.25},
                    {"criterion": "Giúp nâng cao hiệu quả học tập, lao động, tối ưu hóa quy trình làm việc", "score": 0.25},
                    {"criterion": "Là phương tiện để giao tiếp, hợp tác, kết nối cộng đồng và phát triển văn minh", "score": 0.25}
                ]
            },
            {
                "id": "9_ES_02",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "level": "Thông hiểu",
                "score": 1.0,
                "content": "Nêu hai ví dụ cụ thể thể hiện hành vi có đạo đức và văn hóa khi tham gia hoạt động trong môi trường số.",
                "rubric": [
                    {"criterion": "Ví dụ 1: Không đăng tải, chia sẻ thông tin bịa đặt, sai sự thật; kiểm chứng trước khi chia sẻ", "score": 0.5},
                    {"criterion": "Ví dụ 2: Tôn trọng quyền riêng tư, không phát tán hình ảnh hoặc dữ liệu cá nhân của người khác khi chưa được phép", "score": 0.5}
                ]
            },
            {
                "id": "9_ES_03",
                "unit": "1. Một số vấn đề pháp lý về sử dụng dịch vụ Internet",
                "level": "Vận dụng",
                "score": 1.0,
                "content": "Một học sinh phổ thông thể hiện vai trò là 'Công dân số có trách nhiệm' thông qua những hành vi cụ thể nào? (Em hãy nêu rõ ít nhất 4 hành vi thiết thực).",
                "rubric": [
                    {"criterion": "Không chia sẻ thông tin giả mạo, chưa kiểm chứng; kiểm tra nguồn gốc tin tức", "score": 0.25},
                    {"criterion": "Bảo vệ thông tin cá nhân và tài khoản trực tuyến của bản thân và người khác", "score": 0.25},
                    {"criterion": "Giao tiếp, ứng xử lịch sự, văn minh và tôn trọng sự khác biệt trên không gian mạng", "score": 0.25},
                    {"criterion": "Tuân thủ luật pháp về bản quyền phần mềm, sở hữu trí tuệ và an ninh mạng", "score": 0.25}
                ]
            }
        ]
    }
}


# ==============================================================================
# 2. BỘ ĐỘNG CƠ TÍNH TOÁN MA TRẬN VÀ BẢN ĐẶC TẢ SƯ PHẠM
# ==============================================================================

class ExamMatrixEngine:
    """
    Tính toán phân bổ câu hỏi theo tỷ lệ nhận thức (% NB, % TH, % VD, % VDC)
    và chuẩn hóa cấu trúc đề thi theo ma trận đặc tả chuẩn Bộ GD&ĐT.
    """

    def __init__(
        self,
        grade_level: str = "Lớp 9",
        exam_title: str = "KIỂM TRA GIỮA HỌC KỲ I",
        school_name: str = "TRƯỜNG TH&THCS ĐĂK TRÔI",
        governing_body: str = "UBND XÃ KON CHIÊNG",
        school_year: str = "2026 - 2027",
        teacher_name: str = "Trần Đức Tình",
        head_name: str = "Hoàng Thị Kim Tuyết",
        pct_nb: float = 30.0,
        pct_th: float = 40.0,
        pct_vd: float = 30.0,
        pct_vdc: float = 0.0,
        num_mcq: int = 12,     # 3.0 điểm
        num_tf: int = 4,       # 4.0 điểm (16 lệnh hỏi x 0.25đ)
        num_essay: int = 3,    # 3.0 điểm (3 câu x 1.0đ)
        duration_minutes: int = 45,
        topic_ratios: Optional[Dict[str, float]] = None,
        allowed_units: Optional[List[str]] = None,
        custom_requirements: Optional[List[Dict[str, Any]]] = None
    ):
        self.grade_level = grade_level if grade_level in CURRICULUM_DATA else "Lớp 9"
        self.exam_title = exam_title
        self.school_name = school_name
        self.governing_body = governing_body
        self.school_year = school_year
        self.teacher_name = teacher_name
        self.head_name = head_name
        self.duration_minutes = duration_minutes
        self.topic_ratios = topic_ratios or {}
        self.allowed_units = allowed_units
        self.custom_requirements = custom_requirements

        # Chuẩn hóa % nhận thức về 100%
        total_pct = pct_nb + pct_th + pct_vd + pct_vdc
        if total_pct <= 0:
            total_pct = 100.0
            pct_nb, pct_th, pct_vd, pct_vdc = 30.0, 40.0, 30.0, 0.0
        self.pct_nb = round((pct_nb / total_pct) * 100, 1)
        self.pct_th = round((pct_th / total_pct) * 100, 1)
        self.pct_vd = round((pct_vd / total_pct) * 100, 1)
        self.pct_vdc = round(100.0 - (self.pct_nb + self.pct_th + self.pct_vd), 1)

        self.num_mcq = num_mcq
        self.num_tf = num_tf
        self.num_essay = num_essay

        self.data_store = CURRICULUM_DATA.get(self.grade_level, CURRICULUM_DATA["Lớp 9"])

    def _is_unit_allowed(self, unit_str: str) -> bool:
        """Kiểm tra bài học có nằm trong phạm vi đã dạy theo KHGD không."""
        if not self.allowed_units:
            return True
        import re
        u_clean = unit_str.strip().lower()
        for al in self.allowed_units:
            al_clean = al.strip().lower()
            if al_clean in u_clean or u_clean in al_clean:
                return True
            num_al = re.findall(r'\d+', al_clean)
            num_u = re.findall(r'\d+', u_clean)
            if num_al and num_u and num_al[0] == num_u[0]:
                return True
        return False

    def _parse_yccd_levels(self, req_str: str) -> Dict[str, str]:
        """
        Tách văn bản YCCĐ thành 3 mức: Nhận biết, Thông hiểu, Vận dụng.
        Hỗ trợ cả định dạng mục (Nhận biết: ...) và định dạng gạch đầu dòng chứa động từ tư duy
        (˗ Nhận biết được..., - Nêu được..., - Giải thích được..., - Vận dụng được...).
        """
        if not req_str:
            return {
                "nb": "Nêu được các khái niệm và nhận biết kiến thức cốt lõi của bài học.",
                "th": "Giải thích, phân biệt và làm rõ được các nội dung trọng tâm của bài học.",
                "vd": "Vận dụng kiến thức bài học để giải quyết vấn đề trong thực tế."
            }

        # Nếu có tiền tố rõ ràng như 'Nhận biết:'
        req_clean = req_str.replace("\r\n", "\n").replace("\r", "\n")
        lines = req_clean.splitlines()
        has_named_levels = any(
            any(k in l.lower() for k in ["nhận biết:", "thông hiểu:", "vận dụng:", "1. nhận biết", "2. thông hiểu", "3. vận dụng"])
            for l in lines
        )

        if has_named_levels:
            res = {"nb": "", "th": "", "vd": ""}
            current_key = "nb"
            cur_lines = []
            for line in lines:
                l_low = line.strip().lower()
                if l_low.startswith("nhận biết") or l_low.startswith("- nhận biết") or l_low.startswith("1. nhận biết"):
                    if cur_lines:
                        res[current_key] = "\n".join(cur_lines).strip()
                        cur_lines = []
                    current_key = "nb"
                    if ":" in line:
                        part = line.split(":", 1)[1].strip()
                        if part: cur_lines.append(part)
                elif l_low.startswith("thông hiểu") or l_low.startswith("- thông hiểu") or l_low.startswith("2. thông hiểu"):
                    if cur_lines:
                        res[current_key] = "\n".join(cur_lines).strip()
                        cur_lines = []
                    current_key = "th"
                    if ":" in line:
                        part = line.split(":", 1)[1].strip()
                        if part: cur_lines.append(part)
                elif l_low.startswith("vận dụng cao") or l_low.startswith("- vận dụng cao") or l_low.startswith("4. vận dụng cao") or l_low.startswith("vận dụng") or l_low.startswith("- vận dụng") or l_low.startswith("3. vận dụng"):
                    if cur_lines:
                        res[current_key] = "\n".join(cur_lines).strip()
                        cur_lines = []
                    current_key = "vd"
                    if ":" in line:
                        part = line.split(":", 1)[1].strip()
                        if part: cur_lines.append(part)
                else:
                    if line.strip():
                        cur_lines.append(line.strip())
            if cur_lines:
                res[current_key] = "\n".join(cur_lines).strip()
            if not res["nb"]: res["nb"] = "Nêu được các khái niệm và nhận biết kiến thức cốt lõi của bài học."
            if not res["th"]: res["th"] = "Giải thích, phân biệt và làm rõ được các nội dung trọng tâm của bài học."
            if not res["vd"]: res["vd"] = "Vận dụng kiến thức bài học để giải quyết vấn đề trong thực tế."
            return res

        # Bóc tách gạch đầu dòng theo động từ tư duy chuẩn GDPT 2018
        normalized = re.sub(r'[\n\r]*[\˗\-\–\—\•\*\+]\s*', '\n- ', req_clean).strip()
        items = [l.strip().lstrip('- ').strip() for l in normalized.split('\n') if l.strip().lstrip('- ').strip()]

        levels = {"nb": [], "th": [], "vd": []}
        nb_prefixes = ("nhận biết", "nêu", "chỉ ra", "liệt kê", "kể tên", "nhận ra", "biết", "nhận diện", "xác định")
        th_prefixes = ("hiểu", "giải thích", "trình bày", "phân biệt", "so sánh", "mô tả", "làm rõ", "phân tích", "làm quen")
        vd_prefixes = ("vận dụng", "sử dụng", "thực hiện", "tạo", "tìm kiếm", "đánh giá", "khai thác", "áp dụng", "thực hành", "giải quyết", "hoàn thiện")

        for it in items:
            it_low = it.lower()
            first_words = " ".join(it_low.split()[:4])
            if any(first_words.startswith(p) for p in nb_prefixes):
                levels["nb"].append(it)
            elif any(first_words.startswith(p) for p in th_prefixes):
                levels["th"].append(it)
            elif any(first_words.startswith(p) for p in vd_prefixes):
                levels["vd"].append(it)
            else:
                if any(p in it_low for p in ["vận dụng", "thực hành", "giải quyết", "tạo được", "sử dụng được"]):
                    levels["vd"].append(it)
                elif any(p in it_low for p in ["giải thích", "trình bày", "phân biệt", "hiểu được"]):
                    levels["th"].append(it)
                else:
                    levels["nb"].append(it)

        res = {
            "nb": "\n".join([f"- {i}" for i in levels["nb"]]).strip(),
            "th": "\n".join([f"- {i}" for i in levels["th"]]).strip(),
            "vd": "\n".join([f"- {i}" for i in levels["vd"]]).strip()
        }
        if not res["nb"]: res["nb"] = "- Nêu được các khái niệm và nhận biết kiến thức cốt lõi của bài học."
        if not res["th"]: res["th"] = "- Giải thích, phân biệt và làm rõ được các nội dung trọng tâm của bài học."
        if not res["vd"]: res["vd"] = "- Vận dụng kiến thức bài học để giải quyết vấn đề trong thực tế."
        return res

    def _prepare_active_units(self) -> List[Dict[str, Any]]:
        """Lấy danh sách các bài học hợp lệ từ KHGD (tải lên/phân tích AI) hoặc SGK chuẩn."""
        active_units = []
        if self.custom_requirements:
            for idx, cr in enumerate(self.custom_requirements, 1):
                ch_name = cr.get("chapter", "").strip() or f"Chủ đề {idx}"
                u_name = cr.get("unit", "").strip() or f"Bài {idx}"
                raw_req = cr.get("requirement", "").strip()
                
                # Ưu tiên các mức độ đã bóc tách sẵn từ file KHGD
                if cr.get("req_nb") and cr.get("req_th") and cr.get("req_vd"):
                    r_nb = cr.get("req_nb")
                    r_th = cr.get("req_th")
                    r_vd = cr.get("req_vd")
                else:
                    lvl_dict = self._parse_yccd_levels(raw_req)
                    r_nb = lvl_dict["nb"]
                    r_th = lvl_dict["th"]
                    r_vd = lvl_dict["vd"]

                active_units.append({
                    "tt": str(idx),
                    "chapter": ch_name,
                    "unit": u_name,
                    "req_nb": r_nb,
                    "req_th": r_th,
                    "req_vd": r_vd
                })

        if not active_units:
            topics = self.data_store.get("topics", [])
            tt_counter = 1
            for top in topics:
                ch_name = top["chapter"]
                for u in top["units"]:
                    u_name = u["unit_name"]
                    if self.allowed_units and not self._is_unit_allowed(u_name):
                        continue
                    active_units.append({
                        "tt": str(tt_counter),
                        "chapter": ch_name,
                        "unit": u_name,
                        "req_nb": u.get("req_nb", ""),
                        "req_th": u.get("req_th", ""),
                        "req_vd": u.get("req_vd", "")
                    })
                    tt_counter += 1

        if not active_units:
            topics = self.data_store.get("topics", [])
            tt_counter = 1
            for top in topics:
                for u in top["units"]:
                    active_units.append({
                        "tt": str(tt_counter),
                        "chapter": top["chapter"],
                        "unit": u["unit_name"],
                        "req_nb": u.get("req_nb", ""),
                        "req_th": u.get("req_th", ""),
                        "req_vd": u.get("req_vd", "")
                    })
                    tt_counter += 1
                    if len(active_units) >= 4:
                        break
                if len(active_units) >= 4:
                    break
        return active_units

    @staticmethod
    def _distribute_items(count: int, n: int, start_idx: int = 0) -> List[int]:
        """Phân bổ đều số lượng câu hỏi vào n bài học."""
        res = [0] * n
        if n <= 0:
            return res
        for k in range(count):
            idx = (start_idx + k) % n
            res[idx] += 1
        return res

    def _plan_matrix(self, active_units: List[Dict[str, Any]]) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        """
        Tính toán ma trận phân bổ số câu, điểm số và bản đặc tả chuẩn xác 100% toán học.
        Mọi phép cộng ngang và dọc đều khớp tuyệt đối, không hardcode.
        """
        n = len(active_units)

        def fmt_pts(val: float) -> str:
            if abs(val - round(val)) < 1e-4:
                return f"{int(round(val))},0"
            return f"{val:.2f}".rstrip("0").replace(".", ",")

        def fmt_pct(val: float) -> str:
            if abs(val - round(val)) < 1e-4:
                return f"{int(round(val))}%"
            return f"{val:.1f}".replace(".", ",") + "%"

        # Phân bổ số câu chuẩn GDPT 2018 Bộ GD&ĐT:
        # Phần I (MCQ 12 câu = 3.0đ): 8 Biết (2.0đ), 4 Hiểu (1.0đ)
        # Phần II (Đúng - Sai 4 câu = 4.0đ): 1 Biết (1.0đ), 1 Hiểu (1.0đ), 2 Vận dụng (2.0đ)
        # Phần III (Tự luận 3 câu = 3.0đ): 1 Biết/Hiểu (1.0đ), 1 Hiểu (1.0đ), 1 Vận dụng (1.0đ)
        target_mcq_nb = 8
        target_mcq_th = 4
        target_mcq_vd = 0

        target_tf_nb = 1
        target_tf_th = 1
        target_tf_vd = 2

        if self.pct_nb >= 40.0:
            target_es_nb = 1
            target_es_th = 1
            target_es_vd = 1
        else:
            target_es_nb = 0
            target_es_th = 1
            target_es_vd = 2

        mcq_nb = self._distribute_items(target_mcq_nb, n, 0)
        mcq_th = self._distribute_items(target_mcq_th, n, 1 % n)
        mcq_vd = [0] * n

        tf_nb = self._distribute_items(target_tf_nb, n, 0)
        tf_th = self._distribute_items(target_tf_th, n, 1 % n)
        tf_vd = self._distribute_items(target_tf_vd, n, min(2, n - 1))

        es_nb = self._distribute_items(target_es_nb, n, min(1, n - 1)) if target_es_nb > 0 else [0] * n
        es_th = self._distribute_items(target_es_th, n, min(2, n - 1))
        es_vd = self._distribute_items(target_es_vd, n, n - 1)

        mat_rows = []
        spec_rows = []

        for i, u in enumerate(active_units):
            m_nb = mcq_nb[i]; m_th = mcq_th[i]; m_vd = mcq_vd[i]
            t_nb = tf_nb[i]; t_th = tf_th[i]; t_vd = tf_vd[i]
            e_nb = es_nb[i]; e_th = es_th[i]; e_vd = es_vd[i]

            row_nb_cnt = m_nb + t_nb + e_nb
            row_nb_pts = m_nb * 0.25 + t_nb * 1.0 + e_nb * 1.0

            row_th_cnt = m_th + t_th + e_th
            row_th_pts = m_th * 0.25 + t_th * 1.0 + e_th * 1.0

            row_vd_cnt = m_vd + t_vd + e_vd
            row_vd_pts = m_vd * 0.25 + t_vd * 1.0 + e_vd * 1.0

            row_score = row_nb_pts + row_th_pts + row_vd_pts
            row_pct = (row_score / 10.0) * 100.0

            mat_item = {
                "tt": u["tt"],
                "chapter": u["chapter"],
                "unit": u["unit"],
                "raw_mcq_nb": m_nb, "raw_mcq_th": m_th, "raw_mcq_vd": m_vd,
                "raw_tf_nb": t_nb, "raw_tf_th": t_th, "raw_tf_vd": t_vd,
                "raw_es_nb": e_nb, "raw_es_th": e_th, "raw_es_vd": e_vd,
                "row_nb_cnt": row_nb_cnt, "row_nb_score": row_nb_pts,
                "row_th_cnt": row_th_cnt, "row_th_score": row_th_pts,
                "row_vd_cnt": row_vd_cnt, "row_vd_score": row_vd_pts,
                "row_score": row_score, "row_pct": row_pct,
                # Chuỗi hiển thị ma trận
                "mcq_nb": f"{m_nb} ({fmt_pts(m_nb * 0.25)})" if m_nb > 0 else "",
                "mcq_th": f"{m_th} ({fmt_pts(m_th * 0.25)})" if m_th > 0 else "",
                "mcq_vd": f"{m_vd} ({fmt_pts(m_vd * 0.25)})" if m_vd > 0 else "",
                "tf_nb": f"{t_nb} ({fmt_pts(t_nb * 1.0)})" if t_nb > 0 else "",
                "tf_th": f"{t_th} ({fmt_pts(t_th * 1.0)})" if t_th > 0 else "",
                "tf_vd": f"{t_vd} ({fmt_pts(t_vd * 1.0)})" if t_vd > 0 else "",
                "es_nb": f"{e_nb} ({fmt_pts(e_nb * 1.0)})" if e_nb > 0 else "",
                "es_th": f"{e_th} ({fmt_pts(e_th * 1.0)})" if e_th > 0 else "",
                "es_vd": f"{e_vd} ({fmt_pts(e_vd * 1.0)})" if e_vd > 0 else "",
                "tot_nb": f"{row_nb_cnt} ({fmt_pts(row_nb_pts)})" if row_nb_cnt > 0 else "",
                "tot_th": f"{row_th_cnt} ({fmt_pts(row_th_pts)})" if row_th_cnt > 0 else "",
                "tot_vd": f"{row_vd_cnt} ({fmt_pts(row_vd_pts)})" if row_vd_cnt > 0 else "",
                "total_score": f"{fmt_pct(row_pct)} ({fmt_pts(row_score)}đ)"
            }
            mat_rows.append(mat_item)

            # Xây dựng các dòng con cho Bản đặc tả (mỗi mức nhận thức 1 dòng, số câu nằm NGANG HÀNG với YCCĐ)
            level_rows = []
            if row_nb_cnt > 0 or u.get("req_nb"):
                level_rows.append({
                    "level_name": "Nhận biết",
                    "requirement": f"Nhận biết:\n{u.get('req_nb', '')}",
                    "mcq_nb": str(m_nb) if m_nb > 0 else "", "mcq_th": "", "mcq_vd": "",
                    "tf_nb": str(t_nb) if t_nb > 0 else "", "tf_th": "", "tf_vd": "",
                    "es_nb": str(e_nb) if e_nb > 0 else "", "es_th": "", "es_vd": ""
                })
            if row_th_cnt > 0 or u.get("req_th"):
                level_rows.append({
                    "level_name": "Thông hiểu",
                    "requirement": f"Thông hiểu:\n{u.get('req_th', '')}",
                    "mcq_nb": "", "mcq_th": str(m_th) if m_th > 0 else "", "mcq_vd": "",
                    "tf_nb": "", "tf_th": str(t_th) if t_th > 0 else "", "tf_vd": "",
                    "es_nb": "", "es_th": str(e_th) if e_th > 0 else "", "es_vd": ""
                })
            if row_vd_cnt > 0 or u.get("req_vd"):
                level_rows.append({
                    "level_name": "Vận dụng",
                    "requirement": f"Vận dụng:\n{u.get('req_vd', '')}",
                    "mcq_nb": "", "mcq_th": "", "mcq_vd": str(m_vd) if m_vd > 0 else "",
                    "tf_nb": "", "tf_th": "", "tf_vd": str(t_vd) if t_vd > 0 else "",
                    "es_nb": "", "es_th": "", "es_vd": str(e_vd) if e_vd > 0 else ""
                })

            spec_rows.append({
                "tt": u["tt"],
                "chapter": u["chapter"],
                "unit": u["unit"],
                "requirement": f"Nhận biết:\n{u.get('req_nb', '')}\n\nThông hiểu:\n{u.get('req_th', '')}\n\nVận dụng:\n{u.get('req_vd', '')}",
                "mcq_nb": str(m_nb) if m_nb > 0 else "",
                "mcq_th": str(m_th) if m_th > 0 else "",
                "mcq_vd": str(m_vd) if m_vd > 0 else "",
                "tf_nb": str(t_nb) if t_nb > 0 else "",
                "tf_th": str(t_th) if t_th > 0 else "",
                "tf_vd": str(t_vd) if t_vd > 0 else "",
                "es_nb": str(e_nb) if e_nb > 0 else "",
                "es_th": str(e_th) if e_th > 0 else "",
                "es_vd": str(e_vd) if e_vd > 0 else "",
                "level_rows": level_rows
            })

        sum_mcq_nb = sum(r["raw_mcq_nb"] for r in mat_rows)
        sum_mcq_th = sum(r["raw_mcq_th"] for r in mat_rows)
        sum_mcq_vd = sum(r["raw_mcq_vd"] for r in mat_rows)

        sum_tf_nb = sum(r["raw_tf_nb"] for r in mat_rows)
        sum_tf_th = sum(r["raw_tf_th"] for r in mat_rows)
        sum_tf_vd = sum(r["raw_tf_vd"] for r in mat_rows)

        sum_es_nb = sum(r["raw_es_nb"] for r in mat_rows)
        sum_es_th = sum(r["raw_es_th"] for r in mat_rows)
        sum_es_vd = sum(r["raw_es_vd"] for r in mat_rows)

        all_mcq_cnt = sum_mcq_nb + sum_mcq_th + sum_mcq_vd
        all_tf_cnt = sum_tf_nb + sum_tf_th + sum_tf_vd
        all_es_cnt = sum_es_nb + sum_es_th + sum_es_vd

        tot_nb_cnt = sum_mcq_nb + sum_tf_nb + sum_es_nb
        tot_th_cnt = sum_mcq_th + sum_tf_th + sum_es_th
        tot_vd_cnt = sum_mcq_vd + sum_tf_vd + sum_es_vd
        all_q_cnt = tot_nb_cnt + tot_th_cnt + tot_vd_cnt

        pts_mcq_nb = sum_mcq_nb * 0.25; pts_mcq_th = sum_mcq_th * 0.25; pts_mcq_vd = sum_mcq_vd * 0.25
        pts_tf_nb = sum_tf_nb * 1.0; pts_tf_th = sum_tf_th * 1.0; pts_tf_vd = sum_tf_vd * 1.0
        pts_es_nb = sum_es_nb * 1.0; pts_es_th = sum_es_th * 1.0; pts_es_vd = sum_es_vd * 1.0

        tot_mcq_pts = pts_mcq_nb + pts_mcq_th + pts_mcq_vd
        tot_tf_pts = pts_tf_nb + pts_tf_th + pts_tf_vd
        tot_es_pts = pts_es_nb + pts_es_th + pts_es_vd

        tot_nb_pts = pts_mcq_nb + pts_tf_nb + pts_es_nb
        tot_th_pts = pts_mcq_th + pts_tf_th + pts_es_th
        tot_vd_pts = pts_mcq_vd + pts_tf_vd + pts_es_vd
        all_pts = tot_nb_pts + tot_th_pts + tot_vd_pts

        pct_mcq = (tot_mcq_pts / 10.0) * 100.0
        pct_tf = (tot_tf_pts / 10.0) * 100.0
        pct_es = (tot_es_pts / 10.0) * 100.0

        pct_nb = (tot_nb_pts / 10.0) * 100.0
        pct_th = (tot_th_pts / 10.0) * 100.0
        pct_vd = (tot_vd_pts / 10.0) * 100.0

        totals = {
            "total_mcq": f"{all_mcq_cnt} câu ({fmt_pts(tot_mcq_pts)}đ)",
            "total_tf": f"{all_tf_cnt} câu ({fmt_pts(tot_tf_pts)}đ)",
            "total_es": f"{all_es_cnt} câu ({fmt_pts(tot_es_pts)}đ)",
            "sum_score": f"{fmt_pts(all_pts)} điểm",
            "ratio_str": f"{fmt_pct(pct_nb)} Biết - {fmt_pct(pct_th)} Hiểu - {fmt_pct(pct_vd)} Vận dụng",
            "sum_mcq_nb": sum_mcq_nb, "sum_mcq_th": sum_mcq_th, "sum_mcq_vd": sum_mcq_vd,
            "sum_tf_nb": sum_tf_nb, "sum_tf_th": sum_tf_th, "sum_tf_vd": sum_tf_vd,
            "sum_es_nb": sum_es_nb, "sum_es_th": sum_es_th, "sum_es_vd": sum_es_vd,
            "tot_nb_cnt": tot_nb_cnt, "tot_th_cnt": tot_th_cnt, "tot_vd_cnt": tot_vd_cnt,
            "all_q_cnt": all_q_cnt,
            "pts_mcq_nb": pts_mcq_nb, "pts_mcq_th": pts_mcq_th, "pts_mcq_vd": pts_mcq_vd,
            "pts_tf_nb": pts_tf_nb, "pts_tf_th": pts_tf_th, "pts_tf_vd": pts_tf_vd,
            "pts_es_nb": pts_es_nb, "pts_es_th": pts_es_th, "pts_es_vd": pts_es_vd,
            "tot_mcq_pts": tot_mcq_pts, "tot_tf_pts": tot_tf_pts, "tot_es_pts": tot_es_pts,
            "tot_nb_pts": tot_nb_pts, "tot_th_pts": tot_th_pts, "tot_vd_pts": tot_vd_pts,
            "all_pts": all_pts,
            "pct_mcq": pct_mcq, "pct_tf": pct_tf, "pct_es": pct_es,
            "pct_nb": pct_nb, "pct_th": pct_th, "pct_vd": pct_vd
        }

        matrix_rows = {"units": mat_rows, "totals": totals}
        return matrix_rows, spec_rows

    def generate_exam_dossier(self) -> Dict[str, Any]:
        """Tạo trọn bộ hồ sơ đề thi: Ma trận, Đặc tả, Đề TH01, Đề TH02, Đáp án và Biểu điểm chuẩn xác 100%."""
        # 1. Chuẩn bị danh sách bài học active từ KHGD / custom_requirements / SGK
        active_units = self._prepare_active_units()

        # 2. Xây dựng Ma trận và Bản đặc tả trước, tính toán số câu chuẩn xác tuyệt đối
        matrix_rows, spec_rows = self._plan_matrix(active_units)

        # 3. Lấy câu hỏi từ ngân hàng BÁM SÁT MA TRẬN VÀ CÁC BÀI ĐÃ HỌC THEO KHGD
        all_mcqs = list(self.data_store.get("mcq_bank", []))
        all_tfs = list(self.data_store.get("tf_bank", []))
        all_essays = list(self.data_store.get("essay_bank", []))

        chosen_mcqs = []
        chosen_tfs = []
        chosen_essays = []

        for u in matrix_rows["units"]:
            u_name = u["unit"]
            u_mcqs = [q for q in all_mcqs if self._is_unit_allowed(q.get("unit", "")) and (u_name.lower() in q.get("unit", "").lower() or q.get("unit", "").lower() in u_name.lower())]
            if not u_mcqs:
                u_mcqs = all_mcqs
            u_tfs = [q for q in all_tfs if self._is_unit_allowed(q.get("unit", "")) and (u_name.lower() in q.get("unit", "").lower() or q.get("unit", "").lower() in u_name.lower())]
            if not u_tfs:
                u_tfs = all_tfs
            u_essays = [q for q in all_essays if self._is_unit_allowed(q.get("unit", "")) and (u_name.lower() in q.get("unit", "").lower() or q.get("unit", "").lower() in u_name.lower())]
            if not u_essays:
                u_essays = all_essays

            for q in u_mcqs:
                if len([c for c in chosen_mcqs if c.get("id") == q.get("id")]) == 0:
                    chosen_mcqs.append(copy.deepcopy(q))
                    if len(chosen_mcqs) >= self.num_mcq:
                        break
            for q in u_tfs:
                if len([c for c in chosen_tfs if c.get("id") == q.get("id")]) == 0:
                    chosen_tfs.append(copy.deepcopy(q))
                    if len(chosen_tfs) >= self.num_tf:
                        break
            for q in u_essays:
                if len([c for c in chosen_essays if c.get("id") == q.get("id")]) == 0:
                    chosen_essays.append(copy.deepcopy(q))
                    if len(chosen_essays) >= self.num_essay:
                        break

        # Bù từ ngân hàng chung nếu bài đó thiếu câu để luôn đủ đúng số lượng
        while len(chosen_mcqs) < self.num_mcq:
            for q in all_mcqs:
                chosen_mcqs.append(copy.deepcopy(q))
                if len(chosen_mcqs) >= self.num_mcq:
                    break
        chosen_mcqs = chosen_mcqs[:self.num_mcq]

        while len(chosen_tfs) < self.num_tf:
            for q in all_tfs:
                chosen_tfs.append(copy.deepcopy(q))
                if len(chosen_tfs) >= self.num_tf:
                    break
        chosen_tfs = chosen_tfs[:self.num_tf]

        while len(chosen_essays) < self.num_essay:
            for q in all_essays:
                chosen_essays.append(copy.deepcopy(q))
                if len(chosen_essays) >= self.num_essay:
                    break
        chosen_essays = chosen_essays[:self.num_essay]

        # 4. Sinh biến thể Đề hoán vị (MÃ ĐỀ TH02)
        th02_mcqs = []
        for q in chosen_mcqs:
            q_copy = copy.deepcopy(q)
            opt_keys = ["A", "B", "C", "D"]
            opt_vals = [q["options"][k] for k in opt_keys]
            correct_val = q["options"][q["correct"]]

            shuffled_vals = list(opt_vals)
            random.seed(len(q["content"]))
            random.shuffle(shuffled_vals)
            new_opts = {k: v for k, v in zip(opt_keys, shuffled_vals)}
            new_correct = "A"
            for k, v in new_opts.items():
                if v == correct_val:
                    new_correct = k
                    break
            q_copy["options"] = new_opts
            q_copy["correct"] = new_correct
            th02_mcqs.append(q_copy)

        th02_mcqs_shuffled = list(th02_mcqs)
        random.seed(42)
        random.shuffle(th02_mcqs_shuffled)

        th02_tfs = []
        for tf_item in chosen_tfs:
            tf_copy = copy.deepcopy(tf_item)
            stmts = list(tf_copy["statements"])
            random.seed(len(tf_item["topic"]))
            random.shuffle(stmts)
            for idx, lbl in enumerate(["a", "b", "c", "d"]):
                stmts[idx]["label"] = lbl
            tf_copy["statements"] = stmts
            th02_tfs.append(tf_copy)

        th02_essays = list(chosen_essays)

        return {
            "meta": {
                "grade": self.grade_level,
                "exam_title": self.exam_title,
                "school_name": self.school_name,
                "governing_body": self.governing_body,
                "school_year": self.school_year,
                "teacher_name": self.teacher_name,
                "head_name": self.head_name,
                "duration": self.duration_minutes,
                "pct_nb": self.pct_nb,
                "pct_th": self.pct_th,
                "pct_vd": self.pct_vd,
                "pct_vdc": self.pct_vdc
            },
            "matrix_rows": matrix_rows,
            "spec_rows": spec_rows,
            "th01": {
                "code": "TH01",
                "mcqs": chosen_mcqs,
                "tfs": chosen_tfs,
                "essays": chosen_essays
            },
            "th02": {
                "code": "TH02",
                "mcqs": th02_mcqs_shuffled,
                "tfs": th02_tfs,
                "essays": th02_essays
            }
        }

class ExamDocxExporter:
    """
    Xuất bộ hồ sơ đề thi thành file Word (.docx) chuẩn Thông tư Bộ GD&ĐT 100%:
    - Section 1 (Landscape): Bảng Khung Ma Trận Đề Kiểm Tra chuẩn 18 cột có cell merge chuẩn xác
    - Section 2 (Landscape): Bảng Đặc Tả Đề Kiểm Tra chuẩn 13 cột & Ký duyệt
    - Section 3 (Portrait): 
        + Đề thi TH01 & TH02 gồm:
            * Bảng Header thông tin học sinh chuẩn (Table 3 & Table 9)
            * Bảng Khung điểm & Giám khảo & Phách chuẩn (Table 4 & Table 10)
            * Phần I (MCQ 12 câu, 3.0 điểm)
            * Phần II (Đúng/Sai 4 câu, 4.0 điểm) với 2 BẢNG ĐIỀN ĐÁP ÁN 5x6 ĐẶT NGAY ĐẦU PHẦN II
            * Tờ Tự luận cắt phách (3.0 điểm) với các câu tự luận đánh số tiếp nối (Câu 17, 18, 19)
            * BÀI LÀM và dòng kẻ chấm chấm chuẩn
        + Hướng dẫn chấm & Biểu điểm chi tiết cho cả 2 mã đề:
            * Bảng đáp án MCQ 2x13
            * 2 Bảng đáp án Đúng/Sai 5x6
            * Bảng thang điểm Tự luận 4x3 chi tiết
            * Chữ ký phê duyệt
    """

    @staticmethod
    def set_cell_border(cell, **kwargs):
        """Thiết lập viền bảng Word bằng XML (top, bottom, left, right, tl2br, tr2bl)."""
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = tcPr.first_child_found_in("w:tcBorders")
        if tcBorders is None:
            tcBorders = parse_xml(r'<w:tcBorders {} />'.format(nsdecls('w')))
            tcPr.append(tcBorders)
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'):
            edge_data = kwargs.get(edge)
            if edge_data:
                tag = 'w:{}'.format(edge)
                element = tcBorders.find(qn(tag))
                if element is None:
                    element = parse_xml(r'<w:{} {} />'.format(edge, nsdecls('w')))
                    tcBorders.append(element)
                for key, val in edge_data.items():
                    element.set(qn('w:{}'.format(key)), str(val))

    @staticmethod
    def set_cell_shading(cell, color_hex: str = "FFFFFF"):
        """Đặt màu nền ô bảng (Shading) - Luôn để bảng trắng tinh (không màu) theo yêu cầu."""
        tcPr = cell._tc.get_or_add_tcPr()
        for shd in tcPr.findall(qn('w:shd')):
            tcPr.remove(shd)

    @staticmethod
    def _apply_table_grid_and_widths(table, col_widths_dxa: list, total_width_dxa: int = None):
        """
        Thiết lập tblGrid, tblW và tcW cho từng ô trong bảng Word chuẩn OpenXML.
        Hỗ trợ chính xác các ô đã merge (gridSpan), đảm bảo bảng căng đều 100% lề trang.
        """
        widths_dxa = []
        for w in col_widths_dxa:
            if hasattr(w, 'twips'):
                widths_dxa.append(int(w.twips))
            elif hasattr(w, 'pt'):
                widths_dxa.append(int(w.pt * 20))
            else:
                widths_dxa.append(int(w))

        total_w = total_width_dxa if total_width_dxa is not None else sum(widths_dxa)
        table.autofit = False
        tblPr = table._tbl.tblPr

        # 1. Thiết lập tblW trên tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w}" w:type="dxa"/>')
            tblPr.append(tblW)
        else:
            tblW.set(qn('w:w'), str(total_w))
            tblW.set(qn('w:type'), 'dxa')

        # 1.1 Kích hoạt AutoFit Content (tự động co giãn theo nội dung văn bản)
        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is None:
            tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>'))
        else:
            tblLayout.set(qn('w:type'), 'autofit')

        # 2. Tạo hoặc thay thế tblGrid
        tblGrid = table._tbl.find(qn('w:tblGrid'))
        if tblGrid is not None:
            table._tbl.remove(tblGrid)
        new_tblGrid = parse_xml(
            f'<w:tblGrid {nsdecls("w")}>' +
            ''.join([f'<w:gridCol {nsdecls("w")} w:w="{w}"/>' for w in widths_dxa]) +
            '</w:tblGrid>'
        )
        table._tbl.insert(table._tbl.index(tblPr) + 1, new_tblGrid)

        # 3. Gán tcW cho từng ô dựa theo gridSpan
        for row in table.rows:
            tr = row._tr
            col_idx = 0
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.get_or_add_tcPr()
                gridSpan = tcPr.find(qn('w:gridSpan'))
                span = int(gridSpan.get(qn('w:val'))) if gridSpan is not None else 1

                if col_idx < len(widths_dxa):
                    cell_w = sum(widths_dxa[col_idx : col_idx + span])
                else:
                    cell_w = widths_dxa[-1] if widths_dxa else 1000
                col_idx += span

                # Thiết lập căn giữa dọc cho ô
                vAlign = tcPr.find(qn('w:vAlign'))
                if vAlign is None:
                    tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                else:
                    vAlign.set(qn('w:val'), 'center')

                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{cell_w}" w:type="dxa"/>')
                    tcPr.append(tcW)
                else:
                    tcW.set(qn('w:w'), str(cell_w))
                    tcW.set(qn('w:type'), 'dxa')

    @classmethod
    def _set_col_widths(cls, table, widths, total_width_dxa: int = 10206):
        """Thiết lập độ rộng cột cho bảng Word đồng nhất và full chiều rộng vùng in."""
        cls._apply_table_grid_and_widths(table, widths, total_width_dxa=total_width_dxa)

    @staticmethod
    def _set_row_height_exact(row, height_dxa: int):
        """Đặt chiều cao chính xác (exact height) cho hàng bảng Word."""
        trPr = row._tr.get_or_add_trPr()
        trHeight = trPr.find(qn('w:trHeight'))
        if trHeight is None:
            trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{height_dxa}" w:hRule="exact"/>')
            trPr.append(trHeight)
        else:
            trHeight.set(qn('w:val'), str(height_dxa))
            trHeight.set(qn('w:hRule'), 'exact')

    @staticmethod
    def _format_cell(cell, text: str = "", font_size: float = 11.0, bold: bool = False,
                     italic: bool = False, align=WD_ALIGN_PARAGRAPH.CENTER, shading: str = None):
        """Helper định dạng ô bảng chuẩn Times New Roman."""
        if text != "":
            cell.text = text
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(font_size)
            r.bold = bold
            r.italic = italic
        if shading:
            ExamDocxExporter.set_cell_shading(cell, shading)

    @classmethod
    def _ensure_compatibility_settings(cls, doc):
        """
        Thiết lập Word Compatibility Settings chuẩn:
        - Thêm <w:doNotExpandShiftReturn/> để triệt tiêu hoàn toàn lỗi dãn dạt chữ ra 2 bên mép
          khi kết thúc dòng bằng ký tự xuống dòng / Shift+Return / newlines.
        - Thiết lập compatibilityMode = 15 (Word 2013+) để tối ưu bộ engine hiển thị bố cục.
        """
        settings_elm = doc.settings._element
        compat = settings_elm.find(qn('w:compat'))
        if compat is None:
            compat = parse_xml(
                f'<w:compat {nsdecls("w")}>'
                f'<w:doNotExpandShiftReturn/>'
                f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
                f'</w:compat>'
            )
            settings_elm.append(compat)
        else:
            if compat.find(qn('w:doNotExpandShiftReturn')) is None:
                compat.append(parse_xml(f'<w:doNotExpandShiftReturn {nsdecls("w")}/>'))

    @staticmethod
    def _cant_split(row):
        """Chống ngắt trang chia đôi hàng bảng."""
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

    @classmethod
    def _get_exam_dates(cls, meta: Dict[str, Any]) -> Tuple[str, str]:
        """
        Xác định ngày ký biên soạn bản đặc tả và ngày chấm thi phù hợp với kỳ kiểm tra và năm học:
        - Giữa HK1: Ngày làm ma trận 20/10, ngày ra đề 21/10.
        - Hoặc nhận trực tiếp từ config/giao diện: meta['matrix_date'], meta['exam_date']
        """
        import re, datetime
        school_year = meta.get("school_year", "2026 - 2027")
        years = [int(y) for y in re.findall(r"\d{4}", school_year)]
        start_year = years[0] if years else datetime.datetime.now().year
        end_year = years[1] if len(years) > 1 else start_year + 1

        exam_title = meta.get("exam_title", "").upper()

        def _fmt(val, def_d, def_m, def_y, prefix=""):
            word_ngay = "ngày" if prefix else "Ngày"
            if val:
                val = str(val).strip()
                if "tháng" in val.lower() and "năm" in val.lower():
                    return f"{prefix}{val}" if prefix and not val.startswith(prefix) else val
                m = re.search(r'(\d{1,2})[/\-\.](\d{1,2})(?:[/\-\.](\d{4}))?', val)
                if m:
                    d = int(m.group(1))
                    m_ = int(m.group(2))
                    y = int(m.group(3)) if m.group(3) else def_y
                    return f"{prefix}{word_ngay} {d:02d} tháng {m_:02d} năm {y}"
                return f"{prefix}{val}"
            return f"{prefix}{word_ngay} {def_d:02d} tháng {def_m:02d} năm {def_y}"

        if "HỌC KÌ II" in exam_title or "HK II" in exam_title or "HK2" in exam_title:
            if "GIỮA" in exam_title:
                def_mat = (15, 3, end_year)
                def_ex = (16, 3, end_year)
            else:
                def_mat = (10, 5, end_year)
                def_ex = (11, 5, end_year)
        elif "CUỐI HỌC KÌ I" in exam_title or "CUỐI HK1" in exam_title:
            def_mat = (15, 12, start_year)
            def_ex = (16, 12, start_year)
        else:
            # Mặc định Giữa Học kỳ I: Ngày làm ma trận 20/10, ngày ra đề 21/10
            def_mat = (20, 10, start_year)
            def_ex = (21, 10, start_year)

        raw_mat = meta.get("matrix_date")
        raw_ex = meta.get("exam_date")

        spec_date = _fmt(raw_mat, def_mat[0], def_mat[1], def_mat[2])
        guide_date = _fmt(raw_ex, def_ex[0], def_ex[1], def_ex[2], prefix="Kon Chiêng, ")

        return spec_date, guide_date


    @classmethod
    def _render_x_cut_box(cls, doc):
        """
        Render ô chữ nhật gạch chéo chữ X ở đầu trang mặt sau (ô che phách mặt sau khi in 2 mặt):
        Khớp nhau chính xác 100% với Header mặt trước (chiều cao 2376 dxa, bề rộng 9922 dxa),
        để khi rọc phách xuyên qua 2 mặt thì đường kéo ở mặt sau cắt đúng ô gạch chéo X này,
        hoàn toàn không bị phạm vào nội dung câu hỏi!
        """
        t_box = doc.add_table(rows=1, cols=1)
        t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_box.autofit = False
        cls._cant_split(t_box.rows[0])
        # Chiều cao chính xác 2640 dxa (132 pt) để đường kéo mặt sau khớp 100% với mặt trước (Y = 177.00 pt)
        cls._set_row_height_exact(t_box.rows[0], 2745)

        c = t_box.cell(0, 0)
        c.width = Inches(7.087)
        cls.set_cell_border(
            c,
            top={'val': 'single', 'sz': 4, 'space': 0, 'color': 'auto'},
            left={'val': 'single', 'sz': 4, 'space': 0, 'color': 'auto'},
            bottom={'val': 'single', 'sz': 4, 'space': 0, 'color': 'auto'},
            right={'val': 'single', 'sz': 4, 'space': 0, 'color': 'auto'},
            tl2br={'val': 'single', 'sz': 4, 'space': 0, 'color': 'auto'},
            tr2bl={'val': 'single', 'sz': 4, 'space': 0, 'color': 'auto'}
        )

        # 1 đoạn văn giữ cấu trúc XML chuẩn của Word (không chèn các đoạn văn rỗng gây phình to khoảng trống)
        p_in = c.paragraphs[0]
        p_in.paragraph_format.space_before = Pt(0)
        p_in.paragraph_format.space_after = Pt(0)
        r_in = p_in.add_run()
        r_in.font.size = Pt(4)

        cls._set_col_widths(t_box, [10206], total_width_dxa=10206)

        # Đường kéo cắt phách ngay dưới ô gạch chéo
        cls._render_cut_line(doc)
        return t_box

    @classmethod
    def _render_cut_line(cls, doc):
        """Render đường kéo cắt phách mép giấy sát mép trái, chuỗi chấm lửng rõ ràng chuẩn 100% file mẫu, không tràn dòng 2."""
        p_cut = doc.add_paragraph()
        p_cut.paragraph_format.space_before = Pt(2)
        p_cut.paragraph_format.space_after = Pt(2)
        p_cut.paragraph_format.keep_with_next = True
        p_cut.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_c = p_cut.add_run("✂" + "…" * 40)
        r_c.font.name = "Times New Roman"
        r_c.font.size = Pt(12)
        r_c.font.color.rgb = RGBColor(100, 100, 100)

    @classmethod
    def _render_dot_lines(cls, doc, count: int = 18, width_inches: float = 7.087):
        """
        Tạo các dòng chấm thực tế chuẩn 100% file đề mẫu, lấp đầy full trang giấy A4:
        - Sử dụng Tab Stop Right với Dot Leader (w:leader="dot") tại đúng tọa độ lề phải (7.087 in / 10206 dxa).
        - Phủ kín 100% chiều ngang sát mép lề phải giấy A4, không bao giờ bị hụt hay rớt dòng.
        - Chiều cao các dòng chấm được giãn cách cân đối (line_spacing 1.35) phủ kín trang đến sát mép đáy.
        """
        for i in range(count):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2.0)
            p.paragraph_format.line_spacing = 1.35
            p.paragraph_format.keep_with_next = (i < count - 1)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(width_inches), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            r = p.add_run("\t")
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(100, 100, 100)

    @classmethod
    def _render_header_box(cls, doc, meta: Dict[str, Any], code: str, is_essay: bool = False):
        """
        Render bảng tiêu đề hành chính chuẩn 100% file đề mẫu:
        - Căn lề full chiều ngang trang A4 (9922 dxa ~ 6.89 inches).
        - Đặt chiều cao chính xác cho 4 hàng để tổng chiều cao đúng bằng 2376 dxa (~4.19 cm).
        - Ô Số phách được đóng khung hộp riêng biệt có nền xám nhẹ.
        - Dòng thông tin học sinh có đường kẻ phân cách dọc với ô Số phách.
        """
        grade_num = meta['grade'].upper().replace('LỚP ', '').strip()
        t_top = doc.add_table(rows=4, cols=3)
        t_top.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_top.autofit = False

        tblPr = t_top._tbl.tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="10206" w:type="dxa"/>'))
        else:
            tblW.set(qn('w:w'), "10206")
            tblW.set(qn('w:type'), "dxa")

        # Thực hiện Cell Merges trước
        t_top.cell(0, 1).merge(t_top.cell(0, 2))
        t_top.cell(1, 1).merge(t_top.cell(1, 2))
        t_top.cell(2, 1).merge(t_top.cell(2, 2))
        t_top.cell(3, 0).merge(t_top.cell(3, 1))

        # Hàng 0: Tiêu đề hành chính bên trái (UBND, Trường, Gạch chân) & Tên kỳ thi bên phải
        c_gov = t_top.cell(0, 0)
        p_gov = c_gov.paragraphs[0]
        p_gov.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_gov.paragraph_format.line_spacing = 1.05
        p_gov.paragraph_format.space_before = Pt(0)
        p_gov.paragraph_format.space_after = Pt(0.5)
        r_gov = p_gov.add_run(meta['governing_body'].upper())
        r_gov.font.name = "Times New Roman"
        r_gov.font.size = Pt(12.0)
        r_gov.bold = False

        p_school = c_gov.add_paragraph()
        p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_school.paragraph_format.line_spacing = 1.05
        p_school.paragraph_format.space_before = Pt(0)
        p_school.paragraph_format.space_after = Pt(0)
        r_school = p_school.add_run(meta['school_name'].upper())
        r_school.font.name = "Times New Roman"
        r_school.font.size = Pt(12.0)
        r_school.bold = True

        # Dấu gạch ngang nét liền dưới chân TRƯỜNG TH&THCS ĐĂK TRÔI chuẩn Nghị định 30/2020
        p_line = c_gov.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.line_spacing = 0.9
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(0)
        r_line = p_line.add_run("───────────")
        r_line.font.name = "Times New Roman"
        r_line.font.size = Pt(7.5)
        r_line.bold = True

        c_exam = t_top.cell(0, 1)
        p_ex_title = c_exam.paragraphs[0]
        p_ex_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ex_title.paragraph_format.line_spacing = 1.05
        p_ex_title.paragraph_format.space_before = Pt(0)
        p_ex_title.paragraph_format.space_after = Pt(0.5)
        r_ex_title = p_ex_title.add_run("ĐỀ KIỂM TRA GIỮA HỌC KỲ I")
        r_ex_title.font.name = "Times New Roman"
        r_ex_title.font.size = Pt(13.0)
        r_ex_title.bold = True

        p_ex_year = c_exam.add_paragraph()
        p_ex_year.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ex_year.paragraph_format.line_spacing = 1.05
        p_ex_year.paragraph_format.space_before = Pt(0)
        p_ex_year.paragraph_format.space_after = Pt(0)
        r_ex_year = p_ex_year.add_run(f"NĂM HỌC {meta['school_year']}")
        r_ex_year.font.name = "Times New Roman"
        r_ex_year.font.size = Pt(12.5)
        r_ex_year.bold = True

        # Hàng 1:
        t_top.cell(1, 0).paragraphs[0].text = ""
        c_mon = t_top.cell(1, 1)
        p_mon = c_mon.paragraphs[0]
        p_mon.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_mon.paragraph_format.space_before = Pt(0.5)
        p_mon.paragraph_format.space_after = Pt(0.5)
        p_mon.paragraph_format.line_spacing = 1.05
        r_mon = p_mon.add_run(f"MÔN: TIN HỌC {grade_num}")
        r_mon.font.name = "Times New Roman"
        r_mon.font.size = Pt(13.0)
        r_mon.bold = True

        # Hàng 2:
        c_code = t_top.cell(2, 0)
        c_code.paragraphs[0].text = f"ĐỀ CHÍNH THỨC\nMÃ ĐỀ: {code}"
        p_code = c_code.paragraphs[0]
        p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_code.paragraph_format.line_spacing = 1.1
        p_code.paragraph_format.space_before = Pt(1)
        p_code.paragraph_format.space_after = Pt(1)
        if p_code.runs:
            p_code.runs[0].font.name = "Times New Roman"
            p_code.runs[0].font.size = Pt(11)
            p_code.runs[0].bold = True

        c_time = t_top.cell(2, 1)
        if is_essay:
            c_time.paragraphs[0].text = "Phần tự luận: 3 điểm -Thời gian: 13,5 phút\n(không kể thời gian phát đề)"
        else:
            c_time.paragraphs[0].text = "Phần trắc nghiệm: 7 điểm -Thời gian: 31,5 phút\n(không kể thời gian phát đề)"
        p_time = c_time.paragraphs[0]
        p_time.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_time.paragraph_format.line_spacing = 1.1
        p_time.paragraph_format.space_before = Pt(1)
        p_time.paragraph_format.space_after = Pt(1)
        if p_time.runs:
            p_time.runs[0].font.name = "Times New Roman"
            p_time.runs[0].font.size = Pt(10.5)

        # Hàng 3:
        c_stu = t_top.cell(3, 0)
        p_stu0 = c_stu.paragraphs[0]
        p_stu0.text = "Họ và tên: ............................................ - Lớp: ........ - Phòng: ........"
        p_stu0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_stu0.paragraph_format.space_before = Pt(2)
        p_stu0.paragraph_format.space_after = Pt(2)
        p_stu0.paragraph_format.line_spacing = 1.2
        if p_stu0.runs:
            p_stu0.runs[0].font.name = "Times New Roman"
            p_stu0.runs[0].font.size = Pt(10.5)

        p_stu1 = c_stu.add_paragraph()
        p_stu1.text = "SBD: ............................................................................................"
        p_stu1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_stu1.paragraph_format.space_before = Pt(1)
        p_stu1.paragraph_format.space_after = Pt(2)
        p_stu1.paragraph_format.line_spacing = 1.2
        if p_stu1.runs:
            p_stu1.runs[0].font.name = "Times New Roman"
            p_stu1.runs[0].font.size = Pt(10.5)

        c_phach = t_top.cell(3, 2)
        p_phach0 = c_phach.paragraphs[0]
        p_phach0.text = "Số phách"
        p_phach0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_phach0.paragraph_format.space_before = Pt(3)
        p_phach0.paragraph_format.space_after = Pt(1)
        p_phach0.paragraph_format.line_spacing = 1.2
        if p_phach0.runs:
            p_phach0.runs[0].font.name = "Times New Roman"
            p_phach0.runs[0].font.size = Pt(10.5)
            p_phach0.runs[0].bold = True

        p_phach1 = c_phach.add_paragraph()
        p_phach1.text = ""
        p_phach1.paragraph_format.line_spacing = 1.2

        # Thiết lập độ rộng ô chính xác tuyệt đối qua XML:
        # Col 0: 4330 dxa (~3.01 in), Col 1: 4330 dxa, Col 2: 1546 dxa (~1.07 in)
        # Hàng 0, 1, 2: c0 = 4330 dxa, c1 (gộp) = 5876 dxa (~4.08 in)
        for ri in [0, 1, 2]:
            c_left = t_top.cell(ri, 0)
            c_left.width = Inches(3.007)
            tcPr0 = c_left._tc.get_or_add_tcPr()
            tcW0 = tcPr0.find(qn('w:tcW'))
            if tcW0 is None:
                tcPr0.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="4330" w:type="dxa"/>'))
            else:
                tcW0.set(qn('w:w'), "4330")

            c_right = t_top.cell(ri, 1)
            c_right.width = Inches(4.080)
            tcPr1 = c_right._tc.get_or_add_tcPr()
            tcW1 = tcPr1.find(qn('w:tcW'))
            if tcW1 is None:
                tcPr1.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="5876" w:type="dxa"/>'))
            else:
                tcW1.set(qn('w:w'), "5876")

        # Hàng 3: c0 (gộp 0-1) = 8660 dxa (~6.01 in), c2 = 1546 dxa (~1.07 in)
        c30 = t_top.cell(3, 0)
        c30.width = Inches(6.014)
        tcPr30 = c30._tc.get_or_add_tcPr()
        tcW30 = tcPr30.find(qn('w:tcW'))
        if tcW30 is None:
            tcPr30.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="8660" w:type="dxa"/>'))
        else:
            tcW30.set(qn('w:w'), "8660")

        c32 = t_top.cell(3, 2)
        c32.width = Inches(1.074)
        tcPr32 = c32._tc.get_or_add_tcPr()
        tcW32 = tcPr32.find(qn('w:tcW'))
        if tcW32 is None:
            tcPr32.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="1546" w:type="dxa"/>'))
        else:
            tcW32.set(qn('w:w'), "1546")

        # Khai báo tblGrid 3 cột chuẩn
        tblGrid_top = parse_xml(f'<w:tblGrid {nsdecls("w")}><w:gridCol w:w="4330"/><w:gridCol w:w="4330"/><w:gridCol w:w="1546"/></w:tblGrid>')
        t_top._tbl.insert(t_top._tbl.index(t_top._tbl.tblPr) + 1, tblGrid_top)

        # Thiết lập viền: tất cả không viền, riêng ô Số phách có viền hộp 4 cạnh
        for row in t_top.rows:
            for cell in row.cells:
                cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))

        c32._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="6" w:space="0" w:color="auto"/></w:tcBorders>'))

        # Chống ngắt hàng, không đặt exact height để không bao giờ bị cắt chữ
        for r in t_top.rows:
            cls._cant_split(r)

        return t_top

    @classmethod
    def _render_score_box(cls, doc, code: str, is_essay: bool = False):
        """Render khung Điểm & Giám khảo & Phách chuẩn."""
        t_score = doc.add_table(rows=4, cols=4)
        t_score.style = 'Table Grid'
        t_score.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_score.autofit = False

        tblPr_sc = t_score._tbl.tblPr
        tblW_sc = tblPr_sc.find(qn('w:tblW'))
        if tblW_sc is None:
            tblPr_sc.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="10206" w:type="dxa"/>'))
        else:
            tblW_sc.set(qn('w:w'), "10206")
            tblW_sc.set(qn('w:type'), "dxa")

        t_score.cell(0, 0).text = f"MÃ ĐỀ: {code}"
        t_score.cell(0, 1).text = "Họ và tên, chữ ký giám khảo"
        t_score.cell(0, 3).text = "Số phách"
        t_score.cell(1, 0).text = "Điểm"
        t_score.cell(1, 1).text = "Giám khảo 1"
        t_score.cell(1, 2).text = "Giám khảo 2"
        t_score.cell(2, 1).text = "\n........................"
        t_score.cell(2, 2).text = "\n........................"
        
        if is_essay:
            t_score.cell(3, 0).text = "( Phần tự luận, học sinh làm bài trực tiếp trên đề kiểm tra )"
        else:
            t_score.cell(3, 0).text = "( Phần trắc nghiệm gồm 04 trang, học sinh làm bài trực tiếp trên đề kiểm tra )"

        # Merges
        t_score.cell(0, 1).merge(t_score.cell(0, 2))
        t_score.cell(0, 3).merge(t_score.cell(2, 3))
        t_score.cell(1, 0).merge(t_score.cell(2, 0))
        t_score.cell(3, 0).merge(t_score.cell(3, 3))

        for row in t_score.rows:
            for cell in row.cells:
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(2.0)
                p.paragraph_format.space_after = Pt(2.0)
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(12)

        # Định dạng font chuẩn, to rõ ràng đồng bộ cho từng ô trong Khung điểm:
        p_m = t_score.cell(0, 0).paragraphs[0]
        if p_m.runs:
            p_m.runs[0].bold = True
            p_m.runs[0].font.size = Pt(12.5)

        p_gk = t_score.cell(0, 1).paragraphs[0]
        if p_gk.runs:
            p_gk.runs[0].bold = True
            p_gk.runs[0].font.size = Pt(12.5)

        p_sp = t_score.cell(0, 3).paragraphs[0]
        if p_sp.runs:
            p_sp.runs[0].bold = True
            p_sp.runs[0].font.size = Pt(12.5)

        p_d = t_score.cell(1, 0).paragraphs[0]
        if p_d.runs:
            p_d.runs[0].bold = True
            p_d.runs[0].font.size = Pt(12.5)

        p_gk1 = t_score.cell(1, 1).paragraphs[0]
        if p_gk1.runs:
            p_gk1.runs[0].bold = True
            p_gk1.runs[0].font.size = Pt(12)

        p_gk2 = t_score.cell(1, 2).paragraphs[0]
        if p_gk2.runs:
            p_gk2.runs[0].bold = True
            p_gk2.runs[0].font.size = Pt(12)

        # Dòng chấm chữ ký giám khảo
        for c_dots in [t_score.cell(2, 1), t_score.cell(2, 2)]:
            if c_dots.paragraphs[0].runs:
                c_dots.paragraphs[0].runs[0].font.size = Pt(11)

        p_note = t_score.cell(3, 0).paragraphs[0]
        if p_note.runs:
            p_note.runs[0].italic = True
            p_note.runs[0].font.size = Pt(11.5)

        # Thiết lập độ rộng ô chính xác qua XML cho bảng điểm:
        # Col 0: 2400 dxa (1.67 in), Col 1: 3130 dxa (2.17 in), Col 2: 3130 dxa (2.17 in), Col 3: 1546 dxa (1.07 in)
        def _set_tc_w(cell, w_dxa):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))
            else:
                tcW.set(qn('w:w'), str(w_dxa))
            cell.width = Inches(w_dxa / 1440)

        _set_tc_w(t_score.cell(0, 0), 2400)
        _set_tc_w(t_score.cell(0, 1), 6260)
        _set_tc_w(t_score.cell(0, 3), 1546)

        _set_tc_w(t_score.cell(1, 0), 2400)
        _set_tc_w(t_score.cell(1, 1), 3130)
        _set_tc_w(t_score.cell(1, 2), 3130)

        _set_tc_w(t_score.cell(2, 1), 3130)
        _set_tc_w(t_score.cell(2, 2), 3130)

        _set_tc_w(t_score.cell(3, 0), 10206)

        tblGrid_sc = parse_xml(f'<w:tblGrid {nsdecls("w")}><w:gridCol w:w="2400"/><w:gridCol w:w="3130"/><w:gridCol w:w="3130"/><w:gridCol w:w="1546"/></w:tblGrid>')
        t_score._tbl.insert(t_score._tbl.index(t_score._tbl.tblPr) + 1, tblGrid_sc)

        for r in t_score.rows:
            cls._cant_split(r)

        return t_score

    @classmethod
    def _render_mcq_options(cls, doc, options: Dict[str, str], is_p2: bool = False):
        """
        Render 4 phương án A, B, C, D bằng Paragraphs căn lề chuẩn và Tab Stop cố định:
        - Cấu hình 1 (max_opt_len <= 14): Cả 4 phương án trên 1 dòng.
        - Cấu hình 2 (max_opt_len <= 56): 2 phương án/dòng (A-B và C-D).
        - Cấu hình 3 (max_opt_len > 56): Mỗi phương án 1 dòng riêng.
        - is_p2: Tối ưu giãn dòng và cỡ chữ vừa vặn hoàn hảo cho Trang 2.
        """
        opt_keys = ["A", "B", "C", "D"]
        max_opt_len = max(len(str(options.get(k, ""))) for k in opt_keys)
        fs = Pt(13)
        sp_b = Pt(0.2 if is_p2 else 0.3)
        sp_a = Pt(0.4 if is_p2 else 0.6)
        ls = 1.04

        if max_opt_len <= 14:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.20)
            p.paragraph_format.space_before = sp_b
            p.paragraph_format.space_after = sp_a
            p.paragraph_format.line_spacing = ls
            p.paragraph_format.keep_with_next = False
            
            p.paragraph_format.tab_stops.add_tab_stop(Inches(1.85), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(3.60), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(5.35), WD_TAB_ALIGNMENT.LEFT)

            for i, k in enumerate(opt_keys):
                if i > 0:
                    p.add_run("\t")
                r_lbl = p.add_run(f"{k}. ")
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = fs
                r_lbl.bold = True
                
                r_val = p.add_run(str(options.get(k, "")))
                r_val.font.name = "Times New Roman"
                r_val.font.size = fs

        elif max_opt_len <= 56:
            pairs = [("A", "B"), ("C", "D")]
            for p_idx, (k1, k2) in enumerate(pairs):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.20)
                p.paragraph_format.space_before = sp_b
                p.paragraph_format.space_after = sp_a
                p.paragraph_format.line_spacing = ls
                p.paragraph_format.keep_with_next = (p_idx == 0)
                p.paragraph_format.tab_stops.add_tab_stop(Inches(3.60), WD_TAB_ALIGNMENT.LEFT)

                r_lbl1 = p.add_run(f"{k1}. ")
                r_lbl1.font.name = "Times New Roman"
                r_lbl1.font.size = fs
                r_lbl1.bold = True
                r_val1 = p.add_run(str(options.get(k1, "")))
                r_val1.font.name = "Times New Roman"
                r_val1.font.size = fs

                p.add_run("\t")
                r_lbl2 = p.add_run(f"{k2}. ")
                r_lbl2.font.name = "Times New Roman"
                r_lbl2.font.size = fs
                r_lbl2.bold = True
                r_val2 = p.add_run(str(options.get(k2, "")))
                r_val2.font.name = "Times New Roman"
                r_val2.font.size = fs

        else:
            for i, k in enumerate(opt_keys):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.20)
                p.paragraph_format.space_before = sp_b
                p.paragraph_format.space_after = sp_a
                p.paragraph_format.line_spacing = ls
                p.paragraph_format.keep_with_next = (i < 3)

                r_lbl = p.add_run(f"{k}. ")
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = fs
                r_lbl.bold = True
                
                r_val = p.add_run(str(options.get(k, "")))
                r_val.font.name = "Times New Roman"
                r_val.font.size = fs

    @classmethod
    def _estimate_mcq_points(cls, q: Dict[str, Any]) -> float:
        """Ước tính chính xác độ cao tính theo đơn vị points của 1 câu trắc nghiệm."""
        content = q.get("content", "")
        stem_len = len(content) + 12
        stem_lines = max(1, (stem_len + 84) // 85)
        stem_pt = stem_lines * 11.55 + 2.0

        opts = q.get("options", {})
        opt_keys = ["A", "B", "C", "D"]
        max_opt_len = max((len(str(opts.get(k, ""))) for k in opt_keys), default=0)

        if max_opt_len <= 14:
            opt_pt = 11.55 + 1.2
        elif max_opt_len <= 56:
            opt_pt = 2 * (11.55 + 1.2)
        else:
            opt_lines = sum(max(1, (len(str(opts.get(k, ""))) + 80) // 80) for k in opt_keys)
            opt_pt = opt_lines * 11.55 + 4 * 1.2

        return stem_pt + opt_pt

    @classmethod
    def _estimate_tf_points(cls, tf_item: Dict[str, Any]) -> float:
        """Ước tính độ cao (pt) của 1 câu Đúng/Sai."""
        top_txt = tf_item.get("topic", "")[:65]
        top_lines = max(1, (len(top_txt) + 15) // 85)
        top_pt = top_lines * 11.55 + 2.0
        stmts_pt = 0.0
        for stmt in tf_item.get("statements", []):
            s_lines = max(1, (len(stmt.get("content", "")) + 8) // 85)
            stmts_pt += s_lines * 11.55 + 1.2
        return top_pt + stmts_pt

    @classmethod
    def _calculate_optimal_page1_mcq_count(cls, mcqs: List[Dict[str, Any]], tfs: List[Dict[str, Any]] = None) -> int:
        """
        Tính toán thông minh số câu dồn lên Trang 1 sao cho:
        1. Dồn tối đa các câu hỏi lên Trang 1 để lấp đầy trang, không để lại khoảng trống trắng ở cuối.
        2. Cân bằng tối ưu giữa Trang 1 và Trang 2, đảm bảo cả 2 trang vừa vặn hoàn hảo, không bị tràn (dưới 735 pt).
        """
        if not mcqs:
            return 0
        total_mcqs = len(mcqs)
        tf_list = tfs or []
        # Chiều cao cố định trên Trang 2: X-box (93) + Tiêu đề Phần II (16) + Bảng 12 cột (65) + các câu Đúng/Sai
        tf_total_pt = sum(cls._estimate_tf_points(tf) for tf in tf_list)
        fixed_p2_pt = 93.0 + 16.0 + 65.0 + tf_total_pt

        best_k = min(total_mcqs, 7)
        best_diff = 9999.0

        # Thử k từ 10 xuống 5 (dồn tối đa lên Trang 1, ưu tiên dồn nhiều câu lên nếu vừa)
        for k in range(min(total_mcqs, 10), 4, -1):
            p1_mcqs_pt = sum(cls._estimate_mcq_points(q) for q in mcqs[:k])
            p2_mcqs_pt = sum(cls._estimate_mcq_points(q) for q in mcqs[k:])
            p2_total_pt = fixed_p2_pt + p2_mcqs_pt
            p1_full_pt = 208.0 + p1_mcqs_pt

            if p1_full_pt <= 735.0 and p2_total_pt <= 735.0:
                diff = abs(p1_full_pt - p2_total_pt)
                if diff < best_diff:
                    best_diff = diff
                    best_k = k

        return best_k

    @classmethod
    def export(cls, dossier_data: Dict[str, Any], output_path: str) -> str:
        doc = docx.Document()
        cls._ensure_compatibility_settings(doc)
        meta = dossier_data["meta"]
        grade_num = meta['grade'].upper().replace('LỚP ', '').strip()

        # ======================================================================
        # SECTION 1 (LANDSCAPE): KHUNG MA TRẬN ĐỀ KIỂM TRA
        # ======================================================================
        sec1 = doc.sections[0]
        sec1.orientation = WD_ORIENT.LANDSCAPE
        sec1.page_width = Inches(11.69)   # A4 Landscape width
        sec1.page_height = Inches(8.27)   # A4 Landscape height
        sec1.top_margin = Inches(0.79)    # 2.0 cm
        sec1.bottom_margin = Inches(0.39) # 1.0 cm
        sec1.left_margin = Inches(0.39)   # 1.0 cm
        sec1.right_margin = Inches(0.39)  # 1.0 cm

        # Tiêu đề hành chính đầu trang
        t_hdr = doc.add_table(rows=1, cols=3)
        t_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_hdr.autofit = False
        cls._apply_table_grid_and_widths(t_hdr, [4800, 6110, 4800])

        # Ô trái: UBND / Trường
        c0 = t_hdr.rows[0].cells[0]
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(f"{meta['governing_body'].upper()}\n")
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(11)
        r0.bold = True
        r0_2 = p0.add_run(f"{meta['school_name'].upper()}")
        r0_2.font.name = "Times New Roman"
        r0_2.font.size = Pt(11)
        r0_2.bold = True

        # Ô giữa: Tên kỳ kiểm tra & Môn
        c1 = t_hdr.rows[0].cells[1]
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(f"{meta['exam_title'].upper()}, NĂM HỌC {meta['school_year']}\n")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(12)
        r1.bold = True
        r1_2 = p1.add_run(f"MÔN: TIN HỌC {grade_num}\n")
        r1_2.font.name = "Times New Roman"
        r1_2.font.size = Pt(12)
        r1_2.bold = True
        r1_3 = p1.add_run("Thời gian làm bài: 45 phút")
        r1_3.font.name = "Times New Roman"
        r1_3.font.size = Pt(11)
        r1_3.italic = True

        # Ô phải: Tiêu đề khung ma trận
        c2 = t_hdr.rows[0].cells[2]
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run("KHUNG MA TRẬN VÀ BẢN ĐẶC TẢ\nĐỀ KIỂM TRA ĐỊNH KỲ\n")
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(11)
        r2.bold = True
        r2_sub = p2.add_run(f"Tỉ lệ: {meta.get('pct_nb', 40)}% Biết - {meta.get('pct_th', 30)}% Hiểu - {meta.get('pct_vd', 30)}% Vận dụng")
        r2_sub.font.name = "Times New Roman"
        r2_sub.font.size = Pt(10)
        r2_sub.italic = True

        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(4)
        p_sp.paragraph_format.space_after = Pt(4)
        r_sp = p_sp.add_run(f"I. KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KỲ MÔN TIN HỌC, LỚP {grade_num}")
        r_sp.font.name = "Times New Roman"
        r_sp.font.size = Pt(13)
        r_sp.bold = True

        # Bảng Khung Ma Trận (18 cột chuẩn mực Thông tư 22)
        t_mat = doc.add_table(rows=4, cols=18)
        t_mat.style = 'Table Grid'
        t_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_mat.autofit = False

        # 1. Điền text tiêu đề
        t_mat.cell(0, 0).text = "TT"
        t_mat.cell(0, 1).text = "Chương/ chủ đề"
        t_mat.cell(0, 2).text = "Nội dung/đơn vị kiến thức"
        t_mat.cell(0, 3).text = "Mức độ đánh giá"
        t_mat.cell(0, 12).text = "Tổng"
        t_mat.cell(0, 17).text = "Tỉ lệ % điểm"

        t_mat.cell(1, 3).text = "TNKQ"
        t_mat.cell(1, 9).text = "Tự luận"

        t_mat.cell(2, 3).text = "Nhiều lựa chọn"
        t_mat.cell(2, 6).text = "Đúng - Sai"

        t_mat.cell(3, 3).text = "Biết"
        t_mat.cell(3, 4).text = "Hiểu"
        t_mat.cell(3, 5).text = "Vận dụng"
        t_mat.cell(3, 6).text = "Biết"
        t_mat.cell(3, 7).text = "Hiểu"
        t_mat.cell(3, 8).text = "Vận dụng"
        t_mat.cell(3, 9).text = "Biết"
        t_mat.cell(3, 10).text = "Hiểu"
        t_mat.cell(3, 11).text = "Vận dụng"
        t_mat.cell(3, 12).text = "Biết"
        t_mat.cell(3, 14).text = "Hiểu"
        t_mat.cell(3, 16).text = "Vận dụng"

        # 2. Thực hiện Cell Merges chuẩn Thông tư
        t_mat.cell(0, 0).merge(t_mat.cell(3, 0))
        t_mat.cell(0, 1).merge(t_mat.cell(3, 1))
        t_mat.cell(0, 2).merge(t_mat.cell(3, 2))
        t_mat.cell(0, 3).merge(t_mat.cell(0, 11))
        t_mat.cell(0, 12).merge(t_mat.cell(2, 16))
        t_mat.cell(0, 17).merge(t_mat.cell(3, 17))

        t_mat.cell(1, 3).merge(t_mat.cell(1, 8))
        t_mat.cell(1, 9).merge(t_mat.cell(2, 11))

        t_mat.cell(2, 3).merge(t_mat.cell(2, 5))
        t_mat.cell(2, 6).merge(t_mat.cell(2, 8))

        t_mat.cell(3, 12).merge(t_mat.cell(3, 13))
        t_mat.cell(3, 14).merge(t_mat.cell(3, 15))

        # Format toàn bộ ô trong 4 dòng tiêu đề
        for ri in range(4):
            for ci in range(18):
                c = t_mat.cell(ri, ci)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(11.5)
                    p.runs[0].bold = True
                cls.set_cell_shading(c, "F2F2F2")

        # 3. Thêm các hàng dữ liệu bài học (Tính toán động 100%, đồng bộ cỡ chữ 11.5pt)
        def fmt_pts_ui(val: float) -> str:
            if abs(val - round(val)) < 1e-4:
                return f"{int(round(val))},0"
            return f"{val:.2f}".rstrip("0").replace(".", ",")

        def fmt_pct_ui(val: float) -> str:
            if abs(val - round(val)) < 1e-4:
                return f"{int(round(val))}%"
            return f"{val:.1f}".replace(".", ",") + "%"

        mat_units = dossier_data.get("matrix_rows", {}).get("units", [])
        mat_totals = dossier_data.get("matrix_rows", {}).get("totals", {})

        # Tự động tính toán hoặc bổ sung an toàn 100% các trường thống kê cho mat_totals từ mat_units
        def _get_q_cnt(u_obj: Dict[str, Any], raw_key: str, str_key: str) -> int:
            if raw_key in u_obj and isinstance(u_obj[raw_key], (int, float)):
                return int(u_obj[raw_key])
            s = str(u_obj.get(str_key, "")).strip()
            import re
            m = re.search(r'(\d+)', s)
            return int(m.group(1)) if m else 0

        sum_mcq_nb = sum(_get_q_cnt(u, "raw_mcq_nb", "mcq_nb") for u in mat_units)
        sum_mcq_th = sum(_get_q_cnt(u, "raw_mcq_th", "mcq_th") for u in mat_units)
        sum_mcq_vd = sum(_get_q_cnt(u, "raw_mcq_vd", "mcq_vd") for u in mat_units)

        sum_tf_nb = sum(_get_q_cnt(u, "raw_tf_nb", "tf_nb") for u in mat_units)
        sum_tf_th = sum(_get_q_cnt(u, "raw_tf_th", "tf_th") for u in mat_units)
        sum_tf_vd = sum(_get_q_cnt(u, "raw_tf_vd", "tf_vd") for u in mat_units)

        sum_es_nb = sum(_get_q_cnt(u, "raw_es_nb", "es_nb") for u in mat_units)
        sum_es_th = sum(_get_q_cnt(u, "raw_es_th", "es_th") for u in mat_units)
        sum_es_vd = sum(_get_q_cnt(u, "raw_es_vd", "es_vd") for u in mat_units)

        all_mcq_cnt = sum_mcq_nb + sum_mcq_th + sum_mcq_vd or 12
        all_tf_cnt = sum_tf_nb + sum_tf_th + sum_tf_vd or 4
        all_es_cnt = sum_es_nb + sum_es_th + sum_es_vd or 3

        tot_nb_cnt = sum_mcq_nb + sum_tf_nb + sum_es_nb
        tot_th_cnt = sum_mcq_th + sum_tf_th + sum_es_th
        tot_vd_cnt = sum_mcq_vd + sum_tf_vd + sum_es_vd
        all_q_cnt = tot_nb_cnt + tot_th_cnt + tot_vd_cnt or 19

        pts_mcq_nb = sum_mcq_nb * 0.25; pts_mcq_th = sum_mcq_th * 0.25; pts_mcq_vd = sum_mcq_vd * 0.25
        pts_tf_nb = sum_tf_nb * 1.0; pts_tf_th = sum_tf_th * 1.0; pts_tf_vd = sum_tf_vd * 1.0
        pts_es_nb = sum_es_nb * 1.0; pts_es_th = sum_es_th * 1.0; pts_es_vd = sum_es_vd * 1.0

        tot_mcq_pts = pts_mcq_nb + pts_mcq_th + pts_mcq_vd or 3.0
        tot_tf_pts = pts_tf_nb + pts_tf_th + pts_tf_vd or 4.0
        tot_es_pts = pts_es_nb + pts_es_th + pts_es_vd or 3.0

        tot_nb_pts = pts_mcq_nb + pts_tf_nb + pts_es_nb or 4.0
        tot_th_pts = pts_mcq_th + pts_tf_th + pts_es_th or 3.0
        tot_vd_pts = pts_mcq_vd + pts_tf_vd + pts_es_vd or 3.0
        all_pts = tot_nb_pts + tot_th_pts + tot_vd_pts or 10.0

        pct_mcq = (tot_mcq_pts / 10.0) * 100.0
        pct_tf = (tot_tf_pts / 10.0) * 100.0
        pct_es = (tot_es_pts / 10.0) * 100.0

        pct_nb = (tot_nb_pts / 10.0) * 100.0
        pct_th = (tot_th_pts / 10.0) * 100.0
        pct_vd = (tot_vd_pts / 10.0) * 100.0

        computed_totals = {
            "sum_mcq_nb": sum_mcq_nb, "sum_mcq_th": sum_mcq_th, "sum_mcq_vd": sum_mcq_vd,
            "sum_tf_nb": sum_tf_nb, "sum_tf_th": sum_tf_th, "sum_tf_vd": sum_tf_vd,
            "sum_es_nb": sum_es_nb, "sum_es_th": sum_es_th, "sum_es_vd": sum_es_vd,
            "tot_nb_cnt": tot_nb_cnt, "tot_th_cnt": tot_th_cnt, "tot_vd_cnt": tot_vd_cnt,
            "all_q_cnt": all_q_cnt,
            "pts_mcq_nb": pts_mcq_nb, "pts_mcq_th": pts_mcq_th, "pts_mcq_vd": pts_mcq_vd,
            "pts_tf_nb": pts_tf_nb, "pts_tf_th": pts_tf_th, "pts_tf_vd": pts_tf_vd,
            "pts_es_nb": pts_es_nb, "pts_es_th": pts_es_th, "pts_es_vd": pts_es_vd,
            "tot_mcq_pts": tot_mcq_pts, "tot_tf_pts": tot_tf_pts, "tot_es_pts": tot_es_pts,
            "tot_nb_pts": tot_nb_pts, "tot_th_pts": tot_th_pts, "tot_vd_pts": tot_vd_pts,
            "all_pts": all_pts,
            "pct_mcq": pct_mcq, "pct_tf": pct_tf, "pct_es": pct_es,
            "pct_nb": pct_nb, "pct_th": pct_th, "pct_vd": pct_vd
        }
        for k, v in computed_totals.items():
            if k not in mat_totals or not mat_totals[k]:
                mat_totals[k] = v

        for u in mat_units:
            r = t_mat.add_row()
            # Merge cột 12-13 và 14-15
            r.cells[12].merge(r.cells[13])
            r.cells[14].merge(r.cells[15])

            vals = [
                u["tt"], u["chapter"], u["unit"],
                u["mcq_nb"], u["mcq_th"], u["mcq_vd"],
                u["tf_nb"], u["tf_th"], u["tf_vd"],
                u["es_nb"], u["es_th"], u["es_vd"],
                u.get("tot_nb", ""), u.get("tot_th", ""), u.get("tot_vd", ""), u["total_score"]
            ]
            cell_idx_map = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 17]
            for val, c_idx in zip(vals, cell_idx_map):
                c = r.cells[c_idx]
                c.text = val
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx not in [1, 2] else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(1.5)
                p.paragraph_format.space_after = Pt(1.5)
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(11.5)

        # 4. Hàng Tổng số câu (Tính động 100% từ sum các cột)
        r_sum_q = t_mat.add_row()
        r_sum_q.cells[0].merge(r_sum_q.cells[2])
        r_sum_q.cells[12].merge(r_sum_q.cells[13])
        r_sum_q.cells[14].merge(r_sum_q.cells[15])
        q_data = [
            (0, "Tổng số câu"),
            (3, str(mat_totals["sum_mcq_nb"]) if mat_totals["sum_mcq_nb"] > 0 else ""),
            (4, str(mat_totals["sum_mcq_th"]) if mat_totals["sum_mcq_th"] > 0 else ""),
            (5, str(mat_totals["sum_mcq_vd"]) if mat_totals["sum_mcq_vd"] > 0 else ""),
            (6, str(mat_totals["sum_tf_nb"]) if mat_totals["sum_tf_nb"] > 0 else ""),
            (7, str(mat_totals["sum_tf_th"]) if mat_totals["sum_tf_th"] > 0 else ""),
            (8, str(mat_totals["sum_tf_vd"]) if mat_totals["sum_tf_vd"] > 0 else ""),
            (9, str(mat_totals["sum_es_nb"]) if mat_totals["sum_es_nb"] > 0 else ""),
            (10, str(mat_totals["sum_es_th"]) if mat_totals["sum_es_th"] > 0 else ""),
            (11, str(mat_totals["sum_es_vd"]) if mat_totals["sum_es_vd"] > 0 else ""),
            (12, str(mat_totals["tot_nb_cnt"])),
            (14, str(mat_totals["tot_th_cnt"])),
            (16, str(mat_totals["tot_vd_cnt"])),
            (17, f"{mat_totals['all_q_cnt']} câu")
        ]
        for c_idx, val in q_data:
            c = r_sum_q.cells[c_idx]
            c.text = val
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11)
                p.runs[0].bold = True
            cls.set_cell_shading(c, "EAEAEA")

        # 5. Hàng Tổng số điểm (Tính động 100% từ sum các cột)
        r_sum_p = t_mat.add_row()
        r_sum_p.cells[0].merge(r_sum_p.cells[2])
        r_sum_p.cells[12].merge(r_sum_p.cells[13])
        r_sum_p.cells[14].merge(r_sum_p.cells[15])
        p_data = [
            (0, "Tổng số điểm"),
            (3, fmt_pts_ui(mat_totals["pts_mcq_nb"]) if mat_totals["pts_mcq_nb"] > 0 else ""),
            (4, fmt_pts_ui(mat_totals["pts_mcq_th"]) if mat_totals["pts_mcq_th"] > 0 else ""),
            (5, fmt_pts_ui(mat_totals["pts_mcq_vd"]) if mat_totals["pts_mcq_vd"] > 0 else ""),
            (6, fmt_pts_ui(mat_totals["pts_tf_nb"]) if mat_totals["pts_tf_nb"] > 0 else ""),
            (7, fmt_pts_ui(mat_totals["pts_tf_th"]) if mat_totals["pts_tf_th"] > 0 else ""),
            (8, fmt_pts_ui(mat_totals["pts_tf_vd"]) if mat_totals["pts_tf_vd"] > 0 else ""),
            (9, fmt_pts_ui(mat_totals["pts_es_nb"]) if mat_totals["pts_es_nb"] > 0 else ""),
            (10, fmt_pts_ui(mat_totals["pts_es_th"]) if mat_totals["pts_es_th"] > 0 else ""),
            (11, fmt_pts_ui(mat_totals["pts_es_vd"]) if mat_totals["pts_es_vd"] > 0 else ""),
            (12, fmt_pts_ui(mat_totals["tot_nb_pts"])),
            (14, fmt_pts_ui(mat_totals["tot_th_pts"])),
            (16, fmt_pts_ui(mat_totals["tot_vd_pts"])),
            (17, f"{fmt_pts_ui(mat_totals['all_pts'])} điểm")
        ]
        for c_idx, val in p_data:
            c = r_sum_p.cells[c_idx]
            c.text = val
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11)
                p.runs[0].bold = True
            cls.set_cell_shading(c, "EAEAEA")

        # 6. Hàng Tỉ lệ % (Tính động 100%)
        r_pct = t_mat.add_row()
        r_pct.cells[0].merge(r_pct.cells[2])
        r_pct.cells[3].merge(r_pct.cells[5])
        r_pct.cells[6].merge(r_pct.cells[8])
        r_pct.cells[9].merge(r_pct.cells[11])
        r_pct.cells[12].merge(r_pct.cells[16])
        pct_data = [
            (0, "Tỉ lệ %"),
            (3, fmt_pct_ui(mat_totals["pct_mcq"])),
            (6, fmt_pct_ui(mat_totals["pct_tf"])),
            (9, fmt_pct_ui(mat_totals["pct_es"])),
            (12, f"{fmt_pct_ui(mat_totals['pct_nb'])} - {fmt_pct_ui(mat_totals['pct_th'])} - {fmt_pct_ui(mat_totals['pct_vd'])}"),
            (17, "100%")
        ]
        for c_idx, val in pct_data:
            c = r_pct.cells[c_idx]
            c.text = val
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11.5)
                p.runs[0].bold = True
            cls.set_cell_shading(c, "EAEAEA")

        mat_col_widths_dxa = [
            550,   # 0: TT
            1800,  # 1: Chương/chủ đề
            3260,  # 2: Nội dung/đơn vị kiến thức
            650, 650, 650,  # 3,4,5: TNKQ Nhiều lựa chọn (Biết, Hiểu, Vận dụng)
            650, 650, 650,  # 6,7,8: TNKQ Đúng - Sai (Biết, Hiểu, Vận dụng)
            650, 650, 650,  # 9,10,11: Tự luận (Biết, Hiểu, Vận dụng)
            500, 500,       # 12,13: Tổng Biết (Số câu, Điểm)
            500, 500,       # 14,15: Tổng Hiểu (Số câu, Điểm)
            850,            # 16: Tổng Vận dụng
            1400            # 17: Tỉ lệ % điểm
        ]
        cls._apply_table_grid_and_widths(t_mat, mat_col_widths_dxa)

        # ======================================================================
        # SECTION 2 (LANDSCAPE): BẢNG ĐẶC TẢ ĐỀ KIỂM TRA
        # ======================================================================
        sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec2.orientation = WD_ORIENT.LANDSCAPE
        sec2.page_width = Inches(11.69)
        sec2.page_height = Inches(8.27)
        sec2.top_margin = Inches(0.79)
        sec2.bottom_margin = Inches(0.39)
        sec2.left_margin = Inches(0.39)
        sec2.right_margin = Inches(0.39)

        p_sp2 = doc.add_paragraph()
        p_sp2.paragraph_format.space_before = Pt(4)
        p_sp2.paragraph_format.space_after = Pt(4)
        r_sp2 = p_sp2.add_run(f"II. BẢNG ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KỲ MÔN TIN HỌC, LỚP {grade_num}")
        r_sp2.font.name = "Times New Roman"
        r_sp2.font.size = Pt(13)
        r_sp2.bold = True

        # Bảng Đặc Tả (13 cột chuẩn mực)
        t_spec = doc.add_table(rows=4, cols=13)
        t_spec.style = 'Table Grid'
        t_spec.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_spec.autofit = False

        # 1. Điền text tiêu đề
        t_spec.cell(0, 0).text = "TT"
        t_spec.cell(0, 1).text = "Chương/\nchủ đề"
        t_spec.cell(0, 2).text = "Nội dung/đơn vị kiến thức"
        t_spec.cell(0, 3).text = "Yêu cầu cần đạt"
        t_spec.cell(0, 4).text = "Số câu hỏi ở mức độ đánh giá"

        t_spec.cell(1, 4).text = "TNKQ"
        t_spec.cell(1, 10).text = "Tự luận"

        t_spec.cell(2, 4).text = "Nhiều lựa chọn"
        t_spec.cell(2, 7).text = "Đúng - Sai"

        t_spec.cell(3, 4).text = "Biết"
        t_spec.cell(3, 5).text = "Hiểu"
        t_spec.cell(3, 6).text = "Vận dụng"
        t_spec.cell(3, 7).text = "Biết"
        t_spec.cell(3, 8).text = "Hiểu"
        t_spec.cell(3, 9).text = "Vận dụng"
        t_spec.cell(3, 10).text = "Biết"
        t_spec.cell(3, 11).text = "Hiểu"
        t_spec.cell(3, 12).text = "Vận dụng"

        # 2. Merges chuẩn
        t_spec.cell(0, 0).merge(t_spec.cell(3, 0))
        t_spec.cell(0, 1).merge(t_spec.cell(3, 1))
        t_spec.cell(0, 2).merge(t_spec.cell(3, 2))
        t_spec.cell(0, 3).merge(t_spec.cell(3, 3))
        t_spec.cell(0, 4).merge(t_spec.cell(0, 12))

        t_spec.cell(1, 4).merge(t_spec.cell(1, 9))
        t_spec.cell(1, 10).merge(t_spec.cell(2, 12))

        t_spec.cell(2, 4).merge(t_spec.cell(2, 6))
        t_spec.cell(2, 7).merge(t_spec.cell(2, 9))

        for ri in range(4):
            for ci in range(13):
                c = t_spec.cell(ri, ci)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(11.5)
                    p.runs[0].bold = True
                cls.set_cell_shading(c, "F2F2F2")

        # 3. Đổ dữ liệu đặc tả chi tiết (Tách từng dòng theo mức độ nhận thức, số câu ngang hàng chuẩn 100%)
        cur_row_idx = 4  # Dòng dữ liệu đầu tiên sau 4 dòng header

        for s_row in dossier_data["spec_rows"]:
            lvl_rows = s_row.get("level_rows", [])
            if not lvl_rows:
                lvl_rows = [{
                    "level_name": "Nhận biết",
                    "requirement": s_row.get("requirement", ""),
                    "mcq_nb": s_row.get("mcq_nb", ""),
                    "mcq_th": s_row.get("mcq_th", ""),
                    "mcq_vd": s_row.get("mcq_vd", ""),
                    "tf_nb": s_row.get("tf_nb", ""),
                    "tf_th": s_row.get("tf_th", ""),
                    "tf_vd": s_row.get("tf_vd", ""),
                    "es_nb": s_row.get("es_nb", ""),
                    "es_th": s_row.get("es_th", ""),
                    "es_vd": s_row.get("es_vd", "")
                }]

            k_len = len(lvl_rows)
            start_r = cur_row_idx

            for k_idx, lr in enumerate(lvl_rows):
                r = t_spec.add_row()
                if k_idx == 0:
                    r.cells[0].text = str(s_row.get("tt", ""))
                    r.cells[1].text = str(s_row.get("chapter", ""))
                    r.cells[2].text = str(s_row.get("unit", ""))
                else:
                    r.cells[0].text = ""
                    r.cells[1].text = ""
                    r.cells[2].text = ""

                # Cột 3: YCCĐ của mức nhận thức tương ứng
                r.cells[3].text = lr.get("requirement", "")

                # Cột 4..12: Số câu ở mức độ này (NGANG HÀNG TRỰC TIẾP VỚI DÒNG YCCĐ CỦA MỨC ĐÓ)
                vals_k = [
                    lr.get("mcq_nb", ""), lr.get("mcq_th", ""), lr.get("mcq_vd", ""),
                    lr.get("tf_nb", ""), lr.get("tf_th", ""), lr.get("tf_vd", ""),
                    lr.get("es_nb", ""), lr.get("es_th", ""), lr.get("es_vd", "")
                ]
                for ci_off, v_k in enumerate(vals_k):
                    r.cells[4 + ci_off].text = str(v_k) if v_k else ""

                for ci in range(13):
                    c = r.cells[ci]
                    if c.paragraphs:
                        p = c.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if ci == 3 else (WD_ALIGN_PARAGRAPH.LEFT if ci in [1, 2] else WD_ALIGN_PARAGRAPH.CENTER)
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)
                        for run in p.runs:
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(11.5)
                            if ci >= 4 and run.text.strip():
                                run.bold = True

                cur_row_idx += 1

            # Merge dọc các cột TT, Chương, Nội dung qua các dòng của bài đó
            if k_len > 1:
                t_spec.cell(start_r, 0).merge(t_spec.cell(start_r + k_len - 1, 0))
                t_spec.cell(start_r, 1).merge(t_spec.cell(start_r + k_len - 1, 1))
                t_spec.cell(start_r, 2).merge(t_spec.cell(start_r + k_len - 1, 2))

        # 4. Hàng Tổng số câu Bản đặc tả (Khớp 100% với Ma trận)
        r_spec_q = t_spec.add_row()
        r_spec_q.cells[0].merge(r_spec_q.cells[3])
        r_spec_q.cells[0].text = "Tổng số câu"
        p = r_spec_q.cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(11.5)
            p.runs[0].bold = True
        cls.set_cell_shading(r_spec_q.cells[0], "EAEAEA")

        sq_vals = [
            (4, str(mat_totals["sum_mcq_nb"]) if mat_totals["sum_mcq_nb"] > 0 else ""),
            (5, str(mat_totals["sum_mcq_th"]) if mat_totals["sum_mcq_th"] > 0 else ""),
            (6, str(mat_totals["sum_mcq_vd"]) if mat_totals["sum_mcq_vd"] > 0 else ""),
            (7, str(mat_totals["sum_tf_nb"]) if mat_totals["sum_tf_nb"] > 0 else ""),
            (8, str(mat_totals["sum_tf_th"]) if mat_totals["sum_tf_th"] > 0 else ""),
            (9, str(mat_totals["sum_tf_vd"]) if mat_totals["sum_tf_vd"] > 0 else ""),
            (10, str(mat_totals["sum_es_nb"]) if mat_totals["sum_es_nb"] > 0 else ""),
            (11, str(mat_totals["sum_es_th"]) if mat_totals["sum_es_th"] > 0 else ""),
            (12, str(mat_totals["sum_es_vd"]) if mat_totals["sum_es_vd"] > 0 else "")
        ]
        for ci, v in sq_vals:
            c = r_spec_q.cells[ci]
            c.text = v
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11.5)
                p.runs[0].bold = True
            cls.set_cell_shading(c, "EAEAEA")

        # 5. Hàng Tổng số điểm Bản đặc tả (Khớp 100% với Ma trận)
        r_spec_p = t_spec.add_row()
        r_spec_p.cells[0].merge(r_spec_p.cells[3])
        r_spec_p.cells[4].merge(r_spec_p.cells[6])
        r_spec_p.cells[7].merge(r_spec_p.cells[9])
        r_spec_p.cells[10].merge(r_spec_p.cells[12])
        r_spec_p.cells[0].text = "Tổng số điểm"
        r_spec_p.cells[4].text = fmt_pts_ui(mat_totals["tot_mcq_pts"])
        r_spec_p.cells[7].text = fmt_pts_ui(mat_totals["tot_tf_pts"])
        r_spec_p.cells[10].text = fmt_pts_ui(mat_totals["tot_es_pts"])
        for ci in [0, 4, 7, 10]:
            c = r_spec_p.cells[ci]
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11.5)
                p.runs[0].bold = True
            cls.set_cell_shading(c, "EAEAEA")

        # 6. Hàng Tỉ lệ % Bản đặc tả (Khớp 100% với Ma trận)
        r_spec_pct = t_spec.add_row()
        r_spec_pct.cells[0].merge(r_spec_pct.cells[3])
        r_spec_pct.cells[4].merge(r_spec_pct.cells[6])
        r_spec_pct.cells[7].merge(r_spec_pct.cells[9])
        r_spec_pct.cells[10].merge(r_spec_pct.cells[12])
        r_spec_pct.cells[0].text = "Tỉ lệ %"
        r_spec_pct.cells[4].text = fmt_pct_ui(mat_totals["pct_mcq"])
        r_spec_pct.cells[7].text = fmt_pct_ui(mat_totals["pct_tf"])
        r_spec_pct.cells[10].text = fmt_pct_ui(mat_totals["pct_es"])
        for ci in [0, 4, 7, 10]:
            c = r_spec_pct.cells[ci]
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11.5)
                p.runs[0].bold = True
            cls.set_cell_shading(c, "EAEAEA")

        spec_col_widths_dxa = [
            550,   # 0: TT
            1800,  # 1: Chương/chủ đề
            2760,  # 2: Nội dung/đơn vị kiến thức
            4750,  # 3: Yêu cầu cần đạt
            650, 650, 650,  # 4,5,6: Nhiều lựa chọn (Biết, Hiểu, VD)
            650, 650, 650,  # 7,8,9: Đúng - Sai (Biết, Hiểu, VD)
            650, 650, 650   # 10,11,12: Tự luận (Biết, Hiểu, VD)
        ]
        cls._apply_table_grid_and_widths(t_spec, spec_col_widths_dxa)

        # Chữ ký phê duyệt tổ trưởng
        p_sig_space = doc.add_paragraph()
        p_sig_space.paragraph_format.space_before = Pt(8)
        t_sign = doc.add_table(rows=1, cols=2)
        t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_sign.autofit = False
        cls._cant_split(t_sign.rows[0])
        border_none = {'val': 'none'}
        for row in t_sign.rows:
            for c in row.cells:
                cls.set_cell_border(c, top=border_none, bottom=border_none, left=border_none, right=border_none)
        cls._apply_table_grid_and_widths(t_sign, [7855, 7855])

        spec_date, _ = cls._get_exam_dates(meta)

        p_s1 = t_sign.rows[0].cells[0].paragraphs[0]
        p_s1.text = ""  # Cột trái để trống theo yêu cầu bản đặc tả chỉ có tổ CM ký

        p_s2 = t_sign.rows[0].cells[1].paragraphs[0]
        p_s2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_s2.paragraph_format.space_before = Pt(2)
        p_s2.paragraph_format.space_after = Pt(2)
        p_s2.paragraph_format.line_spacing = 1.15
        p_s2.add_run(f"{spec_date}\n").italic = True
        r_tt = p_s2.add_run("TỔ TRƯỞNG CHUYÊN MÔN\n\n\n\n")
        r_tt.bold = True
        p_s2.add_run(meta["head_name"]).bold = True

        for p in [p_s1, p_s2]:
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        # ======================================================================
        # SECTION 3 (PORTRAIT): ĐỀ KIỂM TRA CHÍNH THỨC, BÀI LÀM & ĐÁP ÁN
        # ======================================================================
        sec3 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec3.orientation = WD_ORIENT.PORTRAIT
        sec3.page_width = Inches(8.27)    # 210 mm (A4 chuẩn)
        sec3.page_height = Inches(11.69)  # 297 mm (A4 chuẩn)
        sec3.top_margin = Inches(0.59)    # 15 mm
        sec3.bottom_margin = Inches(0.59) # 15 mm
        sec3.left_margin = Inches(0.59)   # 15 mm
        sec3.right_margin = Inches(0.59)  # 15 mm -> Chiều ngang in ấn 18.0 cm (10206 dxa)

        # Render Đề thi TH01
        cls._render_exam_paper(doc, meta, dossier_data["th01"])

        # Render Đề thi TH02 (hoán vị đảo câu & đáp án)
        doc.add_page_break()
        cls._render_exam_paper(doc, meta, dossier_data["th02"])

        # Render Hướng dẫn chấm và Biểu điểm chi tiết cho cả 2 mã đề
        doc.add_page_break()
        cls._render_scoring_guide(doc, meta, dossier_data["th01"], dossier_data["th02"])

        # Lưu tài liệu ra đĩa
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path

    @classmethod
    def export_exam_only(cls, dossier_data: Dict[str, Any], output_path: str) -> str:
        """
        Xuất riêng file Word (.docx) chỉ gồm ĐỀ KIỂM TRA CHÍNH THỨC (TH01, TH02) và ĐÁP ÁN BIỂU ĐIỂM,
        khổ dọc Portrait chuẩn A4 (18cm / 10206 dxa), không kèm Ma trận hay Đặc tả ngang,
        dành riêng cho giáo viên in ấn và photo phát trực tiếp cho học sinh làm bài.
        """
        doc = docx.Document()
        cls._ensure_compatibility_settings(doc)
        meta = dossier_data["meta"]

        sec = doc.sections[0]
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width = Inches(8.27)    # 210 mm (A4 chuẩn)
        sec.page_height = Inches(11.69)  # 297 mm (A4 chuẩn)
        sec.top_margin = Inches(0.59)    # 15 mm
        sec.bottom_margin = Inches(0.59) # 15 mm
        sec.left_margin = Inches(0.59)   # 15 mm
        sec.right_margin = Inches(0.59)  # 15 mm

        # Render Đề thi TH01
        cls._render_exam_paper(doc, meta, dossier_data["th01"])

        # Render Đề thi TH02 (hoán vị đảo câu & đáp án)
        doc.add_page_break()
        cls._render_exam_paper(doc, meta, dossier_data["th02"])

        # Render Hướng dẫn chấm và Biểu điểm chi tiết cho cả 2 mã đề
        doc.add_page_break()
        cls._render_scoring_guide(doc, meta, dossier_data["th01"], dossier_data["th02"])

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path

    @classmethod
    def _render_exam_paper(cls, doc, meta: Dict[str, Any], exam_data: Dict[str, Any]):
        """
        Render một đề kiểm tra hoàn chỉnh chuẩn khuôn 2 tờ in 2 mặt rọc phách (4 trang A4):
        - TỜ 1 (Trắc nghiệm):
          + Mặt trước (Trang 1): Header hành chính + Kéo cắt phách + Khung điểm + Câu 1..6.
          + Mặt sau (Trang 2): Ô gạch chéo X (cao đúng 2376 dxa) + Kéo cắt phách + Câu 7..12 + Phần Đúng/Sai.
        - TỜ 2 (Tự luận):
          + Mặt trước (Trang 3): Header hành chính + Kéo cắt phách + Khung điểm + Tự luận + BÀI LÀM.
          + Mặt sau (Trang 4): Ô gạch chéo X (cao đúng 2376 dxa) + Kéo cắt phách + BÀI LÀM tiếp theo.
        Tọa độ đường kéo ở cả 2 mặt giấy trùng khít 100%, rọc phách không phạm vào nội dung!
        """
        code = exam_data["code"]

        # ======================================================================
        # TRANG 1: MẶT TRƯỚC TỜ 1 (TRẮC NGHIỆM PHẦN I - CÂU 1 ĐẾN CÂU 6)
        # ======================================================================
        # 1. Bảng tiêu đề hành chính học sinh & số phách
        cls._render_header_box(doc, meta, code, is_essay=False)

        # 2. Đường kéo cắt phách mặt trước (phân cách giữa phần phách học sinh và bài thi)
        cls._render_cut_line(doc)

        # 3. Khung điểm & Giám khảo & Phách
        cls._render_score_box(doc, code, is_essay=False)

        # 4. Phần I: Trắc nghiệm nhiều lựa chọn
        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(2.0)
        p_p1.paragraph_format.space_after = Pt(1.0)
        p_p1.paragraph_format.keep_with_next = True
        r_p1 = p_p1.add_run("I. Khoanh tròn vào chữ cái đứng trước câu trả lời đúng (3,0 điểm)")
        r_p1.font.name = "Times New Roman"
        r_p1.font.size = Pt(13)
        r_p1.bold = True

        # Tính toán phân bổ dồn câu hỏi tối ưu lên Trang 1
        page1_count = cls._calculate_optimal_page1_mcq_count(exam_data["mcqs"], exam_data.get("tfs", []))

        # Render các câu trắc nghiệm đầu tiên trên Trang 1 (vừa vặn, nguyên vẹn cả câu và đáp án)
        for idx, q in enumerate(exam_data["mcqs"][:page1_count], 1):
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(1.5)
            p_q.paragraph_format.space_after = Pt(0.5)
            p_q.paragraph_format.line_spacing = 1.05
            p_q.paragraph_format.keep_with_next = True

            rq_num = p_q.add_run(f"Câu {idx}. (0,25 điểm) ")
            rq_num.font.name = "Times New Roman"
            rq_num.font.size = Pt(13)
            rq_num.bold = True

            rq_cnt = p_q.add_run(q["content"])
            rq_cnt.font.name = "Times New Roman"
            rq_cnt.font.size = Pt(13)

            # Render 4 phương án căn thẳng hàng tuyệt đối qua Tab Stop
            cls._render_mcq_options(doc, q["options"], is_p2=False)

        # ======================================================================
        # TRANG 2: MẶT SAU TỜ 1 (TRẮC NGHIỆM TIẾP THEO NẾU CÓ + PHẦN II ĐÚNG/SAI)
        # ======================================================================
        doc.add_page_break()

        # Ô chữ nhật gạch chéo X (che phách mặt sau) & đường kéo cắt phách chuẩn tọa độ
        cls._render_x_cut_box(doc)

        # Nếu còn câu trắc nghiệm của Phần I chưa đưa hết lên Trang 1 thì hiển thị tiếp ngay dưới đường kéo
        if page1_count < len(exam_data["mcqs"]):
            for idx, q in enumerate(exam_data["mcqs"][page1_count:], page1_count + 1):
                p_q = doc.add_paragraph()
                p_q.paragraph_format.space_before = Pt(1.0)
                p_q.paragraph_format.space_after = Pt(0.2)
                p_q.paragraph_format.line_spacing = 1.04
                p_q.paragraph_format.keep_with_next = True

                rq_num = p_q.add_run(f"Câu {idx}. (0,25 điểm) ")
                rq_num.font.name = "Times New Roman"
                rq_num.font.size = Pt(13)
                rq_num.bold = True

                rq_cnt = p_q.add_run(q["content"])
                rq_cnt.font.name = "Times New Roman"
                rq_cnt.font.size = Pt(13)

                cls._render_mcq_options(doc, q["options"], is_p2=True)

        # TIẾP TỤC NGAY TRÊN TRANG 2: Phần II: Trắc nghiệm Đúng / Sai (Đánh số tiếp nối từ Phần I)
        tf_start = len(exam_data["mcqs"]) + 1
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(0.5)
        p_p2.paragraph_format.space_after = Pt(0.2)
        p_p2.paragraph_format.keep_with_next = True
        r_p2 = p_p2.add_run("II. Viết chữ Đ(đúng) hoặc S(sai) vào các câu trả lời tương ứng trong các câu sau: (4,0 điểm)")
        r_p2.font.name = "Times New Roman"
        r_p2.font.size = Pt(13)
        r_p2.bold = True

        # BẢNG ĐIỀN ĐÁP ÁN ĐÚNG/SAI CHUẨN 2 BẢNG 6 CỘT ĐỀU NHAU (THEO ĐÚNG MẪU NHÀ TRƯỜNG)
        tf_nums = [str(tf_start + i) for i in range(len(exam_data.get("tfs", [])))]
        c1_no = tf_nums[0] if len(tf_nums) > 0 else str(tf_start)
        c2_no = tf_nums[1] if len(tf_nums) > 1 else str(tf_start + 1)
        c3_no = tf_nums[2] if len(tf_nums) > 2 else str(tf_start + 2)
        c4_no = tf_nums[3] if len(tf_nums) > 3 else str(tf_start + 3)

        grid_headers_6 = ["Câu", "Lệnh hỏi", "Đáp án (Đ/S)", "Câu", "Lệnh hỏi", "Đáp án (Đ/S)"]
        col_w_6 = [Inches(0.90), Inches(1.04), Inches(1.60), Inches(0.90), Inches(1.04), Inches(1.60)]

        # Bảng 1: Cho Câu c1_no và c2_no
        t_grid_1 = doc.add_table(rows=5, cols=6)
        t_grid_1.style = 'Table Grid'
        t_grid_1.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r in t_grid_1.rows:
            cls._cant_split(r)
            cls._set_row_height_exact(r, 340)

        for ci, h in enumerate(grid_headers_6):
            c = t_grid_1.cell(0, ci)
            c.text = h
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(13)
            p.runs[0].bold = True
            cls.set_cell_shading(c, "FFFFFF")

        for ri, lbl in enumerate(["A", "B", "C", "D"], 1):
            r_vals = [c1_no, lbl, "", c2_no, lbl, ""]
            for ci, val in enumerate(r_vals):
                c = t_grid_1.cell(ri, ci)
                c.text = val
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(1.5)
                p.paragraph_format.space_after = Pt(1.5)
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(13)
                    if ci in [0, 1, 3, 4]:
                        p.runs[0].bold = True
        cls._set_col_widths(t_grid_1, col_w_6, total_width_dxa=10206)

        # Khoảng cách giữa 2 bảng
        p_sp_grid = doc.add_paragraph()
        p_sp_grid.paragraph_format.space_before = Pt(1)
        p_sp_grid.paragraph_format.space_after = Pt(1)
        p_sp_grid.paragraph_format.line_spacing = 0.5
        p_sp_grid.paragraph_format.keep_with_next = True

        # Bảng 2: Cho Câu c3_no và c4_no
        t_grid_2 = doc.add_table(rows=5, cols=6)
        t_grid_2.style = 'Table Grid'
        t_grid_2.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r in t_grid_2.rows:
            cls._cant_split(r)
            cls._set_row_height_exact(r, 340)

        for ci, h in enumerate(grid_headers_6):
            c = t_grid_2.cell(0, ci)
            c.text = h
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(13)
            p.runs[0].bold = True
            cls.set_cell_shading(c, "FFFFFF")

        for ri, lbl in enumerate(["A", "B", "C", "D"], 1):
            r_vals = [c3_no, lbl, "", c4_no, lbl, ""]
            for ci, val in enumerate(r_vals):
                c = t_grid_2.cell(ri, ci)
                c.text = val
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(1.5)
                p.paragraph_format.space_after = Pt(1.5)
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(13)
                    if ci in [0, 1, 3, 4]:
                        p.runs[0].bold = True
        cls._set_col_widths(t_grid_2, col_w_6, total_width_dxa=10206)

        # Nội dung 4 câu hỏi Đúng / Sai (Đánh số tiếp nối + ghi rõ số điểm 1,0 điểm + cỡ chữ 13)
        for idx, tf_item in enumerate(exam_data["tfs"], tf_start):
            p_tf_q = doc.add_paragraph()
            p_tf_q.paragraph_format.space_before = Pt(1.2)
            p_tf_q.paragraph_format.space_after = Pt(0.2)
            p_tf_q.paragraph_format.line_spacing = 1.05
            p_tf_q.paragraph_format.keep_with_next = True

            top_txt = tf_item["topic"][:65]
            r_tf_num = p_tf_q.add_run(f"Câu {idx}. (1,0 điểm) “{top_txt}”")
            r_tf_num.font.name = "Times New Roman"
            r_tf_num.font.size = Pt(13)
            r_tf_num.bold = True

            for s_idx, stmt in enumerate(tf_item["statements"]):
                p_st = doc.add_paragraph()
                p_st.paragraph_format.space_before = Pt(0)
                p_st.paragraph_format.space_after = Pt(0.2)
                p_st.paragraph_format.left_indent = Inches(0.20)
                p_st.paragraph_format.line_spacing = 1.05
                if s_idx < 3:
                    p_st.paragraph_format.keep_with_next = True

                lbl_val = stmt["label"]
                r_st_lbl = p_st.add_run(f"{lbl_val}) ")
                r_st_lbl.font.name = "Times New Roman"
                r_st_lbl.font.size = Pt(13)
                r_st_lbl.bold = True

                r_st = p_st.add_run(stmt["content"])
                r_st.font.name = "Times New Roman"
                r_st.font.size = Pt(13)

        # ======================================================================
        # TRANG 3: MẶT TRƯỚC TỜ 2 (TỜ LÀM BÀI TỰ LUẬN)
        # ======================================================================
        doc.add_page_break()

        # Header Tờ Tự luận
        cls._render_header_box(doc, meta, code, is_essay=True)

        # Đường kéo cắt phách mặt trước của Tờ Tự luận
        cls._render_cut_line(doc)

        # Khung Điểm & Giám khảo Tự luận
        cls._render_score_box(doc, code, is_essay=True)

                # Nội dung các câu hỏi tự luận (Đánh số tiếp nối sau Phần II + ghi rõ điểm + cỡ chữ 13)
        essay_start = (len(exam_data["mcqs"]) + len(exam_data.get("tfs", [])) + 1)
        for idx, es_q in enumerate(exam_data["essays"], essay_start):
            p_es = doc.add_paragraph()
            p_es.paragraph_format.space_before = Pt(2.5)
            p_es.paragraph_format.space_after = Pt(1.0)
            p_es.paragraph_format.line_spacing = 1.10
            p_es.paragraph_format.keep_with_next = True

            r_es_n = p_es.add_run(f"Câu {idx}. ({es_q['score']} điểm) ")
            r_es_n.font.name = "Times New Roman"
            r_es_n.font.size = Pt(13)
            r_es_n.bold = True

            r_es_c = p_es.add_run(es_q["content"])
            r_es_c.font.name = "Times New Roman"
            r_es_c.font.size = Pt(13)

        # Kết thúc đề thi
        p_end = doc.add_paragraph()
        p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_end.paragraph_format.space_before = Pt(2)
        p_end.paragraph_format.space_after = Pt(2)
        p_end.paragraph_format.keep_with_next = True
        r_end = p_end.add_run("------HẾT------")
        r_end.bold = True
        r_end.font.name = "Times New Roman"
        r_end.font.size = Pt(11)

        # Phần BÀI LÀM của học sinh
        p_bl1 = doc.add_paragraph()
        p_bl1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_bl1.paragraph_format.space_before = Pt(4)
        p_bl1.paragraph_format.space_after = Pt(3)
        p_bl1.paragraph_format.keep_with_next = True
        r_bl1 = p_bl1.add_run("BÀI LÀM")
        r_bl1.bold = True
        r_bl1.font.name = "Times New Roman"
        r_bl1.font.size = Pt(13)

        # Các dòng chấm full 100% lấp đầy toàn bộ nửa dưới Trang 3 đến sát mép đáy
        cls._render_dot_lines(doc, count=17)

        # ======================================================================
        # TRANG 4: MẶT SAU TỜ 2 (BÀI LÀM TỰ LUẬN TIẾP THEO)
        # ======================================================================
        doc.add_page_break()

        # Ô chữ nhật gạch chéo X (che phách mặt sau) & đường kéo cắt phách
        cls._render_x_cut_box(doc)

        p_bl2 = doc.add_paragraph()
        p_bl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_bl2.paragraph_format.space_before = Pt(4)
        p_bl2.paragraph_format.space_after = Pt(3)
        p_bl2.paragraph_format.keep_with_next = True
        r_bl2 = p_bl2.add_run("BÀI LÀM (tiếp theo)")
        r_bl2.bold = True
        r_bl2.font.name = "Times New Roman"
        r_bl2.font.size = Pt(13)

        # 31 dòng chấm full 100% phủ kín toàn bộ trang 4 đến sát mép đáy lề dưới giấy A4
        cls._render_dot_lines(doc, count=28)

    @classmethod
    def _render_scoring_guide(cls, doc, meta: Dict[str, Any], th01: Dict[str, Any], th02: Dict[str, Any]):
        """Render Hướng dẫn chấm và Biểu điểm chi tiết chuẩn mẫu file Đề kiểm tra và công vụ."""
        grade_num = meta['grade'].upper().replace('LỚP ', '').strip()

        for idx_exam, exam in enumerate([th01, th02]):
            if idx_exam > 0:
                doc.add_page_break()

            code = exam["code"]

            # Header Hướng dẫn chấm (3 rows x 2 cols - KHÔNG DÙNG VIỀN LƯỚI ĐEN)
            t_hd_top = doc.add_table(rows=3, cols=2)
            t_hd_top.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_hd_top.autofit = False
            for r in t_hd_top.rows:
                cls._cant_split(r)

            # Xóa viền đen cho header hướng dẫn chấm
            border_none = {'val': 'none'}
            for row in t_hd_top.rows:
                for c in row.cells:
                    cls.set_cell_border(c, top=border_none, bottom=border_none, left=border_none, right=border_none)

            t_hd_top.cell(0, 0).text = f"{meta['governing_body'].upper()}\n{meta['school_name'].upper()}"
            t_hd_top.cell(0, 1).text = f"KIỂM TRA GIỮA HỌC KỲ I\nNĂM HỌC {meta['school_year']}"
            t_hd_top.cell(1, 0).text = ""
            t_hd_top.cell(1, 1).text = f"MÔN: TIN HỌC {grade_num}"
            t_hd_top.cell(2, 0).text = f"MÃ ĐỀ: {code}"
            t_hd_top.cell(2, 1).text = ""

            for row in t_hd_top.rows:
                for c in row.cells:
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(1.5)
                    p.paragraph_format.space_after = Pt(1.5)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(11)

            t_hd_top.cell(0, 0).paragraphs[0].runs[0].bold = True
            t_hd_top.cell(0, 1).paragraphs[0].runs[0].bold = True
            t_hd_top.cell(1, 1).paragraphs[0].runs[0].bold = True
            t_hd_top.cell(2, 0).paragraphs[0].runs[0].bold = True
            cls._set_col_widths(t_hd_top, [Inches(3.40), Inches(3.69)], total_width_dxa=10206)

            # Tiêu đề chính
            p_hd_title = doc.add_paragraph()
            p_hd_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_hd_title.paragraph_format.space_before = Pt(6)
            p_hd_title.paragraph_format.space_after = Pt(4)
            p_hd_title.paragraph_format.keep_with_next = True
            r_hdt = p_hd_title.add_run("HƯỚNG DẪN CHẤM VÀ BIỂU ĐIỂM CHI TIẾT")
            r_hdt.bold = True
            r_hdt.font.name = "Times New Roman"
            r_hdt.font.size = Pt(13.5)

            # A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)
            p_sec_a = doc.add_paragraph()
            p_sec_a.paragraph_format.space_before = Pt(4)
            p_sec_a.paragraph_format.space_after = Pt(2)
            p_sec_a.paragraph_format.keep_with_next = True
            r_sa = p_sec_a.add_run("A. Phần trắc nghiệm: (7 điểm)")
            r_sa.bold = True
            r_sa.font.name = "Times New Roman"
            r_sa.font.size = Pt(13)

            # I. Dạng nhiều lựa chọn
            p_sub_1 = doc.add_paragraph()
            p_sub_1.paragraph_format.space_before = Pt(2)
            p_sub_1.paragraph_format.space_after = Pt(2)
            p_sub_1.paragraph_format.keep_with_next = True
            r_s1 = p_sub_1.add_run("I. Dạng nhiều lựa chọn: Mỗi câu đúng ghi 0,25 điểm")
            r_s1.bold = True
            r_s1.font.name = "Times New Roman"
            r_s1.font.size = Pt(12)

            # Bảng đáp án MCQ (Table 21: 2 rows x 13 cols)
            t_ans1 = doc.add_table(rows=2, cols=len(exam["mcqs"]) + 1)
            t_ans1.style = 'Table Grid'
            t_ans1.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r in t_ans1.rows:
                cls._cant_split(r)

            t_ans1.cell(0, 0).text = "CÂU"
            t_ans1.cell(1, 0).text = "ĐÁP ÁN"
            cls.set_cell_shading(t_ans1.cell(0, 0), "EAEAEA")
            cls.set_cell_shading(t_ans1.cell(1, 0), "EAEAEA")

            for idx, q in enumerate(exam["mcqs"], 1):
                c0 = t_ans1.cell(0, idx)
                c1 = t_ans1.cell(1, idx)
                c0.text = str(idx)
                c1.text = q.get("correct", "A")

            for row in t_ans1.rows:
                for c in row.cells:
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(1.5)
                    p.paragraph_format.space_after = Pt(1.5)
                    if p.runs:
                        p.runs[0].font.name = "Times New Roman"
                        p.runs[0].font.size = Pt(12)
                        p.runs[0].bold = True
            for r in t_ans1.rows: cls._set_row_height_exact(r, 380)
            cls._set_col_widths(t_ans1, [Inches(0.97)] + [Inches(0.51)] * len(exam["mcqs"]), total_width_dxa=10206)

            # II. Dạng đúng/sai
            p_sub_2 = doc.add_paragraph()
            p_sub_2.paragraph_format.space_before = Pt(6)
            p_sub_2.paragraph_format.space_after = Pt(2)
            p_sub_2.paragraph_format.keep_with_next = True
            r_s2 = p_sub_2.add_run("II. Dạng đúng/sai: Mỗi câu đúng ghi 0,25 điểm")
            r_s2.bold = True
            r_s2.font.name = "Times New Roman"
            r_s2.font.size = Pt(12)

            tf_list = exam["tfs"]
            c1_stmts = tf_list[0]["statements"] if len(tf_list) > 0 else []
            c2_stmts = tf_list[1]["statements"] if len(tf_list) > 1 else []
            c3_stmts = tf_list[2]["statements"] if len(tf_list) > 2 else []
            c4_stmts = tf_list[3]["statements"] if len(tf_list) > 3 else []

            # Đánh số tiếp nối theo Phần I
            sg_tf_start = len(exam["mcqs"]) + 1
            c1_num = str(sg_tf_start)
            c2_num = str(sg_tf_start + 1)
            c3_num = str(sg_tf_start + 2)
            c4_num = str(sg_tf_start + 3)

            grid_headers = ["Câu", "Lệnh hỏi", "Đáp án (Đ/S)", "Câu", "Lệnh hỏi", "Đáp án (Đ/S)"]

            # Bảng đáp án Câu tiếp nối (Table 22)
            t_ans2_1 = doc.add_table(rows=5, cols=6)
            t_ans2_1.style = 'Table Grid'
            t_ans2_1.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r in t_ans2_1.rows:
                cls._cant_split(r)

            for ci, h in enumerate(grid_headers):
                c = t_ans2_1.cell(0, ci)
                c.text = h
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11)
                p.runs[0].bold = True
                cls.set_cell_shading(c, "FFFFFF")

            for ri, lbl in enumerate(["A", "B", "C", "D"], 1):
                ans1 = c1_stmts[ri-1]["correct"] if len(c1_stmts) >= ri else "Đ"
                ans2 = c2_stmts[ri-1]["correct"] if len(c2_stmts) >= ri else "S"
                row_vals = [c1_num, lbl, ans1, c2_num, lbl, ans2]
                for ci, val in enumerate(row_vals):
                    c = t_ans2_1.cell(ri, ci)
                    c.text = val
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(1.5)
                    p.paragraph_format.space_after = Pt(1.5)
                    if p.runs:
                        p.runs[0].font.name = "Times New Roman"
                        p.runs[0].font.size = Pt(10)
                        if ci in [2, 5]: p.runs[0].bold = True
            for r in t_ans2_1.rows: cls._set_row_height_exact(r, 340)
            cls._set_col_widths(t_ans2_1, [Inches(0.90), Inches(1.04), Inches(1.60), Inches(0.90), Inches(1.04), Inches(1.60)], total_width_dxa=10206)

            doc.add_paragraph().paragraph_format.space_before = Pt(1)

            # Bảng đáp án Câu tiếp nối (Table 23)
            t_ans2_2 = doc.add_table(rows=5, cols=6)
            t_ans2_2.style = 'Table Grid'
            t_ans2_2.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r in t_ans2_2.rows:
                cls._cant_split(r)

            for ci, h in enumerate(grid_headers):
                c = t_ans2_2.cell(0, ci)
                c.text = h
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11)
                p.runs[0].bold = True
                cls.set_cell_shading(c, "FFFFFF")

            for ri, lbl in enumerate(["A", "B", "C", "D"], 1):
                ans3 = c3_stmts[ri-1]["correct"] if len(c3_stmts) >= ri else "Đ"
                ans4 = c4_stmts[ri-1]["correct"] if len(c4_stmts) >= ri else "Đ"
                row_vals = [c3_num, lbl, ans3, c4_num, lbl, ans4]
                for ci, val in enumerate(row_vals):
                    c = t_ans2_2.cell(ri, ci)
                    c.text = val
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(1.5)
                    p.paragraph_format.space_after = Pt(1.5)
                    if p.runs:
                        p.runs[0].font.name = "Times New Roman"
                        p.runs[0].font.size = Pt(10)
                        if ci in [2, 5]: p.runs[0].bold = True
            for r in t_ans2_2.rows: cls._set_row_height_exact(r, 340)
            cls._set_col_widths(t_ans2_2, [Inches(0.90), Inches(1.04), Inches(1.60), Inches(0.90), Inches(1.04), Inches(1.60)], total_width_dxa=10206)

            # B. PHẦN TỰ LUẬN (3 ĐIỂM)
            p_sec_b = doc.add_paragraph()
            p_sec_b.paragraph_format.space_before = Pt(6)
            p_sec_b.paragraph_format.space_after = Pt(2)
            p_sec_b.paragraph_format.keep_with_next = True
            r_sb = p_sec_b.add_run("B. Phần tự luận: (3 điểm)")
            r_sb.bold = True
            r_sb.font.name = "Times New Roman"
            r_sb.font.size = Pt(13)

            # Bảng thang điểm tự luận (Table 24: 4 rows x 3 cols)
            t_ans3 = doc.add_table(rows=1, cols=3)
            t_ans3.style = 'Table Grid'
            t_ans3.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_ans3.rows[0].cells[0].width = Inches(1.2)
            t_ans3.rows[0].cells[1].width = Inches(4.9)
            t_ans3.rows[0].cells[2].width = Inches(0.99)

            t_ans3.cell(0, 0).text = "Câu"
            t_ans3.cell(0, 1).text = "Nội dung"
            t_ans3.cell(0, 2).text = "Điểm"
            for c in t_ans3.rows[0].cells:
                cls.set_cell_shading(c, "EAEAEA")
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11)
                p.runs[0].bold = True

            sg_es_start = len(exam["mcqs"]) + len(exam.get("tfs", [])) + 1
            for idx, es_q in enumerate(exam["essays"], sg_es_start):
                r = t_ans3.add_row()
                cls._cant_split(r)
                r.cells[0].text = f"{idx} ({es_q['score']}đ)"
                p0 = r.cells[0].paragraphs[0]
                p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p0.runs[0].font.name = "Times New Roman"
                p0.runs[0].font.size = Pt(11)
                p0.runs[0].bold = True

                # Tách từng ý chấm thành các Paragraph riêng biệt, triệt tiêu lỗi dãn dòng dạt 2 bên
                cell_c = r.cells[1]
                cell_s = r.cells[2]

                rubrics = es_q.get("rubric", [])
                for rub_idx, rubric in enumerate(rubrics):
                    p_c = cell_c.paragraphs[0] if rub_idx == 0 else cell_c.add_paragraph()
                    p_c.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p_c.paragraph_format.space_before = Pt(1.5)
                    p_c.paragraph_format.space_after = Pt(1.5)
                    p_c.paragraph_format.line_spacing = 1.15
                    r_c = p_c.add_run(f"- {rubric['criterion']}")
                    r_c.font.name = "Times New Roman"
                    r_c.font.size = Pt(11.5)

                    p_s = cell_s.paragraphs[0] if rub_idx == 0 else cell_s.add_paragraph()
                    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_s.paragraph_format.space_before = Pt(1.5)
                    p_s.paragraph_format.space_after = Pt(1.5)
                    p_s.paragraph_format.line_spacing = 1.15
                    r_s = p_s.add_run(f"{rubric['score']}đ")
                    r_s.font.name = "Times New Roman"
                    r_s.font.size = Pt(11.5)
            cls._set_col_widths(t_ans3, [Inches(0.90), Inches(5.21), Inches(0.97)], total_width_dxa=10206)

            # Lưu ý và Chữ ký
            p_note_ch = doc.add_paragraph()
            p_note_ch.paragraph_format.space_before = Pt(8)
            p_note_ch.paragraph_format.space_after = Pt(4)
            p_note_ch.paragraph_format.keep_with_next = True
            r_ch = p_note_ch.add_run("Lưu ý: Học sinh trình bày cách giải khác, đúng vẫn chấm theo thang điểm trên.")
            r_ch.font.name = "Times New Roman"
            r_ch.font.size = Pt(11)
            r_ch.italic = True

            # Chữ ký Người ra đề, đáp án - Sử dụng bảng 2 cột căn giữa đồng trục 100% để không bao giờ bị lệch
            _, guide_date = cls._get_exam_dates(meta)
            t_sign_hd = doc.add_table(rows=1, cols=2)
            t_sign_hd.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_sign_hd.autofit = False
            cls._cant_split(t_sign_hd.rows[0])
            for row in t_sign_hd.rows:
                for c in row.cells:
                    cls.set_cell_border(c, top=border_none, bottom=border_none, left=border_none, right=border_none)
            cls._set_col_widths(t_sign_hd, [Inches(3.30), Inches(3.79)], total_width_dxa=10206)

            p_sign = t_sign_hd.rows[0].cells[1].paragraphs[0]
            p_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_sign.paragraph_format.space_before = Pt(6)
            p_sign.paragraph_format.space_after = Pt(2)
            p_sign.paragraph_format.line_spacing = 1.15
            p_sign.paragraph_format.keep_with_next = True

            r_loc = p_sign.add_run(f"{guide_date}\n")
            r_loc.italic = True
            r_loc.font.name = "Times New Roman"
            r_loc.font.size = Pt(11)

            r_pos = p_sign.add_run("Người ra đề, đáp án\n\n\n\n")
            r_pos.bold = True
            r_pos.font.name = "Times New Roman"
            r_pos.font.size = Pt(11)

            r_name = p_sign.add_run(meta["teacher_name"])
            r_name.bold = True
            r_name.font.name = "Times New Roman"
            r_name.font.size = Pt(11)
