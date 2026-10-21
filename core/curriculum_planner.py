"""
core/curriculum_planner.py
Quản lý Kế hoạch Giáo dục (KHGD) và Phân phối Chương trình (PPCT) môn Tin học THCS GDPT 2018.
Hỗ trợ:
1. Ngân hàng KHGD 35 tuần chuẩn GDPT 2018 (SGK Kết nối tri thức) đầy đủ cả 4 khối: Lớp 6, 7, 8, 9.
2. Tự động bóc tách file Word 97-2003 (.doc) lẫn OpenXML (.docx), Excel (.xlsx, .xls), PDF (.pdf), CSV, JSON, TXT.
3. Nhận diện và phân tách gộp cả 4 khối lớp (Tin học 6, 7, 8, 9) trong cùng một file Kế hoạch tổ chuyên môn.
4. Trích xuất chính xác: Tuần, Tiết, Tên bài học, Yêu cầu cần đạt (YCCĐ), Thời điểm kiểm tra định kỳ (GK1, CK1, GK2, CK2).
5. Lọc và khóa chặt chẽ phạm vi bài học & YCCĐ đã dạy theo tuần kiểm tra, ngăn chặn tuyệt đối việc ra nhầm bài chưa học.
"""

import os
import io
import re
import json
import logging
import tempfile
import shutil
from typing import List, Dict, Any, Optional, Tuple, Union
import pandas as pd

logger = logging.getLogger(__name__)

# ======================================================================================
# 1. NGÂN HÀNG KHGD 35 TUẦN CHUẨN GDPT 2018 (SGK KẾT NỐI TRI THỨC) CHO CẢ 4 KHỐI LỚP
# ======================================================================================

STANDARD_KHGD_DATA: Dict[str, List[Dict[str, Any]]] = {
    # ------------------------------- LỚP 6 -------------------------------
    "Lớp 6": [
        # HỌC KỲ I (Tuần 1 - 18)
        {"week": 1, "lesson_index": 1, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Thông tin và dữ liệu", "content": "Khái niệm thông tin, dữ liệu, vật mang tin; tầm quan trọng của thông tin trong đời sống.", "assessment": ""},
        {"week": 2, "lesson_index": 2, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Thông tin và dữ liệu (tiếp)", "content": "Các bước cơ bản trong quá trình xử lí thông tin của con người và máy tính.", "assessment": ""},
        {"week": 3, "lesson_index": 3, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 2. Xử lí thông tin trong máy tính", "content": "Biểu diễn thông tin trong máy tính bằng dãy bit 0 và 1, các đơn vị đo dung lượng dữ liệu.", "assessment": ""},
        {"week": 4, "lesson_index": 4, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 3. Thông tin trong máy tính", "content": "Đơn vị đo thông tin (Byte, KB, MB, GB, TB) và khả năng lưu trữ của thiết bị nhớ.", "assessment": ""},
        {"week": 5, "lesson_index": 5, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 4. Mạng máy tính", "content": "Khái niệm mạng máy tính, các thành phần chính (thiết bị đầu cuối, thiết bị kết nối) và lợi ích của mạng.", "assessment": ""},
        {"week": 6, "lesson_index": 6, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 4. Mạng máy tính (tiếp)", "content": "Mạng có dây và mạng không dây, các thiết bị kết nối mạng thông dụng (Switch, Modem, Access Point).", "assessment": ""},
        {"week": 7, "lesson_index": 7, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 5. Internet", "content": "Khái niệm Internet, đặc điểm chính và các dịch vụ cơ bản thông dụng trên Internet.", "assessment": ""},
        {"week": 8, "lesson_index": 8, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 6. Mạng thông tin toàn cầu", "content": "World Wide Web (WWW), website, trang web, địa chỉ trang web (URL) và trình duyệt web.", "assessment": ""},
        {"week": 9, "lesson_index": 9, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Ôn tập & Kiểm tra giữa Học kỳ I", "content": "Đánh giá YCCĐ Chủ đề 1 và Chủ đề 2 (từ Bài 1 đến Bài 6).", "assessment": "KIỂM TRA GIỮA HỌC KỲ I"},
        {"week": 10, "lesson_index": 10, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 7. Tìm kiếm thông tin trên Internet", "content": "Máy tìm kiếm, cách sử dụng từ khóa tìm kiếm và các bước tra cứu thông tin hiệu quả.", "assessment": ""},
        {"week": 11, "lesson_index": 11, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 8. Thư điện tử", "content": "Khái niệm thư điện tử (email), ưu nhược điểm so với thư truyền thống, cấu trúc địa chỉ email.", "assessment": ""},
        {"week": 12, "lesson_index": 12, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Mạng máy tính và Internet", "unit": "Bài 8. Thư điện tử (tiếp)", "content": "Thực hành đăng nhập, soạn thảo, gửi, nhận thư và đính kèm tệp tin an toàn.", "assessment": ""},
        {"week": 13, "lesson_index": 13, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin", "unit": "Bài 9. An toàn thông tin trên Internet", "content": "Bảo vệ thông tin cá nhân, đặt mật khẩu mạnh và nhận biết các nguy cơ rủi ro trên không gian số.", "assessment": ""},
        {"week": 14, "lesson_index": 14, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 10. Sơ đồ tư duy", "content": "Khái niệm sơ đồ tư duy, các thành phần chính và quy tắc biểu diễn thông tin trực quan.", "assessment": ""},
        {"week": 15, "lesson_index": 15, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 11. Định dạng văn bản", "content": "Định dạng kí tự (phông, cỡ, kiểu, màu) và định dạng đoạn văn bản trong phần mềm soạn thảo.", "assessment": ""},
        {"week": 16, "lesson_index": 16, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 12. Trình bày thông tin ở dạng bảng", "content": "Tạo bảng biểu, chèn/xóa hàng cột, căn chỉnh văn bản và số liệu trong ô bảng.", "assessment": ""},
        {"week": 17, "lesson_index": 17, "semester": "Học kỳ I", "chapter": "Ôn tập cuối kỳ", "unit": "Ôn tập học kỳ I", "content": "Hệ thống hóa toàn bộ kiến thức và kỹ năng thực hành Học kỳ I.", "assessment": ""},
        {"week": 18, "lesson_index": 18, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ I", "content": "Đánh giá tổng hợp YCCĐ toàn bộ Học kỳ I môn Tin học 6.", "assessment": "KIỂM TRA CUỐI HỌC KỲ I"},
        # HỌC KỲ II (Tuần 19 - 35)
        {"week": 19, "lesson_index": 19, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Ứng dụng tin học", "unit": "Bài 13. Thực hành: Tìm kiếm và thay thế", "content": "Sử dụng công cụ Find and Replace để tìm kiếm và sửa đổi văn bản nhanh chóng.", "assessment": ""},
        {"week": 20, "lesson_index": 20, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Ứng dụng tin học", "unit": "Bài 14. Hoàn thiện sổ lưu niệm", "content": "Dự án nhỏ tổng hợp kỹ năng soạn thảo văn bản, chèn ảnh, tạo bảng và căn chỉnh trang in.", "assessment": ""},
        {"week": 21, "lesson_index": 21, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Thuật toán", "content": "Khái niệm thuật toán, mô tả thuật toán bằng phương pháp liệt kê các bước và vẽ sơ đồ khối.", "assessment": ""},
        {"week": 22, "lesson_index": 22, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Thuật toán (tiếp)", "content": "Các cấu trúc điều khiển cơ bản trong thuật toán: Cấu trúc tuần tự, cấu trúc rẽ nhánh và lặp.", "assessment": ""},
        {"week": 23, "lesson_index": 23, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 16. Các cấu trúc điều khiển", "content": "Biểu diễn thuật toán có rẽ nhánh (Nếu... thì) và vòng lặp trong các bài toán thực tiễn.", "assessment": ""},
        {"week": 24, "lesson_index": 24, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 17. Chương trình máy tính", "content": "Làm quen môi trường lập trình trực quan Scratch, các khối lệnh chuyển động và hội thoại.", "assessment": ""},
        {"week": 25, "lesson_index": 25, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 17. Chương trình máy tính (tiếp)", "content": "Tạo kịch bản cho nhân vật, điều khiển nhân vật theo sự kiện bấm chuột và phím.", "assessment": ""},
        {"week": 26, "lesson_index": 26, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 18. Thực hành: Làm việc với nhân vật", "content": "Chỉnh sửa trang phục, tạo âm thanh và thay đổi phông nền sân khấu trong Scratch.", "assessment": ""},
        {"week": 27, "lesson_index": 27, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra giữa Học kỳ II", "content": "Đánh giá YCCĐ từ Tuần 19 đến Tuần 26 (Thuật toán và lập trình trực quan Scratch cơ bản).", "assessment": "KIỂM TRA GIỮA HỌC KỲ II"},
        {"week": 28, "lesson_index": 28, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 19. Cấu trúc lặp trong chương trình", "content": "Sử dụng khối lệnh lặp liên tục (forever) và lặp với số lần biết trước trong Scratch.", "assessment": ""},
        {"week": 29, "lesson_index": 29, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 20. Cấu trúc rẽ nhánh trong chương trình", "content": "Khối lệnh nếu... thì (if... then) và nếu... thì... không thì trong kịch bản chương trình.", "assessment": ""},
        {"week": 30, "lesson_index": 30, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 21. Thực hành: Tạo trò chơi đơn giản", "content": "Xây dựng trò chơi nhỏ hứng táo hoặc vượt chướng ngại vật có tính điểm số.", "assessment": ""},
        {"week": 31, "lesson_index": 31, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 22. Dự án học tập cuối năm", "content": "Thực hiện dự án lập trình sản phẩm mô phỏng hoặc trò chơi giáo dục theo nhóm.", "assessment": ""},
        {"week": 32, "lesson_index": 32, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 22. Dự án học tập cuối năm (tiếp)", "content": "Thuyết trình, đóng gói và chia sẻ sản phẩm số của nhóm.", "assessment": ""},
        {"week": 33, "lesson_index": 33, "semester": "Học kỳ II", "chapter": "Ôn tập cuối năm", "unit": "Ôn tập học kỳ II", "content": "Tổng kết kiến thức thuật toán và kỹ năng lập trình giải quyết vấn đề.", "assessment": ""},
        {"week": 34, "lesson_index": 34, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ II", "content": "Đánh giá toàn diện kiến thức và năng lực môn Tin học 6 cả năm.", "assessment": "KIỂM TRA CUỐI HỌC KỲ II"},
        {"week": 35, "lesson_index": 35, "semester": "Học kỳ II", "chapter": "Tổng kết", "unit": "Đánh giá & Trả bài kiểm tra", "content": "Nhận xét, đánh giá quá trình học tập và công bố kết quả năm học.", "assessment": ""}
    ],

    # ------------------------------- LỚP 7 -------------------------------
    "Lớp 7": [
        # HỌC KỲ I (Tuần 1 - 18)
        {"week": 1, "lesson_index": 1, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Thiết bị vào - ra", "content": "Phân biệt thiết bị vào (chuột, bàn phím, webcam) và thiết bị ra (màn hình, loa, máy in). Khái niệm thiết bị số thông minh.", "assessment": ""},
        {"week": 2, "lesson_index": 2, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Thiết bị vào - ra (tiếp)", "content": "Các cổng kết nối thông dụng (USB, HDMI, Jack 3.5mm) và quy tắc sử dụng thiết bị an toàn.", "assessment": ""},
        {"week": 3, "lesson_index": 3, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 2. Phần mềm máy tính", "content": "Phân biệt phần mềm hệ thống (Hệ điều hành) và phần mềm ứng dụng, mối quan hệ giữa phần cứng và phần mềm.", "assessment": ""},
        {"week": 4, "lesson_index": 4, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 3. Quản lí dữ liệu trong máy tính", "content": "Tổ chức tệp và thư mục dạng cây phân cấp, đường dẫn, đuôi mở rộng của tệp (.docx, .xlsx, .pdf, .mp4).", "assessment": ""},
        {"week": 5, "lesson_index": 5, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 4. Phân loại tệp và bảo vệ dữ liệu", "content": "Biện pháp sao lưu dự phòng dữ liệu (Backup), đặt mật khẩu và phòng tránh virus máy tính.", "assessment": ""},
        {"week": 6, "lesson_index": 6, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin", "unit": "Bài 5. Ứng xử trên mạng", "content": "Văn hóa giao tiếp trên mạng xã hội, phòng tránh bắt nạt trực tuyến và tôn trọng quyền riêng tư.", "assessment": ""},
        {"week": 7, "lesson_index": 7, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 6. Làm quen với phần mềm bảng tính", "content": "Giao diện Excel, khái niệm bảng tính, hàng, cột, ô (cell), khối ô (range) và thanh công thức.", "assessment": ""},
        {"week": 8, "lesson_index": 8, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 6. Làm quen với phần mềm bảng tính (tiếp)", "content": "Nhập dữ liệu số, văn bản, thời gian; căn chỉnh và điều chỉnh độ rộng hàng cột.", "assessment": ""},
        {"week": 9, "lesson_index": 9, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Ôn tập & Kiểm tra giữa Học kỳ I", "content": "Đánh giá YCCĐ Chủ đề 1 đến Chủ đề 4 (từ Bài 1 đến Bài 6).", "assessment": "KIỂM TRA GIỮA HỌC KỲ I"},
        {"week": 10, "lesson_index": 10, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 7. Tính toán tự động trên bảng tính", "content": "Viết công thức tính toán có sử dụng dấu bằng (=), phép cộng, trừ, nhân, chia và địa chỉ ô.", "assessment": ""},
        {"week": 11, "lesson_index": 11, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 8. Sử dụng các hàm tính toán cơ bản", "content": "Hàm tính tổng SUM, tính trung bình cộng AVERAGE, tìm giá trị lớn nhất MAX, nhỏ nhất MIN.", "assessment": ""},
        {"week": 12, "lesson_index": 12, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 8. Sử dụng các hàm tính toán cơ bản (tiếp)", "content": "Áp dụng các hàm SUM, AVERAGE, MAX, MIN để tính điểm tổng kết và thống kê số liệu thực tế.", "assessment": ""},
        {"week": 13, "lesson_index": 13, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 9. Định dạng trang tính", "content": "Kẻ đường viền bảng tính, tô màu nền, định dạng số thập phân, phần trăm và căn lề ô tính.", "assessment": ""},
        {"week": 14, "lesson_index": 14, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 9. Định dạng trang tính (tiếp)", "content": "Gộp ô (Merge and Center), tự động xuống dòng (Wrap Text) và hoàn thiện bảng tổng hợp số liệu.", "assessment": ""},
        {"week": 15, "lesson_index": 15, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 10. Hoàn thiện bảng tính", "content": "Sắp xếp dữ liệu tăng dần/giảm dần, thiết lập vùng in và in bảng tính chuyên nghiệp.", "assessment": ""},
        {"week": 16, "lesson_index": 16, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 10. Hoàn thiện bảng tính (tiếp)", "content": "Thực hành dự án quản lí chi tiêu gia đình hoặc theo dõi thành tích học tập cá nhân.", "assessment": ""},
        {"week": 17, "lesson_index": 17, "semester": "Học kỳ I", "chapter": "Ôn tập cuối kỳ", "unit": "Ôn tập học kỳ I", "content": "Hệ thống hóa toàn bộ kiến thức thiết bị, tệp tin và kỹ năng bảng tính Excel Học kỳ I.", "assessment": ""},
        {"week": 18, "lesson_index": 18, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ I", "content": "Đánh giá toàn diện kiến thức và thực hành Học kỳ I môn Tin học 7.", "assessment": "KIỂM TRA CUỐI HỌC KỲ I"},
        # HỌC KỲ II (Tuần 19 - 35)
        {"week": 19, "lesson_index": 19, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 11. Tạo bài trình chiếu", "content": "Khái niệm bài trình chiếu, cấu trúc phân cấp (tiêu đề, ý chính, ý phụ), trang chiếu tiêu đề.", "assessment": ""},
        {"week": 20, "lesson_index": 20, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 11. Tạo bài trình chiếu (tiếp)", "content": "Áp dụng mẫu bố cục (Layout), chọn mẫu định dạng (Theme) và căn chỉnh văn bản trên slide.", "assessment": ""},
        {"week": 21, "lesson_index": 21, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 12. Định dạng đối tượng trên trang chiếu", "content": "Chèn hình ảnh, sơ đồ hình khối (SmartArt), điều chỉnh kích thước và vị trí đối tượng.", "assessment": ""},
        {"week": 22, "lesson_index": 22, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 12. Định dạng đối tượng trên trang chiếu (tiếp)", "content": "Thiết lập hiệu ứng chuyển trang (Transition) và hiệu ứng xuất hiện đối tượng (Animation) hợp lý.", "assessment": ""},
        {"week": 23, "lesson_index": 23, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 13. Thực hành tổng hợp: Hoàn thiện bài trình chiếu", "content": "Xây dựng bài thuyết trình giới thiệu về một danh lam thắng cảnh hoặc chủ đề khoa học.", "assessment": ""},
        {"week": 24, "lesson_index": 24, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 14. Thuật toán tìm kiếm tuần tự", "content": "Ý tưởng và nguyên lý thuật toán tìm kiếm tuần tự (Sequential Search) trên danh sách.", "assessment": ""},
        {"week": 25, "lesson_index": 25, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 14. Thuật toán tìm kiếm tuần tự (tiếp)", "content": "Mô tả thuật toán tìm kiếm tuần tự bằng sơ đồ khối và đánh giá số bước thực hiện.", "assessment": ""},
        {"week": 26, "lesson_index": 26, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Thuật toán tìm kiếm nhị phân", "content": "Điều kiện tiên quyết áp dụng tìm kiếm nhị phân (danh sách đã sắp xếp), nguyên lý chia đôi khoảng tìm kiếm.", "assessment": ""},
        {"week": 27, "lesson_index": 27, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra giữa Học kỳ II", "content": "Đánh giá kỹ năng thiết kế bài trình chiếu và thuật toán tìm kiếm (Tuần 19 đến Tuần 26).", "assessment": "KIỂM TRA GIỮA HỌC KỲ II"},
        {"week": 28, "lesson_index": 28, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Thuật toán tìm kiếm nhị phân (tiếp)", "content": "Mô phỏng từng bước thuật toán tìm kiếm nhị phân trên dãy số cụ thể bằng bảng theo dõi.", "assessment": ""},
        {"week": 29, "lesson_index": 29, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 16. Thuật toán sắp xếp", "content": "Mục đích của việc sắp xếp dữ liệu, giới thiệu ý tưởng thuật toán sắp xếp nổi bọt (Bubble Sort).", "assessment": ""},
        {"week": 30, "lesson_index": 30, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 16. Thuật toán sắp xếp (tiếp)", "content": "Ý tưởng thuật toán sắp xếp chọn (Selection Sort), mô phỏng từng bước đổi chỗ các phần tử.", "assessment": ""},
        {"week": 31, "lesson_index": 31, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 16. Thuật toán sắp xếp (luyện tập)", "content": "So sánh hiệu quả của tìm kiếm nhị phân so với tuần tự sau khi dãy đã được sắp xếp.", "assessment": ""},
        {"week": 32, "lesson_index": 32, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Thực hành tổng hợp thuật toán", "content": "Bài tập tổng hợp mô tả và thực thi thuật toán tìm kiếm và sắp xếp.", "assessment": ""},
        {"week": 33, "lesson_index": 33, "semester": "Học kỳ II", "chapter": "Ôn tập cuối năm", "unit": "Ôn tập học kỳ II", "content": "Hệ thống hóa toàn bộ kiến thức bài trình chiếu và thuật toán Tin học 7.", "assessment": ""},
        {"week": 34, "lesson_index": 34, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ II", "content": "Đánh giá tổng hợp toàn diện cuối năm học môn Tin học 7.", "assessment": "KIỂM TRA CUỐI HỌC KỲ II"},
        {"week": 35, "lesson_index": 35, "semester": "Học kỳ II", "chapter": "Tổng kết", "unit": "Đánh giá & Trả bài kiểm tra", "content": "Tổng kết điểm số, đánh giá năng lực và chuẩn bị chương trình Lớp 8.", "assessment": ""}
    ],

    # ------------------------------- LỚP 8 -------------------------------
    "Lớp 8": [
        # HỌC KỲ I (Tuần 1 - 18)
        {"week": 1, "lesson_index": 1, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Lược sử công cụ tính toán", "content": "Lịch sử phát triển các công cụ tính toán: Bàn tính, máy tính cơ học Pascaline, các thế hệ máy tính điện tử.", "assessment": ""},
        {"week": 2, "lesson_index": 2, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 2. Thông tin trong môi trường số", "content": "Đặc điểm thông tin số (dễ sao chép, lan truyền nhanh), tiêu chuẩn đánh giá độ tin cậy của thông tin.", "assessment": ""},
        {"week": 3, "lesson_index": 3, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 3. Thực hành: Khai thác thông tin số", "content": "Khai thác thông tin trên Internet an toàn, kiểm chứng nguồn tin và tìm kiếm thông tin học thuật.", "assessment": ""},
        {"week": 4, "lesson_index": 4, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 4. Đạo đức và pháp luật trong môi trường số", "content": "Bản quyền phần mềm, luật sở hữu trí tuệ, tác hại của việc sử dụng phần mềm bẻ khóa.", "assessment": ""},
        {"week": 5, "lesson_index": 5, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 5. Sử dụng bảng tính giải quyết bài toán thực tế", "content": "Khái niệm địa chỉ tương đối và địa chỉ tuyệt đối ($A$1) trong công thức bảng tính Excel.", "assessment": ""},
        {"week": 6, "lesson_index": 6, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 5. Sử dụng bảng tính giải quyết bài toán thực tế (tiếp)", "content": "Sử dụng địa chỉ hỗn hợp và sao chép công thức chứa địa chỉ tuyệt đối trong tính toán tài chính.", "assessment": ""},
        {"week": 7, "lesson_index": 7, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 6. Sắp xếp và lọc dữ liệu trong bảng tính", "content": "Lọc tự động AutoFilter, lọc theo điều kiện số liệu, lọc dữ liệu lớn phục vụ báo cáo.", "assessment": ""},
        {"week": 8, "lesson_index": 8, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 7. Trực quan hoá dữ liệu bằng biểu đồ", "content": "Tạo các loại biểu đồ: Biểu đồ cột (so sánh), biểu đồ đường (xu hướng), biểu đồ tròn (cơ cấu tỉ lệ).", "assessment": ""},
        {"week": 9, "lesson_index": 9, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Ôn tập & Kiểm tra giữa Học kỳ I", "content": "Đánh giá YCCĐ từ Tuần 1 đến Tuần 8 (Lịch sử máy tính, đạo đức số và bảng tính Excel nâng cao).", "assessment": "KIỂM TRA GIỮA HỌC KỲ I"},
        {"week": 10, "lesson_index": 10, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 7. Trực quan hoá dữ liệu bằng biểu đồ (tiếp)", "content": "Tùy biến tiêu đề biểu đồ, chú giải, nhãn dữ liệu và xuất biểu đồ ra báo cáo.", "assessment": ""},
        {"week": 11, "lesson_index": 11, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 8a. Làm việc với danh sách dạng liệt kê và hình ảnh", "content": "Tạo danh sách dấu đầu dòng (Bullet), danh sách đánh số (Numbering) và ngắt trang trong văn bản.", "assessment": ""},
        {"week": 12, "lesson_index": 12, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 9a. Tạo đầu trang và chân trang cho văn bản", "content": "Thiết lập Header, Footer, chèn số trang tự động và phân chia phân đoạn (Section) trong tài liệu dài.", "assessment": ""},
        {"week": 13, "lesson_index": 13, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 10a. Định dạng nâng cao cho bài trình chiếu", "content": "Sử dụng Slide Master để tạo mẫu giao diện đồng nhất cho toàn bộ bài thuyết trình.", "assessment": ""},
        {"week": 14, "lesson_index": 14, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 11a. Thực hành: Tạo sản phẩm số hoàn thiện", "content": "Tạo bài trình chiếu đa phương tiện đính kèm video, âm thanh và liên kết Hyperlink.", "assessment": ""},
        {"week": 15, "lesson_index": 15, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 8b. Phần mềm chỉnh sửa ảnh cơ bản", "content": "Làm quen giao diện phần mềm chỉnh sửa ảnh GIMP/Paint.net, các công cụ chọn vùng.", "assessment": ""},
        {"week": 16, "lesson_index": 16, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Ứng dụng tin học", "unit": "Bài 9b. Cắt ghép, thay đổi kích thước và hiệu ứng ảnh", "content": "Cắt ảnh (Crop), xoay ảnh, chỉnh độ sáng tối và áp dụng các bộ lọc màu nghệ thuật.", "assessment": ""},
        {"week": 17, "lesson_index": 17, "semester": "Học kỳ I", "chapter": "Ôn tập cuối kỳ", "unit": "Ôn tập học kỳ I", "content": "Hệ thống hóa toàn bộ kiến thức xử lý văn bản, bảng tính, trình chiếu và chỉnh sửa ảnh.", "assessment": ""},
        {"week": 18, "lesson_index": 18, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ I", "content": "Đánh giá toàn diện kiến thức và kỹ năng thực hành Học kỳ I môn Tin học 8.", "assessment": "KIỂM TRA CUỐI HỌC KỲ I"},
        # HỌC KỲ II (Tuần 19 - 35)
        {"week": 19, "lesson_index": 19, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 12. Từ thuật toán đến chương trình", "content": "Mối quan hệ giữa thuật toán và chương trình máy tính. Khái niệm môi trường lập trình trực quan.", "assessment": ""},
        {"week": 20, "lesson_index": 20, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 13. Biến và kiểu dữ liệu trong chương trình", "content": "Khái niệm biến số, hằng số, các kiểu dữ liệu số nguyên, số thực, xâu kí tự và giá trị logic.", "assessment": ""},
        {"week": 21, "lesson_index": 21, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 13. Biến và kiểu dữ liệu (tiếp)", "content": "Tạo biến, gán giá trị cho biến và sử dụng biến trong các phép toán số học và so sánh.", "assessment": ""},
        {"week": 22, "lesson_index": 22, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 14. Cấu trúc điều khiển rẽ nhánh", "content": "Câu lệnh điều kiện đầy đủ, biểu thức điều kiện kết hợp toán tử VÀ (AND), HOẶC (OR), PHỦ ĐỊNH (NOT).", "assessment": ""},
        {"week": 23, "lesson_index": 23, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 14. Cấu trúc điều khiển lặp", "content": "Vòng lặp với số lần biết trước, vòng lặp dừng theo điều kiện và ứng dụng tính tổng dãy số.", "assessment": ""},
        {"week": 24, "lesson_index": 24, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Gỡ lỗi chương trình", "content": "Nhận biết các loại lỗi: Lỗi cú pháp, lỗi thực thi và lỗi logic. Kỹ thuật lần vết giá trị biến để sửa lỗi.", "assessment": ""},
        {"week": 25, "lesson_index": 25, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Gỡ lỗi chương trình (tiếp)", "content": "Thực hành phát hiện và sửa lỗi trong các kịch bản chương trình tính toán và trò chơi.", "assessment": ""},
        {"week": 26, "lesson_index": 26, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Thực hành lập trình giải quyết bài toán", "content": "Xây dựng chương trình tính tiền điện hoặc kiểm tra số nguyên tố bằng lập trình khối.", "assessment": ""},
        {"week": 27, "lesson_index": 27, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra giữa Học kỳ II", "content": "Đánh giá kỹ năng lập trình, biến số, cấu trúc rẽ nhánh, lặp và gỡ lỗi (Tuần 19 đến Tuần 26).", "assessment": "KIỂM TRA GIỮA HỌC KỲ II"},
        {"week": 28, "lesson_index": 28, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Hướng nghiệp với Tin học", "unit": "Bài 16. Tin học và nghề nghiệp trong xã hội số", "content": "Vai trò của Tin học trong các ngành nghề: Y tế, ngân hàng, giáo dục, nông nghiệp công nghệ cao.", "assessment": ""},
        {"week": 29, "lesson_index": 29, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Hướng nghiệp với Tin học", "unit": "Bài 16. Tin học và nghề nghiệp trong xã hội số (tiếp)", "content": "Các nghề thuộc ngành công nghệ thông tin và truyền thông, cơ hội việc làm và bình đẳng giới.", "assessment": ""},
        {"week": 30, "lesson_index": 30, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Dự án lập trình: Trò chơi giáo dục", "content": "Thiết kế dự án trò chơi hỏi đáp kiến thức trắc nghiệm có chấm điểm tự động.", "assessment": ""},
        {"week": 31, "lesson_index": 31, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Dự án lập trình: Trò chơi giáo dục (tiếp)", "content": "Hoàn thiện âm thanh, hiệu ứng thắng thua và tối ưu kịch bản dự án.", "assessment": ""},
        {"week": 32, "lesson_index": 32, "semester": "Học kỳ II", "chapter": "Chủ đề 4. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Báo cáo sản phẩm dự án", "content": "Báo cáo demo sản phẩm trước lớp và phản biện nhóm.", "assessment": ""},
        {"week": 33, "lesson_index": 33, "semester": "Học kỳ II", "chapter": "Ôn tập cuối năm", "unit": "Ôn tập học kỳ II", "content": "Hệ thống hóa toàn bộ kiến thức lập trình giải quyết vấn đề và hướng nghiệp Tin học 8.", "assessment": ""},
        {"week": 34, "lesson_index": 34, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ II", "content": "Đánh giá tổng hợp toàn diện năng lực môn Tin học 8 cả năm học.", "assessment": "KIỂM TRA CUỐI HỌC KỲ II"},
        {"week": 35, "lesson_index": 35, "semester": "Học kỳ II", "chapter": "Tổng kết", "unit": "Đánh giá & Trả bài kiểm tra", "content": "Tổng kết, rút kinh nghiệm và định hướng môn Tin học Lớp 9.", "assessment": ""}
    ],

    # ------------------------------- LỚP 9 -------------------------------
    "Lớp 9": [
        # HỌC KỲ I (Tuần 1 - 18)
        {"week": 1, "lesson_index": 1, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Thế giới kĩ thuật số", "content": "Vai trò của thiết bị số trong cuộc sống, tính chất thông tin trong xã hội tri thức, thiết bị số thông minh.", "assessment": ""},
        {"week": 2, "lesson_index": 2, "semester": "Học kỳ I", "chapter": "Chủ đề 1. Máy tính và cộng đồng", "unit": "Bài 1. Thế giới kĩ thuật số (tiếp)", "content": "Tác động của thiết bị số đối với y tế, giáo dục, sản xuất tự động và giao thông thông minh.", "assessment": ""},
        {"week": 3, "lesson_index": 3, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin", "unit": "Bài 2. Thông tin trong môi trường số", "content": "Chất lượng thông tin, độ tin cậy và tiêu chí đánh giá độ tin cậy của thông tin trên mạng Internet.", "assessment": ""},
        {"week": 4, "lesson_index": 4, "semester": "Học kỳ I", "chapter": "Chủ đề 2. Tổ chức lưu trữ, tìm kiếm và trao đổi thông tin", "unit": "Bài 2. Thông tin trong môi trường số (tiếp)", "content": "Nhận diện tin giả (Fake news), tin độc hại và cách phòng tránh các chiêu trò lừa đảo trực tuyến.", "assessment": ""},
        {"week": 5, "lesson_index": 5, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 3. Thực hành: Đánh giá chất lượng thông tin", "content": "Thực hành tra cứu, kiểm chứng chéo và đánh giá tính xác thực của thông tin số trước khi chia sẻ.", "assessment": ""},
        {"week": 6, "lesson_index": 6, "semester": "Học kỳ I", "chapter": "Chủ đề 3. Đạo đức, pháp luật và văn hoá trong môi trường số", "unit": "Bài 4. Một số vấn đề pháp lí về sử dụng dịch vụ Internet", "content": "Bản quyền số, quyền sở hữu trí tuệ, luật an ninh mạng và trách nhiệm công dân số trên mạng.", "assessment": ""},
        {"week": 7, "lesson_index": 7, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 5. Sử dụng phần mềm mô phỏng", "content": "Khái niệm phần mềm mô phỏng, lợi ích của việc học tập qua thí nghiệm ảo và mô phỏng thực tế.", "assessment": ""},
        {"week": 8, "lesson_index": 8, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 6. Thực hành: Khai thác phần mềm mô phỏng", "content": "Thực hành khám phá hiện tượng vật lí, hóa học hoặc mạch điện qua phần mềm mô phỏng PhET.", "assessment": ""},
        {"week": 9, "lesson_index": 9, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Ôn tập & Kiểm tra giữa Học kỳ I", "content": "Kiểm tra đánh giá YCCĐ từ Bài 1 đến Bài 6 (Chủ đề 1 đến Chủ đề 4).", "assessment": "KIỂM TRA GIỮA HỌC KỲ I"},
        {"week": 10, "lesson_index": 10, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 7. Trình bày thông tin trong trao đổi và hợp tác", "content": "Nguyên tắc thiết kế sơ đồ tư duy và bài trình chiếu chuyên nghiệp phục vụ báo cáo nhóm.", "assessment": ""},
        {"week": 11, "lesson_index": 11, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 8. Thực hành: Sử dụng công cụ trực quan trình bày thông tin", "content": "Sử dụng công cụ trực quan để thuyết trình báo cáo dự án học tập trước tập thể.", "assessment": ""},
        {"week": 12, "lesson_index": 12, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 9a. Sử dụng công cụ xác thực dữ liệu (Data Validation)", "content": "Thiết lập quy tắc xác thực dữ liệu đầu vào trong bảng tính Excel, cảnh báo lỗi nhập liệu.", "assessment": ""},
        {"week": 13, "lesson_index": 13, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 10a. Sử dụng hàm đếm có điều kiện COUNTIF", "content": "Cú pháp và ứng dụng hàm COUNTIF trong thống kê số lượng học sinh theo tiêu chí xếp loại.", "assessment": ""},
        {"week": 14, "lesson_index": 14, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 11a. Sử dụng hàm tính tổng có điều kiện SUMIF", "content": "Cú pháp và ứng dụng hàm SUMIF trong quản lí tài chính, tính tổng chi phí theo từng khoản mục.", "assessment": ""},
        {"week": 15, "lesson_index": 15, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 12a. Sử dụng hàm điều kiện IF", "content": "Cú pháp hàm điều kiện IF, hàm IF lồng nhau để phân loại học lực và xét điều kiện khen thưởng.", "assessment": ""},
        {"week": 16, "lesson_index": 16, "semester": "Học kỳ I", "chapter": "Chủ đề 4. Ứng dụng tin học", "unit": "Bài 13a. Hoàn thiện bảng tính quản lí tài chính gia đình", "content": "Dự án tổng hợp kết hợp Data Validation, COUNTIF, SUMIF, IF và biểu đồ phân tích.", "assessment": ""},
        {"week": 17, "lesson_index": 17, "semester": "Học kỳ I", "chapter": "Ôn tập cuối kỳ", "unit": "Ôn tập học kỳ I", "content": "Hệ thống hóa toàn bộ kiến thức thiết bị số, đạo đức số và kỹ năng bảng tính hàm nâng cao.", "assessment": ""},
        {"week": 18, "lesson_index": 18, "semester": "Học kỳ I", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ I", "content": "Đánh giá tổng hợp toàn diện Học kỳ I môn Tin học 9.", "assessment": "KIỂM TRA CUỐI HỌC KỲ I"},
        # HỌC KỲ II (Tuần 19 - 35)
        {"week": 19, "lesson_index": 19, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 14. Giải quyết vấn đề bằng máy tính", "content": "Quy trình giải quyết vấn đề gồm 4 bước: Xác định bài toán, chia nhỏ vấn đề, mô hình hóa và cài đặt.", "assessment": ""},
        {"week": 20, "lesson_index": 20, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 15. Bài toán trong tin học và thuật toán", "content": "Xác định rõ Input, Output của bài toán. Mô tả thuật toán bằng ngôn ngữ tự nhiên và lưu đồ.", "assessment": ""},
        {"week": 21, "lesson_index": 21, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 16. Thực hành: Lập trình máy tính", "content": "Làm quen với ngôn ngữ lập trình Python, môi trường IDLE/VS Code, lệnh print() và phép toán cơ bản.", "assessment": ""},
        {"week": 22, "lesson_index": 22, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Bài 16. Thực hành: Biến và lệnh gán trong Python", "content": "Khai báo biến, lệnh gán (=), hàm input() để nhập dữ liệu từ bàn phím và ép kiểu int(), float().", "assessment": ""},
        {"week": 23, "lesson_index": 23, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Cấu trúc rẽ nhánh trong Python", "content": "Câu lệnh if, if-else, thụt đầu dòng (indentation) và biểu thức điều kiện so sánh trong Python.", "assessment": ""},
        {"week": 24, "lesson_index": 24, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Cấu trúc lặp for trong Python", "content": "Vòng lặp for với hàm range(n), duyệt biến đếm và tính tổng dãy số nguyên liên tiếp.", "assessment": ""},
        {"week": 25, "lesson_index": 25, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Vòng lặp while trong Python", "content": "Vòng lặp while với điều kiện kiểm tra, điều kiện dừng và cách phòng tránh lặp vô hạn.", "assessment": ""},
        {"week": 26, "lesson_index": 26, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Thực hành lập trình thuật toán cơ bản", "content": "Cài đặt thuật toán tìm ước chung lớn nhất (UCLN) hoặc kiểm tra số nguyên tố trong Python.", "assessment": ""},
        {"week": 27, "lesson_index": 27, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra giữa Học kỳ II", "content": "Đánh giá YCCĐ giải quyết vấn đề và kỹ năng lập trình Python cơ bản (Tuần 19 đến Tuần 26).", "assessment": "KIỂM TRA GIỮA HỌC KỲ II"},
        {"week": 28, "lesson_index": 28, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Dữ liệu kiểu danh sách (List) trong Python", "content": "Khái niệm danh sách, truy cập phần tử qua chỉ số (index), thêm phần tử với append().", "assessment": ""},
        {"week": 29, "lesson_index": 29, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Thuật toán tìm kiếm tuần tự trên List", "content": "Cài đặt thuật toán tìm kiếm tuần tự một phần tử trong danh sách bằng vòng lặp for.", "assessment": ""},
        {"week": 30, "lesson_index": 30, "semester": "Học kỳ II", "chapter": "Chủ đề 5. Giải quyết vấn đề với sự trợ giúp của máy tính", "unit": "Thuật toán sắp xếp danh sách cơ bản", "content": "Mô phỏng và cài đặt thuật toán sắp xếp nổi bọt (Bubble Sort) trên mảng Python.", "assessment": ""},
        {"week": 31, "lesson_index": 31, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Hướng nghiệp với Tin học", "unit": "Bài 17. Tin học và thế giới nghề nghiệp", "content": "Các nhóm nghề nghiệp trong ngành CNTT: Lập trình viên, chuyên viên an ninh mạng, nhà khoa học dữ liệu.", "assessment": ""},
        {"week": 32, "lesson_index": 32, "semester": "Học kỳ II", "chapter": "Chủ đề 6. Hướng nghiệp với Tin học", "unit": "Bài 17. Tin học và thế giới nghề nghiệp (tiếp)", "content": "Bình đẳng giới trong lĩnh vực tin học, lập kế hoạch học tập và chuẩn bị hành trang vào THPT.", "assessment": ""},
        {"week": 33, "lesson_index": 33, "semester": "Học kỳ II", "chapter": "Ôn tập cuối năm", "unit": "Ôn tập học kỳ II", "content": "Tổng kết toàn bộ kiến thức lập trình giải quyết vấn đề và hướng nghiệp Tin học 9.", "assessment": ""},
        {"week": 34, "lesson_index": 34, "semester": "Học kỳ II", "chapter": "Đánh giá định kỳ", "unit": "Kiểm tra cuối Học kỳ II", "content": "Đánh giá tổng hợp toàn diện năng lực môn Tin học 9 tốt nghiệp cấp THCS.", "assessment": "KIỂM TRA CUỐI HỌC KỲ II"},
        {"week": 35, "lesson_index": 35, "semester": "Học kỳ II", "chapter": "Tổng kết", "unit": "Đánh giá & Trả bài kiểm tra", "content": "Tổng kết điểm số, đánh giá xếp loại học tập và bế mạc năm học.", "assessment": ""}
    ]
}


# ======================================================================================
# 2. BỘ PHÂN TÍCH ĐA ĐỊNH DẠNG (DOC, DOCX, XLSX, PDF, CSV, JSON) & PHÂN TÁCH ĐA KHỐI LỚP
# ======================================================================================

class CurriculumPlanner:
    """Quản lý logic Kế hoạch giáo dục và lọc phạm vi ra đề theo tuần/bài."""

    @staticmethod
    def get_standard_plan(grade: str) -> List[Dict[str, Any]]:
        """Lấy KHGD 35 tuần chuẩn GDPT 2018 của khối lớp."""
        normalized_grade = grade.strip()
        if normalized_grade in STANDARD_KHGD_DATA:
            return STANDARD_KHGD_DATA[normalized_grade]
        for g_key in STANDARD_KHGD_DATA:
            if normalized_grade in g_key or g_key in normalized_grade:
                return STANDARD_KHGD_DATA[g_key]
        return STANDARD_KHGD_DATA.get("Lớp 6", [])

    @staticmethod
    def parse_multimodal_school_plan(file_bytes: bytes, filename: str) -> Dict[str, List[Dict[str, Any]]]:
        """
        Bóc tách thông minh file KHGD của trường hỗ trợ đa định dạng:
        .doc (Word 97-2003), .docx (Word OpenXML), .xlsx, .xls, .pdf, .csv, .json, .txt.
        Tự động nhận diện và phân chia theo từng Khối lớp (Lớp 6, Lớp 7, Lớp 8, Lớp 9).
        Trả về dict: {"Lớp 6": [...], "Lớp 7": [...], "Lớp 8": [...], "Lớp 9": [...]}
        """
        ext = os.path.splitext(filename)[1].lower()
        is_cfbf = file_bytes.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")

        logger.info(f"Đang bóc tách file KHGD: {filename} ({ext}, CFBF={is_cfbf})")

        # Xử lý tệp Word nhị phân cũ .doc (Compound File Binary Format)
        if ext == ".doc" or is_cfbf:
            return CurriculumPlanner._parse_legacy_doc_plan(file_bytes, filename=filename)
        elif ext == ".docx":
            return CurriculumPlanner._parse_docx_plan(file_bytes, filename=filename)
        elif ext in [".xlsx", ".xls"]:
            return CurriculumPlanner._parse_excel_plan(file_bytes, filename=filename)
        elif ext == ".pdf":
            return CurriculumPlanner._parse_pdf_plan(file_bytes, filename=filename)
        elif ext == ".csv":
            return CurriculumPlanner._parse_csv_plan(file_bytes, filename=filename)
        elif ext == ".json":
            return CurriculumPlanner._parse_json_plan(file_bytes, filename=filename)
        elif ext == ".txt":
            return CurriculumPlanner._parse_text_plan(file_bytes.decode("utf-8", errors="ignore"), filename=filename)
        else:
            raise ValueError(f"Định dạng {ext} chưa được hỗ trợ. Vui lòng tải file Word (.docx, .doc), Excel (.xlsx), PDF (.pdf), CSV hoặc JSON.")

    # ------------------- Chuyển đổi & Bóc tách Word (.doc / .docx) -------------------

    @staticmethod
    def parse_yccd_into_levels(raw_text: str) -> Dict[str, str]:
        """
        Phân tích thông minh các gạch đầu dòng YCCĐ trong KHGD thành 3 mức độ nhận thức:
        - Nhận biết (nb)
        - Thông hiểu (th)
        - Vận dụng (vd)
        Bám sát các động từ chỉ mức độ tư duy chuẩn Thông tư GDPT 2018.
        """
        if not raw_text:
            return {
                "nb": "- Nhận biết kiến thức cốt lõi và các khái niệm cơ bản của bài học.",
                "th": "- Giải thích, phân biệt và làm rõ được các nội dung trọng tâm của bài học.",
                "vd": "- Vận dụng kiến thức bài học để giải quyết vấn đề thực tiễn."
            }

        raw_clean = raw_text.replace("\r\n", "\n").replace("\r", "\n")
        # Chuẩn hóa các loại gạch đầu dòng ˗ (U+02D7), -, –, —, •, *, +
        normalized = re.sub(r'[\n\r]*[\˗\-\–\—\•\*\+]\s*', '\n- ', raw_clean).strip()
        lines = [l.strip().lstrip('- ').strip() for l in normalized.split('\n') if l.strip().lstrip('- ').strip()]

        levels = {"nb": [], "th": [], "vd": []}

        # Từ khóa nhận biết
        nb_prefixes = (
            "nhận biết", "nêu", "chỉ ra", "liệt kê", "kể tên", "nhận ra", "biết", "nhận diện", "xác định"
        )
        # Từ khóa thông hiểu
        th_prefixes = (
            "hiểu", "giải thích", "trình bày", "phân biệt", "so sánh", "mô tả", "làm rõ", "phân tích", "làm quen"
        )
        # Từ khóa vận dụng
        vd_prefixes = (
            "vận dụng", "sử dụng", "thực hiện", "tạo", "tìm kiếm", "đánh giá", "khai thác", "áp dụng", "thực hành", "giải quyết", "hoàn thiện"
        )

        for line in lines:
            l_lower = line.lower()
            first_words = " ".join(l_lower.split()[:4])
            if any(first_words.startswith(p) for p in nb_prefixes):
                levels["nb"].append(line)
            elif any(first_words.startswith(p) for p in th_prefixes):
                levels["th"].append(line)
            elif any(first_words.startswith(p) for p in vd_prefixes):
                levels["vd"].append(line)
            else:
                # Kiểm tra theo ngữ nghĩa trong câu
                if any(p in l_lower for p in ["vận dụng", "thực hành", "giải quyết", "tạo được", "sử dụng được"]):
                    levels["vd"].append(line)
                elif any(p in l_lower for p in ["giải thích", "trình bày", "phân biệt", "hiểu được"]):
                    levels["th"].append(line)
                else:
                    levels["nb"].append(line)

        res = {
            "nb": "\n".join([f"- {item}" for item in levels["nb"]]).strip(),
            "th": "\n".join([f"- {item}" for item in levels["th"]]).strip(),
            "vd": "\n".join([f"- {item}" for item in levels["vd"]]).strip()
        }

        if not res["nb"]:
            res["nb"] = "- Nhận biết kiến thức cốt lõi và các khái niệm cơ bản của bài học."
        if not res["th"]:
            res["th"] = "- Giải thích, phân biệt và làm rõ được các nội dung trọng tâm của bài học."
        if not res["vd"]:
            res["vd"] = "- Vận dụng kiến thức bài học để giải quyết vấn đề thực tiễn."

        return res

    @staticmethod
    def _detect_grade_for_table(tbl, doc=None, filename="", fallback_index=0) -> str:
        """
        Nhận diện chính xác khối lớp (Lớp 6, Lớp 7, Lớp 8, Lớp 9) của bảng PPCT:
        1. Phân tích tên file: ví dụ 'KHGD9.docx', 'Tin 9', 'Khoi 9'...
        2. Phân tích nội dung tiêu đề và văn bản trong tài liệu.
        3. Phân tích từ khóa bài học đặc thù của từng khối lớp môn Tin học THCS.
        """
        scores = {"Lớp 6": 0, "Lớp 7": 0, "Lớp 8": 0, "Lớp 9": 0}

        # 1. Tên file
        fn_lower = filename.lower()
        fn_match = re.search(r'(?:khgd|tin|khoi|lop|k|_|-|\s)?\s*([6789])', fn_lower)
        if fn_match:
            fn_grade = f"Lớp {fn_match.group(1)}"
            scores[fn_grade] += 12

        # 2. Toàn bộ chữ trong bảng
        table_text = " ".join([c.text.lower() for r in tbl.rows for c in r.cells])

        # Từ khóa đặc trưng của từng khối lớp theo SGK Tin học Kết nối tri thức:
        signatures = {
            "Lớp 9": [
                "thế giới kĩ thuật số", "phần mềm mô phỏng", "pháp lí về sử dụng dịch vụ internet",
                "countif", "sumif", "hàm if", "tin học và thế giới nghề nghiệp",
                "thông tin trong giải quyết vấn đề", "bảng tính quản lí tài chính", "bài toán tin học",
                "khối 9", "lớp 9", "tin học 9", "tin 9"
            ],
            "Lớp 8": [
                "lịch sử phát triển máy tính", "soạn thảo văn bản nâng cao", "thuật toán mảng",
                "lập trình trực quan", "khối 8", "lớp 8", "tin học 8", "tin 8", "vẽ hình trong lập trình"
            ],
            "Lớp 7": [
                "thiết bị vào - ra", "phần mềm ứng dụng", "định dạng bảng tính", "công thức trong bảng tính",
                "khối 7", "lớp 7", "tin học 7", "tin 7"
            ],
            "Lớp 6": [
                "thông tin và dữ liệu", "xử lý thông tin", "mạng máy tính", "sơ đồ tư duy",
                "an toàn trên internet", "khối 6", "lớp 6", "tin học 6", "tin 6"
            ]
        }

        for g, kws in signatures.items():
            for kw in kws:
                if kw in table_text:
                    scores[g] += 4

        # 3. Quét các đoạn văn bản trong doc (nếu có)
        if doc and hasattr(doc, 'paragraphs'):
            full_doc_text = " ".join([p.text.lower() for p in doc.paragraphs])
            for g_num in ["6", "7", "8", "9"]:
                g_str = f"Lớp {g_num}"
                if f"tin học {g_num}" in full_doc_text or f"lớp {g_num}" in full_doc_text or f"khối {g_num}" in full_doc_text:
                    scores[g_str] += 6

        # Lấy khối lớp có điểm cao nhất
        max_score = max(scores.values())
        if max_score > 0:
            best_grade = max(scores.items(), key=lambda x: x[1])[0]
            return best_grade

        # Fallback theo thứ tự
        grade_order = ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"]
        return grade_order[fallback_index % len(grade_order)]


    @staticmethod
    def _parse_legacy_doc_plan(file_bytes: bytes, filename: str = "") -> Dict[str, List[Dict[str, Any]]]:
        """
        Chuyển đổi tệp Word 97-2003 (.doc nhị phân) sang .docx bằng Microsoft Word COM
        rồi tiến hành bóc tách cấu trúc sư phạm chi tiết.
        """
        converted_docx_bytes = None
        try:
            import win32com.client
            temp_dir = tempfile.gettempdir()
            tmp_doc_path = os.path.join(temp_dir, f"khgd_in_{os.getpid()}.doc")
            tmp_docx_path = os.path.join(temp_dir, f"khgd_out_{os.getpid()}.docx")

            with open(tmp_doc_path, "wb") as f_doc:
                f_doc.write(file_bytes)

            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0  # wdAlertsNone
            try:
                doc = word_app.Documents.Open(
                    FileName=tmp_doc_path,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False
                )
                doc.SaveAs2(tmp_docx_path, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
                doc.Close(SaveChanges=False)
            finally:
                word_app.Quit()

            if os.path.exists(tmp_docx_path):
                with open(tmp_docx_path, "rb") as f_docx:
                    converted_docx_bytes = f_docx.read()
                try:
                    os.remove(tmp_doc_path)
                    os.remove(tmp_docx_path)
                except Exception:
                    pass
        except Exception as ex_com:
            logger.warning(f"Chuyển đổi Word COM thất bại ({ex_com}), chuyển sang trích xuất chuỗi thô.")

        if converted_docx_bytes:
            return CurriculumPlanner._parse_docx_plan(converted_docx_bytes, filename=filename)

        # Fallback khi không dùng được Word COM: trích xuất chuỗi text từ binary doc
        text_content = file_bytes.decode("latin1", errors="ignore")
        return CurriculumPlanner._parse_text_plan(text_content)

    @staticmethod
    def _parse_docx_plan(file_bytes: bytes, filename: str = "") -> Dict[str, List[Dict[str, Any]]]:
        """Bóc tách file Word (.docx) chứa bảng phân phối chương trình của trường."""
        import docx
        doc = docx.Document(io.BytesIO(file_bytes))
        result_by_grade: Dict[str, List[Dict[str, Any]]] = {
            "Lớp 6": [],
            "Lớp 7": [],
            "Lớp 8": [],
            "Lớp 9": []
        }

        # 1. Tìm tất cả các bảng PPCT trong tài liệu
        # Loại trừ các bảng danh mục thiết bị dạy học và bảng phòng máy
        ppct_tables = []
        for idx, tbl in enumerate(doc.tables):
            if len(tbl.rows) < 2:
                continue
            hdr_text = " ".join([c.text.strip().lower() for c in tbl.rows[0].cells])
            # Bỏ qua bảng thiết bị, bảng phòng thực hành
            if "thiết bị" in hdr_text or "tên phòng" in hdr_text or ("số lượng" in hdr_text and "bài học" not in hdr_text):
                continue
            # Bảng PPCT bắt buộc có cột Bài học / Tên bài VÀ Yêu cầu cần đạt
            if ("bài học" in hdr_text or "tên bài" in hdr_text or "bài dạy" in hdr_text) and ("yêu cầu" in hdr_text or "yccd" in hdr_text):
                ppct_tables.append((idx + 1, tbl))

        grade_order = ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"]

        # Nếu tìm thấy đúng các bảng PPCT
        if ppct_tables:
            for t_idx, (tbl_num, tbl) in enumerate(ppct_tables):
                # Nhận diện thông minh khối lớp bám sát tên file, tiêu đề và nội dung bài học
                target_grade = CurriculumPlanner._detect_grade_for_table(tbl, doc, filename=filename, fallback_index=t_idx)
                parsed_rows = CurriculumPlanner._extract_ppct_from_table(tbl)
                if parsed_rows:
                    if target_grade in result_by_grade:
                        result_by_grade[target_grade] = parsed_rows
                    else:
                        result_by_grade["Lớp 9"] = parsed_rows
        else:
            # Fallback duyệt tổng quát nếu bảng có tiêu đề tùy biến
            for idx, tbl in enumerate(doc.tables):
                rows = CurriculumPlanner._extract_ppct_from_table(tbl)
                if len(rows) >= 5:
                    target_grade = CurriculumPlanner._detect_grade_for_table(tbl, doc, filename=filename, fallback_index=idx)
                    if target_grade in result_by_grade and len(result_by_grade[target_grade]) == 0:
                        result_by_grade[target_grade] = rows
                    else:
                        for g in grade_order:
                            if len(result_by_grade[g]) == 0:
                                result_by_grade[g] = rows
                                break

        total_items = sum(len(v) for v in result_by_grade.values())
        if total_items == 0:
            full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            return CurriculumPlanner._parse_text_plan(full_text)

        return result_by_grade

    @staticmethod
    def _extract_ppct_from_table(tbl) -> List[Dict[str, Any]]:
        """Trích xuất danh sách bài học và YCCĐ chuẩn xác từ một bảng PPCT."""
        rows_data = []
        if len(tbl.rows) < 2:
            return rows_data

        curr_chapter = "Chủ đề môn Tin học"
        curr_sem = "Học kỳ I"
        curr_week = 1
        lesson_counter = 1

        for r_idx in range(1, len(tbl.rows)):
            cells = [c.text.strip() for c in tbl.rows[r_idx].cells]
            if not any(cells):
                continue

            row_str = " ".join(cells).strip()

            # Nhận diện dòng chuyển Học kỳ
            if "HỌC KÌ I" in row_str.upper() or "HỌC KỲ I" in row_str.upper():
                curr_sem = "Học kỳ I"
                continue
            elif "HỌC KÌ II" in row_str.upper() or "HỌC KỲ II" in row_str.upper():
                curr_sem = "Học kỳ II"
                continue

            # Nhận diện dòng tiêu đề Chủ đề / Chương (dòng gộp hoặc bắt đầu bằng Chủ đề)
            if row_str.lower().startswith("chủ đề") or "chủ đề" in cells[0].lower():
                curr_chapter = cells[0].replace("\n", " ").strip()
                continue

            # Kiểm tra xem hàng này có phải là hàng bài học không
            stt_val = cells[0].strip()
            # Thông thường cột 0 là STT số (1, 2, 3...) hoặc cột 1 là Tên bài
            is_lesson_row = False
            unit_title = ""
            time_val = ""
            yccd_val = ""

            if stt_val.isdigit() and len(cells) >= 2:
                is_lesson_row = True
                unit_title = cells[1].replace("\n", " ").strip()
                # Cột thời điểm thường là cột 3 (tiết/tuần)
                time_val = cells[3] if len(cells) > 3 else ""
                yccd_val = cells[-1].replace("\n", " ").strip()
            elif any(k in cells[0].lower() for k in ["bài ", "ôn tập", "kiểm tra"]):
                is_lesson_row = True
                unit_title = cells[0].replace("\n", " ").strip()
                time_val = cells[2] if len(cells) > 2 else ""
                yccd_val = cells[-1].replace("\n", " ").strip()

            if is_lesson_row and unit_title:
                # Phân tích tuần học từ cột thời điểm (ví dụ: '1,2', 'Tuần 1', '3,4', '8')
                digits = re.findall(r"\d+", time_val)
                if digits:
                    parsed_w = int(digits[-1])
                    # Nếu số tiết lớn hơn 35 (tiết thứ 50...) thì quy đổi tuần = (tiết + 1) // 2
                    if parsed_w > 35:
                        curr_week = max(1, (parsed_w + 1) // 2)
                    else:
                        curr_week = parsed_w
                else:
                    curr_week = lesson_counter

                lesson_counter += 1

                # Nhận diện đánh giá định kỳ
                assess_val = ""
                u_upper = unit_title.upper()
                if "GIỮA HỌC KỲ I" in u_upper or "GIỮA KỲ 1" in u_upper or "GIỮA KÌ I" in u_upper:
                    assess_val = "KIỂM TRA GIỮA HỌC KỲ I"
                elif "CUỐI HỌC KỲ I" in u_upper or "CUỐI KỲ 1" in u_upper or "CUỐI KÌ I" in u_upper:
                    assess_val = "KIỂM TRA CUỐI HỌC KỲ I"
                elif "GIỮA HỌC KỲ II" in u_upper or "GIỮA KỲ 2" in u_upper or "GIỮA KÌ II" in u_upper:
                    assess_val = "KIỂM TRA GIỮA HỌC KỲ II"
                elif "CUỐI HỌC KỲ II" in u_upper or "CUỐI KỲ 2" in u_upper or "CUỐI KÌ II" in u_upper:
                    assess_val = "KIỂM TRA CUỐI HỌC KỲ II"

                # Bóc tách thông minh các mức độ nhận thức từ nội dung YCCĐ
                lvl_dict = CurriculumPlanner.parse_yccd_into_levels(yccd_val)
                rows_data.append({
                    "week": curr_week,
                    "lesson_index": lesson_counter,
                    "semester": curr_sem if curr_week <= 18 else "Học kỳ II",
                    "chapter": curr_chapter,
                    "unit": unit_title,
                    "content": yccd_val or f"Yêu cầu cần đạt của bài học: {unit_title}",
                    "assessment": assess_val,
                    "req_nb": lvl_dict["nb"],
                    "req_th": lvl_dict["th"],
                    "req_vd": lvl_dict["vd"]
                })

        return rows_data

    # ------------------- Các Parser Excel, PDF, CSV, JSON -------------------

    @staticmethod
    def _parse_excel_plan(file_bytes: bytes) -> Dict[str, List[Dict[str, Any]]]:
        """Bóc tách file Excel (.xlsx, .xls) có thể chứa nhiều sheets cho từng khối lớp."""
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        result_by_grade: Dict[str, List[Dict[str, Any]]] = {
            "Lớp 6": [],
            "Lớp 7": [],
            "Lớp 8": [],
            "Lớp 9": []
        }

        for sheet_name in xl.sheet_names:
            s_upper = sheet_name.upper()
            target_grade = "Lớp 6"
            if "6" in s_upper:
                target_grade = "Lớp 6"
            elif "7" in s_upper:
                target_grade = "Lớp 7"
            elif "8" in s_upper:
                target_grade = "Lớp 8"
            elif "9" in s_upper:
                target_grade = "Lớp 9"

            df = xl.parse(sheet_name)
            parsed_items = CurriculumPlanner._parse_dataframe_plan(df)
            if parsed_items:
                result_by_grade[target_grade].extend(parsed_items)

        total_items = sum(len(v) for v in result_by_grade.values())
        if total_items == 0 and len(xl.sheet_names) > 0:
            df = xl.parse(xl.sheet_names[0])
            col_grade = None
            for c in df.columns:
                if any(k in str(c).lower() for k in ["lớp", "lop", "khối", "khoi", "grade"]):
                    col_grade = c
                    break

            if col_grade:
                for g_val, sub_df in df.groupby(col_grade):
                    g_str = str(g_val).strip()
                    matched_g = "Lớp 6"
                    for k in ["6", "7", "8", "9"]:
                        if k in g_str:
                            matched_g = f"Lớp {k}"
                            break
                    result_by_grade[matched_g] = CurriculumPlanner._parse_dataframe_plan(sub_df)
            else:
                result_by_grade["Lớp 6"] = CurriculumPlanner._parse_dataframe_plan(df)

        return result_by_grade

    @staticmethod
    def _parse_dataframe_plan(df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Chuẩn hóa DataFrame thành danh sách bài học."""
        col_map = {}
        for c in df.columns:
            c_low = str(c).strip().lower()
            if any(k in c_low for k in ["tuần", "tuan", "week", "thời điểm"]):
                col_map[c] = "week"
            elif any(k in c_low for k in ["tiết", "tiet", "stt", "lesson"]):
                col_map[c] = "lesson_index"
            elif any(k in c_low for k in ["học kỳ", "hoc ky", "semester"]):
                col_map[c] = "semester"
            elif any(k in c_low for k in ["chủ đề", "chu de", "chương", "chapter"]):
                col_map[c] = "chapter"
            elif any(k in c_low for k in ["tên bài", "bài", "ten bai", "unit", "bài học"]):
                col_map[c] = "unit"
            elif any(k in c_low for k in ["yêu cầu", "yccd", "mục tiêu", "content"]):
                col_map[c] = "content"
            elif any(k in c_low for k in ["kiểm tra", "đánh giá", "assessment", "ghi chú"]):
                col_map[c] = "assessment"

        df_renamed = df.rename(columns=col_map)
        rows = []
        curr_week = 1

        for _, r in df_renamed.iterrows():
            unit_val = str(r.get("unit", "")).strip()
            if not unit_val or unit_val == "nan":
                continue

            raw_w = r.get("week")
            if pd.notna(raw_w) and str(raw_w).replace(".", "").isdigit():
                curr_week = int(float(raw_w))

            assess_val = str(r.get("assessment", "")).strip() if pd.notna(r.get("assessment")) else ""
            u_upper = unit_val.upper()
            if "GIỮA HỌC KỲ I" in u_upper or "GIỮA KỲ 1" in u_upper:
                assess_val = "KIỂM TRA GIỮA HỌC KỲ I"
            elif "CUỐI HỌC KỲ I" in u_upper or "CUỐI KỲ 1" in u_upper:
                assess_val = "KIỂM TRA CUỐI HỌC KỲ I"
            elif "GIỮA HỌC KỲ II" in u_upper or "GIỮA KỲ 2" in u_upper:
                assess_val = "KIỂM TRA GIỮA HỌC KỲ II"
            elif "CUỐI HỌC KỲ II" in u_upper or "CUỐI KỲ 2" in u_upper:
                assess_val = "KIỂM TRA CUỐI HỌC KỲ II"

            rows.append({
                "week": curr_week,
                "lesson_index": int(float(r.get("lesson_index", 1))) if pd.notna(r.get("lesson_index")) and str(r.get("lesson_index")).replace(".", "").isdigit() else 1,
                "semester": str(r.get("semester", "Học kỳ I" if curr_week <= 18 else "Học kỳ II")).strip(),
                "chapter": str(r.get("chapter", "Chủ đề môn Tin học")).strip(),
                "unit": unit_val,
                "content": str(r.get("content", f"Yêu cầu cần đạt bài {unit_val}")).strip(),
                "assessment": assess_val
            })
        return rows

    @staticmethod
    def _parse_pdf_plan(file_bytes: bytes) -> Dict[str, List[Dict[str, Any]]]:
        """Bóc tách PDF kế hoạch giáo dục sử dụng pypdf hoặc pdfplumber."""
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        full_text = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

        return CurriculumPlanner._parse_text_plan(full_text)

    @staticmethod
    def _parse_csv_plan(file_bytes: bytes) -> Dict[str, List[Dict[str, Any]]]:
        """Bóc tách file CSV."""
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8")
        except Exception:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="cp1258")

        items = CurriculumPlanner._parse_dataframe_plan(df)
        return {"Lớp 6": items}

    @staticmethod
    def _parse_json_plan(file_bytes: bytes) -> Dict[str, List[Dict[str, Any]]]:
        """Bóc tách file JSON."""
        data = json.loads(file_bytes.decode("utf-8"))
        result: Dict[str, List[Dict[str, Any]]] = {"Lớp 6": [], "Lớp 7": [], "Lớp 8": [], "Lớp 9": []}

        if isinstance(data, dict):
            for k, v in data.items():
                for g in ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"]:
                    if g.lower() in k.lower():
                        result[g] = v
                        break
        elif isinstance(data, list):
            result["Lớp 6"] = data

        return result

    @staticmethod
    def _parse_text_plan(text: str) -> Dict[str, List[Dict[str, Any]]]:
        """Phân tích văn bản tự do bằng biểu thức chính quy (Regex) và phân tách khối lớp."""
        result: Dict[str, List[Dict[str, Any]]] = {
            "Lớp 6": [],
            "Lớp 7": [],
            "Lớp 8": [],
            "Lớp 9": []
        }

        current_grade = "Lớp 6"
        lines = text.split("\n")
        week_counter = 1

        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            l_upper = line_str.upper()
            if "TIN HỌC 6" in l_upper or "LỚP 6" in l_upper or "KHỐI 6" in l_upper:
                current_grade = "Lớp 6"
                week_counter = 1
                continue
            elif "TIN HỌC 7" in l_upper or "LỚP 7" in l_upper or "KHỐI 7" in l_upper:
                current_grade = "Lớp 7"
                week_counter = 1
                continue
            elif "TIN HỌC 8" in l_upper or "LỚP 8" in l_upper or "KHỐI 8" in l_upper:
                current_grade = "Lớp 8"
                week_counter = 1
                continue
            elif "TIN HỌC 9" in l_upper or "LỚP 9" in l_upper or "KHỐI 9" in l_upper:
                current_grade = "Lớp 9"
                week_counter = 1
                continue

            match_bai = re.search(r"(Bài\s+\d+[a-zA-Z]?[\.:\s]+[^\n\r]+)", line_str, re.IGNORECASE)
            match_tuan = re.search(r"Tuần\s+(\d+)", line_str, re.IGNORECASE)

            if match_tuan:
                week_counter = int(match_tuan.group(1))

            if match_bai:
                unit_title = match_bai.group(1).strip()
                assess = ""
                u_up = unit_title.upper()
                if "GIỮA HỌC KỲ I" in u_up:
                    assess = "KIỂM TRA GIỮA HỌC KỲ I"
                elif "CUỐI HỌC KỲ I" in u_up:
                    assess = "KIỂM TRA CUỐI HỌC KỲ I"
                elif "GIỮA HỌC KỲ II" in u_up:
                    assess = "KIỂM TRA GIỮA HỌC KỲ II"
                elif "CUỐI HỌC KỲ II" in u_up:
                    assess = "KIỂM TRA CUỐI HỌC KỲ II"

                result[current_grade].append({
                    "week": week_counter,
                    "lesson_index": week_counter,
                    "semester": "Học kỳ I" if week_counter <= 18 else "Học kỳ II",
                    "chapter": "Kế hoạch giáo dục nhà trường",
                    "unit": unit_title,
                    "content": f"Yêu cầu cần đạt: {unit_title}",
                    "assessment": assess
                })

        return result

    # ------------------- Lọc phạm vi ra đề & Xuất file mẫu -------------------

    @staticmethod
    def filter_by_exam_period(
        plan: List[Dict[str, Any]],
        exam_title: str,
        max_week: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        Lọc danh sách các bài học học sinh đã học tới thời điểm kiểm tra.
        Đảm bảo TUYỆT ĐỐI không lấy nhầm bài học của các tuần/kỳ sau.
        """
        ex_upper = exam_title.upper()
        if max_week is None:
            if "GIỮA HỌC KỲ I" in ex_upper or "GIỮA KỲ 1" in ex_upper:
                max_week = 9
            elif "CUỐI HỌC KỲ I" in ex_upper or "CUỐI KỲ 1" in ex_upper:
                max_week = 18
            elif "GIỮA HỌC KỲ II" in ex_upper or "GIỮA KỲ 2" in ex_upper:
                max_week = 27
            elif "CUỐI HỌC KỲ II" in ex_upper or "CUỐI KỲ 2" in ex_upper or "CẢ NĂM" in ex_upper:
                max_week = 35
            else:
                max_week = 9

        filtered = []
        for item in plan:
            w = item.get("week", 1)
            u = item.get("unit", "")
            # Bỏ qua tiết kiểm tra thuần túy khi lấy danh mục bài giảng để tạo câu hỏi
            is_exam_slot = "kiểm tra" in u.lower() or "ôn tập" in u.lower() or "trả bài" in u.lower()
            if w <= max_week:
                if not is_exam_slot or len(filtered) == 0:
                    filtered.append(item)
        return filtered

    @staticmethod
    def export_sample_excel(grade: str = "Lớp 6") -> bytes:
        """Xuất file Excel KHGD mẫu 35 tuần cho giáo viên tải về tham khảo."""
        plan = CurriculumPlanner.get_standard_plan(grade)
        df = pd.DataFrame(plan)
        df = df.rename(columns={
            "week": "Tuần",
            "lesson_index": "Tiết theo PPCT",
            "semester": "Học kỳ",
            "chapter": "Chủ đề / Chương SGK",
            "unit": "Tên bài học",
            "content": "Yêu cầu cần đạt / Nội dung cốt lõi",
            "assessment": "Thời điểm kiểm tra đánh giá"
        })
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=f"KHGD_{grade.replace(' ', '_')}")
        return output.getvalue()
