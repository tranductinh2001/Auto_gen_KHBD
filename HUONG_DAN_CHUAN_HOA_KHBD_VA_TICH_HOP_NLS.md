# HƯỚNG DẪN CHUẨN HÓA KẾ HOẠCH BÀI DẠY (KHBD) & TÍCH HỢP NĂNG LỰC SỐ (NLS)
> **Tài liệu kỹ thuật và Prompt quy chuẩn dùng để tích hợp vào ứng dụng tạo giáo án tự động từ Sách giáo khoa (SGK)**  
> **Căn cứ pháp lý:**  
> - **Công văn 5512/BGDĐT-GDTrH** về việc xây dựng và tổ chức thực hiện kế hoạch giáo dục của nhà trường.  
> - **Thông tư 02/2025/TT-BGDĐT** quy định Khung năng lực số cho người học trong hệ thống giáo dục quốc dân.  
> - **Chương trình Giáo dục phổ thông 2018 (GDPT 2018)**.

---

## MỤC LỤC
1. [Tổng quan Kiến trúc Giáo án Chuẩn](#1-tổng-quan-kiến-trúc-giáo-án-chuẩn)
2. [Quy chuẩn Bảng Tiêu đề (Table 0 - Header Box)](#2-quy-chuẩn-bảng-tiêu-đề-table-0---header-box)
3. [Cấu trúc Chi tiết 5 Phần của KHBD](#3-cấu-trúc-chi-tiết-5-phần-của-khbd)
4. [Bảng Tham chiếu Khung Năng Lực Số (Thông tư 02/2025/TT-BGDĐT)](#4-bảng-tham-chiếu-khung-năng-lực-số-thông-tư-022025tt-bgdđt)
5. [Quy tắc Nhận diện & Đánh dấu Chữ đỏ NLS](#5-quy-tắc-nhận-diện--đánh-dấu-chữ-đỏ-nls)
6. [Quy chuẩn Khối Chữ ký & Kỹ thuật Chống Tách trang (CantSplit)](#6-quy-chuẩn-khối-chữ-ký--kỹ-thuật-chống-tách-trang-cantsplit)
7. [System Prompt & Prompt Mẫu Cho Ứng Dụng AI Sinh Giáo Án Từ SGK](#7-system-prompt--prompt-mẫu-cho-ứng-dụng-ai-sinh-giáo-án-từ-sgk)
8. [Code Mẫu Python (docx-python) Chuẩn Hóa Định Dạng Tự Động](#8-code-mẫu-python-docx-python-chuẩn-hóa-định-dạng-tự-động)

---

## 1. TỔNG QUAN KIẾN TRÚC GIÁO ÁN CHUẨN

Một Kế hoạch bài dạy (KHBD) hoàn chỉnh gồm 6 khối thành phần:
```
┌────────────────────────────────────────────────────────┐
│  BẢNG TIÊU ĐỀ (Table 0): Tuần, Tiết, Ngày dạy | Họ tên, Tổ │
├────────────────────────────────────────────────────────┤
│  TÊN BÀI DẠY: [TÊN BÀI IN HOA, MÀU XANH ĐẬM]           │
├────────────────────────────────────────────────────────┤
│  I. MỤC TIÊU                                           │
│     1. Kiến thức                                       │
│     2. Năng lực (Chung, Đặc thù, Năng lực số chữ đỏ)  │
│     3. Phẩm chất                                       │
├────────────────────────────────────────────────────────┤
│  II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU                      │
├────────────────────────────────────────────────────────┤
│  III. TIẾN TRÌNH DẠY HỌC (Bảng 4 hoạt động 5512)      │
│     - HĐ 1: Mở đầu / Khởi động                         │
│     - HĐ 2: Hình thành kiến thức mới                   │
│     - HĐ 3: Luyện tập                                  │
│     - HĐ 4: Vận dụng                                   │
├────────────────────────────────────────────────────────┤
│  IV. BẢNG TỔNG HỢP NĂNG LỰC SỐ (NLS) TÍCH HỢP          │
├────────────────────────────────────────────────────────┤
│  V. KHỐI PHÊ DUYỆT (Duyệt Chuyên môn | Duyệt Tổ CM)   │
└────────────────────────────────────────────────────────┘
```

---

## 2. QUY CHUẨN BẢNG TIÊU ĐỀ (TABLE 0 - HEADER BOX)

Để tránh lỗi **"Họ và tên giáo viên: Trần Đức Tình bị thọt thành 2 dòng"**, bảng tiêu đề đầu bài phải tuân thủ nghiêm ngặt kích thước sau:

### Tỷ lệ kích thước cột (Khổ A4 đứng, lề trái 2.0 cm, lề phải 1.5 cm):
- **Tổng chiều rộng vùng in**: `~490pt - 496pt`
- **Cột trái (`c0`)**: Chiều rộng **`190pt`** (`3800 dxa`).
- **Cột phải (`c1`)**: Chiều rộng **`300pt`** (`6000 dxa`) $\rightarrow$ *đảm bảo tên giáo viên dài đến 35 ký tự vẫn nằm trọn vẹn trên 1 dòng duy nhất*.

### Bảng minh họa cấu trúc:
| Cột trái (Rộng 190pt) | Cột phải (Rộng 300pt) |
| :--- | :--- |
| **Tuần**: [Số tuần, VD: 1 hoặc 1 - 2]<br>**Tiết**: [Số tiết PPCT, VD: 1 hoặc 1 - 3]<br>*Ngày dạy: [Khoảng ngày theo PPCT]* | **Họ và tên giáo viên**: **Trần Đức Tình** *(Chắc chắn 1 dòng)*<br>**Tổ chuyên môn**: **Tổ Tự Nhiên** |

- **Font chữ**: Times New Roman.
- **Cỡ chữ**: 12pt (Riêng dòng Ngày dạy in nghiêng: 11.5pt).
- **Giãn dòng**: 1.15 line spacing, `space_before = 0pt`, `space_after = 2pt`.
- **Khóa ngắt trang**: Hàng này phải gán thuộc tính `<w:cantSplit/>`.

---

## 3. CẤU TRÚC CHI TIẾT 5 PHẦN CỦA KHBD

### TÊN BÀI DẠY
- Căn giữa, in hoa đậm, màu xanh đậm (`#003366` hoặc `RGBColor(0, 51, 102)`), cỡ chữ 14.5pt.
- Dòng phụ chú: `Môn học: ...; Lớp: ... | Thời gian thực hiện: ... tiết` (in nghiêng, căn giữa, 12pt).

---

### I. MỤC TIÊU
#### 1. Về kiến thức:
- Liệt kê các chuẩn kiến thức cốt lõi học sinh chiếm lĩnh được theo yêu cầu cần đạt của SGK/CT GDPT 2018.

#### 2. Về năng lực:
- **Năng lực chung**: Tự chủ và tự học, Giao tiếp và hợp tác, Giải quyết vấn đề và sáng tạo.
- **Năng lực đặc thù**: Năng lực nhận thức tin học/khoa học/thể chất, thực hành kỹ năng chuyên biệt.
- **Năng lực số (NLS) tích hợp theo Thông tư 02/2025/TT-BGDĐT** *(Bắt buộc in đậm chữ đỏ - `#FF0000`)*:
  ```text
  2.3. Năng lực số (NLS) tích hợp theo Thông tư 02/2025/TT-BGDĐT:
  - [Mã_Chỉ_Số_1]: Mô tả hoạt động số học sinh thực hiện (VD: [1.2.TC2a]: Sử dụng phần mềm trình chiếu/bảng tính để xử lý dữ liệu học tập).
  - [Mã_Chỉ_Số_2]: Mô tả hoạt động số học sinh thực hiện (VD: [2.1.TC2a]: Tìm kiếm và trích dẫn thông tin an toàn từ Internet).
  ```

#### 3. Về phẩm chất:
- Rèn luyện các phẩm chất cốt lõi: Yêu nước, Nhân ái, Chăm chỉ, Trung thực, Trách nhiệm.

---

### II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
- **Giáo viên**: Kế hoạch bài dạy, bài giảng điện tử (Canva, PowerPoint), máy tính kết nối Internet, máy chiếu/Tivi, phiếu học tập số (Google Forms, Padlet, Liveworksheets).
- **Học sinh**: Sách giáo khoa, vở ghi, thiết bị số cá nhân/phòng máy, tài liệu nhóm.

---

### III. TIẾN TRÌNH DẠY HỌC (CHUẨN 5512)
Gồm 4 hoạt động cơ bản. Mỗi hoạt động trình bày bằng **Bảng 2 cột**:
- **Cột 1 (Rộng 310pt)**: `Hoạt động của GV và HS`
- **Cột 2 (Rộng 190pt)**: `Dự kiến sản phẩm`

#### Cấu trúc 4 bước trong mỗi hoạt động:
1. **Bước 1: Chuyển giao nhiệm vụ**: GV giao nhiệm vụ học tập rõ ràng (nếu có ứng dụng số thì bôi đỏ).
2. **Bước 2: Thực hiện nhiệm vụ**: HS làm việc cá nhân/nhóm, tra cứu, thao tác.
3. **Bước 3: Báo cáo, thảo luận**: HS trình bày sản phẩm, đại diện nhóm thuyết trình.
4. **Bước 4: Kết luận, nhận định**: GV chuẩn hóa kiến thức, đánh giá quá trình.

> **Quy ước NLS trong tiến trình:** Mọi vị trí học sinh hoặc giáo viên sử dụng thiết bị số, phần mềm, Internet, gửi bài qua mạng... phải được **in chữ đỏ** kèm mã năng lực tương ứng.

---

### IV. BẢNG TỔNG HỢP NĂNG LỰC SỐ (NLS) TÍCH HỢP TRONG BÀI DẠY
Đặt ở cuối phần bài dạy, trước phần khối duyệt:
- Bảng 3 cột:
  - Cột 1: **TT** (Rộng 40pt, căn giữa)
  - Cột 2: **Nội dung hoạt động số của học sinh** (Rộng 320pt)
  - Cột 3: **Mã NLS đạt được** (Rộng 140pt, in đậm chữ đỏ, căn giữa)

*Ví dụ mẫu:*
| TT | Nội dung hoạt động số của học sinh | Mã NLS đạt được |
| :---: | :--- | :---: |
| 1 | HS truy cập liên kết học tập để tìm kiếm thông tin về chủ đề bài học theo hướng dẫn của GV. | **2.1.TC2a**<br>**1.2.TC2a** |
| 2 | Các nhóm tạo bài thuyết trình trên PowerPoint/Canva và chia sẻ qua nhóm Zalo/Google Drive. | **3.2.TC2a**<br>**4.1.TC2a** |

---

### V. KHỐI PHÊ DUYỆT (CHỐNG TÁCH TRANG)
Bảng 1 hàng, 2 cột đặt ở cuối cùng:
- **Cột 1 (Trái)**: Duyệt của Ban Giám hiệu / Phụ trách Chuyên môn.
- **Cột 2 (Phải)**: Duyệt của Tổ trưởng Chuyên môn.

| Duyệt của Chuyên môn<br>*(Ký và ghi rõ họ tên)*<br><br><br><br>**Lưu Văn Chí** | Duyệt của tổ chuyên môn<br>*(Ký và ghi rõ họ tên)*<br><br><br><br>**Hoàng Thị Kim Tuyết** |
| :---: | :---: |

---

## 4. BẢNG THAM CHIẾU KHUNG NĂNG LỰC SỐ (THÔNG TƯ 02/2025/TT-BGDĐT)
*(Dành cho cấp Trung học cơ sở - Tiêu chuẩn 2: TC2)*

| Miền năng lực | Mã chỉ số | Yêu cầu cần đạt / Biểu hiện của học sinh |
| :--- | :---: | :--- |
| **1. Vận hành thiết bị và phần mềm** | **1.1.TC2a** | Xác định các thiết bị phần cứng, phần mềm và dịch vụ số cần thiết cho nhiệm vụ học tập. |
| | **1.1.TC2b** | Vận hành thành thạo thiết bị số thông dụng và khắc phục sự cố kỹ thuật đơn giản. |
| | **1.2.TC2a** | Sử dụng phần mềm chuyên dụng/công cụ số phù hợp để thu thập, xử lý dữ liệu. |
| | **1.2.TC2b** | Cài đặt, tùy biến các ứng dụng học tập số an toàn và đúng mục đích. |
| **2. Thông tin và dữ liệu** | **2.1.TC2a** | Xây dựng từ khóa, áp dụng chiến lược tìm kiếm thông tin hiệu quả trên Internet. |
| | **2.2.TC2a** | Lựa chọn, tổ chức và lưu trữ dữ liệu số khoa học (thư mục, đám mây). |
| | **2.3.TC2a** | Đánh giá độ tin cậy, tính mới, độ chính xác của thông tin số từ nhiều nguồn. |
| | **2.3.TC2b** | Tham gia các hoạt động trao đổi, học tập trên môi trường số có trách nhiệm. |
| **3. Giao tiếp và hợp tác trong môi trường số** | **3.1.TC2a** | Sử dụng các kênh truyền thông số (email, nhóm trực tuyến, diễn đàn) để trao đổi phù hợp. |
| | **3.2.TC2a** | Chia sẻ dữ liệu và học liệu số với bạn học, giáo viên an toàn. |
| | **3.3.TC2a** | Hợp tác làm việc nhóm thông qua các công cụ trực tuyến để hoàn thành sản phẩm số. |
| | **3.4.TC2a** | Tuân thủ các quy tắc ứng xử, văn hóa trên môi trường mạng (Netiquette). |
| **4. Sáng tạo nội dung số** | **4.1.TC2a** | Tạo mới/chỉnh sửa nội dung số (văn bản, slide, sơ đồ tư duy, infographic, video). |
| | **4.2.TC2a** | Tích hợp, tái cấu trúc học liệu số sáng tạo và hợp pháp. |
| | **4.3.TC2a** | Tuân thủ quy định về bản quyền số, sở hữu trí tuệ và trích dẫn nguồn học liệu. |
| | **4.4.TC2a** | Áp dụng tư duy thuật toán, khối lệnh lập trình để giải quyết vấn đề. |
| **5. An toàn và an ninh số** | **5.1.TC2a** | Bảo vệ thiết bị số cá nhân khỏi mã độc và nguy cơ mất an toàn thông tin. |
| | **5.2.TC2a** | Bảo vệ dữ liệu cá nhân, nhận diện các hình thức lừa đảo trên không gian mạng. |
| | **5.3.TC2a** | Bảo vệ sức khỏe thể chất, tinh thần khi dùng thiết bị số (tư thế, thời lượng). |
| | **5.4.TC2c** | Chủ động cập nhật kiến thức, kỹ năng để thích ứng với sự phát triển của công nghệ số. |
| **6. Giải quyết vấn đề trong môi trường số** | **6.1.TC2a** | Nhận diện thiết bị thông minh, tính năng tự động hóa và ứng dụng ban đầu của AI. |
| | **6.2.TC2a** | Sử dụng công nghệ số và công cụ tìm kiếm để giải quyết bài toán/tình huống thực tiễn. |
| | **6.3.TC2a** | Nhận biết khoảng trống năng lực số của bản thân để tự học và nâng cao kỹ năng. |

---

## 5. QUY TẮC NHẬN DIỆN & ĐÁNH DẤU CHỮ ĐỎ NLS

1. **Màu sắc chuẩn**: Mã RGB `(255, 0, 0)` hoặc mã Hex `#FF0000` (Đỏ tươi nổi bật).
2. **Các vị trí BẮT BUỘC có màu đỏ**:
   - Mục `2.3. Năng lực số (NLS)` trong phần I. Mục tiêu.
   - Các câu văn mô tả thao tác học tập có yếu tố số trong phần III. Tiến trình dạy học:
     - *Ví dụ*: `<font color="red">- HS sử dụng máy tính bảng quét mã QR để truy cập phiếu khảo sát trên Google Forms [1.2.TC2a].</font>`
   - Cột **Mã NLS đạt được** trong Bảng tổng hợp NLS ở phần IV.
3. **Nguyên tắc tích hợp tự nhiên**: Không khiên cưỡng. Mỗi tiết học (45 phút) tích hợp từ **2 đến 4 mã chỉ số NLS** phù hợp nhất với nội dung bài học.

---

## 6. QUY CHUẨN KHỐI CHỮ KÝ & KỸ THUẬT CHỐNG TÁCH TRANG (CANTSPLIT)

### 3 lỗi ngắt trang thường gặp & cách khắc phục:
1. **Lỗi 1: Tách đôi bảng ký giữa 2 trang**: Tiêu đề ký ở cuối trang 1, họ tên rơi sang trang 2.
   $\rightarrow$ **Khắc phục**: Gán thuộc tính `<w:cantSplit/>` vào thẻ `<w:trPr>` của tất cả các hàng trong bảng chữ ký.
2. **Lỗi 2: Bảng ký đứng cô độc 1 mình ở trang mới**:
   $\rightarrow$ **Khắc phục**: Gán thuộc tính `keep_with_next = True` cho các đoạn văn bản trong bảng ký và đoạn văn bản liền kề trước bảng.
3. **Lỗi 3: Khoảng cách ký quá lớn làm tràn trang**:
   $\rightarrow$ **Khắc phục**: Thiết lập `space_after = Pt(36)` (~1.3 cm) vừa vặn cho 1 chữ ký đẹp, không dùng các dòng xuống hàng `\n\n\n\n` vô tội vạ.

---

## 7. SYSTEM PROMPT & PROMPT MẪU CHO ỨNG DỤNG AI SINH GIÁO ÁN TỪ SGK

Khi tích hợp vào ứng dụng tạo giáo án khác, sử dụng cấu trúc Prompt 2 tầng như sau:

### SYSTEM PROMPT (Cấu hình Persona và Luật sinh):
```text
Bạn là Chuyên gia Phương pháp Giảng dạy và Soạn giáo án chuẩn Bộ Giáo dục & Đào tạo Việt Nam.
Nhiệm vụ của bạn là nhận nội dung bài học từ Sách giáo khoa (SGK) và biên soạn một Kế hoạch bài dạy (KHBD) hoàn chỉnh, tuân thủ nghiêm ngặt:
1. Cấu trúc Công văn 5512/BGDĐT-GDTrH (4 hoạt động: Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng; Bảng 2 cột Hoạt động GV-HS và Dự kiến sản phẩm).
2. Tích hợp Khung Năng lực số (NLS) theo Thông tư 02/2025/TT-BGDĐT cấp THCS (Tiêu chuẩn 2 - TC2).
3. Đánh dấu [MÀU ĐỎ] hoặc ghi chú rõ [NLS - Mã_Chỉ_Số] cho tất cả mục tiêu NLS, hoạt động số của học sinh trong tiến trình và Bảng tổng hợp NLS cuối bài.
4. Bảng thông tin đầu bài (Table 0):
   - Cột trái: Tuần, Tiết, Ngày dạy theo phân phối chương trình.
   - Cột phải: "Họ và tên giáo viên: Trần Đức Tình" (Chắc chắn trên 1 dòng duy nhất) và "Tổ chuyên môn: Tổ Tự Nhiên".
5. Khối duyệt 2 cột:
   - Cột trái: "Duyệt của Chuyên môn" - Người duyệt: "Lưu Văn Chí".
   - Cột phải: "Duyệt của tổ chuyên môn" - Người duyệt: "Hoàng Thị Kim Tuyết".
Tuyệt đối không bỏ sót bất kỳ phần nào trong cấu trúc 5 phần.
```

### USER PROMPT TEMPLATE (Mẫu truyền dữ liệu SGK):
```text
Dựa vào nội dung bài học trong Sách giáo khoa dưới đây, hãy soạn Kế hoạch bài dạy hoàn chỉnh theo đúng quy chuẩn đã đặt ra:

[THÔNG TIN BÀI DẠY]
- Môn học: {MON_HOC} (VD: Tin học 8 / Hoạt động trải nghiệm 8 / Khoa học tự nhiên 6)
- Lớp: {LOP}
- Tên bài dạy: {TEN_BAI_DAY}
- Tuần theo PPCT: {TUAN} (VD: Tuần 5)
- Tiết theo PPCT: {TIET} (VD: Tiết 5 hoặc Tiết 13 - 15)
- Khoảng ngày dạy: {NGAY_DAY} (VD: 05/10/2026 - 09/10/2026)
- Thời lượng: {THOI_LUONG} tiết

[NỘI DUNG SÁCH GIÁO KHOA]
{NOI_DUNG_VAN_BAN_HOAC_TOM_TAT_SGK}

[YÊU CẦU ĐẶC BIỆT VỀ NĂNG LỰC SỐ]
- Hãy chọn lọc 2-4 mã năng lực số phù hợp nhất trong Thông tư 02/2025/TT-BGDĐT (VD: 1.2.TC2a, 2.1.TC2a, 3.2.TC2a, 4.1.TC2a).
- Đưa các mã này vào mục 2.3 trong I. Mục tiêu.
- Thể hiện rõ hoạt động số của học sinh trong Bước 1 và Bước 2 của các hoạt động dạy học.
- Lập Bảng tổng hợp NLS (IV) ở cuối bài.
```

---

## 8. CODE MẪU PYTHON (DOCX-PYTHON) CHUẨN HÓA ĐỊNH DẠNG TỰ ĐỘNG

Dưới đây là module Python chuẩn (`khbd_styler.py`) có thể copy trực tiếp vào ứng dụng khác để xuất file Word `.docx` hoàn hảo:

```python
# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

COLOR_RED = RGBColor(255, 0, 0)
COLOR_DARK_BLUE = RGBColor(0, 51, 102)

def set_cell_width_dxa(cell, width_pt):
    cell.width = Pt(width_pt)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(width_pt * 20)))
    tcW.set(qn('w:type'), 'dxa')

def make_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None:
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def create_table0_header(doc, week_str, period_str, date_str, teacher_name="Trần Đức Tình", to_cm="Tổ Tự Nhiên"):
    """Tạo bảng tiêu đề chuẩn xác, tên GV 1 dòng duy nhất"""
    tbl0 = doc.add_table(rows=1, cols=2)
    tbl0.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl0.autofit = False
    make_row_cant_split(tbl0.rows[0])
    
    c0 = tbl0.cell(0, 0)
    c1 = tbl0.cell(0, 1)
    set_cell_width_dxa(c0, 190) # Cột trái 190pt
    set_cell_width_dxa(c1, 300) # Cột phải 300pt (rộng rãi, chống rớt dòng)

    # Cột trái: Tuần, Tiết, Ngày dạy
    p0 = c0.paragraphs[0]
    p0.paragraph_format.line_spacing = 1.15
    p0.paragraph_format.space_after = Pt(2)
    p0.add_run("Tuần: ").font.size = Pt(12)
    r_w = p0.add_run(week_str)
    r_w.bold = True
    r_w.font.size = Pt(12)

    p1 = c0.add_paragraph()
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.space_after = Pt(2)
    p1.add_run("Tiết: ").font.size = Pt(12)
    r_p = p1.add_run(period_str)
    r_p.bold = True
    r_p.font.size = Pt(12)

    p2 = c0.add_paragraph()
    p2.paragraph_format.line_spacing = 1.15
    p2.paragraph_format.space_after = Pt(0)
    p2.add_run("Ngày dạy: ").font.size = Pt(11.5)
    r_d = p2.add_run(date_str)
    r_d.italic = True
    r_d.font.size = Pt(11.5)

    # Cột phải: Họ và tên GV (1 dòng duy nhất) & Tổ chuyên môn
    p0_1 = c1.paragraphs[0]
    p0_1.paragraph_format.line_spacing = 1.15
    p0_1.paragraph_format.space_after = Pt(2)
    p0_1.paragraph_format.keep_with_next = True
    p0_1.add_run("Họ và tên giáo viên: ").font.size = Pt(12)
    r_t = p0_1.add_run(teacher_name)
    r_t.bold = True
    r_t.font.size = Pt(12)

    p1_1 = c1.add_paragraph()
    p1_1.paragraph_format.line_spacing = 1.15
    p1_1.paragraph_format.space_after = Pt(0)
    p1_1.add_run("Tổ chuyên môn: ").font.size = Pt(12)
    r_to = p1_1.add_run(to_cm)
    r_to.bold = True
    r_to.font.size = Pt(12)

def create_signature_block(doc, signer_cm="Lưu Văn Chí", signer_to="Hoàng Thị Kim Tuyết"):
    """Tạo khối phê duyệt 2 cột chống tách trang tuyệt đối"""
    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(8)
    p_sig.paragraph_format.keep_with_next = True

    tbl_s = doc.add_table(rows=1, cols=2)
    tbl_s.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_s.autofit = False
    make_row_cant_split(tbl_s.rows[0])
    
    c0 = tbl_s.cell(0, 0)
    c1 = tbl_s.cell(0, 1)
    set_cell_width_dxa(c0, 245)
    set_cell_width_dxa(c1, 245)

    # Cột trái: Duyệt của Chuyên môn - Lưu Văn Chí
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_after = Pt(2)
    p0.paragraph_format.keep_with_next = True
    r = p0.add_run("Duyệt của Chuyên môn\n")
    r.bold = True
    r.font.size = Pt(13)

    p1 = c0.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(36) # Khoảng trống ký tay chuẩn
    p1.paragraph_format.keep_with_next = True
    r = p1.add_run("(Ký và ghi rõ họ tên)")
    r.italic = True
    r.font.size = Pt(11)

    p2 = c0.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(0)
    r = p2.add_run(signer_cm)
    r.bold = True
    r.font.size = Pt(13)

    # Cột phải: Duyệt của tổ chuyên môn - Hoàng Thị Kim Tuyết
    p0_1 = c1.paragraphs[0]
    p0_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0_1.paragraph_format.space_after = Pt(2)
    p0_1.paragraph_format.keep_with_next = True
    r = p0_1.add_run("Duyệt của tổ chuyên môn\n")
    r.bold = True
    r.font.size = Pt(13)

    p1_1 = c1.add_paragraph()
    p1_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1_1.paragraph_format.space_after = Pt(36)
    p1_1.paragraph_format.keep_with_next = True
    r = p1_1.add_run("(Ký và ghi rõ họ tên)")
    r.italic = True
    r.font.size = Pt(11)

    p2_1 = c1.add_paragraph()
    p2_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2_1.paragraph_format.space_after = Pt(0)
    r = p2_1.add_run(signer_to)
    r.bold = True
    r.font.size = Pt(13)
```

---

## 9. TỔNG KẾT CHECKLIST KIỂM TRA TRƯỚC KHI XUẤT BẢN

Trước khi bàn giao hoặc trình ký giáo án, ứng dụng cần tự động chạy kiểm tra checklist:
- [x] **Table 0**: Tỷ lệ `190pt / 300pt`, Họ và tên GV `Trần Đức Tình` nằm trên đúng **1 dòng**.
- [x] **Tổ chuyên môn**: Ghi đúng `Tổ Tự Nhiên` (đồng bộ cho cả Tin học, GDTC và HĐTN).
- [x] **Thời gian**: Tuần, Tiết, Ngày dạy khớp 100% với file PPCT và Lịch tuần của nhà trường.
- [x] **NLS**: Có mục `2.3 Năng lực số` bôi đỏ, có gắn mã `[X.X.TC2x]` trong hoạt động dạy học, có Bảng tổng hợp NLS ở phần IV.
- [x] **Chữ ký**: Cột trái `Lưu Văn Chí` (Duyệt của Chuyên môn), cột phải `Hoàng Thị Kim Tuyết` (Duyệt của tổ chuyên môn).
- [x] **Chống tách trang**: Bảng chữ ký có thuộc tính `cantSplit` và `keep_with_next`, không bị vỡ/cắt đôi giữa 2 trang.
