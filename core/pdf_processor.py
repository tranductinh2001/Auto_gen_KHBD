import os
from typing import List, Tuple, Optional
import fitz  # PyMuPDF


class PDFPageProcessor:
    """Xử lý tệp PDF: phân đoạn theo phạm vi trang (start_page, end_page) và chuyển đổi sang ảnh chất lượng cao."""

    @staticmethod
    def convert_pdf_to_images(
        pdf_path: str,
        temp_dir: str = "temp_images",
        start_page: Optional[int] = None,
        end_page: Optional[int] = None,
        dpi: int = 200
    ) -> List[Tuple[str, int]]:
        """
        Quét và chuyển đổi các trang trong khoảng [start_page, end_page] (1-indexed).
        Nếu start_page hoặc end_page để trống thì tự động quét từ đầu hoặc đến hết sách.
        """
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir, exist_ok=True)

        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            return []

        # Chuẩn hóa dải trang 1-indexed
        s = 1 if (start_page is None or start_page < 1) else start_page
        e = total_pages if (end_page is None or end_page > total_pages) else end_page
        if s > e:
            s, e = e, s

        image_tasks = []
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]

        for page_idx in range(s - 1, e):
            page_num = page_idx + 1
            page = doc.load_page(page_idx)
            pix = page.get_pixmap(dpi=dpi)

            img_filename = f"{base_name}_page_{page_num}.png"
            img_path = os.path.join(temp_dir, img_filename)
            pix.save(img_path)
            image_tasks.append((img_path, page_num))

        return image_tasks
