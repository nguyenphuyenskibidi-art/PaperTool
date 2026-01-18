#  PaperTool

**PaperTool** là một giải pháp phần mềm chuyên dụng nhằm tối ưu hóa quy trình **chuyển đổi, phân tích và quản lý tệp PDF**, tập trung vào **độ chính xác định dạng**, **hiệu suất cao** và **tính ổn định**.

Công cụ được phát triển độc lập, hướng đến người dùng kỹ thuật lẫn người dùng phổ thông cần xử lý PDF chuyên sâu trên Windows.

---

##  Tính năng cốt lõi

###  Chuyển đổi hiệu suất cao
- Chuyển đổi **PDF → PowerPoint (.pptx)** với tốc độ xử lý nhanh  
- Duy trì bố cục gốc, hạn chế sai lệch nội dung khi trình chiếu  

###  Render High-DPI (300 DPI)
- Áp dụng **Matrix 3×3** để render hình ảnh chất lượng cao  
- Khắc phục triệt để lỗi **vỡ font, lệch font, font lạ**  
- Đảm bảo hình ảnh sắc nét khi trình chiếu trên màn hình lớn  

###  Phân tích cấu trúc tài liệu PDF
- Thống kê chi tiết trên từng trang:
  - Số lượng từ
  - Số dòng văn bản
  - Số đối tượng hình ảnh

###  Quản lý PDF tích hợp
- Ghép nhiều PDF thành một tệp duy nhất  
- Tách trang PDF theo nhu cầu  
- Thiết lập **mật khẩu bảo mật** với chuẩn mã hóa  

###  Giao diện đa ngôn ngữ (CLI)
- Hỗ trợ **Tiếng Việt 🇻🇳** và **Tiếng Anh 🇬🇧**  
- Điều khiển hoàn toàn qua **Command Line Interface**

---

##  Hướng dẫn vận hành

PaperTool được đóng gói dưới dạng **tệp thực thi độc lập (.exe)**.

- Không cần cài Python  
- Không cần thư viện phụ trợ  
- Chạy trực tiếp trên Windows  

### Các bước sử dụng

1. Truy cập mục **Releases**
2. Tải phiên bản mới nhất: `PaperTool_v1.0.0.exe`
3. Chạy file và chọn ngôn ngữ:
   - `1` → Tiếng Việt
   - `2` → English
4. **Kéo & thả** file PDF vào cửa sổ chương trình
5. Nhấn **Enter** để bắt đầu xử lý

---

## 🧠 Mã nguồn tham khảo

Đoạn mã sau mô tả **logic cốt lõi** trong quá trình render PDF và chuyển đổi sang PowerPoint:

```python
import fitz
from pptx import Presentation

def process_document(pdf_path, pptx_output):
    """
    Quy trình render tài liệu với độ phân giải cao và chuyển đổi sang PPTX
    """
    presentation = Presentation()
    document = fitz.open(pdf_path)
    
    for page_index in range(len(document)):
        page = document.load_page(page_index)
        
        # Thiết lập ma trận render để đảm bảo chất lượng hình ảnh sắc nét
        render_matrix = fitz.Matrix(3, 3)
        pixmap = page.get_pixmap(matrix=render_matrix)
        
        temp_image = f"page_cache_{page_index}.png"
        pixmap.save(temp_image)
        
        # Khởi tạo slide và chèn hình ảnh vào PowerPoint
        slide_layout = presentation.slide_layouts[6]
        slide = presentation.slides.add_slide(slide_layout)
        slide.shapes.add_picture(
            temp_image,
            0,
            0,
            width=presentation.slide_width,
            height=presentation.slide_height
        )
        
    presentation.save(pptx_output)
    document.close()
