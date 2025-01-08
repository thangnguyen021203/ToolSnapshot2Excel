import cv2
import numpy as np
from openpyxl import Workbook
import os

def Convert2Excel(image_path, excel_path, ocr):
    cells, results = extractText(image_path, ocr)
    rows = group_and_sort_cells_by_row(cells, results)
    write_to_excel(rows, excel_path)

def extractText(img_path, ocr):
    """
    Xử lý ảnh đầu vào, phát hiện các cells theo đường thẳng và đường ngang, và OCR từng ô.

    Args:
        img__path: Đường dẫn ảnh đầu vào.

    Returns:
        cells: Danh sách các tọa độ ô (x, y, w, h).
        results: Danh sách kết quả OCR theo thứ tự.
    """
    # Đọc ảnh
    image = cv2.imread(img_path)

    # Lấy kích thước của ảnh
    height, width = image.shape[:2]

    #Preprocessing ảnh
    # Vẽ đường viền màu đen quanh toàn bộ ảnh (góc trên trái và góc dưới phải)
    top_left = (0, 0)  # Góc trên trái
    bottom_right = (width - 1, height - 1)  # Góc dưới phải
    cv2.rectangle(image, top_left, bottom_right, (0, 0, 0), 2)

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Làm mờ ảnh bằng Gaussian để giảm nhiễu trước khi làm nét
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    denoised = cv2.fastNlMeansDenoising(blurred, None, 10, 7, 21)
    sharpened = cv2.addWeighted(gray, 1.5, denoised, -0.5, 0)
    clahe = cv2.createCLAHE(clipLimit=75.0, tileGridSize=(8, 8))
    sharpened = clahe.apply(sharpened)
    
    # Ngưỡng nhị phân thích nghi
    binary = cv2.adaptiveThreshold(sharpened, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 49, 1)

    # Làm nổi bật đường kẻ
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (100, 1))  # Kernel ngang
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 100))   # Kernel dọc

    # Tìm số đường dọc và ngang
    horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    contours_horizontal, _ = cv2.findContours(horizontal_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    num_horizontal_lines = len(contours_horizontal)
    vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)
    contours_vertical, _ = cv2.findContours(vertical_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    num_vertical_lines = len(contours_vertical)
    # print(num_horizontal_lines, num_vertical_lines) để dành debug

    # Phát hiện các đường ngang và dọc
    horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
    # Kết hợp các đường ngang và dọc
    table_lines = cv2.add(horizontal_lines, vertical_lines)

    # Dilate để làm đầy các khoảng trống nhỏ
    table_lines = cv2.dilate(table_lines, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)), iterations=1)

    # Tìm contours để phát hiện ô
    contours, _ = cv2.findContours(table_lines, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    cells = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if 20 < w < 1000 and 20 < h < 1000:  # Lọc các vùng nhỏ/quá lớn
            cells.append((x, y, w, h))
    
    # Loại bỏ các ô chứa các ô khác
    filtered_cells = []
    for i, (x1, y1, w1, h1) in enumerate(cells):
        is_container = False
        for j, (x2, y2, w2, h2) in enumerate(cells):
            if i != j:
                # Kiểm tra nếu ô (x1, y1, w1, h1) chứa ô (x2, y2, w2, h2)
                if x1 <= x2 and y1 <= y2 and (x1 + w1) >= (x2 + w2) and (y1 + h1) >= (y2 + h2):
                    is_container = True
                    break
        if not is_container:
            filtered_cells.append((x1, y1, w1, h1))

    # Sắp xếp các ô (theo hàng trước, cột sau)
    cells = sorted(filtered_cells, key=lambda c: (c[1], c[0]))  # Dùng hàng trước

    # Hiển thị các ô đã phát hiện (Debugging)
    debug_image = image.copy()
    for (x, y, w, h) in cells:
        cv2.rectangle(debug_image, (x, y), (x + w, y + h), (0, 500, 0), 1)

    # Tạo folder lưu debug image
    current_dir = os.path.dirname(os.path.abspath(__file__))
    debug_folder = os.path.join(current_dir, 'debug_detected_cells')
    if not os.path.exists(debug_folder):
        os.makedirs(debug_folder)
    file_name_without_ext = os.path.splitext(os.path.basename(img_path))[0]
    debug_image_path = os.path.join(debug_folder, f'{file_name_without_ext}_debug.jpg')
    cv2.imwrite(debug_image_path, debug_image)

    # OCR từng ô
    results = []
    for (x, y, w, h) in cells:
        # Cắt từng ô (không padding để tránh chạm vào ô khác)
        cell_img = gray[y:y+h, x:x+w]

        # Làm mịn ảnh để giảm nhiễu
        cell_img = cv2.medianBlur(cell_img, 3)

        # Chuyển sang BGR cho PaddleOCR
        cell_img = cv2.cvtColor(cell_img, cv2.COLOR_GRAY2BGR)
        
        # OCR ô hiện tại
        ocr_result = ocr.ocr(cell_img, cls=False)
        
        # Lấy text từ kết quả OCR
        if ocr_result and isinstance(ocr_result[0], list):  # Kiểm tra kết quả hợp lệ
            text = ''.join([line[1][0] for line in ocr_result[0] if line and line[1]])  # Lọc kết quả hợp lệ
        else:
            text = ''  # Nếu ô trống
        
        results.append(text)  # Lưu kết quả, ô trống thì là ''
    return cells, results


def group_and_sort_cells_by_row(cells, results, row_threshold=20):
    """
    Nhóm các cells thành các hàng, liên kết với results và sắp xếp từng hàng theo trục x.

    Args:
        cells: Danh sách các tọa độ ô (x, y, w, h).
        results: Danh sách kết quả OCR theo thứ tự.
        row_threshold: Ngưỡng khoảng cách y để coi các ô thuộc cùng một hàng.

    Returns:
        List[List[Tuple, str]]: Danh sách các hàng với các cells và giá trị OCR được sắp xếp.
    """
    # Bước 1: Sắp xếp cells và liên kết với results
    cells_with_results = sorted(
        zip(cells, results), key=lambda c: (c[0][1], c[0][0])
    )  # Sắp xếp theo y trước, x sau

    rows = []
    current_row = []

    for (cell, result) in cells_with_results:
        x, y, w, h = cell  # Giải nén cell
        if not current_row:
            current_row.append((cell, result))
        else:
            # Lấy tọa độ y của ô cuối cùng trong hàng hiện tại
            last_cell, _ = current_row[-1]
            _, last_y, _, last_h = last_cell
            if abs(y - last_y) < row_threshold:
                current_row.append((cell, result))
            else:
                # Thêm hàng cũ vào danh sách và tạo hàng mới
                rows.append(current_row)
                current_row = [(cell, result)]

    # Thêm hàng cuối cùng
    if current_row:
        rows.append(current_row)

    # Bước 2: Sắp xếp từng hàng theo tọa độ x
    for row in rows:
        row.sort(key=lambda c: c[0][0])  # Sắp xếp theo x

    return rows



def write_to_excel(rows, output_file="duongke.xlsx"):
    """
    Điền dữ liệu OCR vào Excel.

    Args:
        rows: Danh sách các hàng, mỗi hàng là danh sách các cells và giá trị OCR.
        output_file: Đường dẫn file Excel đầu ra.
    """
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active

    # Điền dữ liệu vào từng ô Excel
    for row_idx, row in enumerate(rows):
        for col_idx, item in enumerate(row):
            # Đảm bảo item là một tuple (cell, value)
            if isinstance(item, tuple) and len(item) == 2:
                cell, value = item
                ws.cell(row=row_idx + 1, column=col_idx + 1, value=value)
            else:
                raise ValueError(f"Invalid row item format: {item}")

    wb.save(output_file)
    # print("Convert to Excel successfully!") để dành debug

