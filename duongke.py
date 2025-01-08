from paddleocr import PaddleOCR
from tkinter import Tk, filedialog, ttk, Button, messagebox
from threading import Thread
from utils_dk import extractText, group_and_sort_cells_by_row, write_to_excel

class ImageToExcelConverter:
    def __init__(self, ocr):
        self.ocr = ocr
        self.progress_bar = None

    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.progress_bar.update()

    def run_conversion(self, image_path, excel_name):
        try:
            progress = 0
            self.update_progress(progress)
            
            # Bước 1: Đọc ảnh và xử lý OCR
            progress += 30
            self.update_progress(progress)
            cells, results = self.extract_text(image_path)
            
            # Bước 2: Nhóm và sắp xếp các ô theo hàng
            progress += 40
            self.update_progress(progress)
            rows = self.group_and_sort_cells_by_row(cells, results)
            
            # Bước 3: Ghi dữ liệu vào file Excel
            progress += 30
            self.update_progress(progress)
            self.write_to_excel(rows, output_file=excel_name)
            
            # Thông báo thành công
            messagebox.showinfo("Thành công", f"Chuyển đổi hoàn tất! File lưu tại: {excel_name}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")
        finally:
            self.update_progress(100)

    def extract_text(self, image_path):
        # Gọi hàm OCR để lấy dữ liệu
        return extractText(image_path, self.ocr)
    
    def group_and_sort_cells_by_row(self, cells, results):
        # Giả sử đây là hàm nhóm và sắp xếp ô
        return group_and_sort_cells_by_row(cells, results)
    
    def write_to_excel(self, rows, output_file):
        # Gọi hàm ghi dữ liệu vào Excel
        write_to_excel(rows, output_file=output_file)

    def choose_file_and_run(self):
        # Chọn file hình ảnh
        image_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
        if not image_path:
            return  # Người dùng đóng hộp thoại mà không chọn file
        
        # Yêu cầu người dùng nhập tên file đầu ra
        excel_name = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not excel_name:
            return  # Người dùng đóng hộp thoại mà không nhập tên
        
        # Chạy xử lý trong thread riêng
        Thread(target=self.run_conversion, args=(image_path, excel_name)).start()

    def create_gui(self):
        # Tạo cửa sổ chính
        root = Tk()
        root.title("Chuyển đổi hình ảnh sang Excel")
        
        # Kích thước cửa sổ
        window_width = 500
        window_height = 200

        # Lấy kích thước màn hình
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # Tính toán vị trí x, y để đặt cửa sổ ở giữa
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)

        # Cập nhật vị trí và kích thước cửa sổ
        root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # Nhãn tiêu đề
        title_label = ttk.Label(root, text="Chuyển đổi hình ảnh sang Excel", font=("Arial", 16))
        title_label.pack(pady=10)
        
        # Nút chọn file
        select_button = Button(root, text="Chọn file hình ảnh", command=self.choose_file_and_run, font=("Arial", 12))
        select_button.pack(pady=10)
        
        # Thanh tiến độ
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=20)
        
        # Chạy giao diện
        root.mainloop()



# Khởi tạo OCR và chạy giao diện
ocr = PaddleOCR(use_angle_cls=False, lang='ch')
converter = ImageToExcelConverter(ocr)
converter.create_gui()
