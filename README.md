# Công cụ Tự động Nhập điểm Writing IUH (Đa trình duyệt)

Công cụ hỗ trợ giảng viên tự động nhập điểm Writing (`Q17 - Part 1` và `Q18 - Part 2`) từ file Excel lên **trang luyenthinn** trên mọi trình duyệt nhân Chromium (**Google Chrome, Microsoft Edge, Cốc Cốc, Brave**).

---

## 📥 Tải về bộ công cụ (Dùng ngay không cần cài Python)

👉 **[BẤM VÀO ĐÂY ĐỂ TẢI FILE `Share_Auto_Diem_IUH.zip`](https://github.com/tntrung8321/Auto_Nhap_Diem_IUH/releases/latest/download/Share_Auto_Diem_IUH.zip)**

*(Hoặc bấm nút màu xanh **`<> Code`** ở góc trên bên phải $\rightarrow$ chọn **`Download ZIP`**)*.

---

## 🚀 Hướng dẫn sử dụng (Chỉ 3 bước)

### Bước 1: Chuẩn bị
1. Sau khi tải file `.zip` về máy, nhấn chuột phải chọn **Giải nén (Extract)** ra một thư mục bình thường trên máy (không bấm chạy trực tiếp bên trong cửa sổ file nén).
2. Tạo một **bản Copy** của file Excel chấm điểm Writing (nếu có sinh viên nào vắng thi hoàn toàn không có tên trên web thì xóa dòng của sinh viên đó trong bản Copy).
   - **Tự động nhận diện cột & dòng:** Công cụ tự động tìm đúng cột `MSSV`, `Họ tên`, điểm `Q17 (Part 1)`, điểm `Q18 (Part 2)` và dòng sinh viên đầu tiên dù file Excel của mỗi giảng viên lệch cột hay khác dòng bắt đầu.
3. Mở sẵn **trang luyenthinn** của lớp trên trình duyệt (**Chrome, Edge, Cốc Cốc, Brave** đều được).
   - ⚠️ **Lưu ý quan trọng:** Giao diện **trang luyenthinn** phải để ngôn ngữ **Tiếng Việt** (để hiển thị các chữ *"Yêu cầu chấm điểm"*, *"Tạo bình luận"*, *"trên 5,00"*).

### Bước 2: Chạy công cụ
- **Cách 1 (Khuyên dùng - Không cần cài Python):**
  Nhấp đúp vào file **`Auto_Nhap_Diem_IUH.exe`**.
  *(Nếu Windows hiện bảng xanh SmartScreen hỏi xác nhận, bấm **More info** $\rightarrow$ **Run anyway**)*.
- **Cách 2 (Dự phòng nếu máy chặn file `.exe`):**
  1. Cài đặt **Python 3.12** từ Microsoft Store (hoặc từ `python.org`, nhớ tích ô *"Add Python to PATH"*).
  2. Nhấp đúp vào file **`1_Cai_Dat_Thu_Vien_Lan_Dau.bat`** (chỉ chạy 1 lần đầu tiên để tải thư viện).
  3. Nhấp đúp vào file **`2_Chay_Auto_Nhap_Diem.bat`** để bắt đầu.

Sau khi cửa sổ chương trình hiện lên:
1. Chọn file Excel điểm (bản Copy) $\rightarrow$ Nhấn **Open**.
2. Kiểm tra thông tin sinh viên đầu tiên được chương trình tự động nhận diện trên màn hình:
   - Nhấn phím **Enter** để bắt đầu chấm từ đầu danh sách.
   - *(Hoặc gõ **Số dòng Excel** / **Mã số sinh viên** nếu muốn chấm tiếp từ giữa danh sách)*.
3. Trong **5 giây** đầu tiên sau khi nhấn Enter: **Bấm chuột 1 cái vào khoảng trắng trên trang luyenthinn** của trình duyệt rồi thả tay ra khỏi chuột và bàn phím.

### Bước 3: Dừng khẩn cấp bất cứ lúc nào
- Khi muốn dừng ngay lập tức: **Vẩy mạnh chuột kịch vào 1 trong 4 góc màn hình** và giữ 1 giây.
- Chương trình sẽ dừng lại ngay, báo rõ đang dừng ở sinh viên nào và lưu toàn bộ lịch sử đã chấm vào file `nhat_ky_cham_diem.txt` nằm cùng thư mục với file Excel của bạn.
