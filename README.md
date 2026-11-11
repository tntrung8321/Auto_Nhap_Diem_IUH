# Hệ Thống Chấm, Kiểm Duyệt & Nhập Điểm Writing IUH (`IUH_Writing_AI_Suite`)

Ứng dụng Desktop độc lập ([`IUH_Writing_AI_Suite.exe`](file:///d:/IUH/Auto_Nhap_Diem_IUH/IUH_Writing_AI_Suite.exe)) hỗ trợ giảng viên tiếng Anh tại IUH tự động hóa quy trình chấm thi viết (**English for Beginners — Writing Part 1 & Part 2**) trên **1 màn hình làm việc duy nhất (Apple Minimalist Single-Screen Workspace)** kết hợp **Hội Đồng Chấm 2 Vòng Đa Mô Hình (`gemini-3.8-flash ⇄ gemini-3.7-flash`)** và giảng viên kiểm duyệt ở giữa (**Human-in-the-Loop**).

---

## 1. Khởi Động Nhanh (Không cần cài đặt Python)

- Tải về và nhấp đúp trực tiếp vào file **[`IUH_Writing_AI_Suite.exe`](file:///d:/IUH/Auto_Nhap_Diem_IUH/IUH_Writing_AI_Suite.exe)** ngay tại thư mục gốc dự án.
- Ứng dụng tự động tạo/khôi phục phiên làm việc gần nhất (các lớp đang mở, điểm 8 cột `C - O - G - V` đã chấm/duyệt, cỡ chữ đã chọn) từ `config.json` và file `<MãLớp>_review.json`.

---

## 2. Quy Trình Làm Việc Trên 1 Màn Hình Duy Nhất

### Bước 1: Hút đề & bài làm của lớp từ Moodle (`F12`)
1. Mở trang Bảng điểm lớp trên trình duyệt (`luyenthinn.iuh.edu.vn`).
2. Trên thanh công cụ trên cùng của phần mềm, bấm nút **`Hút dữ liệu`** (mã JavaScript 8 luồng tự động được sao chép vào Clipboard).
3. Trên trình duyệt, nhấn `F12` $\rightarrow$ chọn thẻ **Console** $\rightarrow$ dán (`Ctrl + V`) và nhấn `Enter`.
4. Sau vài giây, trình duyệt tự động tải về file `bai_lam_iuh_<MãLớp>.json` (chứa Mã đề, Đề bài thực tế, Hình ảnh đề thi và Bài làm nguyên bản của cả lớp).

### Bước 2: Mở Lớp, Cấu Hình Hội Đồng Chấm 2 Vòng & Chấm Tự Động
1. Bấm vào nút **`Lớp: ...`** ở góc trái trên cùng để mở **Menu Lớp Dọc Nổi (Floating Popover)**:
   - Chọn nhanh các lớp có sẵn trong thư mục hoặc bấm **`Mở thêm lớp học mới (Chọn JSON hoặc Excel)...`**.
   - Mỗi lớp có bộ nhớ phiên độc lập (`_review.json`) và chốt chặn chống ghi đè nhầm file gốc giữa các lớp.
2. Bấm nút **`Cài đặt`** để:
   - Nhập khóa kết nối Gemini API (hỗ trợ nhập nhiều API Key cách nhau bởi dấu phẩy `,` hoặc chấm phẩy `;` để tự động xoay vòng khi hết hạn mức Free Tier).
   - Chọn Model Giám khảo Vòng 1 (mặc định `gemini-3.8-flash` · `High Thinking`) và Model Trọng tài Vòng 2 (mặc định `gemini-3.7-flash`, hỗ trợ `gemini-3.1-pro-preview`, `gemini-4-argon`...).
   - Chọn đúng **Bộ quy tắc chấm theo Sách & Đợt thi**:
     - **`Cuối kì - Sách Cambridge`** ([`MARKING_RULES_CAMBRIDGE_CUOI_KI.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/MARKING_RULES_CAMBRIDGE_CUOI_KI.md) — `Phiên bản v7 Calibrated on Gold Standard`).
     - **`Giữa kì - Sách Cambridge`** (`MARKING_RULES_CAMBRIDGE_GIUA_KI.md`).
     - **`Giữa kì - Sách Oxford`** (`MARKING_RULES_OXFORD_GIUA_KI.md`).
     - **`Cuối kì - Sách Oxford`** (`MARKING_RULES_OXFORD_CUOI_KI.md`).
3. Bấm nút **`Chấm tự động`** (hoặc bấm **`Tiến trình`**) để theo dõi **Thanh Tiến Trình Nổi (Floating Progress Bar)** chạy liền mạch từ `0%` $\rightarrow$ `100%` qua 5 mốc:
   - **`Phân đề`** *(Gom ngầm 20 mã đề & bài mẫu chuẩn Gold Standard trong bộ nhớ — giữ nguyên 100% thứ tự dòng Excel của sinh viên)*
   - **`Chấm vòng 1`** *(Chấm 4 cụm đề song song với `High Thinking`)*
   - **`Đối chiếu chuẩn`** *(Quét 12 quy tắc cứng của Khoa bằng Python Rule Engine)*
   - **`Thẩm định chéo`** *(Model thứ 2 độc lập phúc khảo các bài nhạy cảm/lệch điểm — chỉ tốn tổng cộng tối đa 5 lượt gọi API / 1 lớp)*
   - **`Hoàn tất`** *(Trích dẫn dẫn chứng nguyên văn `"..."` & trả điểm đúng thứ tự sinh viên)*

### Bước 3: Kiểm Duyệt Trực Quan & Sửa Điểm Tốc Độ Cao (`C - O - G - V`)
- **Hiển thị song song 2 tầng (Part 1 ở trên, Part 2 ở dưới):**
  - Bên trái (`~76%` diện tích): **Đề bài** (tự động in đậm các câu hỏi `?` và gạch đầu dòng yêu cầu chính) đặt cạnh **Bài làm nguyên bản của sinh viên**.
  - Bên phải (`~24%` diện tích): 4 ô điểm **`C - O - G - V`** (`0–5`) kèm **Lý do Đạt / Lỗi cụ thể có trích dẫn `"..."`** và nhãn xác thực hội đồng (`Thẩm định 2 vòng (3.8-flash ⇄ 3.7-flash)` hoặc `Chuẩn hội đồng (3.8-flash + Rule Engine)`).
- **Tổ hợp phím tắt kiểm duyệt 100% không cần chuột:**
  - `↑` / `↓`: Chuyển sinh viên (khi ở danh sách trái) hoặc di chuyển dọc qua 8 ô điểm (`P1: C, O, G, V` $\leftrightarrow$ `P2: C, O, G, V`).
  - `←` / `→`: Chuyển qua lại giữa Danh sách sinh viên bên trái (`←`) và Cột điểm bên phải (`→`).
  - Phím số `0` – `5`: Ghi đè trực tiếp con điểm mới vào ô đang chọn và tính lại tức thì điểm Part 1, Part 2, Điểm Tổng (khớp 100% công thức `=MROUND(I+O, 0.5)` của Excel).
  - `Space`: Đánh dấu `✓` đã duyệt bài sinh viên hiện tại và tự động chuyển xuống sinh viên tiếp theo.

### Bước 4: Xuất Điểm Ra Excel & Đồng Bộ Hai Chiều Với Giảng Viên 2 (GV2)
- Bấm nút **`Lưu điểm ra Excel`** ở góc dưới bên trái:
  - Ghi chính xác 8 cột điểm `C - O - G - V` vào file Excel Phiếu chấm, giữ nguyên 100% công thức tính tổng (`H, I, N, O, P`) và **bảo toàn nguyên vẹn cột `Điểm tổng GV2` (cột `Q`) cùng `Điểm thống nhất` (cột `R`)**.
- **Khi cần sửa điểm trong file Excel để khớp với GV2:**
  - Sau khi sửa và lưu file `.xlsx`, bấm nút **`Nạp lại Excel`** — phần mềm tự động nhận diện file Excel mới hơn và đồng bộ ngược điểm mới nhất từ Excel vào phần mềm.

### Bước 5: Nhập Điểm Chạy Ngầm Lên Moodle (`JS F12` — Nhịp 3–5s/câu, Bảo Vệ 2 Tầng, Tách Riêng Từng Lớp)
1. Bấm nút **`Sao chép mã nhập Moodle`** ở góc dưới bên trái (chỉ đóng gói điểm của đúng lớp đang mở trên màn hình).
2. Mở trang Bảng điểm Moodle của lớp đó trên trình duyệt $\rightarrow$ nhấn `F12` (thẻ **Console**) $\rightarrow$ dán (`Ctrl + V`) và nhấn `Enter`.
3. **Cơ chế Bảo vệ 2 Tầng tự động:**
  - Khớp chính xác từng dòng `<tr>` theo `MSSV` và lấy đúng link `slot=17` (`Q17`) / `slot=18` (`Q18`), **tuyệt đối không bao giờ bị nhảy dòng hay lệch sinh viên**.
  - **Tự động phát hiện & bỏ qua (`0 giây`, không ghi đè)** nếu giữa chừng có sinh viên đã được người khác chấm `Q17` hoặc `Q18` trước đó.

---

## 3. Cấu Trúc Thư Mục Chuẩn

- [`IUH_Writing_AI_Suite.exe`](file:///d:/IUH/Auto_Nhap_Diem_IUH/IUH_Writing_AI_Suite.exe): File chạy chính duy nhất của hệ thống (đã đóng gói sẵn Python & thư viện).
- [`MARKING_RULES_CAMBRIDGE_CUOI_KI.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/MARKING_RULES_CAMBRIDGE_CUOI_KI.md): Bộ quy tắc chấm chuẩn hóa cho **Cuối kì — Sách Cambridge** (`v7`).
- [`PROJECT_MEMORY.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/PROJECT_MEMORY.md): Bộ nhớ kiến trúc và lịch sử nâng cấp toàn diện của dự án (Giai đoạn 1 $\rightarrow$ Giai đoạn 11).
- [`src/app_gui.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py), [`src/auto_iuh.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/auto_iuh.py), [`src/test_auto_iuh.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/test_auto_iuh.py), [`src/build_exe.bat`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/build_exe.bat): Mã nguồn, bộ 24 kịch bản kiểm thử tự động và script đóng gói `.exe`.
