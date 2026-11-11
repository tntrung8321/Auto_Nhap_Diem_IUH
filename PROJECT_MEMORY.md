# BỘ NHỚ DỰ ÁN & LỊCH SỬ CẬP NHẬT (PROJECT MEMORY & CHANGELOG)
**Dự án:** `IUH Writing Grading Suite` (`d:\IUH\Auto_Nhap_Diem_IUH`)  
**Cập nhật lần cuối:** `2026-11-07`

---

## 1. Cấu Trúc Các File Lưu Trữ & Bộ Nhớ Trong Dự Án

1. **[`PROJECT_MEMORY.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/PROJECT_MEMORY.md) *(File này)*:**
   - Lưu trữ toàn bộ lịch sử nâng cấp, các quyết định thiết kế giao diện, quy tắc làm tròn điểm Excel, các ca chấm điểm chuẩn (Benchmark cases) và kế hoạch cải tiến tiếp theo để tra cứu xuyên suốt mọi phiên làm việc.
2. **[`MARKING_RULES_CAMBRIDGE_CUOI_KI.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/MARKING_RULES_CAMBRIDGE_CUOI_KI.md) *(Cùng hệ thống Đa Bộ Quy Tắc theo Sách & Đợt thi)*:**
   - Bộ quy tắc chấm viết **Cuối kì — Sách Cambridge** (`v7 - Calibrated on 2607AVCB005 & 2607AVCB006 Gold Standard`), tách bạch rõ ràng với các đợt thi khác (`MARKING_RULES_CAMBRIDGE_GIUA_KI.md`, `MARKING_RULES_OXFORD_GIUA_KI.md`, `MARKING_RULES_OXFORD_CUOI_KI.md`).
3. **[`config.json`](file:///d:/IUH/Auto_Nhap_Diem_IUH/config.json):**
   - Lưu cấu hình người dùng: `api_key`, `model_name`, `rule_profile`, `last_json_path`, `last_excel_path`, `open_classes`, `active_class_code`, `class_bar_visible`, `font_size_pt`.
4. **File phiên kiểm duyệt từng lớp (`<thư_mục_lớp>/bai_lam_iuh_<MãLớp>_review.json`):**
   - Tự động lưu tức thì (`auto-save`) toàn bộ điểm 8 cột `C - O - G - V`, trạng thái tick duyệt (`checked`), và lý do chi tiết mỗi khi giảng viên gõ phím sửa điểm hoặc bấm `Space`.
   - Đã hoàn tất 100% 2 lớp chuẩn vàng (Gold Standard):
     - Lớp `2607AVCB005`: [`2607AVCB005/bai_lam_iuh_2607AVCB005_review.json`](file:///d:/IUH/Auto_Nhap_Diem_IUH/2607AVCB005/bai_lam_iuh_2607AVCB005_review.json) (`34/34` SV).
     - Lớp `2607AVCB006`: [`2607AVCB006/bai_lam_iuh_2607AVCB006_review.json`](file:///d:/IUH/Auto_Nhap_Diem_IUH/2607AVCB006/bai_lam_iuh_2607AVCB006_review.json) (`38/38` SV).
5. **Mã nguồn chính & Kiểm thử:**
   - [`src/app_gui.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py): Toàn bộ giao diện 3 bước, bộ hút JS F12 (`JS_SCRAPER_CODE`), bộ nhập điểm chạy ngầm JS F12 (`build_js_paced_grade_submitter`), thuật toán hậu kiểm (`_calibrate_part_scores`), và tổng hợp lý do (`_synthesize_part_cogv_reasons`).
   - [`src/auto_iuh.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/auto_iuh.py): Tự động nhận diện cấu trúc cột Excel (`detect_excel_schema`, `detect_columns_and_parse`) và điều khiển tự động bàn phím (`cham_mot_cau_web`).
   - [`src/test_auto_iuh.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/test_auto_iuh.py): Bộ 16+ kịch bản unit test tự động kiểm tra toàn diện.

---

## 2. Toàn Bộ Những Gì Đã Thực Hiện & Hoàn Tất

### Giai đoạn 1: Điều hướng bàn phím tốc độ cao tại Tab 2 (Kiểm duyệt & Sửa điểm)
- **Phím `↑` / `↓`:** Chuyển sinh viên (khi đang ở Danh sách SV) hoặc chuyển qua lại giữa 8 ô điểm từ trên xuống dưới (`P1: C → O → G → V` $\rightarrow$ `P2: C → O → G → V`).
- **Phím `←` / `→`:** Chuyển vùng điều khiển giữa **Danh sách SV bên trái** (`←`) và **Cột ô điểm bên phải** (`→`).
- **Gõ trực tiếp `0–5`:** Ghi đè ngay con điểm vào ô đang chọn mà không cần xóa số cũ.
- **Phím `Space`:** Tự động đánh dấu Tick duyệt (`✓`), chuyển sang sinh viên tiếp theo và trả tiêu điểm về danh sách.
- **Khung viền nổi bật:** Ô điểm đang chọn có viền xanh đậm `#2563eb` và nền `#dbeafe`.

### Giai đoạn 2: Sửa lỗi định dạng Đề bài (Prompt) & Hiển thị song song Đề + Bài làm
- **Xử lý đứt dòng đề bài (`clean_prompt_text` & `JS_SCRAPER_CODE`):**
  - Gộp các ký tự xuống dòng mềm (`\n`) nằm giữa câu trong HTML của Moodle (`PART 1: Write\nan email...`, `• what\nsport you like to play`), giữ nguyên các gạch đầu dòng `•` và loại bỏ tiêu đề lặp.
- **Hiển thị Đề bài (Prompt) và Bài làm (Essay) song song cùng lúc:**
  - Chia đôi khung nội dung mỗi Part để giảng viên đối chiếu trực tiếp câu hỏi trên đề và câu trả lời của sinh viên mà không bị lớp phủ che khuất.

### Giai đoạn 3: Giải thích chi tiết từng con điểm `C - O - G - V` & Tự động đồng bộ khi sửa điểm
- Cột **Lý do chi tiết** nêu cụ thể:
  - **Content (`C`):** Số từ, đạt mấy trên tổng số câu hỏi của đề, thiếu câu hỏi nào, lỗi xưng hô với người nhận hay thiếu kết thư/ký tên.
  - **Organization (`O`):** Số lỗi dùng dấu phẩy `,` thay dấu chấm `.` (comma splice), lỗi xuống dòng rời rạc từng câu, hoặc lặp đầu câu.
  - **Grammar (`G`):** Trích dẫn chính xác các cụm sai ngữ pháp trong bài (VD: `"Thank you for reply"`, `"Can you watch the movie title"`, `"I'm interesting my party"`).
  - **Vocabulary (`V`):** Chỉ rõ các từ sai chính tả cụ thể (VD: `"verry"`, `"strees"`, `"busetyful"`) và các cụm từ tốt bù trừ.
- Khi giảng viên gõ phím `0–5` sửa điểm: Lý do tự động tính toán và cập nhật lại tức thì khớp với con điểm mới. Nếu gõ `Content = 0` (bài lạc đề `>= 12` từ), tự động điền `0 - 1 - 1 - 1`.

### Giai đoạn 4: Hiệu chỉnh AI từ góc nhìn Người Nhận Thư (Target Reader) & Chuẩn hóa Xuất Excel
- **Các ca chuẩn hóa thực tế (`2607AVCB005`):**
  - **Lê Nguyễn Tấn Phong (Part 2 — `AVCB.FT.W.P2.L002`):** Đạt 2/3 câu hỏi (giờ/rạp hẹn + ăn Phở/Bánh mì) kèm chào/kết thư, chỉ hỏng câu tên phim $\rightarrow$ Chấm `4 - 3 - 3 - 3` (`3.25đ`).
  - **Đặng Thị Ngọc Hà (Part 2 — `AVCB.FT.W.P2.L005`):** Không chào người nhận, gọi trực tiếp người nhận `Emma` ở ngôi thứ 3 (`"want Emma to bring..."`, `"Emma can't help you prepare..."`) $\rightarrow$ Chấm `1 - 1 - 1 - 1` (`1.0đ`).
  - **Trương Thị Như Ý (Part 2 — `AVCB.FT.W.P2.L005`):** Đề đóng vai chủ tiệc trả lời `Emma`, nhưng học tủ văn mẫu đóng vai người được mời hẹn đi `"HighLands coffee"` uống `"milk tea"` $\rightarrow$ Khóa cứng Lạc đề `0 - 1 - 1 - 1` (`0.75đ`).
  - **Lương Thị Thu Nga (Part 1 & Part 2):** Không bắt nhầm từ đúng `"stress"`; phát hiện chào nhầm tên `"Hi Son"` và ghép câu văn mẫu `"exam next Monday"` ở Part 2 (`3 - 2 - 3 - 3`).
- **Khớp 100% công thức làm tròn `=MROUND(I+O, 0.5)` của Excel:**
  - Dùng công thức `((raw1 + raw2 + 1) // 2) / 2.0` trong `_calc_item_scores` để các mức tổng `.25` và `.75` (`3.25`, `6.25`, `7.25`, `8.25`) luôn làm tròn lên `.5` (`3.5`, `6.5`, `7.5`, `8.5`) giống hệt cột màu vàng trong Excel.
- **Xuất điểm Excel an toàn tuyệt đối:**
  - Ghi đúng 8 cột `C - O - G - V` (Part 1: `D, E, F, G`; Part 2: `J, K, L, M`), bảo toàn 100% công thức Excel ở các cột tổng (`H, I, N, O, P`), có hộp thoại `Retry / Cancel` khi file Excel đang mở.

### Giai đoạn 5: Đa lớp học & Nhập điểm chạy ngầm F12 có nhịp nghỉ tự nhiên 3–5s/câu
- Hỗ trợ gom điểm nhiều lớp vào 1 đoạn mã JS F12 (`build_js_paced_grade_submitter`), chạy tuần tự 1 luồng gửi form `comment.php` với nhịp nghỉ ngẫu nhiên `3.0–5.0s/câu` và hiển thị tiến độ trực tiếp trên trang Moodle.

---

### Giai đoạn 6: Nhúng Tab Lớp vào Topbar (0px Khoảng Trống Chết), Thu Gọn Khu Vực Chấm Điểm (`16 : 5`), Bộ Nhớ Độc Lập Từng Lớp & Khóa Bảo Vệ File Gốc
1. **Nhúng trực tiếp Thanh Tab Lớp Học lên `topbar` (`height=36px`, `0px` khoảng trống thừa):**
   - Xóa bỏ hoàn toàn `self.class_bar_container` từng gây dải màu xám trống khi ẩn thanh lớp.
   - Các tab lớp (`2607AVCB005 (34/34 ✓) ×`, `+ Mở lớp`, `◀ Ẩn tab` / `▶ Hiện tab (N lớp)`) nằm gọn ngay trên thanh tiêu đề tối màu `#0f172a`, tiết kiệm 100% chiều dọc màn hình.
2. **Thu hẹp khu vực Chấm điểm & Báo lỗi (`16 : 5` $\approx$ `76% : 24%`):**
   - Trong [`_build_part_row`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py#L3465-L3718), chia tỷ lệ `weight=16` (`~76%` chiều ngang) cho **Đề bài + Bài làm** và `weight=5, minsize=230` (`~24%` chiều ngang) cho **Chấm điểm + Báo lỗi**.
   - Thu gọn cột `MỤC` (`24px`), cột `ĐIỂM` (`30px`) và `sidebar` (`212px`), tự động xuống dòng lý do theo chiều rộng thực tế (`<Configure>`).
3. **Bộ nhớ phiên độc lập cho từng Tab Lớp (`self.class_sessions[code]`) + Chấm AI chạy ngầm đa lớp + Tự khôi phục 1 file:**
   - Mỗi tab lớp lưu độc lập trong RAM (`self.class_sessions[code]`) và trên đĩa (`bai_lam_iuh_<MãLớp>_review.json`): danh sách `review_items`, vị trí sinh viên đang dò (`current_review_pos`), ô điểm đang chọn (`active_score_slot_idx`), bộ lọc (`filter_mode`), và tiến trình AI riêng (`ai_running`, `ai_queued`, `ai_done`, `ai_total`).
   - Giảng viên có thể mở lại lớp đã chấm (như `2607AVCB005`) để kiểm lại Part 1 hoặc Part 2 **chỉ bằng 1 file duy nhất** (chỉ chọn file Excel hoặc chỉ chọn file JSON) trong khi một lớp khác (VD: `2607AVCB004`) đang được AI chấm ngầm ở hậu cảnh (`2607AVCB004 (AI: 15/35)`).
4. **Chốt chặn bảo vệ tuyệt đối File Gốc trên máy tính (`Strict File Isolation Guard`):**
   - [`validate_class_files_not_mixed`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py#L1636-L1650): Chặn ghép nhầm file JSON của lớp A với file Excel của lớp B.
   - [`_route_selected_file_safely`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py#L4498-L4530): Nếu đang đứng ở Tab lớp `2607AVCB005` mà chọn file của lớp khác (VD: `2607AVCB004`), hệ thống **tự động mở/chuyển sang Tab riêng cho lớp `2607AVCB004`** thay vì ghi đè lên đường dẫn của lớp `2607AVCB005`.
   - [`save_review_session`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py#L1747-L1785) & [`export_approved_scores_to_excel`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py#L2264-L2365): Chặn tuyệt đối việc ghi đè nếu danh sách `MSSV` giữa bộ nhớ và file đích có `0` sinh viên trùng khớp.
   - Cô lập biến môi trường `IUH_SUITE_CONFIG_PATH` trong [`src/test_auto_iuh.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/test_auto_iuh.py#L32-L44) để khi chạy test không bao giờ chạm vào `config.json` thật của giảng viên.

---

### Giai đoạn 7: Giao Diện Tối Giản Phong Cách Apple (Single-Screen Workspace — Zero Icons, Zero "AI" Buzzwords), Floating Class Popover, Tăng Giảm Cỡ Chữ `A- / A+`, Bảo Toàn GV2 & Bảo Vệ JS Moodle Submitter

1. **Giao diện 1 Màn hình duy nhất (Apple macOS Minimalist Single-Screen UI — `y = 0`):**
   - Loại bỏ hoàn toàn thanh `topbar` đen và thanh 3 Tab khỏi màn hình để không gian chấm & kiểm duyệt bắt đầu ngay tại đỉnh cửa sổ (`y = 0`), thu hồi thêm **`+70px` chiều dọc** cho Đề bài và Bài làm.
   - Bảng màu chuẩn Apple (`#F5F5F7` nền phụ, `#FFFFFF` thẻ bài làm giấy trắng, `#D2D2D7` / `#E5E5EA` đường kẻ `1px` mảnh, `#1D1D1F` chữ mực đậm, `#86868B` nhãn phụ, `#0071E3` nút lưu chính).
   - **100% không dùng Icon/Emoji và không dùng chữ "AI"** trên toàn bộ nút bấm giao diện: `Lớp: 2607AVCB005 (34/34)`, `Trước`, `Tiếp`, `Hút dữ liệu`, `Chấm tự động`, `Cài đặt`, `A- | 11pt | A+`, `Ẩn cả 2 đề`, `Đổi file`, `Nạp lại Excel`, `Lưu điểm ra Excel`, `Sao chép mã nhập Moodle`.
2. **Menu Lớp Dọc Nổi (`self.class_overlay_panel` dùng `.place()` + `.lift()`):**
   - Bấm vào nút `Lớp: ...` sẽ xổ menu dọc đè lên bề mặt giao diện (`0px` đẩy lệch bố cục), tự động đóng khi click ra ngoài hoặc bấm `Escape`, đồng thời tự động quét các lớp khác trong cùng thư mục để mở nhanh 1-click.
3. **Tăng / Giảm Cỡ Chữ Tức Thì (`A- | 11pt | A+`) & Tự Động In Đậm Câu Hỏi Đề Bài (`prompt_question`):**
   - Nâng cỡ chữ bài làm sinh viên lên `Segoe UI 11pt` mặc định, cho phép chỉnh nóng từ `9pt` đến `16pt` và tự lưu vào `config.json` (`font_size_pt`).
   - Tự động in đậm (`bold #1D1D1F`) các dòng câu hỏi (`?`) và gạch đầu dòng yêu cầu chính trong khung Đề bài.
4. **Hỗ trợ Giảng viên 2 (GV2) Chấm Độc Lập & Đồng Bộ Ngược Từ Excel (`Nạp lại Excel`):**
   - Giữ nguyên 100% các cột ngoài 8 cột `C-O-G-V` của GV1 (bao gồm cột `Q` Điểm tổng GV2 và cột `R` Điểm thống nhất).
   - Khi giảng viên sửa điểm trong file `.xlsx` để thống nhất với GV2, hệ thống tự động nhận diện file `.xlsx` mới hơn `_review.json` (hoặc qua nút `Nạp lại Excel`) để cập nhật điểm mới nhất trước khi xuất mã JS nhập lên Moodle.
5. **Cơ chế Bảo vệ 2 Tầng khi Nhập Điểm Moodle (`build_js_paced_grade_submitter`) gặp Sinh Viên Đã Được Người Khác Chấm Giữa Chừng:**
   - **Không bao giờ lệch dòng:** Khớp chính xác theo `MSSV` trên từng thẻ `<tr>` và lấy đúng link `slot=17` / `slot=18` của dòng đó.
   - **Lớp bảo vệ 1 (Trên bảng điểm Gradebook — `isSlotAlreadyGradedOnTable`) & Lớp bảo vệ 2 (Kiểm tra trực tiếp trong form `comment.php` — `hasExistingServerGrade`):** Mặc định (`overwriteExisting = false`) tự động **BỎ QUA (`0s`, không chờ 3–5s, không ghi đè)** bất kỳ câu `Q17`/`Q18` nào đã có điểm do người khác chấm trước đó hoặc vừa chấm xong giữa chừng.
   - Có ô tích tùy chọn *"Ghi đè cả những câu đã có người chấm trước đó"* ngay trên thanh trạng thái nổi của Moodle nếu giảng viên chủ động muốn cập nhật đè điểm mới.
   - Tự động thử lại bằng định dạng dấu thập phân thay thế (`7.5` $\leftrightarrow$ `7,5`) nếu máy chủ Moodle trả về cảnh báo định dạng số.

---

### Giai đoạn 8: Dọn Gọn Cấu Trúc Dự Án & Hợp Nhất Quy Tắc Chấm Chuẩn Hóa Mới Nhất (`MARKING_RULES_IUH.md` — `v6`)
1. **Hợp nhất 3 bản `MARKING_RULES_IUH.md` thành 1 bản chuẩn duy nhất tại gốc ([`MARKING_RULES_IUH.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/MARKING_RULES_IUH.md)):**
   - Kết hợp trình bày Markdown rõ đẹp của bản `v5` với toàn bộ quy tắc mới nhất của bản **`v6 - Target Reader & Role-Reversal Calibration`** (bao gồm bắt đảo ngược vai chủ tiệc $\leftrightarrow$ khách mời, Mẫu tham chiếu #7 cập nhật `4 - 3 - 3 - 3` cho bài đạt 2/3 câu hỏi, và Mẫu tham chiếu #9 Emma xưng hô ngôi thứ 3 `1 - 1 - 1 - 1`), đồng bộ vào [`config.json`](file:///d:/IUH/Auto_Nhap_Diem_IUH/config.json) và [`src/app_gui.py`](file:///d:/IUH/Auto_Nhap_Diem_IUH/src/app_gui.py).
2. **Đóng gói 1 file chạy duy nhất tại thư mục gốc ([`d:\IUH\Auto_Nhap_Diem_IUH\IUH_Writing_AI_Suite.exe`](file:///d:/IUH/Auto_Nhap_Diem_IUH/IUH_Writing_AI_Suite.exe)) & Xóa bỏ toàn bộ file thừa:**
   - Xóa các thư mục tạm `build/`, `dist/`, `IUH_Writing_AI_Suite.spec`, 2 file `MARKING_RULES_IUH.md` trùng lặp, file test rác `bai_lam_iuh_2607AVCB004_review.json`, file `DEMO_PHIEU_CHAM_2607AVCB005.xlsx`, `HUONG_DAN_SU_DUNG.txt`, `src/docs/superpowers/`, và 2 file `.bat` cũ.
   - Cập nhật [`README.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/README.md) với hướng dẫn sử dụng mới nhất cho giao diện Apple Minimalist 1 màn hình.

---

### Giai đoạn 9: Huấn Luyện Chuẩn Vàng 72 Sinh Viên (`2607AVCB005` & `2607AVCB006`), Chuẩn Hóa Nhận Xét Kiểm Chứng Cho Giám Khảo & Phân Tách Bộ Quy Tắc Theo Sách/Đợt Thi (`MARKING_RULES_CAMBRIDGE_CUOI_KI.md` — `v7`)

1. **Khử lỗi tự ý xuống hàng khi Hút dữ liệu F12 (`JS_SCRAPER_CODE`):**
   - Khử toàn bộ ký tự xuống dòng mềm (`\n`) bên trong các `TEXT_NODE` của HTML Moodle và tự động nối lại các dòng bị ngắt giữa cụm từ (`"what\ntime you"`, `"how\noften you"`, `"at the\nrestaurant"`), chấm dứt hiện tượng phạt oan `XUONG_DONG_TUNG_CAU` do trình duyệt tự wrap chữ.
2. **Huấn luyện quy tắc `v7` từ 144 bài viết thực tế của 2 lớp `2607AVCB005` & `2607AVCB006`:**
   - **Part 1 — Phân biệt 2 bẫy câu hỏi:**
     - Quên dấu `?` cuối câu hỏi đúng chủ đề (*"What time do you wake up. And what your usually does in the morning."*, Huỳnh Gia Huy) hoặc dùng câu hỏi gián tiếp (*"Let me know what..."*) $\rightarrow$ Vẫn tính là **đã hỏi đủ 2 câu (`Content = 5`)**.
     - Chỉ dùng câu hỏi xã giao/mời mọc (*"How are you?"*, *"Would you like to my house?"*, *"OK for you?"*, Phạm Thái Dương) hoặc kể về bạn ở ngôi thứ 3 (*"My friend's favourite sport is..."*, Nguyễn Thị Yến Nhi) $\rightarrow$ Loại khỏi câu hỏi đề bài, trừ về **`Content = 3`**.
   - **Part 1 — Học thuộc văn mẫu hẹn giờ (`Would you like to... I'm free at...`) hoặc biến toàn bộ gạch đầu dòng thành câu hỏi ngược cho bạn (Huỳnh Đan Trường) $\rightarrow$ Khóa cứng `0 - 1 - 1 - 1`.**
   - **Part 2 — Phân hóa 3 mức Lộn vai / Đảo vai giao tiếp:**
     - Lộn vai nhẹ ở câu mở đầu (*"Would you like to come..."*) nhưng trả lời đúng 3 câu hỏi sau $\rightarrow$ `Content = 4`.
     - Mở đầu *"Thanks you for inviting"* KÈM trả lời ngược chủ thể (*"And I can bring some food and drinks"*, Phạm Thị Bình Chi) hoặc ghép văn mẫu bận học thi (Bùi Phước Khoa) $\rightarrow$ `Content = 3`.
     - Ký tên cuối thư bằng chính tên người gửi trên đề (`Sophie` / `Alex`) hoặc kể chuyện quá khứ đi Vũng Tàu hoặc **Tự biên tự diễn liên thông P1 $\rightarrow$ P2** (ở P2 viết thư trả lời cho chính bức thư P1 của mình: Lưu Gia Hân, Huỳnh Đan Trường, Lê Văn Chương) $\rightarrow$ Khóa cứng **`0 - 1 - 1 - 1`**.
   - **Kết thư chỉ ghi `"Your friend."` không có tên riêng $\rightarrow$ Trừ 1đ Content (`Content = 4`, Hoàng Kiều Anh).**
   - **Lặp $\ge 5$ câu đơn mở đầu bằng `"I..."` liên tiếp không có từ nối $\rightarrow$ Hạ `Organization = 3`.**
   - **Mắc $\ge 2$ lỗi cấu trúc ngữ pháp cơ bản $\rightarrow$ Chấm `Grammar = 3` ngay từ đầu và trích dẫn trực tiếp từng cụm sai.**
3. **Chuẩn hóa Nhận xét Kiểm chứng cho Giám khảo (Examiner Verification Format):**
   - Cấm nhận xét chung chung; bắt buộc trích dẫn trực tiếp cụm từ sai `"..."` trong bài sinh viên.
   - Khi Giám khảo sửa điểm 1 tiêu chí trên giao diện (`_on_score_var_changed`), chỉ tổng hợp lại lý do của đúng tiêu chí vừa sửa và giữ nguyên dẫn chứng của 3 tiêu chí còn lại.
4. **Đổi tên & Phân tách Hệ thống Quy tắc chấm theo Sách và Đợt thi (`MARKING_RULE_PROFILES`):**
   - Đổi tên file quy tắc hiện tại thành **[`MARKING_RULES_CAMBRIDGE_CUOI_KI.md`](file:///d:/IUH/Auto_Nhap_Diem_IUH/MARKING_RULES_CAMBRIDGE_CUOI_KI.md)** (`Cuối kì - Sách Cambridge`).
   - Tích hợp hộp chọn **Bộ quy tắc chấm** trong cửa sổ `Cài đặt` và `Tab 1` cho phép chuyển đổi độc lập giữa 4 bộ quy tắc:
     1. `Cuối kì - Sách Cambridge` (`MARKING_RULES_CAMBRIDGE_CUOI_KI.md`)
     2. `Giữa kì - Sách Cambridge` (`MARKING_RULES_CAMBRIDGE_GIUA_KI.md`)
     3. `Giữa kì - Sách Oxford` (`MARKING_RULES_OXFORD_GIUA_KI.md`)
     4. `Cuối kì - Sách Oxford` (`MARKING_RULES_OXFORD_CUOI_KI.md`)

---

### Giai đoạn 10: Chấm Theo Cụm Mã Đề Ngầm (`4 Lượt Gọi API / 1 Lớp`), Đính Kèm Bài Mẫu Gold Standard Đúng Mã Đề, Nâng Cấp `gemini-3.8-flash` (`thinking_level="high"`) & Tách Riêng Từng Lớp Khi Sao Chép Mã Nhập Moodle

1. **Chấm Theo Cụm Mã Đề Ngầm Trong Bộ Nhớ (`_build_prompt_clusters_for_part`, `cham_1_cum_ma_de`, `_grade_tasks_by_prompt_clusters`):**
   - Thay vì trộn 7 sinh viên có 14 đề bài khác nhau vào 1 lần gọi khiến AI phân tán sự chú ý, hệ thống tự động gom nhóm ngầm trong RAM các bài làm theo đúng **Mã đề** (`AVCB.FT.W.P1.L001`..`L010` và `AVCB.FT.W.P2.L001`..`L010`) và đóng gói thành tối đa **2 cụm cho Part 1** + **2 cụm cho Part 2** = **tối đa 4 lượt gọi API cho cả 1 lớp (~35–40 sinh viên)**.
   - Tiết kiệm số lượt gọi API (chấm được **5 lớp/ngày/1 model** trong hạn mức 20 RPD Free Tier, hoặc **15–20 lớp/ngày** khi xoay vòng qua các model/API key miễn phí).
   - **Bảo toàn tuyệt đối 100% thứ tự sinh viên (`row`):** Việc gom nhóm theo mã đề chỉ diễn ra ngầm trong bộ nhớ khi gửi cho AI; sau khi nhận kết quả, hệ thống ghép điểm theo `mssv` trả về đúng thứ tự dòng Excel ban đầu (`[items_by_mssv[it["mssv"]] for it in existing_items]`). Danh sách sinh viên bên cột trái giao diện và trong file Excel không bao giờ bị xáo trộn.
2. **Đính Kèm Bài Mẫu Chuẩn Giảng Viên (Gold Standard Few-Shot) Của Đúng Mã Đề (`_load_gold_standard_examples_by_code`, `_format_gold_examples_for_code`) & Đối Chiếu Chéo P1 $\rightarrow$ P2:**
   - Phủ đủ **20/20 mã đề** (`P1.L001..L010` và `P2.L001..L010`) từ dữ liệu chuẩn vàng `2607AVCB005` & `2607AVCB006` (kèm bộ dữ liệu dự phòng `_FALLBACK_GOLD_EXAMPLES_BY_CODE` nhúng sẵn trong code).
   - Ngay dưới mỗi ngăn mã đề trong prompt, AI được cung cấp đề bài gốc và 1–3 bài mẫu chuẩn giảng viên của chính mã đề đó để so sánh trực tiếp.
   - Khi chấm cụm Part 2, hệ thống tự động đính kèm dòng `Đối chiếu chéo Part 1 của SV` (`_extract_p1_cross_ref_note`) ghi rõ lời chào và chữ ký ở Part 1 để AI và bộ hậu kiểm phát hiện ngay lỗi sinh viên tự viết thư trả lời cho bài Part 1 của chính mình.
3. **Nâng Cấp Model Mặc Định Lên `gemini-3.8-flash` (`thinking_level="high"`) & Xoay Vòng Đa API Key / Đa Model Free Tier (`_call_gemini_interactions`, `_parse_api_keys`):**
   - Chuyển `DEFAULT_MODEL` từ `gemini-3.5-flash` (đã deprecated) sang **`gemini-3.8-flash`** với danh sách xoay vòng `["gemini-3.8-flash", "gemini-3.6-flash", "gemini-2.5-pro", "gemini-2.5-flash", "gemini-3.5-flash"]`.
   - Nâng `thinking_level` từ `"low"` lên **`"high"`** (tự động hạ xuống `"low"` nếu gặp model không hỗ trợ `"high"`) và tăng `timeout_sec` lên `120s`.
   - Hỗ trợ nhập nhiều API Key phân cách bởi dấu phẩy `,`, chấm phẩy `;` hoặc khoảng trắng trong ô API Key để tự động xoay vòng khi chạm trần 20 RPD (`HTTP 429`).
4. **Tách Riêng Từng Lớp Khi Bấm `Sao chép mã nhập Moodle` (`_collect_multi_class_grades(only_active_class=True)`):**
   - Mặc định khi bấm `Sao chép mã nhập Moodle`, hệ thống **chỉ đóng gói điểm của đúng lớp đang mở trên màn hình (`self.active_class_code`)**, không gộp tràn sang các lớp khác đang mở, giúp bảng tiến độ trên trình duyệt F12 hiển thị chính xác sĩ số của lớp đó (`38/38` thay vì `38/80`).

---

### Giai đoạn 11: Hội Đồng Chấm Đa Mô Hình 2 Vòng (`gemini-3.8-flash` $\rightleftarrows$ `gemini-3.7-flash` / `gemini-3.1-pro-preview`) & Thanh Tiến Trình Nổi (Floating Progress Bar — Dynamic Island Style)

1. **Hội Đồng Chấm 2 Vòng Đa Mô Hình (`_select_borderline_cases_for_arbitration`, `cham_vong_2_trong_tai_phan_bien`):**
   - **Vòng 1 (Giám khảo 1 — Mặc định `gemini-3.8-flash` · `thinking_level="high"`):** Chấm chuyên sâu 4 Cụm Mã Đề (`2 cụm P1 + 2 cụm P2`) đối chiếu với 20/20 bài mẫu Gold Standard.
   - **Đối chiếu chuẩn (`_calibrate_scores` — Python Rule Engine):** Quét tức thì 12 quy tắc cứng (chữ ký cuối thư, lỗi Comma splice, xuống dòng rời rạc, câu hỏi chủ đề P1, lộn vai giao tiếp P1 $\rightarrow$ P2).
   - **Vòng 2 (Trọng tài Thẩm định chéo Độc lập — Mặc định `gemini-3.7-flash` · `thinking_level="high"`, hỗ trợ chọn `gemini-3.1-pro-preview`, `gemini-4-argon`...):**
     - Tự động lọc ra **tối đa 6 hồ sơ nhạy cảm/xung đột** trong lớp (bài có độ lệch giữa GK1 và Quy tắc Khoa, bài ở ranh giới lạc đề `Content <= 2`, chênh lệch lớn giữa P1 và P2, hoặc bài ngưỡng điểm cao `G/V >= 4`).
     - Chỉ tốn **đúng 1 lượt gọi API bổ sung/lớp** (tổng cộng **5 lượt gọi API / 1 lớp** cho cả 2 vòng) để Trọng tài Vòng 2 đọc độc lập Đề gốc + Bài mẫu Gold Standard + Bài làm + Điểm đề xuất của GK1 + Cảnh báo Quy tắc Khoa và ra phán quyết chốt điểm (`KetQuaHoiDongTrongTai`).
     - Gắn nhãn xác thực **`Thẩm định 2 vòng (3.8-flash ⇄ 3.7-flash)`** hoặc **`Chuẩn hội đồng (3.8-flash + Rule Engine)`** ngay trên thanh tiêu đề của từng sinh viên để giảng viên dễ dàng kiểm chứng.
2. **Thanh Tiến Trình Nổi & Line Tiến Trình Liền Mạch (`_build_pipeline_hud_contents`, `_redraw_pipeline_canvas`, `_show_pipeline_hud`, `_update_pipeline_hud_step`):**
   - Tích hợp nút **`Tiến trình`** trên thanh công cụ chính và khung **Floating Progress Bar** định vị bằng `.place(relx=0.5, y=42, anchor="n", relwidth=0.78)` nổi ở chính giữa phía trên (`0px` đẩy lệch bố cục, không làm co hẹp khung đọc bài Part 1 & Part 2).
   - Vẽ **Line Tiến Trình Liền Mạch (`tk.Canvas`)** chạy mượt từ `0%` $\rightarrow$ `100%` (xanh dương `#0071E3` khi đang chạy $\rightarrow$ xanh lá `#10B981` khi hoàn tất) nối qua **5 mốc ngắn gọn, chuyên nghiệp**:
     1. `Phân đề` (*20 mã đề*)
     2. `Chấm vòng 1` (*4 cụm đề*)
     3. `Đối chiếu chuẩn` (*Quy tắc Khoa*)
     4. `Thẩm định chéo` (*Model vòng 2*)
     5. `Hoàn tất` (*Khớp dòng Excel*)
   - Kèm **Dòng trạng thái trực tiếp 1 dòng (`lbl_pipeline_live_line`)** và nút **`Chi tiết`** để mở/đóng ngăn nhật ký 4 dòng (`pipeline_stream_text`) trích xuất trực tiếp `thinking_summaries="auto"` từ Gemini Interactions API (`_emit_thought_summary`).
