import json
import os
import shutil
import tempfile
import unittest
import openpyxl
from auto_iuh import (
    detect_columns_and_parse,
    format_score,
    normalize_mssv,
    filter_start_students,
)
from app_gui import (
    DEFAULT_RUBRIC,
    IUHWritingAISuiteApp,
    KetQuaCham,
    _calibrate_scores,
    build_review_items_from_excel_and_json,
    cham_1_sinh_vien,
    chay_cham_8_luong_ghi_excel,
    chay_cham_ai_tao_ban_duyet,
    clean_prompt_text,
    ensure_cogv_reasons,
    export_approved_scores_to_excel,
    get_review_session_path,
    inspect_json_dataset,
    load_review_session,
    save_review_session,
)


class TestAutoIUH(unittest.TestCase):
    def setUp(self):
        self._tmp_cfg_dir = tempfile.mkdtemp(prefix="iuh_suite_test_cfg_")
        self._old_cfg_env = os.environ.get("IUH_SUITE_CONFIG_PATH")
        os.environ["IUH_SUITE_CONFIG_PATH"] = os.path.join(self._tmp_cfg_dir, "test_config.json")

    def tearDown(self):
        if self._old_cfg_env is None:
            os.environ.pop("IUH_SUITE_CONFIG_PATH", None)
        else:
            os.environ["IUH_SUITE_CONFIG_PATH"] = self._old_cfg_env
        shutil.rmtree(self._tmp_cfg_dir, ignore_errors=True)

    def test_normalize_mssv_and_score(self):
        self.assertEqual(normalize_mssv(25677201), '25677201')
        self.assertEqual(normalize_mssv('sv_24707841'), '24707841')
        self.assertEqual(normalize_mssv('24707841.0'), '24707841')
        self.assertEqual(normalize_mssv('MSSV'), '')
        self.assertEqual(format_score(None), '0')
        self.assertEqual(format_score('  '), '0')
        self.assertEqual(format_score(3.0), '3')
        self.assertEqual(format_score(2.75), '2.75')

    def test_detect_real_shifted_file(self):
        real_path = r'D:\IUH\Auto_Nhap_Diem_IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx'
        if os.path.exists(real_path):
            cols, students = detect_columns_and_parse(real_path)
            self.assertEqual(cols['mssv'], 3)   # Col D
            self.assertEqual(cols['q17'], 10)   # Col K
            self.assertEqual(cols['q18'], 16)   # Col Q
            self.assertEqual(cols['p1_cols'], [5, 6, 7, 8])
            self.assertEqual(cols['p2_cols'], [11, 12, 13, 14])
            self.assertEqual(cols['comment'], 20)  # Col U
            self.assertEqual(students[0]['row'], 10)
            self.assertEqual(students[0]['mssv'], '25677201')
            self.assertEqual(students[0]['ten'], 'Trần Quốc Duy')
            self.assertEqual(students[0]['q17'], '1')
            self.assertEqual(students[0]['q18'], '2.75')

    def test_detect_real_department_template(self):
        real_path = r'D:\IUH\Tài liệu giảng dạy AVCB\AVCB\2606AVCB029\2606AVCB029_GV1+GV2.xlsx'
        if os.path.exists(real_path):
            cols, students = detect_columns_and_parse(real_path)
            self.assertEqual(cols['mssv'], 1)   # Col B
            self.assertEqual(cols['q17'], 8)    # Col I
            self.assertEqual(cols['q18'], 14)   # Col O
            self.assertEqual(cols['p1_cols'], [3, 4, 5, 6])
            self.assertEqual(cols['p2_cols'], [9, 10, 11, 12])
            self.assertEqual(cols['comment'], 18)  # Col S
            self.assertEqual(students[0]['row'], 10)
            self.assertEqual(students[0]['mssv'], '24631181')
            self.assertEqual(students[0]['ten'], 'Võ Quốc Tường')
            self.assertEqual(students[0]['q17'], '3.5')
            self.assertEqual(students[0]['q18'], '2')

    def test_detect_custom_layout_split_name_and_start_row_5(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['STT', 'Họ đệm', 'Tên', 'Mã SV', 'Q17', 'Q18'])
        ws.append([])
        ws.append([])
        ws.append([])
        ws.append([1, 'Nguyễn Văn', 'An', 'sv_24631181', 3.0, None])
        ws.append([2, 'Lê Thị', 'Bình', 24727581, 4.25, 3.5])
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name
        wb.save(tmp_path)
        wb.close()
        try:
            cols, students = detect_columns_and_parse(tmp_path)
            self.assertEqual(cols['mssv'], 3)
            self.assertEqual(cols['q17'], 4)
            self.assertEqual(cols['q18'], 5)
            self.assertEqual(len(students), 2)
            self.assertEqual(students[0], {
                'row': 5, 'mssv': '24631181', 'ten': 'Nguyễn Văn An', 'q17': '3', 'q18': '0'
            })
            self.assertEqual(len(filter_start_students(students, '6')), 1)
            self.assertEqual(filter_start_students(students, '24727581')[0]['mssv'], '24727581')
        finally:
            os.remove(tmp_path)

    def test_fallback_data_scan_when_mssv_header_custom(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['No', 'StudentCode', 'Full Name', 'Part 1', 'Part 2'])
        ws.append([1, 25677201, 'Trần Quốc Duy', 2.5, 3.75])
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp_path = tmp.name
        wb.save(tmp_path)
        wb.close()
        try:
            cols, students = detect_columns_and_parse(tmp_path)
            self.assertEqual(cols['mssv'], 1)
            self.assertEqual(cols['q17'], 3)
            self.assertEqual(cols['q18'], 4)
            self.assertEqual(students[0]['mssv'], '25677201')
            self.assertEqual(students[0]['q17'], '2.5')
            self.assertEqual(students[0]['q18'], '3.75')
        finally:
            os.remove(tmp_path)

    def test_blank_student_essay_zero_score_without_api(self):
        sv_data = {
            "mssv": "25677201",
            "p1": {"code": "AVCB.P1.01", "prompt": "Write about your family", "essay": "   "},
            "p2": {"code": "AVCB.P2.01", "prompt": "Write an email", "essay": ""},
        }
        mssv, kq = cham_1_sinh_vien("fake-key", sv_data, DEFAULT_RUBRIC)
        self.assertEqual(mssv, "25677201")
        self.assertEqual(kq.p1_content, 0)
        self.assertEqual(kq.p2_vocab, 0)
        self.assertIn("bỏ trắng", kq.nhan_xet)

    def test_inspect_json_dataset_with_text_and_images(self):
        sample_data = [
            {
                "mssv": "25677201",
                "p1": {
                    "code": "AVCB.FT.W.P1.L004",
                    "prompt": "The picture is yours. Write a post of about 50 words.",
                    "images": [{"mime_type": "image/png", "data": "iVBORw0KGgo="}],
                    "essay": "My name is Duy and this is my family photo.",
                },
                "p2": {
                    "code": "AVCB.FT.W.P2.L002",
                    "prompt": "Write an email to your friend Alex.",
                    "images": [],
                    "essay": "Hi Alex, I would love to go to the cinema with you.",
                },
            }
        ]
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", encoding="utf-8", delete=False) as f:
            json.dump(sample_data, f)
            json_path = f.name
        try:
            info = inspect_json_dataset(json_path)
            self.assertEqual(info["total"], 1)
            self.assertEqual(info["non_empty_students"], 1)
            self.assertEqual(info["p1_codes"], ["AVCB.FT.W.P1.L004"])
            self.assertEqual(info["p2_codes"], ["AVCB.FT.W.P2.L002"])
            self.assertEqual(info["p1_img_count"], 1)
            self.assertEqual(info["p2_img_count"], 0)
        finally:
            os.remove(json_path)

    def test_chay_cham_8_luong_ghi_excel_preserves_formulas_and_fallback_reads_final_score(self):
        real_path = r'D:\IUH\Auto_Nhap_Diem_IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx'
        if not os.path.exists(real_path):
            self.skipTest("Không tìm thấy file Excel mẫu để test.")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_excel:
            excel_copy = tmp_excel.name
        shutil.copy2(real_path, excel_copy)

        sample_json = [
            {
                "mssv": "25677201",
                "p1": {"code": "P1.01", "prompt": "Topic 1", "essay": "Hello world sample essay one"},
                "p2": {"code": "P2.01", "prompt": "Topic 2", "essay": "Hello world sample essay two"},
            }
        ]
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", encoding="utf-8", delete=False) as tmp_json:
            json.dump(sample_json, tmp_json)
            json_path = tmp_json.name

        def mock_grader(api_key, sv_data, rubric_text):
            return sv_data["mssv"], KetQuaCham(
                p1_content=4, p1_org=3, p1_grammar=3, p1_vocab=4,   # Sum = 14 -> Final 1 = 3.5
                p2_content=5, p2_org=4, p2_grammar=4, p2_vocab=4,   # Sum = 17 -> Final 2 = 4.25
                nhan_xet="P1 (35 từ): Đúng chủ đề. P2 (48 từ): Tốt."
            )

        try:
            summary = chay_cham_8_luong_ghi_excel(
                api_key="test-key",
                json_path=json_path,
                excel_path=excel_copy,
                rubric_text=DEFAULT_RUBRIC,
                grader_fn=mock_grader,
            )
            self.assertEqual(summary["graded_count"], 1)

            # 1. Kiểm tra công thức Excel được bảo toàn 100%
            wb = openpyxl.load_workbook(excel_copy, data_only=False)
            ws = wb[wb.sheetnames[0]]
            self.assertEqual(ws.cell(10, 6).value, 4)
            self.assertEqual(ws.cell(10, 7).value, 3)
            self.assertEqual(ws.cell(10, 8).value, 3)
            self.assertEqual(ws.cell(10, 9).value, 4)
            self.assertEqual(ws.cell(10, 10).value, "=F10+G10+H10+I10")
            self.assertEqual(ws.cell(10, 11).value, "=J10/4")
            self.assertEqual(ws.cell(10, 12).value, 5)
            self.assertEqual(ws.cell(10, 16).value, "=L10+M10+N10+O10")
            self.assertEqual(ws.cell(10, 17).value, "=P10/4")
            self.assertEqual(ws.cell(10, 18).value, "=MROUND(K10+Q10,0.5)")
            self.assertEqual(ws.cell(10, 21).value, "P1 (35 từ): Đúng chủ đề. P2 (48 từ): Tốt.")
            wb.close()

            # 2. Kiểm tra đọc lại bằng detect_columns_and_parse (kể cả khi chưa mở Excel bấm Ctrl+S)
            _, students = detect_columns_and_parse(excel_copy)
            self.assertEqual(students[0]["mssv"], "25677201")
            self.assertEqual(students[0]["q17"], "3.5")
            self.assertEqual(students[0]["q18"], "4.25")
        finally:
            os.remove(excel_copy)
            os.remove(json_path)

    def test_chay_cham_8_luong_ghi_excel_department_template_col_b(self):
        real_path = r'D:\IUH\Tài liệu giảng dạy AVCB\AVCB\2606AVCB029\2606AVCB029_GV1+GV2.xlsx'
        if not os.path.exists(real_path):
            self.skipTest("Không tìm thấy file mẫu gốc Khoa để test.")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_excel:
            excel_copy = tmp_excel.name
        shutil.copy2(real_path, excel_copy)

        sample_json = [
            {
                "mssv": "24631181",
                "p1": {"code": "P1.01", "prompt": "Topic 1", "essay": "Hello world sample essay one"},
                "p2": {"code": "P2.01", "prompt": "Topic 2", "essay": "Hello world sample essay two"},
            }
        ]
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", encoding="utf-8", delete=False) as tmp_json:
            json.dump(sample_json, tmp_json)
            json_path = tmp_json.name

        def mock_grader(api_key, sv_data, rubric_text):
            return sv_data["mssv"], KetQuaCham(
                p1_content=5, p1_org=4, p1_grammar=4, p1_vocab=5,   # Sum = 18 -> Final 1 = 4.5
                p2_content=3, p2_org=3, p2_grammar=3, p2_vocab=3,   # Sum = 12 -> Final 2 = 3
                nhan_xet="Nhận xét mẫu Khoa Cột B."
            )

        try:
            summary = chay_cham_8_luong_ghi_excel(
                api_key="test-key",
                json_path=json_path,
                excel_path=excel_copy,
                rubric_text=DEFAULT_RUBRIC,
                grader_fn=mock_grader,
            )
            self.assertEqual(summary["graded_count"], 1)

            wb = openpyxl.load_workbook(excel_copy, data_only=False)
            ws = wb[wb.sheetnames[0]]
            self.assertEqual(ws.cell(10, 4).value, 5)
            self.assertEqual(ws.cell(10, 8).value, "=D10+E10+F10+G10")
            self.assertEqual(ws.cell(10, 9).value, "=H10/4")
            self.assertEqual(ws.cell(10, 19).value, "Nhận xét mẫu Khoa Cột B.")
            wb.close()

            _, students = detect_columns_and_parse(excel_copy)
            self.assertEqual(students[0]["mssv"], "24631181")
            self.assertEqual(students[0]["q17"], "4.5")
            self.assertEqual(students[0]["q18"], "3")
        finally:
            os.remove(excel_copy)
            os.remove(json_path)

    def test_calibrate_scores_matches_lecturer_benchmark_rules(self):
        # 1. Phạm Nguyễn Bảo Nhi (25677521): P1 = 5-5-5-5; P2 = 4-5-5-4 (Content trừ 1 vì thiếu ký tên kết thư, Vocab trừ 1 vì talking/wil/saturday, Grammar vẫn 5)
        sv_nhi = {
            "mssv": "25677521",
            "p1": {
                "code": "AVCB.FT.W.P1.L001",
                "prompt": "Write an email (about 50 words) to a friend.\ndocument.querySelectorAll('textarea').forEach(function(el){});",
                "essay": "Hi Mary,\nTo relax in my city, I often go to Crescent Mall. I usually walk around and buy some fresh fruit and snacks. I like this place because it is very big and beautiful. Where do you usually go to relax? What do you do there?\nBest,\nNhi",
            },
            "p2": {
                "code": "AVCB.FT.W.P2.L004",
                "prompt": "Hello, I'm so excited about the beach trip! Do you like to take photos? Let me know what you want for the picnic on the beach.",
                "essay": "Hi, Lucy\nI am very excited too! I really like talking photos, so we wil take many pictures. For the picnic, I would like some sandwiches and fresh fruit. I can cook and bring some fried chicken for us.\nSee you on saturday!",
            },
        }
        raw_kq = KetQuaCham(
            nhan_xet="Nhận xét thử nghiệm",
            p1_content=5, p1_org=5, p1_grammar=5, p1_vocab=5,
            p2_content=5, p2_org=5, p2_grammar=5, p2_vocab=5,
        )
        cal_kq = _calibrate_scores(sv_nhi, raw_kq)
        self.assertEqual((cal_kq.p1_content, cal_kq.p1_org, cal_kq.p1_grammar, cal_kq.p1_vocab), (5, 5, 5, 5))
        self.assertEqual((cal_kq.p2_content, cal_kq.p2_org, cal_kq.p2_grammar, cal_kq.p2_vocab), (4, 5, 5, 4))

        # 2. Sao chép nguyên văn đề bài (Nguyễn Thị Tuyết Nhi 24736241 P2) & Viết câu cụt <12 từ (Nguyễn Phước Huy 25754951 P1) -> 0 toàn bộ
        sv_copy = {
            "mssv": "24736241",
            "p1": {"code": "L003", "prompt": "Write an email", "essay": "Hi Quan , Would you like to get up a"},
            "p2": {
                "code": "AVCB.FT.W.P2.L007",
                "prompt": "Hi, Thanks for inviting me to your picnic! I would love to come. What time should I be there? Is your mother going to take us to the park? Can I bring some food or drinks? Looking forward to relaxing with you! Ben",
                "essay": "Hi Selena! Thanks for inviting me to your picnic! I would love to come. What time should I be there? I free 7pm. Is your mother going to take us to the park? Can I bring some food or drinks? Looking forward to relaxing with you! Ben",
            },
        }
        cal_copy = _calibrate_scores(
            sv_copy,
            KetQuaCham(
                nhan_xet="Test copy",
                p1_content=2, p1_org=2, p1_grammar=2, p1_vocab=2,
                p2_content=4, p2_org=4, p2_grammar=3, p2_vocab=3,
            ),
        )
        self.assertEqual((cal_copy.p1_content, cal_copy.p1_org, cal_copy.p1_grammar, cal_copy.p1_vocab), (0, 0, 0, 0))
        self.assertEqual((cal_copy.p2_content, cal_copy.p2_org, cal_copy.p2_grammar, cal_copy.p2_vocab), (0, 0, 0, 0))
        self.assertIn("chép lại", cal_copy.nhan_xet)
        self.assertNotIn("querySelectorAll", clean_prompt_text(sv_nhi["p1"]["prompt"]))

        # 3. Lạc đề hoàn toàn (Phạm Lê Tuyết Nhi 25741971: Content = 0 nhưng có viết >= 12 từ) -> Khóa 0 - 1 - 1 - 1
        sv_offtopic = {
            "mssv": "25741971",
            "p1": {
                "code": "AVCB.FT.W.P1.L005",
                "prompt": "Remember to give information about what you like to do on the weekend.",
                "essay": "Hi Minh, Would you like to come with me to the cinema at the weekend? I'm free at 3 o'clock at the afternoon. It that OK for you? Can we meet at Kinh Duong Vuong Galaxy Cinema.",
            },
            "p2": {
                "code": "AVCB.FT.W.P2.L008",
                "prompt": "Thank you for inviting me to the museum on Saturday!",
                "essay": "Hi Nhi, Thank you for your invitation. I'm very happy. I'm busy at 3 o'clock at the afternoon because I'm busy cleaning my house and doing homework.",
            },
        }
        cal_off = _calibrate_scores(
            sv_offtopic,
            KetQuaCham(
                nhan_xet="Lạc đề",
                p1_content=0, p1_org=3, p1_grammar=3, p1_vocab=3,
                p2_content=0, p2_org=3, p2_grammar=3, p2_vocab=3,
            ),
        )
        self.assertEqual((cal_off.p1_content, cal_off.p1_org, cal_off.p1_grammar, cal_off.p1_vocab), (0, 1, 1, 1))
        self.assertEqual((cal_off.p2_content, cal_off.p2_org, cal_off.p2_grammar, cal_off.p2_vocab), (0, 1, 1, 1))

    def test_cogv_reasons_and_review_session_persist_and_clean_excel_export(self):
        real_path = r'D:\IUH\Auto_Nhap_Diem_IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx'
        if not os.path.exists(real_path):
            self.skipTest("Không tìm thấy file Excel mẫu để test.")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_excel:
            excel_copy = tmp_excel.name
        shutil.copy2(real_path, excel_copy)

        sample_json = [
            {
                "mssv": "25677201",
                "ho_ten": "Trần Quốc Duy",
                "p1": {
                    "code": "AVCB.FT.W.P1.L001",
                    "prompt": "Write an email (about 50 words) to a friend about where you go to relax.",
                    "essay": "Hi Kiet,\nI go to Da Nang city\nI visit Ba Na Hills\nWhere you go to relax?\nWhat you likes to do?\nBest,\nDuy",
                    "essay_lines": [
                        "Hi Kiet,",
                        "I go to Da Nang city",
                        "I visit Ba Na Hills",
                        "Where you go to relax?",
                        "What you likes to do?",
                        "Best,",
                        "Duy",
                    ],
                    "word_count": 28,
                    "line_count": 7,
                    "format_type": "XUONG_DONG_TUNG_CAU",
                },
                "p2": {
                    "code": "AVCB.FT.W.P2.L005",
                    "prompt": "Write a reply (about 50 words) to Emma's party invitation.",
                    "essay": "Hi Emma,\nThe party starts at 8pm. I can bring some fruit and drinks for us.\nSee you soon,\nDuy",
                    "essay_lines": [
                        "Hi Emma,",
                        "The party starts at 8pm. I can bring some fruit and drinks for us.",
                        "See you soon,",
                        "Duy",
                    ],
                    "word_count": 20,
                    "line_count": 4,
                    "format_type": "BINH_THUONG",
                },
            }
        ]
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", encoding="utf-8", delete=False) as tmp_json:
            json.dump(sample_json, tmp_json)
            json_path = tmp_json.name

        review_path = get_review_session_path(json_path)
        try:
            kq = KetQuaCham(
                p1_content=5, p1_org=2, p1_grammar=3, p1_vocab=3,
                p1_c_reason="Đạt: Đủ ý bản thân và câu hỏi.",
                p1_o_reason="Lỗi: Xuống dòng rời rạc từng câu.",
                p1_g_reason="Lỗi: Where you go, What you likes.",
                p1_v_reason="Giới hạn: Từ vựng cơ bản.",
                p2_content=4, p2_org=4, p2_grammar=4, p2_vocab=3,
                p2_c_reason="Đạt: Trả lời giờ và đồ mang theo.",
                p2_o_reason="Đạt: Viết liền mạch.",
                p2_g_reason="Đạt: Cấu trúc câu đơn đúng.",
                p2_v_reason="Giới hạn: Từ vựng đơn giản.",
                nhan_xet="Nhận xét tổng hợp",
            )
            kq = ensure_cogv_reasons(sample_json[0], kq)
            self.assertTrue(kq.p1_c_reason)
            self.assertTrue(kq.p2_v_reason)

            # Giả lập giảng viên chỉnh sửa điểm trên Tab 2 (ví dụ sửa p1_org từ 2 lên 3) và đánh dấu Tick checked=True
            review_items = [
                {
                    "row": 10,
                    "mssv": "25677201",
                    "ten": "Trần Quốc Duy",
                    "checked": True,
                    "absent": False,
                    "note": "",
                    "p1": {
                        **sample_json[0]["p1"],
                        "c": 5, "o": 3, "g": 3, "v": 3,
                        "rc": kq.p1_c_reason, "ro": kq.p1_o_reason, "rg": kq.p1_g_reason, "rv": kq.p1_v_reason,
                    },
                    "p2": {
                        **sample_json[0]["p2"],
                        "c": 4, "o": 4, "g": 4, "v": 3,
                        "rc": kq.p2_c_reason, "ro": kq.p2_o_reason, "rg": kq.p2_g_reason, "rv": kq.p2_v_reason,
                    },
                }
            ]
            save_review_session(json_path, excel_copy, review_items)
            loaded = load_review_session(json_path)
            self.assertEqual(len(loaded["items"]), 1)
            self.assertTrue(loaded["items"][0]["checked"])
            self.assertEqual(loaded["items"][0]["p1"]["o"], 3)

            # Khi bấm Approve: chỉ xuất 8 cột điểm C-O-G-V sang Excel, KHÔNG ghi nhận xét (xóa sạch cột nhận xét nếu có)
            res = export_approved_scores_to_excel(excel_copy, loaded["items"])
            self.assertEqual(res["exported_count"], 1)

            wb = openpyxl.load_workbook(excel_copy, data_only=False)
            ws = wb[wb.sheetnames[0]]
            # Cột F, G, H, I (6, 7, 8, 9) là P1 C-O-G-V
            self.assertEqual(ws.cell(10, 6).value, 5)
            self.assertEqual(ws.cell(10, 7).value, 3)
            self.assertEqual(ws.cell(10, 8).value, 3)
            self.assertEqual(ws.cell(10, 9).value, 3)
            # Công thức tổng bảo toàn
            self.assertEqual(ws.cell(10, 10).value, "=F10+G10+H10+I10")
            self.assertEqual(ws.cell(10, 11).value, "=J10/4")
            # Cột nhận xét (cột 21 / U) hoàn toàn trống (không ghi nhận xét AI)
            self.assertIsNone(ws.cell(10, 21).value)
            self.assertIsNone(ws.cell(9, 21).value)
            wb.close()
        finally:
            if os.path.exists(excel_copy):
                os.remove(excel_copy)
            if os.path.exists(json_path):
                os.remove(json_path)
            if os.path.exists(review_path):
                os.remove(review_path)

    def _find_2607_files(self):
        folder = r'D:\IUH\Auto_Nhap_Diem_IUH\2607AVCB005'
        if not os.path.isdir(folder):
            return None, None
        excel_src = None
        json_src = os.path.join(folder, 'bai_lam_iuh_2607AVCB005.json')
        rev_src = os.path.join(folder, 'bai_lam_iuh_2607AVCB005_review.json')
        for fn in sorted(os.listdir(folder)):
            if fn.endswith('.xlsx') and not fn.startswith('~$') and 'grades' not in fn.lower() and '(2)' not in fn:
                excel_src = os.path.join(folder, fn)
                break
        if not (excel_src and os.path.exists(excel_src) and os.path.exists(json_src)):
            return None, None

        # Sao chép sang thư mục tạm riêng của từng test để bảo vệ tuyệt đối 100% dữ liệu thật của giảng viên
        sandbox_cls_dir = os.path.join(self._tmp_cfg_dir, '2607AVCB005')
        os.makedirs(sandbox_cls_dir, exist_ok=True)
        excel_copy = os.path.join(sandbox_cls_dir, os.path.basename(excel_src))
        json_copy = os.path.join(sandbox_cls_dir, 'bai_lam_iuh_2607AVCB005.json')
        rev_copy = os.path.join(sandbox_cls_dir, 'bai_lam_iuh_2607AVCB005_review.json')
        if not os.path.exists(excel_copy):
            shutil.copy2(excel_src, excel_copy)
        if not os.path.exists(json_copy):
            shutil.copy2(json_src, json_copy)
        if os.path.exists(rev_src) and not os.path.exists(rev_copy):
            shutil.copy2(rev_src, rev_copy)
        return excel_copy, json_copy

    def test_build_review_items_from_real_class_2607avcb005_and_export_clean_excel(self):
        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005 để test.")

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_copy = os.path.join(tmpdir, "PHIEU_CHAM_2607AVCB005.xlsx")
            json_copy = os.path.join(tmpdir, "bai_lam_iuh_2607AVCB005.json")
            shutil.copy2(excel_real, excel_copy)
            shutil.copy2(json_real, json_copy)

            items = build_review_items_from_excel_and_json(excel_copy, json_copy)
            self.assertEqual(len(items), 34)
            # Kiểm tra đề bài thực tế từ file JSON được gắn vào từng sinh viên
            first_active = next(it for it in items if not it.get("absent"))
            self.assertIn("AVCB.FT.W.P1", first_active["p1"]["code"])
            self.assertTrue(len(first_active["p1"]["prompt"]) > 20)
            self.assertTrue(len(first_active["p1"]["rc"]) > 5)
            self.assertTrue(len(first_active["p1"]["ro"]) > 5)
            self.assertTrue(len(first_active["p1"]["rg"]) > 5)
            self.assertTrue(len(first_active["p1"]["rv"]) > 5)

            # Giả lập sửa điểm SV đầu tiên và Approve xuất ra Excel
            first_active["p1"]["c"] = 4
            first_active["checked"] = True
            res = export_approved_scores_to_excel(excel_copy, items)
            self.assertGreaterEqual(res["exported_count"], 30)

            wb = openpyxl.load_workbook(excel_copy, data_only=False)
            ws = wb[wb.sheetnames[0]]
            row_idx = first_active["row"]
            self.assertEqual(ws.cell(row_idx, 4).value, 4)
            # Công thức cột H (RAW SCORE 1), I (FINAL SCORE 1), P (Điểm tổng GV1) giữ nguyên
            self.assertEqual(ws.cell(row_idx, 8).value, f"=D{row_idx}+E{row_idx}+F{row_idx}+G{row_idx}")
            self.assertEqual(ws.cell(row_idx, 9).value, f"=H{row_idx}/4")
            self.assertEqual(ws.cell(row_idx, 16).value, f"=MROUND(I{row_idx}+O{row_idx},0.5)")
            # Cột S (19) trước đây chứa Nhận xét AI nay phải được xóa sạch hoàn toàn
            self.assertIsNone(ws.cell(9, 19).value)
            self.assertIsNone(ws.cell(row_idx, 19).value)
            wb.close()

    def test_chay_cham_ai_tao_ban_duyet_does_not_touch_excel_until_approved(self):
        real_path = r'D:\IUH\Auto_Nhap_Diem_IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx'
        if not os.path.exists(real_path):
            self.skipTest("Không tìm thấy file Excel mẫu.")

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_copy = os.path.join(tmpdir, "test_excel.xlsx")
            json_copy = os.path.join(tmpdir, "bai_lam_test.json")
            shutil.copy2(real_path, excel_copy)
            # Lấy 1 SV trống bài làm để chạy không cần gọi API thật
            with open(json_copy, "w", encoding="utf-8") as f:
                json.dump(
                    [
                        {
                            "mssv": "25677201",
                            "ho_ten": "Trần Quốc Duy",
                            "p1": {"code": "P1_01", "prompt": "Write an email.", "essay": ""},
                            "p2": {"code": "P2_01", "prompt": "Write a reply.", "essay": ""},
                        }
                    ],
                    f,
                )

            wb_before = openpyxl.load_workbook(excel_copy, data_only=False)
            ws_before = wb_before[wb_before.sheetnames[0]]
            orig_val = ws_before.cell(10, 6).value
            wb_before.close()

            res = chay_cham_ai_tao_ban_duyet(
                api_key="DUMMY_KEY",
                json_path=json_copy,
                excel_path=excel_copy,
                rubric_text=DEFAULT_RUBRIC,
                only_ungraded=False,
            )
            self.assertEqual(res["graded_count"], 1)
            self.assertTrue(os.path.exists(get_review_session_path(json_copy)))

            # File Excel hoàn toàn chưa bị ghi đè khi chưa bấm Approve ở Tab 2
            wb_after = openpyxl.load_workbook(excel_copy, data_only=False)
            ws_after = wb_after[wb_after.sheetnames[0]]
            self.assertEqual(ws_after.cell(10, 6).value, orig_val)
            wb_after.close()

    def test_gui_3_tabs_smoke_and_interactions(self):
        import tkinter as tk

        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005.")

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_copy = os.path.join(tmpdir, "PHIEU_CHAM_2607AVCB005.xlsx")
            json_copy = os.path.join(tmpdir, "bai_lam_iuh_2607AVCB005.json")
            shutil.copy2(excel_real, excel_copy)
            shutil.copy2(json_real, json_copy)

            root = tk.Tk()
            root.withdraw()
            try:
                app = IUHWritingAISuiteApp(root)
                app.excel1_path_var.set(excel_copy)
                app.excel2_path_var.set(excel_copy)
                app.json_path_var.set(json_copy)
                app._refresh_data_matching(silent=True)

                self.assertEqual(len(app.review_items), 34)
                self.assertEqual(app.current_review_pos, 0)

                # Mặc định hiển thị Đề bài & Bài làm song song cùng lúc (Side-by-Side)
                self.assertTrue(app.prompt_visible["p1"])
                self.assertTrue(app.prompt_visible["p2"])
                app._toggle_prompt_overlay("p1")
                self.assertFalse(app.prompt_visible["p1"])
                app._toggle_prompt_overlay("p1")
                self.assertTrue(app.prompt_visible["p1"])
                app._toggle_both_prompts()
                self.assertFalse(app.prompt_visible["p1"])
                self.assertFalse(app.prompt_visible["p2"])
                app._toggle_both_prompts()
                self.assertTrue(app.prompt_visible["p1"])
                self.assertTrue(app.prompt_visible["p2"])

                # Test điều hướng 4 chiều (Cách B) & Khoanh nổi bật ô điểm đang chọn
                self.assertEqual(app.focus_zone, "list")
                self.assertEqual(app.sidebar.cget("highlightbackground"), "#2563eb")

                # Nhấn mũi tên Phải (→) để nhảy từ danh sách SV sang ô điểm P1 - C
                app._on_key_right()
                self.assertEqual(app.focus_zone, "scores")
                self.assertEqual(app.active_score_slot_idx, 0)
                self.assertEqual(app.part_widgets["p1"]["score_boxes"]["c"].cget("highlightbackground"), "#2563eb")
                self.assertEqual(app.part_widgets["p1"]["score_boxes"]["c"].cget("bg"), "#dbeafe")

                # Nhấn mũi tên Xuống (↓) để đi dọc xuống ô điểm P1 - O
                app._on_key_down()
                self.assertEqual(app.active_score_slot_idx, 1)
                self.assertEqual(app.part_widgets["p1"]["score_boxes"]["c"].cget("highlightbackground"), "#cbd5e1")
                self.assertEqual(app.part_widgets["p1"]["score_boxes"]["o"].cget("highlightbackground"), "#2563eb")

                # Test gõ trực tiếp số 0-5 ghi đè ngay mà không cần Backspace (chặn số ngoài 0-5) và tự động cập nhật Lý do chi tiết
                class DummyEvent:
                    def __init__(self, keysym="", char=""):
                        self.keysym = keysym
                        self.char = char

                app._on_score_entry_keypress(DummyEvent(keysym="4", char="4"), "p1", "o")
                self.assertEqual(app.review_items[0]["p1"]["o"], 4)
                self.assertIn("(4đ)", app.part_widgets["p1"]["reason_labels"]["o"].cget("text"))
                app._on_score_entry_keypress(DummyEvent(keysym="9", char="9"), "p1", "o")
                self.assertEqual(app.review_items[0]["p1"]["o"], 4)

                # Test khi giảng viên sửa Content = 0 cho bài >= 12 từ: tự động khóa 0 - 1 - 1 - 1 và đổi lý do sang Lạc đề / Khóa điểm
                app._on_score_entry_keypress(DummyEvent(keysym="0", char="0"), "p1", "c")
                self.assertEqual(app.review_items[0]["p1"]["c"], 0)
                self.assertEqual(app.review_items[0]["p1"]["o"], 1)
                self.assertEqual(app.review_items[0]["p1"]["g"], 1)
                self.assertEqual(app.review_items[0]["p1"]["v"], 1)
                self.assertIn("(0đ)", app.part_widgets["p1"]["reason_labels"]["c"].cget("text"))
                self.assertIn("Khóa điểm (1đ)", app.part_widgets["p1"]["reason_labels"]["o"].cget("text"))

                # Test bấm Space ngay khi đang đứng trong ô điểm: tự động Tick ✓, qua SV kế tiếp và trả tiêu điểm về danh sách
                app._on_score_entry_keypress(DummyEvent(keysym="space", char=" "), "p1", "o")
                self.assertTrue(app.review_items[0]["checked"])
                self.assertEqual(app.current_review_pos, 1)
                self.assertEqual(app.focus_zone, "list")

                # Test nạp bảng xem trước ở Tab 3
                app._inspect_excel_tab2()
                self.assertEqual(len(app.web_tree.get_children()), 34)
            finally:
                root.destroy()

    def test_prompt_cleaning_and_target_reader_calibration(self):
        from app_gui import clean_prompt_text, _calibrate_part_scores

        broken_prompt = (
            "AVCB. FINAL WRITING TEST\n\n"
            "PART 1: Write\nan email (about 50 words) to a friend.\n\n"
            "Write\nan email (about 50 words) to a friend.\n\n"
            "Remember to give information about:\n\n"
            "• what\nsport you like to play\n\n"
            "• how\noften you play it"
        )
        cleaned = clean_prompt_text(broken_prompt)
        self.assertNotIn("AVCB. FINAL WRITING TEST", cleaned)
        self.assertIn("PART 1: Write an email (about 50 words) to a friend.", cleaned)
        self.assertEqual(cleaned.lower().count("write an email (about 50 words) to a friend."), 1)
        self.assertIn("• what sport you like to play", cleaned)
        self.assertIn("• how often you play it", cleaned)

        # 1. Lê Nguyễn Tấn Phong P2 -> 4-3-3-3
        pr_phong = "What time and where should we meet? Can you let me know the movie title? Would you like to eat something with me before the movie?"
        es_phong = (
            "Hi Tom,\nThank you for reply me! Would you like to meet at 9 o'clock in the evening, at Tan Son Nhat Galaxy Cinema. "
            "It's on Nguyen Van Bao street. Can you watch the movie title happy. Yes, I do. would like to eat Pho or Banh mi. "
            "Is that ok for you?\nSee you soon,\nPhong"
        )
        c, o, g, v, _ = _calibrate_part_scores(pr_phong, es_phong, 4, 4, 4, 4, is_part1=False)
        self.assertEqual((c, o, g, v), (4, 3, 3, 3))

        # 2. Đặng Thị Ngọc Hà P2 -> 1-1-1-1
        pr_emma = "Thank you for inviting me to your party. What time does the party start? Do you want me to bring any food or drinks? Also, can I help you prepare for the party?"
        es_ha = (
            "I'm interesting my party , 6:30 evening the party start , I invited many new friend meet Emma. "
            "I want Emma to bring any food and drink , I only Emma is best friend , I hope Emma not come late. "
            "Emma can't help you prepare for the party. You only with everyone happy.\nLove you"
        )
        c, o, g, v, _ = _calibrate_part_scores(pr_emma, es_ha, 3, 2, 2, 3, is_part1=False)
        self.assertEqual((c, o, g, v), (1, 1, 1, 1))

        # 3. Trương Thị Như Ý P2 -> 0-1-1-1 (Lạc đề văn mẫu Highlands Coffee / đảo ngược vai chủ tiệc)
        es_nhu_y = (
            "HI Y,\nThank you for your invitation. I'm verry happy. I'd love I'm free on Sunday. "
            "We can meet at HighLands coffee Nguyen Van Dung at 8 o'clock in the evening. "
            "We can drink milk tea and talk together. Do you want me to bring any food?. I hope we have a coffee so good.\n"
            "See You Soon\nTRANG"
        )
        c, o, g, v, _ = _calibrate_part_scores(pr_emma, es_nhu_y, 3, 4, 4, 3, is_part1=False)
        self.assertEqual((c, o, g, v), (0, 1, 1, 1))

        # 4. Kiểm tra làm tròn tổng điểm khớp 100% với công thức =MROUND(I+O, 0.5) của Excel (làm tròn .25 lên .5)
        for p1_sum, p2_sum, expected_mround in (
            (3, 10, 3.5),   # 0.75 + 2.50 = 3.25 -> 3.5 (Nguyễn Bảo Trân)
            (17, 16, 8.5),  # 4.25 + 4.00 = 8.25 -> 8.5 (Võ Đăng Khôi)
            (16, 17, 8.5),  # 4.00 + 4.25 = 8.25 -> 8.5 (Nguyễn Thị Tú Quỳnh)
            (18, 11, 7.5),  # 4.50 + 2.75 = 7.25 -> 7.5 (Lương Thị Thu Nga)
            (12, 13, 6.5),  # 3.00 + 3.25 = 6.25 -> 6.5 (Trần Nguyễn Yến Nhi)
        ):
            dummy_item = {
                "p1": {"c": p1_sum, "o": 0, "g": 0, "v": 0},
                "p2": {"c": p2_sum, "o": 0, "g": 0, "v": 0},
            }
            _, _, tot = IUHWritingAISuiteApp._calc_item_scores(dummy_item)
            self.assertEqual(tot, expected_mround)

    def test_multi_class_tab_bar_and_paced_js_submitter(self):
        from app_gui import (
            auto_pair_class_files,
            build_js_paced_grade_submitter,
            extract_class_code,
        )

        self.assertEqual(
            extract_class_code(r"D:\IUH\2607AVCB005\bai_lam_iuh_2607AVCB005.json"),
            "2607AVCB005",
        )
        self.assertEqual(
            extract_class_code("", r"D:\IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx"),
            "2606AVCB034",
        )

        # Kiểm tra tự động ghép cặp file JSON <-> Excel cùng thư mục
        cls_dir = r"D:\IUH\Auto_Nhap_Diem_IUH\2607AVCB005"
        if os.path.isdir(cls_dir):
            j_paired, e_paired = auto_pair_class_files(os.path.join(cls_dir, "bai_lam_iuh_2607AVCB005.json"))
            self.assertTrue(j_paired.endswith("bai_lam_iuh_2607AVCB005.json"))
            self.assertTrue(e_paired.lower().endswith(".xlsx"))

        # Kiểm tra tạo mã JS F12 chạy ngầm tuần tự có nhịp tự nhiên 3-5s/câu
        sample_grades = {
            "25677201": {"ten": "Trần Quốc Duy", "q17": "3.25", "q18": "3.75", "class_code": "2607AVCB005"},
            "25688981": {"ten": "Võ Đăng Khôi", "q17": "4.25", "q18": "4", "class_code": "2606AVCB034"},
        }
        js_code = build_js_paced_grade_submitter(
            sample_grades,
            min_delay_sec=3.0,
            max_delay_sec=5.0,
            class_codes=["2607AVCB005", "2606AVCB034"],
        )
        self.assertIn("25677201", js_code)
        self.assertIn("25688981", js_code)
        self.assertIn("const MIN_DELAY_MS = 3000;", js_code)
        self.assertIn("const MAX_DELAY_MS = 5000;", js_code)
        self.assertIn("comment.php", js_code)
        self.assertIn("-mark", js_code)

    def test_per_class_independent_memory_single_file_recovery_and_strict_file_isolation_guard(self):
        import tkinter as tk
        from app_gui import (
            resolve_or_recover_class_files,
            validate_class_files_not_mixed,
        )

        # 1. Kiểm tra Chốt chặn chống ghép nhầm file JSON của lớp A với file Excel của lớp B
        with self.assertRaises(ValueError) as ctx:
            validate_class_files_not_mixed(
                r"D:\IUH\2607AVCB005\bai_lam_iuh_2607AVCB005.json",
                r"D:\IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx",
            )
        self.assertIn("2607AVCB005", str(ctx.exception))
        self.assertIn("2606AVCB034", str(ctx.exception))

        with self.assertRaises(ValueError):
            build_review_items_from_excel_and_json(
                r"D:\IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx",
                r"D:\IUH\2607AVCB005\bai_lam_iuh_2607AVCB005.json",
            )

        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005.")

        # 2. Kiểm tra khôi phục đầy đủ cả Part 1 & Part 2 khi chỉ chọn 1 file duy nhất (chỉ Excel hoặc chỉ JSON)
        items_from_excel_only = build_review_items_from_excel_and_json(excel_real, "", class_code="2607AVCB005")
        self.assertEqual(len(items_from_excel_only), 34)
        self.assertTrue(any(it["p2"]["essay"] for it in items_from_excel_only))
        self.assertTrue(any(it["p2"]["prompt"] for it in items_from_excel_only))

        items_from_json_only = build_review_items_from_excel_and_json("", json_real, class_code="2607AVCB005")
        self.assertEqual(len(items_from_json_only), 34)
        self.assertTrue(any(it["p2"]["rc"] for it in items_from_json_only))

        # 3. Kiểm tra tự động tìm lại đường dẫn thật khi config cũ lưu đường dẫn Temp đã bị xóa
        rec_j, rec_e = resolve_or_recover_class_files(
            "2607AVCB005",
            r"C:\Temp\deleted_folder\bai_lam_iuh_2607AVCB005.json",
            r"C:\Temp\deleted_folder\PHIEU_CHAM_2607AVCB005.xlsx",
        )
        self.assertTrue(os.path.exists(rec_j))
        self.assertTrue(os.path.exists(rec_e))

        # 4. Kiểm tra chống ghi đè nhầm danh sách MSSV của lớp khác vào file *_review.json hoặc file Excel gốc
        with tempfile.TemporaryDirectory() as tmpdir:
            rev_json_a = os.path.join(tmpdir, "bai_lam_iuh_2607AVCB005.json")
            save_review_session(
                rev_json_a,
                "",
                [{"mssv": "25757651", "ten": "SV Lớp 005", "p1": {"c": 3}, "p2": {"c": 4}}],
                class_code="2607AVCB005",
            )
            with self.assertRaises(ValueError):
                save_review_session(
                    rev_json_a,
                    "",
                    [{"mssv": "99999999", "ten": "SV Lớp Khác", "p1": {"c": 1}, "p2": {"c": 1}}],
                    class_code="2607AVCB005",
                )

            real_034 = r"D:\IUH\Auto_Nhap_Diem_IUH\Copy of 2606AVCB034_Trần Nguyên Trung.xlsx"
            if os.path.exists(real_034):
                excel_034_copy = os.path.join(tmpdir, "2606AVCB034.xlsx")
                shutil.copy2(real_034, excel_034_copy)
                with self.assertRaises(ValueError):
                    export_approved_scores_to_excel(excel_034_copy, items_from_excel_only)

        # 5. Kiểm tra GUI: Thanh Tab Lớp Học nhúng trực tiếp vào Topbar (0px khoảng trống) & Bộ nhớ độc lập từng lớp khi lớp khác đang chạy AI
        root = tk.Tk()
        root.withdraw()
        try:
            app = IUHWritingAISuiteApp(root)
            # Class bar nằm trực tiếp bên trong topbar, không còn class_bar_container gây khoảng trống xám
            self.assertIs(app.class_bar_frame.master, app.topbar)
            self.assertFalse(hasattr(app, "class_bar_container"))

            # Mở lớp 2607AVCB005 và di chuyển tới SV số 18 (index 17 - Lê Nguyễn Tấn Phong) để kiểm lại Part 2
            app._open_or_activate_class_from_file(excel_real)
            self.assertEqual(app.active_class_code, "2607AVCB005")
            self.assertEqual(len(app.review_items), 34)
            app.current_review_pos = 17
            app.focus_zone = "scores"
            app.active_score_slot_idx = 4  # Ô điểm Part 2 - Content
            app._render_review_sidebar()
            app._render_review_detail()

            # Giả lập mở thêm lớp thứ 2 (2607AVCB004) đang chạy AI ở hậu cảnh
            app.open_classes.append({
                "code": "2607AVCB004",
                "json_path": "",
                "excel_path": "",
                "checked_count": 0,
                "total_count": 35,
            })
            sess_004 = app._get_or_create_class_session("2607AVCB004")
            sess_004["ai_running"] = True
            sess_004["ai_done"] = 15
            sess_004["ai_total"] = 35
            sess_004["review_items"] = [
                {
                    "row": 10,
                    "mssv": "24000001",
                    "ten": "Nguyễn Văn A (Lớp 004)",
                    "checked": False,
                    "absent": False,
                    "graded": True,
                    "note": "",
                    "p1": {"code": "P1", "prompt": "P1", "essay": "E1", "c": 4, "o": 4, "g": 4, "v": 4},
                    "p2": {"code": "P2", "prompt": "P2", "essay": "E2", "c": 3, "o": 3, "g": 3, "v": 3},
                }
            ]

            # Chuyển qua tab 2607AVCB004 đang chạy AI rồi chuyển ngược về 2607AVCB005:
            # Cho phép chuyển tab tự do kể cả khi AI đang chấm lớp 004, và giữ nguyên vị trí SV #18 + ô điểm Part 2 của lớp 005!
            app._switch_to_class_tab("2607AVCB004")
            self.assertEqual(app.active_class_code, "2607AVCB004")
            self.assertEqual(len(app.review_items), 1)

            app._switch_to_class_tab("2607AVCB005")
            self.assertEqual(app.active_class_code, "2607AVCB005")
            self.assertEqual(len(app.review_items), 34)
            self.assertEqual(app.current_review_pos, 17)
            self.assertEqual(app.active_score_slot_idx, 4)
        finally:
            root.destroy()

    def test_gv2_columns_remain_unchanged_and_excel_manual_edit_sync(self):
        import time
        import tkinter as tk

        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005.")

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_copy = os.path.join(tmpdir, "PHIEU_CHAM_2607AVCB005.xlsx")
            json_copy = os.path.join(tmpdir, "bai_lam_iuh_2607AVCB005.json")
            shutil.copy2(excel_real, excel_copy)
            shutil.copy2(json_real, json_copy)

            # 1. Đặt điểm tổng GV2 (Cột Q / 17) và công thức Điểm thống nhất (Cột R / 18) vào dòng 10
            wb = openpyxl.load_workbook(excel_copy)
            ws = wb[wb.sheetnames[0]]
            ws.cell(10, 17, value=8.0)
            ws.cell(10, 18, value="=AVERAGE(P10,Q10)")
            wb.save(excel_copy)
            wb.close()

            items = build_review_items_from_excel_and_json(excel_copy, json_copy, class_code="2607AVCB005")
            save_review_session(json_copy, excel_copy, items, class_code="2607AVCB005")

            # Xuất điểm ra Excel -> Cột Q (GV2) và Cột R (Thống nhất) phải giữ nguyên 100% (remain unchanged)
            export_approved_scores_to_excel(excel_copy, items)
            wb_check = openpyxl.load_workbook(excel_copy, data_only=False)
            ws_check = wb_check[wb_check.sheetnames[0]]
            self.assertEqual(ws_check.cell(10, 17).value, 8.0)
            self.assertEqual(ws_check.cell(10, 18).value, "=AVERAGE(P10,Q10)")
            wb_check.close()

            # 2. Giả lập giảng viên mở file Excel sửa lại điểm thành phần dòng 10 để khớp với GV2 (P1: 5-5-4-4 -> Q17=4.5; P2: 4-4-3-3 -> Q18=3.5)
            time.sleep(0.3)
            wb_edit = openpyxl.load_workbook(excel_copy)
            ws_edit = wb_edit[wb_edit.sheetnames[0]]
            ws_edit.cell(10, 4, value=5)
            ws_edit.cell(10, 5, value=5)
            ws_edit.cell(10, 6, value=4)
            ws_edit.cell(10, 7, value=4)
            ws_edit.cell(10, 10, value=4)
            ws_edit.cell(10, 11, value=4)
            ws_edit.cell(10, 12, value=3)
            ws_edit.cell(10, 13, value=3)
            wb_edit.save(excel_copy)
            wb_edit.close()
            now_ts = time.time() + 2.0
            os.utime(excel_copy, (now_ts, now_ts))

            # Khi nạp lại hoặc lấy mã nhập điểm Moodle, phải nhận diện file Excel mới hơn _review.json và lấy đúng Q17=4.5, Q18=3.5!
            reloaded_items = build_review_items_from_excel_and_json(excel_copy, json_copy, class_code="2607AVCB005")
            st_row10 = next(it for it in reloaded_items if it["row"] == 10)
            self.assertEqual((st_row10["p1"]["c"], st_row10["p1"]["o"], st_row10["p1"]["g"], st_row10["p1"]["v"]), (5, 5, 4, 4))
            self.assertEqual((st_row10["p2"]["c"], st_row10["p2"]["o"], st_row10["p2"]["g"], st_row10["p2"]["v"]), (4, 4, 3, 3))

            root = tk.Tk()
            root.withdraw()
            try:
                app = IUHWritingAISuiteApp(root)
                app.open_classes = [{
                    "code": "2607AVCB005",
                    "json_path": json_copy,
                    "excel_path": excel_copy,
                    "checked_count": 34,
                    "total_count": 34,
                }]
                app.active_class_code = "2607AVCB005"
                app.json_path_var.set(json_copy)
                app.excel1_path_var.set(excel_copy)
                app.excel2_path_var.set(excel_copy)
                # Giả sử bộ nhớ RAM đang giữ điểm cũ trước khi sửa Excel, nhưng file Excel trên đĩa mới hơn _review.json
                app.review_items = items
                grades_map, _ = app._collect_multi_class_grades()
                mssv_row10 = st_row10["mssv"]
                self.assertEqual(grades_map[mssv_row10]["q17"], "4.5")
                self.assertEqual(grades_map[mssv_row10]["q18"], "3.5")
            finally:
                root.destroy()

    def test_js_moodle_submitter_skips_already_graded_by_others_and_supports_overwrite(self):
        from app_gui import build_js_paced_grade_submitter

        sample_grades = {
            "25677201": {"ten": "Trần Quốc Duy", "q17": "3.25", "q18": "3.75", "class_code": "2607AVCB005"},
            "25682031": {"ten": "Lê Nguyễn Tấn Phong", "q17": "1.5", "q18": "3.25", "class_code": "2607AVCB005"},
        }
        js_code = build_js_paced_grade_submitter(
            sample_grades,
            min_delay_sec=3.0,
            max_delay_sec=5.0,
            class_codes=["2607AVCB005"],
        )
        # Kiểm tra cơ chế bảo vệ kép khi giữa chừng có 1 SV đã được người khác chấm Q17/Q18:
        # 1) Nhận diện trên bảng điểm (isSlotAlreadyGradedOnTable)
        self.assertIn("isSlotAlreadyGradedOnTable", js_code)
        # 2) Nhận diện trực tiếp trong form comment.php (hasExistingServerGrade) ngay trước khi POST
        self.assertIn("hasExistingServerGrade", js_code)
        self.assertIn("skipped_existing", js_code)
        self.assertIn("skipped_same", js_code)
        # 3) Tùy chọn bật/tắt ghi đè ngay trên thanh nổi Moodle (mặc định tắt để không đè điểm người khác)
        self.assertIn("iuh-chk-overwrite-existing", js_code)
        # 4) Tự động đổi dấu chấm <-> dấu phẩy nếu Moodle tiếng Việt yêu cầu định dạng số thập phân dấu phẩy
        self.assertIn("altScore", js_code)

    def test_apple_single_screen_ui_floating_popover_and_font_scaling(self):
        import tkinter as tk

        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005.")

        root = tk.Tk()
        root.withdraw()
        try:
            app = IUHWritingAISuiteApp(root)
            app._open_or_activate_class_from_file(excel_real)

            # 1. Kiểm tra Menu Lớp Dọc Nổi (Floating Popover) dùng .place() không chiếm chiều dọc cố định
            self.assertFalse(app.class_overlay_visible)
            app._toggle_class_overlay()
            self.assertTrue(app.class_overlay_visible)
            self.assertEqual(app.class_overlay_panel.winfo_manager(), "place")
            app._hide_class_overlay()
            self.assertFalse(app.class_overlay_visible)

            # 2. Kiểm tra nút tăng/giảm cỡ chữ A- / A+
            initial_pt = app.font_size_pt
            app._change_font_size(1)
            self.assertEqual(app.font_size_pt, initial_pt + 1)
            app._change_font_size(-1)
            self.assertEqual(app.font_size_pt, initial_pt)

            # 3. Kiểm tra tô đậm câu hỏi chính trong khung Đề bài (tag "prompt_question")
            pr_widget = app.part_widgets["p1"]["prompt_text"]
            self.assertIn("prompt_question", pr_widget.tag_names())
        finally:
            root.destroy()

    def test_v7_gold_standard_calibration_2607avcb005_and_006_and_examiner_reasons(self):
        from app_gui import (
            JS_SCRAPER_CODE,
            KetQuaCham,
            _calibrate_part_scores,
            _calibrate_scores,
            _is_generic_or_stale_reason,
            _synthesize_part_cogv_reasons,
        )

        # 0. Kiểm tra JS_SCRAPER_CODE đã khử xuống dòng mềm trong HTML text node và nối câu bị rớt dòng giữa chừng
        self.assertIn("createTreeWalker", JS_SCRAPER_CODE)
        self.assertIn("prevEndsMidPhrase", JS_SCRAPER_CODE)

        # 1. Huỳnh Gia Huy (P1 L003): Quên dấu '?' cuối câu hỏi ("What time do you wake up. And what your usually does in the morning.")
        # -> KHÔNG trừ Content (C=5), nhưng O=3 (comma splice), G=3 (bắt ngay từ đầu và trích dẫn cụ thể lỗi "I have breakfast is", "My Mom is cooked", "It very delicous", "what your usually does"), V=3
        pr_huy = (
            "PART 1: Write an email (about 50 words) to a friend.\n"
            "Remember to give information about:\n- what time you wake up\n- what you do in the morning\n- what you have for breakfast\n"
            "Remember to ask about:\n- what time your friend wakes up\n- what he/she usually does in the morning"
        )
        es_huy = (
            "Hi Long,\n"
            "I usually wake up at 6:00 AM . Next I wash my face and I have breakfast. I have breakfast is a bread and milk. "
            "And sometimes have a pizza , rice, fish, My Mom is cooked it. It very delicous and very well. In the free time you can go to my home.\n"
            "And you . What time do you wake up. And what your usually does in the morning.\n"
            "See you soon.\nHuy"
        )
        c, o, g, v, notes_huy = _calibrate_part_scores(pr_huy, es_huy, 5, 4, 4, 4, is_part1=True)
        self.assertEqual((c, o, g, v), (5, 3, 3, 3))
        rc, ro, rg, rv = _synthesize_part_cogv_reasons(
            pr_huy, es_huy, c, o, g, v,
            "", "Đạt: Mạch lạc, dùng từ nối tốt.", "Ngữ pháp khá (4đ)", "Đạt: Từ vựng tốt.",
            is_part1=True, notes=notes_huy,
        )
        self.assertIn("(5đ)", rc)
        self.assertIn("(3đ)", ro)
        self.assertNotIn("Đạt:", ro)
        self.assertIn("(3đ)", rg)
        self.assertIn("my mom is cooked", rg.lower())
        self.assertIn("(3đ)", rv)
        self.assertIn("delicous", rv.lower())

        # 2. Ngô Tuấn Tú (P1 L006): Lái sang kể bữa sáng (breakfast) và câu hỏi cụt "This is my opinion and what your favourite food"
        # -> C=3, G=3, V=3, và tuyệt đối không hiển thị "Trả lời đầy đủ..." ở mức 3 điểm
        pr_tu = (
            "PART 1: Write an email (about 50 words) to a friend.\n"
            "Remember to give information about:\n- what your favourite food is\n- how often you eat it\n- where you like to eat it\n"
            "Remember to ask about:\n- what your friend's favourite food is\n- where he/she likes to eat it"
        )
        es_tu = (
            "Hello, Tuan.\n"
            "I write this email because i remember you. My favourite food is meat. In the morning, i usually eat meat and egg for my breakfast. "
            "I like eat breakfast in my house because i feel comfortable and chill. This is my opinion and what your favourite food. "
            "Beaside, Where do you like eat it. I look forward to answering by you.\n"
            "See you soon!\nMax"
        )
        c_tu, o_tu, g_tu, v_tu, _ = _calibrate_part_scores(pr_tu, es_tu, 5, 4, 4, 4, is_part1=True)
        self.assertEqual((c_tu, o_tu, g_tu, v_tu), (3, 4, 3, 3))
        self.assertTrue(_is_generic_or_stale_reason("Trả lời đầy đủ thông tin món ăn yêu thích và hỏi đủ 2 câu hỏi cho bạn.", 3, "c"))
        rc_tu, _, rg_tu, rv_tu = _synthesize_part_cogv_reasons(
            pr_tu, es_tu, c_tu, o_tu, g_tu, v_tu,
            "Trả lời đầy đủ thông tin món ăn yêu thích và hỏi đủ 2 câu hỏi cho bạn.", "", "", "",
            is_part1=True,
        )
        self.assertIn("(3đ)", rc_tu)
        self.assertNotIn("Trả lời đầy đủ", rc_tu)
        self.assertIn("like eat", rg_tu.lower())
        self.assertIn("beaside", rv_tu.lower())

        # 3. Hoàng Kiều Anh (P1): Kết thư "Your friend." không tên (C=4), lặp 5 câu "I..." không từ nối (O=3), G=3 ("I after go", "I sometime plays", "What time you go")
        pr_ka = (
            "PART 1: Write an email (about 50 words) to a friend.\n"
            "Remember to give information about:\n- where you go after school\n- what you do in the evening\n- what time you go to bed\n"
            "Remember to ask about:\n- what your friend does after school\n- what time he/she goes to bed"
        )
        es_ka = (
            "Hi Hoa,\n"
            "How are you? I after go to school in the morning. I study English with my friends . My teacher is very happy. "
            "I go home at 6 o'clock . I sometime plays video games in the evening. I go to bed at 10 o'clock. "
            "What do you do after school? What time you go to the bed?\n"
            "Your friend."
        )
        c_ka, o_ka, g_ka, v_ka, _ = _calibrate_part_scores(pr_ka, es_ka, 5, 4, 4, 3, is_part1=True)
        self.assertEqual((c_ka, o_ka, g_ka, v_ka), (4, 3, 3, 3))

        # 4. Nguyễn Minh Dương (P1 L008): Chỉ hỏi "Would you like to my house?" và "OK for you?" -> không tính là câu hỏi chủ đề (C=3), xuống dòng từng câu (O=3), G=3, V=2 ("tow rooms", "bead room", "My love my house", "Reply soon me")
        pr_duong = (
            "PART 1: Write an email (about 50 words) to a friend.\n"
            "Remember to give information about:\n- where you live\n- how many rooms your house has\n- what your favorite room is\n"
            "Remember to ask about:\n- where your friend lives\n- what his/her favorite room is"
        )
        es_duong = (
            "Hi\nWould you like to my house?\nMy house have tow rooms a bead room a living room\n"
            "My favorite room is bed room because it is clean and comfortable\n"
            "My love my house because I can talk with family everyday\n"
            "OK for you?\nReply soon me\nSee you\nDuong"
        )
        c_d, o_d, g_d, v_d, _ = _calibrate_part_scores(pr_duong, es_duong, 5, 4, 4, 3, is_part1=True, format_type="XUONG_DONG_TUNG_CAU")
        self.assertEqual((c_d, o_d, g_d, v_d), (3, 3, 3, 2))

        # 5. Phạm Thị Bình Chi (P2 L007): Mở đầu "Thanks you for inviting" VÀ trả lời "And I can bring some food and drinks" (đảo vai chủ-khách ở 2 ý -> C=3), xuống dòng từng câu (O=3), G=3
        pr_chi = (
            "PART 2: Write a reply (about 50 words) to the given email.\n"
            "Hi,\nThanks for inviting me to your picnic! I would love to come. What time should I be there? "
            "Is your mother going to take us to the park? Can I bring some food or drinks?\n"
            "Looking forward to relaxing with you!\nBen"
        )
        es_chi = (
            "Hi, Ben\nThanks you for inviting.\nI am happy.\nI am free time at 8:00 am. Is that you ok?\n"
            "My mother is going to take us to the park.\nAnd I can bring some food and drinks.\n"
            "See you soon.\nBest friend,\nChi"
        )
        c_chi, o_chi, g_chi, v_chi, _ = _calibrate_part_scores(pr_chi, es_chi, 5, 3, 4, 3, is_part1=False, format_type="XUONG_DONG_TUNG_CAU")
        self.assertEqual((c_chi, o_chi, g_chi, v_chi), (3, 3, 3, 3))

        # 6. Huỳnh Đan Trường (24684671): P1 biến toàn bộ ý bản thân thành câu hỏi cho Hung + P2 tự đóng vai Hung trả lời cho Truong -> Cả P1 và P2 đều khóa 0-1-1-1
        sv_truong = {
            "mssv": "24684671",
            "p1": {
                "prompt": pr_huy,
                "essay": (
                    "Hi Hung,\nI'm write this letter for you because we are best friend, i want to hear you talk about your day. "
                    "What time do you wake up everyday? I'm always wake up at 6 A.M and i hope you also wake up at that time. "
                    "What do you do in the morning? Besides study, do you play sports? And the last question i want to know from you, "
                    "What do you have for breakfast? Rice or Pho? That's all, i really miss you and when we meet, i hope we can play soccer together.\n"
                    "See you soon!\nTruong"
                ),
                "format_type": "BINH_THUONG",
            },
            "p2": {
                "prompt": pr_chi,
                "essay": (
                    "Hi Truong,\nI'm very happy when i receive your letter. Every morning, i usually wake up at 6:30 A.M. "
                    "After that, i brush my teeth and play badminton with my father. For breakfast, i often eat bread and drink milk.\n"
                    "See you soon!\nHung"
                ),
                "format_type": "BINH_THUONG",
            },
        }
        kq_truong = _calibrate_scores(
            sv_truong,
            KetQuaCham(
                p1_content=4, p1_org=3, p1_grammar=4, p1_vocab=4,
                p2_content=4, p2_org=4, p2_grammar=4, p2_vocab=4,
            ),
        )
        self.assertEqual((kq_truong.p1_content, kq_truong.p1_org, kq_truong.p1_grammar, kq_truong.p1_vocab), (0, 1, 1, 1))
        self.assertEqual((kq_truong.p2_content, kq_truong.p2_org, kq_truong.p2_grammar, kq_truong.p2_vocab), (0, 1, 1, 1))
        self.assertIn("hồi đáp cho chính bài Part 1", kq_truong.p2_c_reason)

    def test_marking_rule_profiles_cambridge_and_oxford_selection(self):
        import tkinter as tk
        from app_gui import (
            DEFAULT_RULE_PROFILE,
            MARKING_RULE_PROFILES,
            get_marking_rules_path,
            load_marking_rules,
        )

        # 1. Mặc định là Cuối kì - Sách Cambridge (MARKING_RULES_CAMBRIDGE_CUOI_KI.md)
        self.assertEqual(DEFAULT_RULE_PROFILE, "CAMBRIDGE_CUOI_KI")
        self.assertIn("CAMBRIDGE_CUOI_KI", MARKING_RULE_PROFILES)
        self.assertIn("CAMBRIDGE_GIUA_KI", MARKING_RULE_PROFILES)
        self.assertIn("OXFORD_GIUA_KI", MARKING_RULE_PROFILES)
        self.assertIn("OXFORD_CUOI_KI", MARKING_RULE_PROFILES)

        cam_final_path = get_marking_rules_path("CAMBRIDGE_CUOI_KI")
        self.assertTrue(cam_final_path.endswith("MARKING_RULES_CAMBRIDGE_CUOI_KI.md"))
        cam_content = load_marking_rules("CAMBRIDGE_CUOI_KI")
        self.assertIn("CUỐI KÌ — SÁCH CAMBRIDGE", cam_content)
        self.assertTrue(os.path.exists(cam_final_path))
        # File tên chung chung cũ MARKING_RULES_IUH.md đã được tự động dọn bỏ
        self.assertFalse(os.path.exists(os.path.join(os.path.dirname(cam_final_path), "MARKING_RULES_IUH.md")))

        # 2. Kiểm tra chuyển đổi qua lại giữa Cuối kì Cambridge <-> Giữa kì Oxford <-> Cuối kì Oxford trên GUI
        ox_mid_path = get_marking_rules_path("OXFORD_GIUA_KI")
        ox_final_path = get_marking_rules_path("OXFORD_CUOI_KI")
        existed_ox_mid = os.path.exists(ox_mid_path)
        existed_ox_final = os.path.exists(ox_final_path)
        root = tk.Tk()
        root.withdraw()
        try:
            app = IUHWritingAISuiteApp(root)
            self.assertEqual(app.rule_profile, "CAMBRIDGE_CUOI_KI")
            self.assertIn("MARKING_RULES_CAMBRIDGE_CUOI_KI.md", app.rules_info_var.get())

            app.rule_profile_label_var.set("Giữa kì - Sách Oxford")
            app._on_rule_profile_selected()
            self.assertEqual(app.rule_profile, "OXFORD_GIUA_KI")
            self.assertIn("MARKING_RULES_OXFORD_GIUA_KI.md", app.rules_info_var.get())
            self.assertIn("GIỮA KÌ — SÁCH OXFORD", app.rubric_content)

            app.rule_profile_label_var.set("Cuối kì - Sách Oxford")
            app._on_rule_profile_selected()
            self.assertEqual(app.rule_profile, "OXFORD_CUOI_KI")
            self.assertIn("MARKING_RULES_OXFORD_CUOI_KI.md", app.rules_info_var.get())
            self.assertIn("CUỐI KÌ — SÁCH OXFORD", app.rubric_content)

            app.rule_profile_label_var.set("Cuối kì - Sách Cambridge")
            app._on_rule_profile_selected()
            self.assertEqual(app.rule_profile, "CAMBRIDGE_CUOI_KI")
            self.assertIn("MARKING_RULES_CAMBRIDGE_CUOI_KI.md", app.rules_info_var.get())
        finally:
            root.destroy()
            if not existed_ox_mid and os.path.exists(ox_mid_path):
                os.remove(ox_mid_path)
            if not existed_ox_final and os.path.exists(ox_final_path):
                os.remove(ox_final_path)

    def test_prompt_clustered_batching_preserves_student_order_and_single_class_moodle_copy(self):
        import re
        import tkinter as tk
        from unittest.mock import patch
        import app_gui
        from app_gui import (
            AVAILABLE_MODELS,
            DEFAULT_MODEL,
            DEFAULT_THINKING_LEVEL,
            KetQuaCham1PhanSV,
            KetQuaChamCumMaDe,
            _build_prompt_clusters_for_part,
            _format_gold_examples_for_code,
            _load_gold_standard_examples_by_code,
            _parse_api_keys,
        )

        # 1. Kiểm tra nâng cấp model mặc định gemini-3.8-flash + thinking_level=high + xoay vòng đa API Key
        self.assertEqual(DEFAULT_MODEL, "gemini-3.8-flash")
        self.assertEqual(DEFAULT_THINKING_LEVEL, "high")
        self.assertIn("gemini-3.8-flash", AVAILABLE_MODELS)
        self.assertIn("gemini-3.6-flash", AVAILABLE_MODELS)
        self.assertIn("gemini-2.5-pro", AVAILABLE_MODELS)
        self.assertIn("gemini-2.5-flash", AVAILABLE_MODELS)
        self.assertEqual(_parse_api_keys("keyA, keyB; keyC\nkeyD"), ["keyA", "keyB", "keyC", "keyD"])

        # 2. Kiểm tra bộ bài mẫu chuẩn (Gold Standard Few-Shot) phủ đủ 20/20 mã đề P1.L001..L010 và P2.L001..L010
        gold_map = _load_gold_standard_examples_by_code()
        for i in range(1, 11):
            p1_code = f"AVCB.FT.W.P1.L{i:03d}"
            p2_code = f"AVCB.FT.W.P2.L{i:03d}"
            self.assertIn(p1_code, gold_map)
            self.assertIn(p2_code, gold_map)
            self.assertTrue(len(gold_map[p1_code]) >= 1)
            self.assertTrue(len(gold_map[p2_code]) >= 1)
        formatted_l001 = _format_gold_examples_for_code("AVCB.FT.W.P1.L001", is_part1=True)
        self.assertIn("Điểm chuẩn GV:", formatted_l001)
        self.assertIn("AVCB.FT.W.P1.L001", formatted_l001)

        # 3. Kiểm tra gom cụm theo mã đề (<= 2 cụm P1 + <= 2 cụm P2 = tối đa 4 lượt gọi API/lớp)
        #    và bảo toàn tuyệt đối 100% số thứ tự sinh viên (row) trên giao diện sau khi chấm
        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005.")

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_copy = os.path.join(tmpdir, "PHIEU_CHAM_2607AVCB005.xlsx")
            json_copy = os.path.join(tmpdir, "bai_lam_iuh_2607AVCB005.json")
            shutil.copy2(excel_real, excel_copy)
            shutil.copy2(json_real, json_copy)

            before_items = build_review_items_from_excel_and_json(excel_copy, json_copy)
            expected_row_mssv_order = [(it["row"], it["mssv"], it["ten"]) for it in before_items]

            with open(json_copy, "r", encoding="utf-8") as f:
                ds_raw = json.load(f)
            map_raw = {normalize_mssv(x["mssv"]): x for x in ds_raw if normalize_mssv(x.get("mssv"))}
            tasks_sim = [
                (it["row"], it["mssv"], it["ten"], map_raw[it["mssv"]])
                for it in before_items
                if it["mssv"] in map_raw
            ]

            p1_clusters = _build_prompt_clusters_for_part(tasks_sim, "p1")
            p2_clusters = _build_prompt_clusters_for_part(tasks_sim, "p2")
            self.assertGreaterEqual(len(p1_clusters), 1)
            self.assertLessEqual(len(p1_clusters), 2)
            self.assertGreaterEqual(len(p2_clusters), 1)
            self.assertLessEqual(len(p2_clusters), 2)
            self.assertLessEqual(len(p1_clusters) + len(p2_clusters), 4)

            # Giả lập _call_gemini_interactions để kiểm chứng:
            # - Chỉ gọi đúng <= 4 lần cho cả lớp 34 SV
            # - Trong prompt gửi đi có chia ngăn mã đề, có Bài mẫu Gold Standard và có Đối chiếu chéo Part 1
            # - Ngay cả khi AI trả danh sách theo thứ tự mã đề bị xáo trộn, chay_cham_ai_tao_ban_duyet vẫn trả về đúng thứ tự dòng Excel gốc 100%!
            prompts_sent: list[str] = []

            def fake_call_gemini(api_key, input_data, schema_dict, preferred_model=DEFAULT_MODEL, timeout_sec=120, thinking_level=DEFAULT_THINKING_LEVEL):
                prompt_str = input_data if isinstance(input_data, str) else str(input_data)
                prompts_sent.append(prompt_str)
                mssvs_in_prompt = re.findall(r"MSSV:\s*(\d{7,10})", prompt_str)
                # Cố tình đảo ngược thứ tự trả về từ AI để chứng minh bộ điều phối map lại đúng thứ tự SV ban đầu
                resp_items = [
                    KetQuaCham1PhanSV(
                        mssv=m,
                        content=4,
                        org=4,
                        grammar=3,
                        vocab=3,
                        c_reason="Đạt đủ ý chính.",
                        o_reason="Bố cục rõ ràng.",
                        g_reason="Sai nhẹ ngữ pháp.",
                        v_reason="Từ vựng cơ bản.",
                        nhan_xet_ngan="Đạt yêu cầu.",
                    )
                    for m in reversed(mssvs_in_prompt)
                ]
                return KetQuaChamCumMaDe(danh_sach=resp_items).model_dump_json()

            with patch.object(app_gui, "_call_gemini_interactions", side_effect=fake_call_gemini):
                summary = chay_cham_ai_tao_ban_duyet(
                    api_key="test_key_1, test_key_2",
                    json_path=json_copy,
                    excel_path=excel_copy,
                    rubric_text=DEFAULT_RUBRIC,
                    only_ungraded=False,
                )

            self.assertLessEqual(len(prompts_sent), 4)
            self.assertTrue(any("BÀI MẪU CHUẨN GIẢNG VIÊN (GOLD STANDARD FEW-SHOT)" in p for p in prompts_sent))
            self.assertTrue(any("Đối chiếu chéo Part 1 của SV" in p for p in prompts_sent))

            after_items = summary["items"]
            actual_row_mssv_order = [(it["row"], it["mssv"], it["ten"]) for it in after_items]
            self.assertEqual(actual_row_mssv_order, expected_row_mssv_order)

            # 4. Kiểm tra nút 'Sao chép mã nhập Moodle' mặc định chỉ đóng gói đúng 1 lớp đang chọn (only_active_class=True)
            root = tk.Tk()
            root.withdraw()
            try:
                app = IUHWritingAISuiteApp(root)
                app.open_classes = [
                    {
                        "code": "2607AVCB005",
                        "json_path": json_copy,
                        "excel_path": excel_copy,
                        "checked_count": 34,
                        "total_count": 34,
                    },
                    {
                        "code": "2607AVCB999",
                        "json_path": "",
                        "excel_path": "",
                        "checked_count": 2,
                        "total_count": 2,
                    },
                ]
                sess_005 = app._get_or_create_class_session("2607AVCB005")
                sess_005["review_items"] = after_items
                sess_999 = app._get_or_create_class_session("2607AVCB999")
                items_999 = [
                    {
                        "row": 10,
                        "mssv": "25999001",
                        "ten": "Sinh Viên Lớp 999 A",
                        "checked": True,
                        "absent": False,
                        "graded": True,
                        "note": "",
                        "p1": {"c": 5, "o": 4, "g": 4, "v": 3},
                        "p2": {"c": 4, "o": 4, "g": 3, "v": 3},
                    },
                    {
                        "row": 11,
                        "mssv": "25999002",
                        "ten": "Sinh Viên Lớp 999 B",
                        "checked": True,
                        "absent": False,
                        "graded": True,
                        "note": "",
                        "p1": {"c": 3, "o": 3, "g": 3, "v": 3},
                        "p2": {"c": 3, "o": 3, "g": 3, "v": 3},
                    },
                ]
                sess_999["review_items"] = items_999

                # Khi đang đứng ở lớp 2607AVCB999: mặc định chỉ lấy đúng 2 SV của lớp 2607AVCB999, KHÔNG gộp 34 SV của lớp 2607AVCB005!
                app.active_class_code = "2607AVCB999"
                app.review_items = items_999
                app.json_path_var.set("")
                app.excel1_path_var.set("")
                app.excel2_path_var.set("")

                grades_single, classes_single = app._collect_multi_class_grades()
                self.assertEqual(classes_single, ["2607AVCB999"])
                self.assertEqual(len(grades_single), 2)
                self.assertIn("25999001", grades_single)
                self.assertNotIn(after_items[0]["mssv"], grades_single)

                # Khi truyền only_active_class=False thì vẫn hỗ trợ lấy gộp cả 2 lớp
                grades_all, classes_all = app._collect_multi_class_grades(only_active_class=False)
                self.assertIn("2607AVCB005", classes_all)
                self.assertIn("2607AVCB999", classes_all)
                self.assertGreater(len(grades_all), 30)
            finally:
                root.destroy()

    def test_multi_model_pass2_arbitrator_and_live_pipeline_inspector(self):
        import re
        import tkinter as tk
        from unittest.mock import patch
        import app_gui
        from app_gui import (
            AVAILABLE_MODELS,
            DEFAULT_ARBITRATOR_MODEL,
            DEFAULT_MODEL,
            DEFAULT_THINKING_LEVEL,
            KetQuaCham1PhanSV,
            KetQuaChamCumMaDe,
            KetQuaHoiDongTrongTai,
            KetQuaTrongTaiSV,
            _emit_thought_summary,
        )

        # 1. Kiểm tra danh sách model thế hệ mới (gemini-3.8-flash + gemini-3.7-flash + gemini-3.1-pro-preview + gemini-4-argon)
        self.assertEqual(DEFAULT_MODEL, "gemini-3.8-flash")
        self.assertEqual(DEFAULT_ARBITRATOR_MODEL, "gemini-3.7-flash")
        self.assertIn("gemini-3.8-flash", AVAILABLE_MODELS)
        self.assertIn("gemini-3.7-flash", AVAILABLE_MODELS)
        self.assertIn("gemini-3.1-pro-preview", AVAILABLE_MODELS)
        self.assertIn("gemini-4-argon", AVAILABLE_MODELS)

        excel_real, json_real = self._find_2607_files()
        if not excel_real:
            self.skipTest("Không tìm thấy dữ liệu lớp 2607AVCB005.")

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_copy = os.path.join(tmpdir, "PHIEU_CHAM_2607AVCB005.xlsx")
            json_copy = os.path.join(tmpdir, "bai_lam_iuh_2607AVCB005.json")
            shutil.copy2(excel_real, excel_copy)
            shutil.copy2(json_real, json_copy)

            before_items = build_review_items_from_excel_and_json(excel_copy, json_copy)
            expected_order = [(it["row"], it["mssv"], it["ten"]) for it in before_items]

            models_called: list[str] = []
            stage_events: list[tuple[int, str, str, str]] = []

            def record_stage(step_idx: int, state: str, badge_txt: str, detail_msg: str):
                stage_events.append((step_idx, state, badge_txt, detail_msg))

            def fake_multi_model_gemini(
                api_key,
                input_data,
                schema_dict,
                preferred_model=DEFAULT_MODEL,
                timeout_sec=120,
                thinking_level=DEFAULT_THINKING_LEVEL,
            ):
                models_called.append(preferred_model)
                prompt_str = input_data if isinstance(input_data, str) else str(input_data)
                _emit_thought_summary(preferred_model, f"Đang suy luận chuyên sâu với {preferred_model} (thinking={thinking_level})...")

                # Phân biệt lượt gọi Vòng 2 Trọng tài (KetQuaHoiDongTrongTai) vs Vòng 1 Cụm Mã Đề (KetQuaChamCumMaDe)
                if "nhan_xet_hoi_dong" in (schema_dict.get("properties") or {}):
                    mssvs_arb = re.findall(r"MSSV:\s*(\d{7,10})", prompt_str)
                    unique_mssvs = list(dict.fromkeys(mssvs_arb))
                    arb_items = [
                        KetQuaTrongTaiSV(
                            mssv=m,
                            verdict="Đồng thuận Quy tắc Khoa: Chốt điểm sau phản biện 2 vòng",
                            p1_content=3,
                            p1_org=3,
                            p1_grammar=3,
                            p1_vocab=3,
                            p1_c_reason="Đạt 2/3 ý, đã thẩm định vòng 2.",
                            p1_o_reason="Bố cục cơ bản.",
                            p1_g_reason="Có lỗi ngữ pháp nhẹ.",
                            p1_v_reason="Từ vựng cơ bản.",
                            p2_content=3,
                            p2_org=3,
                            p2_grammar=3,
                            p2_vocab=3,
                            p2_c_reason="Trả lời đạt 2/3 câu hỏi.",
                            p2_o_reason="Bố cục cơ bản.",
                            p2_g_reason="Có lỗi ngữ pháp nhẹ.",
                            p2_v_reason="Từ vựng cơ bản.",
                        )
                        for m in unique_mssvs
                    ]
                    return KetQuaHoiDongTrongTai(
                        nhan_xet_hoi_dong="Đã phúc khảo độc lập các hồ sơ nhạy cảm.",
                        danh_sach=arb_items,
                    ).model_dump_json()

                mssvs_in_prompt = re.findall(r"MSSV:\s*(\d{7,10})", prompt_str)
                resp_items = []
                for idx_m, m in enumerate(mssvs_in_prompt):
                    # Cố tình cho vài SV đầu có điểm Content=2 hoặc Grammar=4 để kích hoạt bộ lọc Trọng tài Vòng 2
                    c_val = 2 if idx_m == 0 else 4
                    g_val = 4 if idx_m < 2 else 3
                    resp_items.append(
                        KetQuaCham1PhanSV(
                            mssv=m,
                            content=c_val,
                            org=4,
                            grammar=g_val,
                            vocab=3,
                            c_reason="Đạt ý chính.",
                            o_reason="Bố cục rõ ràng.",
                            g_reason="Ngữ pháp khá.",
                            v_reason="Từ vựng cơ bản.",
                            nhan_xet_ngan="Đã chấm Vòng 1.",
                        )
                    )
                return KetQuaChamCumMaDe(danh_sach=resp_items).model_dump_json()

            with patch.object(app_gui, "_call_gemini_interactions", side_effect=fake_multi_model_gemini):
                summary = chay_cham_ai_tao_ban_duyet(
                    api_key="test_key_1, test_key_2",
                    json_path=json_copy,
                    excel_path=excel_copy,
                    rubric_text=DEFAULT_RUBRIC,
                    model_name="gemini-3.8-flash",
                    only_ungraded=False,
                    enable_pass2_arbitrator=True,
                    arbitrator_model="gemini-3.7-flash",
                    stage_cb=record_stage,
                )

            # Tổng số lượt gọi API: <= 4 cụm Vòng 1 (gemini-3.8-flash) + 1 lượt Trọng tài Vòng 2 (gemini-3.7-flash) = tối đa 5 lượt!
            self.assertLessEqual(summary["api_calls_used"], 5)
            self.assertIn("gemini-3.8-flash", models_called)
            self.assertIn("gemini-3.7-flash", models_called)

            # Kiểm tra đủ 5 bước tiến trình suy luận (Bước 1 -> Bước 5) + sự kiện thought summary
            steps_seen = {step_idx for step_idx, _, _, _ in stage_events}
            for req_step in (0, 1, 2, 3, 4, 5):
                self.assertIn(req_step, steps_seen)

            after_items = summary["items"]
            actual_order = [(it["row"], it["mssv"], it["ten"]) for it in after_items]
            self.assertEqual(actual_order, expected_order)
            self.assertTrue(any("Thẩm định 2 vòng (3.8-flash ⇄ 3.7-flash)" in (it.get("council_badge") or "") for it in after_items))
            self.assertTrue(any("Chuẩn hội đồng (3.8-flash + Rule Engine)" in (it.get("council_badge") or "") for it in after_items))

            # Kiểm tra giao diện Floating Progress Bar (HUD nổi .place) + Line tiến trình (0% -> 100%)
            root = tk.Tk()
            root.withdraw()
            try:
                app = IUHWritingAISuiteApp(root)
                self.assertEqual(app.btn_toolbar_pipeline.cget("text"), "Tiến trình")
                app._show_pipeline_hud()
                self.assertTrue(app.pipeline_hud_visible)
                self.assertEqual(app.pipeline_hud_frame.winfo_manager(), "place")
                app._reset_pipeline_hud_steps("Bắt đầu kiểm thử HUD...")
                self.assertEqual(app.pipeline_progress_pct, 0)
                self.assertEqual(app.lbl_pipeline_pct.cget("text"), "0%")
                for step_idx, state, badge_txt, detail_msg in stage_events:
                    app._update_pipeline_hud_step(step_idx, state, badge_txt, detail_msg)
                self.assertEqual(app.pipeline_progress_pct, 100)
                self.assertEqual(app.lbl_pipeline_pct.cget("text"), "100%")
                node_titles = [app.pipeline_step_widgets[i]["base_title"] for i in range(1, 6)]
                self.assertEqual(
                    node_titles,
                    ["Phân đề", "Chấm vòng 1", "Đối chiếu chuẩn", "Thẩm định chéo", "Hoàn tất"],
                )
                hud_log = app.pipeline_stream_text.get("1.0", tk.END)
                self.assertIn("Bước 4", hud_log)
                self.assertIn("Suy luận nội bộ", hud_log)
                app._toggle_pipeline_detail_drawer()
                self.assertTrue(app.pipeline_detail_expanded)
                app._toggle_pipeline_detail_drawer()
                self.assertFalse(app.pipeline_detail_expanded)
                app._hide_pipeline_hud()
                self.assertFalse(app.pipeline_hud_visible)
                self.assertEqual(app.pipeline_hud_frame.winfo_manager(), "")
            finally:
                root.destroy()


if __name__ == '__main__':
    unittest.main()
