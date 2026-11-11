import datetime
import os
import sys
import time
import tkinter as tk
from tkinter import filedialog
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
import pyautogui
import pyperclip

pyautogui.FAILSAFE = True


def normalize_mssv(val) -> str:
    if val is None:
        return ''
    s = str(val).strip()
    if s.lower().startswith('sv_'):
        s = s[3:].strip()
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit() and 6 <= len(s) <= 12:
        return s
    return ''


def format_score(val) -> str:
    if val is None:
        return '0'
    s = str(val).strip().replace(',', '.')
    if not s:
        return '0'
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return f'{f:g}'
    except ValueError:
        return '0'


def _compute_component_avg(row: tuple, comp_cols: list[int] | None):
    """Tính dự phòng (Content + Org + Grammar + Vocab) / 4 khi công thức Excel chưa được cache."""
    if not comp_cols or len(comp_cols) != 4:
        return None
    vals = []
    has_any_number = False
    for c_idx in comp_cols:
        if c_idx is None or c_idx >= len(row):
            return None
        v = row[c_idx]
        if v is None or str(v).strip() == '':
            vals.append(0.0)
        else:
            try:
                vals.append(float(str(v).strip().replace(',', '.')))
                has_any_number = True
            except ValueError:
                return None
    if not has_any_number:
        return None
    return sum(vals) / 4.0


def detect_excel_schema(all_rows: list[tuple], override_cols: dict | None = None) -> dict:
    """Tự động nhận diện vị trí các cột trong bảng điểm Excel (0-based index)."""
    col_mssv = None
    col_ten_gop = None
    col_ho_dem = None
    col_ten_rieng = None
    col_q17_primary = None
    col_q17_fallback = None
    col_q18_primary = None
    col_q18_fallback = None
    content_cols = []
    col_thong_nhat = None
    col_comment = None

    scan_limit = min(20, len(all_rows))
    for r in all_rows[:scan_limit]:
        if not r:
            continue
        for j, cell in enumerate(r):
            if cell is None:
                continue
            txt = ' '.join(str(cell).strip().lower().split())
            if not txt:
                continue

            if col_mssv is None and (
                txt in (
                    'mssv', 'mã sv', 'ma sv', 'mã số sv', 'ma so sv',
                    'mã số sinh viên', 'tên tài khoản', 'ten tai khoan',
                    'student id', 'mã số id'
                )
                or 'mssv' in txt
                or 'mã sv' in txt
            ):
                col_mssv = j

            if col_ten_gop is None and txt in (
                'họ và tên', 'ho va ten', 'họ tên', 'ho ten',
                'full name', 'fullname', 'student name'
            ):
                col_ten_gop = j
            elif col_ho_dem is None and txt in (
                'họ và tên đệm', 'ho va ten dem', 'họ đệm', 'ho dem',
                'họ', 'ho', 'first name', 'last name'
            ):
                col_ho_dem = j
            elif col_ten_rieng is None and txt in ('tên', 'ten', 'name'):
                col_ten_rieng = j

            if txt == 'content' and j not in content_cols:
                content_cols.append(j)

            if col_q17_primary is None and (
                'final score 1' in txt
                or txt in ('q17', 'điểm q17', 'diem q17', 'điểm part 1', 'diem part 1', 'task 1', 'writing 1')
            ):
                col_q17_primary = j
            elif col_q17_fallback is None and ('part 1' in txt or 'raw score 1' in txt):
                col_q17_fallback = j

            if col_q18_primary is None and (
                'final score 2' in txt
                or txt in ('q18', 'điểm q18', 'diem q18', 'điểm part 2', 'diem part 2', 'task 2', 'writing 2')
            ):
                col_q18_primary = j
            elif col_q18_fallback is None and ('part 2' in txt or 'raw score 2' in txt):
                col_q18_fallback = j

            if col_thong_nhat is None and ('điểm thống nhất' in txt or 'diem thong nhat' in txt):
                col_thong_nhat = j

            if col_comment is None and (
                'nhận xét' in txt or 'nhan xet' in txt or 'ai comment' in txt or 'cảnh báo' in txt
            ):
                col_comment = j

    col_q17 = col_q17_primary if col_q17_primary is not None else col_q17_fallback
    col_q18 = col_q18_primary if col_q18_primary is not None else col_q18_fallback

    if col_mssv is None and all_rows:
        max_cols = max((len(r) for r in all_rows if r), default=0)
        best_col, best_count = None, 0
        for j in range(max_cols):
            cnt = sum(1 for r in all_rows if r and j < len(r) and normalize_mssv(r[j]) != '')
            if cnt > best_count:
                best_col, best_count = j, cnt
        if best_count >= 1:
            col_mssv = best_col

    col_ten = col_ten_gop if col_ten_gop is not None else col_ten_rieng
    if col_ten is None and col_mssv is not None:
        col_ten = col_mssv + 1

    content_cols.sort()
    if len(content_cols) >= 2:
        p1_cols = [content_cols[0] + k for k in range(4)]
        p2_cols = [content_cols[1] + k for k in range(4)]
    elif col_mssv == 1:
        p1_cols = [3, 4, 5, 6]
        p2_cols = [9, 10, 11, 12]
    elif col_mssv == 3:
        p1_cols = [5, 6, 7, 8]
        p2_cols = [11, 12, 13, 14]
    else:
        p1_cols = None
        p2_cols = None

    if col_comment is None:
        if col_thong_nhat is not None:
            col_comment = col_thong_nhat + 1
        elif col_mssv == 1:
            col_comment = 18  # Cột S (1-based: 19)
        else:
            col_comment = 20  # Cột U (1-based: 21)

    cols = {
        'mssv': col_mssv,
        'ten': col_ten,
        'ho_dem': col_ho_dem if (col_ten_gop is None and col_ten_rieng is not None) else None,
        'q17': col_q17,
        'q18': col_q18,
        'p1_cols': p1_cols,
        'p2_cols': p2_cols,
        'comment': col_comment,
    }
    if override_cols:
        cols.update(override_cols)
    return cols


def detect_columns_and_parse(path: str, override_cols: dict | None = None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    cols = detect_excel_schema(all_rows, override_cols)

    students = []
    c_mssv = cols['mssv']
    c_ten = cols['ten']
    c_ho_dem = cols['ho_dem']
    c_q17 = cols['q17']
    c_q18 = cols['q18']
    p1_cols = cols.get('p1_cols')
    p2_cols = cols.get('p2_cols')

    if c_mssv is not None:
        for row_idx, r in enumerate(all_rows, 1):
            if not r or c_mssv >= len(r):
                continue
            mssv = normalize_mssv(r[c_mssv])
            if not mssv:
                continue

            if c_ho_dem is not None and c_ten is not None and c_ho_dem < len(r) and c_ten < len(r):
                ho = str(r[c_ho_dem] or '').strip()
                ten_r = str(r[c_ten] or '').strip()
                ten = f'{ho} {ten_r}'.strip()
            elif c_ten is not None and c_ten < len(r) and r[c_ten] is not None:
                ten = str(r[c_ten]).strip()
            else:
                ten = ''

            raw_q17 = r[c_q17] if (c_q17 is not None and c_q17 < len(r)) else None
            comp_q17 = _compute_component_avg(r, p1_cols)
            if comp_q17 is not None and (c_q17 is None or (p1_cols and c_q17 > p1_cols[-1])):
                raw_q17 = comp_q17

            raw_q18 = r[c_q18] if (c_q18 is not None and c_q18 < len(r)) else None
            comp_q18 = _compute_component_avg(r, p2_cols)
            if comp_q18 is not None and (c_q18 is None or (p2_cols and c_q18 > p2_cols[-1])):
                raw_q18 = comp_q18

            q17 = format_score(raw_q17)
            q18 = format_score(raw_q18)
            students.append({
                'row': row_idx,
                'mssv': mssv,
                'ten': ten,
                'q17': q17,
                'q18': q18,
            })

    return cols, students


def filter_start_students(students: list[dict], user_input: str) -> list[dict]:
    s = user_input.strip()
    if not s:
        return students
    mssv_cand = normalize_mssv(s)
    if mssv_cand:
        for idx, st in enumerate(students):
            if st['mssv'] == mssv_cand:
                return students[idx:]
    if s.isdigit():
        min_r = int(s)
        return [st for st in students if st['row'] >= min_r]
    return students


def tim_va_dan(tu_khoa: str, doi_sau_dan: float = 1.0, speed_mult: float = 1.0):
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.8 * speed_mult)
    pyperclip.copy(tu_khoa)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(doi_sau_dan * speed_mult)
    pyautogui.press('esc')
    time.sleep(0.8 * speed_mult)


def cham_mot_cau_web(mssv: str, diem, speed_mult: float = 1.0):
    # 1. Tìm dòng sinh viên qua mã MSSV (khớp cả cột Tên tài khoản <MSSV>, Mã số ID sv_<MSSV> và email <MSSV>@iuh.local)
    tim_va_dan(str(mssv).strip(), speed_mult=speed_mult)
    # 2. Tìm chữ "Yêu cầu chấm điểm" đầu tiên trên dòng và mở Tầng 2
    tim_va_dan('Yêu cầu chấm điểm', speed_mult=speed_mult)
    pyautogui.press('enter')
    time.sleep(4.5 * speed_mult)
    # 3. Tìm chữ "Tạo bình luận" và mở Tầng 3
    tim_va_dan('Tạo bình luận', speed_mult=speed_mult)
    pyautogui.press('enter')
    time.sleep(5.0 * speed_mult)
    # 4. Tìm chữ "trên 5,00" và lùi 1 nhịp (Shift+Tab) vào ô Điểm
    tim_va_dan('trên 5,00', speed_mult=speed_mult)
    pyautogui.hotkey('shift', 'tab')
    time.sleep(0.8 * speed_mult)
    # 5. Xóa sạch điểm cũ an toàn và nhập điểm mới
    pyautogui.press('backspace', presses=5, interval=0.1)
    pyautogui.press('delete', presses=5, interval=0.1)
    pyautogui.write(str(diem), interval=0.15)
    time.sleep(2.0 * speed_mult)
    # 6. Nhấn Enter để Lưu trữ (Tầng 3 tự đóng)
    pyautogui.press('enter')
    time.sleep(4.0 * speed_mult)
    # 7. Đóng Tầng 2 (Ctrl+W) và tải lại Bảng điểm (F5) để xóa chữ "Yêu cầu chấm điểm" vừa chấm
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(1.5 * speed_mult)
    pyautogui.press('f5')
    time.sleep(4.0 * speed_mult)


def ask_column_idx(prompt_text: str, default_idx: int | None = None) -> int:
    while True:
        default_letter = get_column_letter(default_idx + 1) if default_idx is not None else ''
        suffix = f' (Mac dinh: {default_letter})' if default_letter else ''
        raw = input(f'{prompt_text}{suffix}: ').strip().upper()
        if not raw and default_idx is not None:
            return default_idx
        try:
            return column_index_from_string(raw) - 1
        except ValueError:
            print('   [!] Vui long nhap chu cai cot Excel hop le (vi du: B, D, I, K, O, Q).')


def col_label(idx: int | None) -> str:
    return f'Cot {get_column_letter(idx + 1)}' if idx is not None else 'Chua ro'


def main():
    print('=================================================================')
    print('     CONG CU TU DONG NHAP DIEM WRITING IUH (DA TRINH DUYET)      ')
    print('=================================================================')
    print('1. Dang mo hop thoai chon file Excel...')
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    path = filedialog.askopenfilename(
        title='Chon file Excel diem Writing (Ban Copy)',
        filetypes=[('Excel files', '*.xlsx')],
    )
    if not path:
        print('\n[!] Ban chua chon file! Chuong trinh ket thuc.')
        input('Nhan phim ENTER de dong cua so...')
        sys.exit(0)

    print(f'-> Da chon file: {path}')
    try:
        cols, students = detect_columns_and_parse(path)
    except Exception as e:
        print(f'\n[!] Loi khi doc file Excel: {e}')
        input('Nhan phim ENTER de dong cua so...')
        sys.exit(1)

    override = {}
    if cols['mssv'] is None:
        print('\n[!] Khong tu tim thay tieu de cot MSSV.')
        override['mssv'] = ask_column_idx('   -> Nhap chu cai cot MSSV (vi du B hoac D)')
    if cols['q17'] is None:
        print('\n[!] Khong tu tim thay tieu de cot Diem Q17 (Part 1).')
        override['q17'] = ask_column_idx('   -> Nhap chu cai cot Diem Q17 / Part 1 (vi du I hoac K)')
    if cols['q18'] is None:
        print('\n[!] Khong tu tim thay tieu de cot Diem Q18 (Part 2).')
        override['q18'] = ask_column_idx('   -> Nhap chu cai cot Diem Q18 / Part 2 (vi du O hoac Q)')

    if override:
        cols, students = detect_columns_and_parse(path, override)

    while True:
        print('\n-----------------------------------------------------------------')
        print('-> TU DONG NHAN DIEN CAU TRUC FILE EXCEL:')
        print(
            f'   • MSSV: {col_label(cols["mssv"])} | '
            f'Ho ten: {col_label(cols["ten"])} | '
            f'Q17 (Part 1): {col_label(cols["q17"])} | '
            f'Q18 (Part 2): {col_label(cols["q18"])}'
        )
        if not students:
            print('   [!] Khong tim thay sinh vien nao co MSSV hop le trong cot da chon!')
            ans = input('   Gõ "DOI" de chon lai cot, hoac nhan ENTER de thoat: ').strip().upper()
            if ans != 'DOI':
                sys.exit(0)
        else:
            st0 = students[0]
            print(f'   • Tim thay: {len(students)} sinh vien (Bat dau tu dong {st0["row"]})')
            print(
                f'   • Sinh vien dau tien (Dong {st0["row"]}): '
                f'{st0["mssv"]} - {st0["ten"]} | Q17 = {st0["q17"]} | Q18 = {st0["q18"]}'
            )
            print('-----------------------------------------------------------------')
            print('2. Hay mo san trang luyenthinn tren trinh duyet (Edge, Chrome, Coc Coc, Brave)')
            print('   [Luu y: Giao dien trang luyenthinn de ngon ngu TIENG VIET]')
            print('   [Quy tac dung khan cap: Vay manh chuot kịch vao 1 trong 4 GOC MAN HINH]')
            ans = input(
                '\n3. Nhan ENTER de bat dau tu dau (hoac nhap So dong / MSSV de chay tiep, hoac go DOI de doi cot): '
            ).strip()
            if ans.upper() != 'DOI':
                students = filter_start_students(students, ans)
                break

        override = {
            'mssv': ask_column_idx('   -> Nhap chu cai cot MSSV', cols['mssv']),
            'ten': ask_column_idx('   -> Nhap chu cai cot Ho ten', cols['ten']),
            'q17': ask_column_idx('   -> Nhap chu cai cot Diem Q17 (Part 1)', cols['q17']),
            'q18': ask_column_idx('   -> Nhap chu cai cot Diem Q18 (Part 2)', cols['q18']),
            'ho_dem': None,
        }
        cols, students = detect_columns_and_parse(path, override)

    if not students:
        print('\n[!] Khong co sinh vien nao trong pham vi da chon!')
        input('Nhan phim ENTER de dong cua so...')
        sys.exit(0)

    print(f'\n-> Se tien hanh nhap diem cho {len(students)} sinh vien (bat dau tu {students[0]["mssv"]} - {students[0]["ten"]}).')
    print('-> Ban co 5 GIAY: HAY BAM CHUOT 1 CAI VAO KHOANG TRANG TREN TRANG LUYENTHINN CUA TRINH DUYET!')
    time.sleep(5)

    log_path = os.path.join(os.path.dirname(path), 'nhat_ky_cham_diem.txt')
    log = open(log_path, 'a', encoding='utf-8')
    current_sv_info = ''
    try:
        for idx, st in enumerate(students, 1):
            mssv, ten, row_num = st['mssv'], st['ten'], st['row']
            for phan, diem in [('Q17 (Part 1)', st['q17']), ('Q18 (Part 2)', st['q18'])]:
                current_sv_info = f'Dong {row_num} | {mssv} ({ten}) - {phan}'
                print(f'[{idx}/{len(students)}] (Dong {row_num}) Dang cham {mssv} ({ten}) - {phan}: diem = {diem}...')
                cham_mot_cau_web(mssv, diem, speed_mult=1.0)
                log.write(f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | Dong {row_num} | {mssv} | {ten} | {phan} = {diem}\n')
                log.flush()
        print('\n=================================================================')
        print('          === DA HOAN TAT TOAN BO DANH SACH! ===                 ')
        print(f'   Nhat ky da luu tai: {log_path}')
        print('=================================================================')
    except pyautogui.FailSafeException:
        print('\n=================================================================')
        print(' [!] DA DUNG KHAN CAP DO BAN VAY CHUOT VAO GOC MAN HINH!')
        print(f' [!] Vi tri dang dung: {current_sv_info}')
        print(f' [!] Nhat ky cac sinh vien da cham luu tai: {log_path}')
        print('=================================================================')
    except Exception as e:
        print(f'\n[!] Co loi xay ra khi dang cham ({current_sv_info}): {e}')
    finally:
        log.close()

    input('\nNhan phim ENTER de dong cua so...')
    sys.exit(0)


if __name__ == '__main__':
    main()
