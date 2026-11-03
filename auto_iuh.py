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


def detect_columns_and_parse(path: str, override_cols: dict | None = None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    col_mssv = None
    col_ten_gop = None
    col_ho_dem = None
    col_ten_rieng = None
    col_q17_primary = None
    col_q17_fallback = None
    col_q18_primary = None
    col_q18_fallback = None

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
                txt in ('mssv', 'mã sv', 'ma sv', 'mã số sv', 'ma so sv', 'mã số sinh viên', 'tên tài khoản', 'ten tai khoan', 'student id', 'mã số id')
                or 'mssv' in txt
                or 'mã sv' in txt
            ):
                col_mssv = j

            if col_ten_gop is None and txt in ('họ và tên', 'ho va ten', 'họ tên', 'ho ten', 'full name', 'fullname', 'student name'):
                col_ten_gop = j
            elif col_ho_dem is None and txt in ('họ và tên đệm', 'ho va ten dem', 'họ đệm', 'ho dem', 'họ', 'ho', 'first name', 'last name'):
                col_ho_dem = j
            elif col_ten_rieng is None and txt in ('tên', 'ten', 'name'):
                col_ten_rieng = j

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

    cols = {
        'mssv': col_mssv,
        'ten': col_ten,
        'ho_dem': col_ho_dem if (col_ten_gop is None and col_ten_rieng is not None) else None,
        'q17': col_q17,
        'q18': col_q18,
    }
    if override_cols:
        cols.update(override_cols)

    students = []
    c_mssv = cols['mssv']
    c_ten = cols['ten']
    c_ho_dem = cols['ho_dem']
    c_q17 = cols['q17']
    c_q18 = cols['q18']

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

            q17 = format_score(r[c_q17] if (c_q17 is not None and c_q17 < len(r)) else None)
            q18 = format_score(r[c_q18] if (c_q18 is not None and c_q18 < len(r)) else None)
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
                pyautogui.hotkey('ctrl', 'f'); time.sleep(0.8)
                pyperclip.copy(f'sv_{mssv}'); pyautogui.hotkey('ctrl', 'v'); time.sleep(1.0)
                pyautogui.press('esc'); time.sleep(0.8)
                pyautogui.hotkey('ctrl', 'f'); time.sleep(0.8)
                pyperclip.copy('Yêu cầu chấm điểm'); pyautogui.hotkey('ctrl', 'v'); time.sleep(1.0)
                pyautogui.press('esc'); time.sleep(0.8)
                pyautogui.press('enter'); time.sleep(4.5)
                pyautogui.hotkey('ctrl', 'f'); time.sleep(0.8)
                pyperclip.copy('Tạo bình luận'); pyautogui.hotkey('ctrl', 'v'); time.sleep(1.0)
                pyautogui.press('esc'); time.sleep(0.8)
                pyautogui.press('enter'); time.sleep(5.0)
                pyautogui.hotkey('ctrl', 'f'); time.sleep(0.8)
                pyperclip.copy('trên 5,00'); pyautogui.hotkey('ctrl', 'v'); time.sleep(1.0)
                pyautogui.press('esc'); time.sleep(0.8)
                pyautogui.hotkey('shift', 'tab'); time.sleep(0.8)
                pyautogui.press('backspace', presses=5, interval=0.1)
                pyautogui.press('delete', presses=5, interval=0.1)
                pyautogui.write(str(diem), interval=0.15); time.sleep(2.0)
                pyautogui.press('enter'); time.sleep(4.0)
                pyautogui.hotkey('ctrl', 'w'); time.sleep(1.5)
                pyautogui.press('f5'); time.sleep(4.0)
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
