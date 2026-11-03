@echo off
chcp 65001 >nul
echo Dang cai dat thu vien Python can thiet (pyautogui, openpyxl, pyperclip)...
python -m pip install --upgrade pip
python -m pip install -r "%~dp0requirements.txt"
echo.
echo === CAI DAT HOAN TAT! BAY GIO BAN CO THE MO FILE SO 2 DE CHAY ===
pause
