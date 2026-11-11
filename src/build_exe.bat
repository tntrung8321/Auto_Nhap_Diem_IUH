@echo off
chcp 65001 >nul
echo Dang dong goi IUH_Writing_AI_Suite.exe bang PyInstaller...
python -m PyInstaller --noconfirm --noconsole --onefile --name "IUH_Writing_AI_Suite" --paths "%~dp0" --add-data "%~dp0..\MARKING_RULES_CAMBRIDGE_CUOI_KI.md;." --distpath "%~dp0.." --workpath "%~dp0build" --specpath "%~dp0" "%~dp0app_gui.py"
rmdir /s /q "%~dp0build" 2>nul
del /q "%~dp0IUH_Writing_AI_Suite.spec" 2>nul
echo === DONG GOI THANH CONG: IUH_Writing_AI_Suite.exe ===
pause
