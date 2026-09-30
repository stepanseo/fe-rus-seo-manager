@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
pyinstaller --noconfirm --clean --onefile --windowed --name FE-RUS_SEO_Manager fe_rus_seo_manager.py
echo.
echo EXE: dist\FE-RUS_SEO_Manager.exe
pause
