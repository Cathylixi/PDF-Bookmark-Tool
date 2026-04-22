@echo off
echo Cleaning up old builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "PDF_Bookmark_Tool.spec" del /f /q "PDF_Bookmark_Tool.spec"

echo Installing requirements...
pip install -r requirements.txt

echo Building standalone executable...
python -m PyInstaller --noconfirm --onedir --windowed --name "PDF_Bookmark_Tool" --clean main.py

echo.
echo ====================================================
echo Build finished! 
echo Your executable is in the 'dist\PDF_Bookmark_Tool' folder.
echo You can zip this folder and send it to others.
echo ====================================================
pause
