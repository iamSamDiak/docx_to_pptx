@echo off
echo ========================================
echo  DOCX to PPTX - Build Script
echo ========================================
echo.

echo [1/4] Installation des dependances...
pip install -r requirements.txt
pip install pyinstaller
echo.

echo [2/4] Creation de l'executable...
pyinstaller --onefile --windowed --name="DOCX to PPTX" --icon=icons/app.ico --collect-all PyQt5 --collect-all docx --collect-all pptx --collect-all lxml main.py
echo.

echo [3/4] Verification d'Inno Setup...
where iscc >nul 2>nul
if %errorlevel% neq 0 (
    echo Inno Setup non trouve. Installation...
    choco install innosetup -y
)
echo.

echo [4/4] Creation de l'installeur...
iscc setup.iss
echo.

echo ========================================
echo  Build termine !
echo  Installeur : installer\DOCX_to_PPTX_Setup.exe
echo ========================================
pause
