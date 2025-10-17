# DOCX to PPTX - Build Script
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " DOCX to PPTX - Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/4] Installation des dependances..." -ForegroundColor Yellow
pip install -r requirements.txt
pip install pyinstaller
Write-Host ""

Write-Host "[2/4] Creation de l'executable..." -ForegroundColor Yellow
pyinstaller --onefile --windowed --name="DOCX to PPTX" --icon=icons/app.ico --collect-all PyQt5 --collect-all docx --collect-all pptx --collect-all lxml main.py
Write-Host ""

Write-Host "[3/4] Verification d'Inno Setup..." -ForegroundColor Yellow
$isccPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $isccPath)) {
    Write-Host "Inno Setup non trouve. Installation..." -ForegroundColor Red
    choco install innosetup -y
    $isccPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
}
Write-Host ""

Write-Host "[4/4] Creation de l'installeur..." -ForegroundColor Yellow
& $isccPath setup.iss
Write-Host ""

Write-Host "========================================" -ForegroundColor Green
Write-Host " Build termine !" -ForegroundColor Green
Write-Host " Installeur : installer\DOCX_to_PPTX_Setup.exe" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Read-Host "Appuyez sur Entree pour continuer"
