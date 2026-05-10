@echo off
title poinmeet - Clean for Share
cd /d "%~dp0"

echo.
echo  ============================================
echo   poinmeet - Bersihkan untuk Share
echo  ============================================
echo.
echo  Yang akan dihapus:
echo  - backend\venv\        (virtual environment)
echo  - backend\__pycache__\ (cache Python)
echo  - backend\uploads\     (gambar yang diupload)
echo  - backend.log / backend_err.log
echo  - *.docx, *.pdf        (file output/test)
echo  - ~$*.docx             (temp Word)
echo.
echo  Yang TIDAK dihapus:
echo  - backend\.env         (API key tetap disertakan)
echo.
set /p CONFIRM=Lanjutkan? (y/n): 
if /i not "%CONFIRM%"=="y" (
    echo  Dibatalkan.
    pause & exit /b
)

echo.
echo  [*] Membersihkan...

if exist "backend\venv"        rmdir /s /q "backend\venv"
if exist "backend\__pycache__" rmdir /s /q "backend\__pycache__"
if exist "backend\uploads" (
    rmdir /s /q "backend\uploads"
    mkdir "backend\uploads"
)
if exist "backend.log"     del /f /q "backend.log"
if exist "backend_err.log" del /f /q "backend_err.log"

for %%f in (*.docx *.pdf) do del /f /q "%%f"
for %%f in (~$*.docx) do del /f /q "%%f"

echo  [OK] Selesai! Siap untuk dishare.
echo.
echo  Penerima cukup jalankan install.bat untuk memulai.
echo.
pause
