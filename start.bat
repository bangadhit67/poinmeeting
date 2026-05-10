@echo off
title poinmeet
color 0A
cd /d "%~dp0"

echo.
echo  ============================================
echo   poinmeet - MOM Generator
echo  ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Python tidak ditemukan. Jalankan install.bat terlebih dahulu.
    pause & exit /b 1
)

:: Check venv
if not exist "%~dp0backend\venv\Scripts\activate.bat" (
    echo  [!] Virtual environment belum ada. Jalankan install.bat terlebih dahulu.
    pause & exit /b 1
)

:: Kill existing backend on port 5000
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5000 "') do (
    taskkill /PID %%a /F >nul 2>&1
)

echo  [*] Menjalankan backend...
start /B "" cmd /c "cd /d "%~dp0backend" && venv\Scripts\activate && python app.py > ..\backend.log 2>&1"

:: Wait for backend
echo  [*] Menunggu backend siap
set /a c=0
:wait
timeout /t 1 /nobreak >nul
curl -s http://localhost:5000/api/health >nul 2>&1
if %errorlevel% equ 0 goto ready
set /a c+=1
if %c% geq 20 (
    echo  [!] Backend gagal start. Cek backend.log untuk detail.
    pause & exit /b 1
)
<nul set /p "=."
goto wait

:ready
echo.
echo  [OK] Backend siap!
echo  [*] Membuka browser...
start http://localhost:5000/app
echo.
echo  Aplikasi berjalan di http://localhost:5000/app
echo  Ketik  restart  lalu Enter untuk restart backend.
echo  Tutup window ini (atau tekan Enter tanpa input) untuk mematikan backend.
echo.

:input_loop
set "cmd_input="
set /p "cmd_input=> "
if /i "%cmd_input%"=="restart" goto do_restart
goto shutdown

:do_restart
echo  [*] Merestart backend...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5000 "') do (
    taskkill /PID %%a /F >nul 2>&1
)
start /B "" cmd /c "cd /d "%~dp0backend" && venv\Scripts\activate && python app.py > ..\backend.log 2>&1"
set /a c=0
:wait_restart
timeout /t 1 /nobreak >nul
curl -s http://localhost:5000/api/health >nul 2>&1
if %errorlevel% equ 0 goto restart_ok
set /a c+=1
if %c% geq 20 (echo  [!] Gagal restart. Cek backend.log. & goto input_loop)
goto wait_restart
:restart_ok
echo  [OK] Backend berhasil direstart!
start http://localhost:5000/app
goto input_loop

:shutdown

:: Cleanup: kill backend
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5000 "') do (
    taskkill /PID %%a /F >nul 2>&1
)
echo  [OK] Backend dihentikan.
