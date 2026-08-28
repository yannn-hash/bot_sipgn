@echo off
title Launcher Aplikasi SIPGN
echo ========================================================
echo MEMULAI APLIKASI SIPGN BOT
echo ========================================================
echo.

:: Simpan path utama folder
set "ROOT_DIR=%~dp0"

echo [1/2] Menghidupkan Server API (Backend)...
:: Membuka CMD baru khusus untuk Backend
start "Server API (Backend)" cmd /k "cd /d "%ROOT_DIR%web_app\backend" && call "%ROOT_DIR%.venv\Scripts\activate.bat" && uvicorn main:app --reload --port 8000"

:: Tunggu 2 detik agar backend siap
timeout /t 2 /nobreak >nul

echo [2/2] Menghidupkan Tampilan Web (Frontend)...
:: Membuka CMD baru khusus untuk Frontend dan otomatis buka Chrome
start "UI Web (Frontend)" cmd /k "cd /d "%ROOT_DIR%web_app\frontend" && npm run dev -- --open"

echo.
echo ========================================================
echo SEMUA SERVER SEDANG BERJALAN.
echo.
echo Anda akan melihat dua jendela hitam (Backend dan Frontend).
echo JANGAN MENUTUP KEDUA JENDELA TERSEBUT selama Anda menggunakan web.
echo.
echo Jika browser tidak terbuka otomatis, buka: http://localhost:5170
echo ========================================================
pause
