@echo off
title Setup Login GitHub - bayi-app
cls
echo ======================================================
echo       SETUP LOGIN GITHUB (bimcuk/bayi-app)
echo ======================================================
echo.
echo Langkah mudah membuat Token GitHub:
echo 1. Buka link: https://github.com/settings/tokens
echo 2. Klik "Generate new token (classic)"
echo 3. Beri Note bebas (contoh: laptop)
echo 4. Centang checklist [repo] (Full control of private repositories)
echo 5. Klik "Generate token" di bawah, lalu copy tokennya.
echo.
echo ------------------------------------------------------
echo Format token biasanya diawali: ghp_...
echo ------------------------------------------------------
echo.

set /p GITHUB_TOKEN="Paste GitHub Token di sini, lalu tekan Enter: "

if "%GITHUB_TOKEN%"=="" (
    echo.
    echo [ERROR] Token tidak boleh kosong! Silakan jalankan ulang script ini.
    echo.
    pause
    exit /b 1
)

echo.
echo ⏳ Mengonfigurasi autentikasi GitHub...
"C:\Users\user\AppData\Local\Programs\MinGit\cmd\git.exe" remote set-url origin https://%GITHUB_TOKEN%@github.com/bimcuk/bayi-app.git

echo 🚀 Mengirim seluruh commit terbaru ke branch main GitHub...
"C:\Users\user\AppData\Local\Programs\MinGit\cmd\git.exe" push origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ======================================================
    echo 🎉 [SUKSES] Berhasil login dan push ke GitHub!
    echo Repositori kamu: https://github.com/bimcuk/bayi-app
    echo.
    echo Untuk update selanjutnya, kamu tinggal double-click:
    echo autopush.bat (akan otomatis push tanpa minta login lagi).
    echo ======================================================
) else (
    echo.
    echo ======================================================
    echo ❌ [GAGAL] Gagal push ke GitHub.
    echo Pastikan token kamu benar dan centang [repo] aktif.
    echo ======================================================
)

echo.
pause
