@echo off
chcp 65001 > nul
title Depo Paneli - Sipariş Toplama Programı
echo ===================================================================
echo     DEPO PANELI | E-TICARET SIPARIS VE DEPO TOPLAMA SISTEMI
echo ===================================================================
echo.
echo Program baslatiliyor, lutfen bekleyin...
echo Tarayiciniz otomatik olarak acilacaktir: http://127.0.0.1:5000
echo.

echo Port 5000 kontrol ediliyor...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5000 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
taskkill /F /IM SiparisToplama.exe >nul 2>&1

cd /d "c:\Users\ugurk\OneDrive\Masaüstü\projem"
"C:\Users\ugurk\AppData\Local\Programs\Python\Python312\python.exe" app.py

pause
