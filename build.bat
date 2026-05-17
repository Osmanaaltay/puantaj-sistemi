@echo off
title Puantaj EXE Builder

echo Baslaniyor...
echo.

:: py launcher kontrolu
py --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python bulunamadi! python.org dan Python 3.11 indirin.
    pause
    exit /b 1
)

:: 3.11 yuklu mu? Yoksa winget ile indir
py -3.11 --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python 3.11 bulunamadi, indiriliyor...
    winget install Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo [HATA] Otomatik kurulum basarisiz!
        echo Manuel olarak su komutu calistirin: winget install Python.Python.3.11
        pause
        exit /b 1
    )
    echo [OK] Python 3.11 kuruldu. Devam ediliyor...
    :: PATH'i yenile
    call refreshenv >nul 2>&1
)

for /f "tokens=*" %%v in ('py -3.11 --version') do echo [OK] %%v bulundu.

echo.
echo [1/3] Paketler yukleniyor...
py -3.11 -m pip install flask flask-cors openpyxl pyinstaller pywebview --quiet
if errorlevel 1 (
    echo [HATA] Paket yuklenemedi!
    pause
    exit /b 1
)
echo [OK] Paketler hazir.

echo.
echo [2/3] EXE olusturuluyor (1-3 dakika surebilir)...
py -3.11 -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "PuantajSistemi" ^
  --add-data "index.html;." ^
  --hidden-import flask ^
  --hidden-import flask_cors ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  --hidden-import sqlite3 ^
  --hidden-import calendar ^
  --hidden-import threading ^
  --collect-all webview ^
  app.py

if errorlevel 1 (
    echo [HATA] EXE olusturulamadi!
    pause
    exit /b 1
)

echo.
echo [3/3] Temizleniyor...
if exist "build" rmdir /s /q build
if exist "PuantajSistemi.spec" del /q PuantajSistemi.spec

echo.
echo TAMAMLANDI! dist\PuantajSistemi.exe hazir.
echo.
start "" dist
pause