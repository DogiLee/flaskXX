@echo off
setlocal
rem ==================================================================
rem  PDGM test ortami: uretim data klasorune DOKUNMAZ.
rem  Uygulama dosyalari ..\gokberk_flask_TEST klasorune kopyalanir,
rem  bos bir data klasoruyle 127.0.0.1:5002 portunda baslatilir.
rem  Kullanici hesaplari (data\kullanicilar.json) uretimle aynidir.
rem  Bastan baslamak icin: sunucuyu kapatin, gokberk_flask_TEST\data
rem  klasorunu silin, bu dosyayi tekrar calistirin.
rem ==================================================================
set "PAKET=%~dp0"
for %%I in ("%PAKET%..\..") do set "PROJE=%%~fI"
for %%I in ("%PROJE%\..") do set "UST=%%~fI"
set "TEST=%UST%\gokberk_flask_TEST"

if not exist "%PROJE%\.venv\Scripts\python.exe" (
    echo HATA: %PROJE%\.venv bulunamadi.
    pause
    exit /b 1
)

robocopy "%PROJE%" "%TEST%" *.py .env /R:1 /W:1 /NFL /NDL /NJH /NJS >nul
robocopy "%PROJE%\templates" "%TEST%\templates" /E /R:1 /W:1 /NFL /NDL /NJH /NJS >nul
robocopy "%PROJE%\static" "%TEST%\static" /E /R:1 /W:1 /NFL /NDL /NJH /NJS >nul
if not exist "%TEST%\data" mkdir "%TEST%\data"
if exist "%PROJE%\data\kullanicilar.json" if not exist "%TEST%\data\kullanicilar.json" (
    copy /Y "%PROJE%\data\kullanicilar.json" "%TEST%\data\" >nul
)

echo.
echo   Test sunucusu : http://127.0.0.1:5002   (uretim sunucusu 5001, etkilenmez)
echo   Test verisi   : %TEST%\data
echo.
set PDGM_PORT=5002
set PDGM_BIND=127.0.0.1
cd /d "%TEST%"
"%PROJE%\.venv\Scripts\python.exe" app.py
pause
