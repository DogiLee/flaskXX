@echo off
setlocal

rem Betigin bulundugu klasore gec (nereden calistirilirsa calistirilsin dogru calissin)
cd /d "%~dp0"

rem Sanal ortam kontrolu
if not exist ".venv\Scripts\python.exe" (
    echo HATA: .venv sanal ortami bulunamadi.
    echo Once su komutlari calistirin:
    echo   uv venv
    echo   uv pip install -r requirements.txt
    pause
    exit /b 1
)

rem .env kontrolu
if not exist ".env" (
    echo HATA: .env dosyasi bulunamadi.
    echo .env.example dosyasini kopyalayip .env olarak duzenleyin.
    pause
    exit /b 1
)

echo ============================================
echo   PDGM Is Takip Sistemi baslatiliyor...
echo ============================================
echo.

".venv\Scripts\python.exe" app.py

set HATA_KODU=%ERRORLEVEL%

echo.
if %HATA_KODU% NEQ 0 (
    echo Sunucu HATA ile kapandi ^(kod: %HATA_KODU%^). Yukaridaki mesaji kontrol edin.
    pause
) else (
    echo Sunucu normal sekilde kapandi.
)

endlocal
