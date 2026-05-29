@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ============================================================
echo  Last Pages Program - Cube Block Creator
echo ============================================================
echo.

REM --- Download latest version from GitHub (silent fail to local files) ---
echo [1/3] Guncel surum indiriliyor...
curl -sSLf -o create_blocks.py.new "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/create_blocks.py" && move /Y create_blocks.py.new create_blocks.py >nul
curl -sSLf -o blocks_data.json.new "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/blocks_data.json" && move /Y blocks_data.json.new blocks_data.json >nul
if exist create_blocks.py.new del create_blocks.py.new >nul 2>&1
if exist blocks_data.json.new del blocks_data.json.new >nul 2>&1
echo     Tamam.
echo.

REM --- Ensure openpyxl is installed ---
echo [2/3] openpyxl kontrol ediliyor...
python -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo     openpyxl yok, kuruluyor...
    python -m pip install --quiet openpyxl
)
echo     Tamam.
echo.

REM --- Get Excel path: from arg (drag-and-drop) or from config or prompt ---
set "EXCEL=%~1"
if "%EXCEL%"=="" (
    if exist config.txt (
        set /p EXCEL=<config.txt
    )
)
if "%EXCEL%"=="" (
    echo Excel dosyasinin tam yolunu girin (veya bu .bat uzerine surukleyin^):
    set /p EXCEL=^>
    if not "!EXCEL!"=="" (
        echo !EXCEL!>config.txt
    )
)

if "%EXCEL%"=="" (
    echo HATA: Excel yolu verilmedi.
    pause
    exit /b 1
)

echo [3/3] Script calistiriliyor...
echo     Excel: %EXCEL%
echo.

python create_blocks.py "%EXCEL%"
set RC=%ERRORLEVEL%

echo.
echo ============================================================
if %RC%==0 (
    echo  BASARILI. Cikti dosyasi ayni klasorde.
) else (
    echo  HATA olustu. Yukaridaki mesaja bak.
)
echo ============================================================
echo.
pause
