@echo off
cd /d "%~dp0"

echo ============================================================
echo  Last Pages Program - Cube Block Creator
echo ============================================================
echo.

echo [1/3] Guncel surum indiriliyor...
curl -sSL -o create_blocks.py.tmp "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/create_blocks.py"
if exist create_blocks.py.tmp move /Y create_blocks.py.tmp create_blocks.py >nul

curl -sSL -o blocks_data.json.tmp "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/blocks_data.json"
if exist blocks_data.json.tmp move /Y blocks_data.json.tmp blocks_data.json >nul
echo     Tamam.
echo.

echo [2/3] openpyxl kontrol ediliyor...
python -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo     openpyxl yok, kuruluyor...
    python -m pip install --quiet openpyxl
)
echo     Tamam.
echo.

set "EXCEL=%~1"
if "%EXCEL%"=="" if exist config.txt set /p EXCEL=<config.txt
if "%EXCEL%"=="" (
    set /p EXCEL=Excel dosyasinin tam yolunu girin: 
)

if "%EXCEL%"=="" (
    echo HATA: Excel yolu verilmedi.
    pause
    exit /b 1
)

echo %EXCEL%>config.txt

echo [3/3] Script calistiriliyor...
echo     Excel: %EXCEL%
echo.

python create_blocks.py "%EXCEL%"

echo.
echo ============================================================
echo  BITTI. Cikmak icin bir tusa basin.
echo ============================================================
pause >nul