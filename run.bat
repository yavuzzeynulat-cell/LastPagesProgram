@echo off
title Cube Block Creator
cd /d "%~dp0"

echo.
echo ============================================================
echo  Last Pages Program - Cube Block Creator
echo ============================================================
echo  Klasor: %CD%
echo  Argument: %~1
echo ============================================================
echo.
echo Devam etmek icin bir tusa basin (bu pencere acik kalacak)...
pause

echo.
echo [1] Python kontrol ediliyor...
where python >nul 2>&1
if errorlevel 1 goto :no_python
python --version
echo     Tamam.

echo.
echo [2] curl kontrol ediliyor...
where curl >nul 2>&1
if errorlevel 1 goto :no_curl
echo     Tamam.

echo.
echo [3] Guncel script indiriliyor...
curl -sSL -o create_blocks.py "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/create_blocks.py"
if not exist create_blocks.py goto :no_download
curl -sSL -o blocks_data.json "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/blocks_data.json"
if not exist blocks_data.json goto :no_download
echo     Indirildi.

echo.
echo [4] openpyxl kontrol ediliyor...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo     Kurulu degil, kuruluyor...
    python -m pip install openpyxl
)
echo     Tamam.

echo.
set "EXCEL=%~1"
if "%EXCEL%"=="" if exist config.txt set /p EXCEL=<config.txt
if "%EXCEL%"=="" goto :ask_excel
goto :run

:ask_excel
echo [5] Excel dosyasinin tam yolunu yapistir (cift tirnaksiz):
set /p EXCEL=^> 
if "%EXCEL%"=="" goto :no_excel

:run
echo %EXCEL%> config.txt
echo.
echo [6] Script calistiriliyor...
echo     Excel: %EXCEL%
echo.
python create_blocks.py "%EXCEL%"
goto :done

:no_python
echo.
echo HATA: Python kurulu degil veya PATH'te degil.
echo Cozum: python.org'dan Python 3.10+ indir, kurulumda
echo        "Add Python to PATH" tikini isaretle.
goto :done

:no_curl
echo.
echo HATA: curl bulunamadi (Windows 10 1803+ gerekli).
goto :done

:no_download
echo.
echo HATA: Script indirilemedi. Internet baglantisi ya da firewall sorunu.
goto :done

:no_excel
echo.
echo HATA: Excel yolu girilmedi.
goto :done

:done
echo.
echo ============================================================
echo  BITTI. Cikmak icin bir tusa basin.
echo ============================================================
pause