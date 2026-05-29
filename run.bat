@echo off
title Last Pages - Cube Block Creator
cd /d "%~dp0"

REM --- Python kurulu mu ve gercek mi (MS Store stub mu) kontrol ---
where python > "%TEMP%\pypath.txt" 2>nul
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  HATA: Python kurulu DEGIL.
    echo ============================================================
    echo.
    echo Cozum:
    echo   1. https://www.python.org/downloads/ git
    echo   2. Python 3.10 veya yenisini indir
    echo   3. Kurulumda "Add python.exe to PATH" tikini MUTLAKA isaretle
    echo   4. Microsoft Store'dan KURMA - calismaz.
    echo.
    pause
    exit /b
)

findstr /i "WindowsApps" "%TEMP%\pypath.txt" >nul
if not errorlevel 1 (
    echo.
    echo ============================================================
    echo  HATA: Python = Microsoft Store stub (gercek degil)
    echo ============================================================
    echo.
    echo Su an "python" yazinca Microsoft Store aciliyor cunku
    echo gercek Python yerine bos bir yer tutucu kurulu.
    echo.
    echo Cozum:
    echo   1. Ayarlar acin: "Uygulama yurutucu takma adlari" ara
    echo      ("App execution aliases")
    echo   2. python.exe ve python3.exe icin KAPALI yap
    echo   3. https://www.python.org/downloads/ adresinden Python 3.10+ indir
    echo   4. Kurulumda "Add python.exe to PATH" tikini ISARETLE
    echo   5. Bu run.bat'i tekrar calistir
    echo.
    pause
    exit /b
)

echo Python bulundu:
type "%TEMP%\pypath.txt"
echo.

echo Launcher indiriliyor...
curl -sSLo launcher.py "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/launcher.py" 2>nul

if not exist launcher.py (
    echo.
    echo HATA: launcher.py indirilemedi. Internet/firewall sorunu.
    pause
    exit /b
)

echo Launcher calisiyor...
python launcher.py
set "PY_EXIT=%ERRORLEVEL%"

if not "%PY_EXIT%"=="0" (
    echo.
    echo HATA: Python exit code = %PY_EXIT%
    echo Olasi sebep: tkinter eksik veya launcher.py'de hata.
    pause
)