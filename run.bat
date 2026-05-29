@echo off
title Last Pages - Cube Block Creator
cd /d "%~dp0"

echo Launcher indiriliyor...
curl -sSLo launcher.py "https://raw.githubusercontent.com/yavuzzeynulat-cell/LastPagesProgram/main/launcher.py" 2>nul

if not exist launcher.py (
    echo.
    echo HATA: launcher.py indirilemedi.
    echo Sebep: internet baglantisi yok, ya da curl/firewall sorunu.
    pause
    exit /b
)

python launcher.py
if errorlevel 1 (
    echo.
    echo HATA: Pencere acilamadi.
    echo Olasi sebep:
    echo   1. Python kurulu degil   ^(python.org'dan 3.10+ kur, "Add to PATH" tikle^)
    echo   2. tkinter eksik         ^(genelde Python ile gelir, eksikse "python -m pip install tk"^)
    echo.
    pause
)