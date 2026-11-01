@echo off
rem LastPagesApp.exe'yi derler: dist\LastPagesApp.exe
rem Tedarik zinciri: build\venv icine yalniz kilitteki surum + parmak iziyle kur
rem (alt paketler dahil), sonra bilinen acik taramasi; acik varsa derleme durur.
cd /d "%~dp0"
set VPY=build\venv\Scripts\python.exe
set TARAYICI=%USERPROFILE%\tools\tedarik\Scripts\python.exe
if not exist "%VPY%" python -m venv build\venv || goto hata
"%VPY%" -m pip install -q --require-hashes --no-deps -r requirements.txt || goto hata
"%VPY%" -m pip check || goto hata
if not exist "%TARAYICI%" (
    echo HATA: acik tarayici yok: %TARAYICI%
    goto hata
)
"%TARAYICI%" -X utf8 "%USERPROFILE%\tools\tedarik\tara.py" --disable-pip -r requirements.txt || goto hata
"%VPY%" -m PyInstaller --noconfirm --clean LastPagesApp.spec || goto hata
echo.
echo Bitti: dist\LastPagesApp.exe
pause
exit /b 0

:hata
echo.
echo HATA: derleme basarisiz.
pause
exit /b 1
