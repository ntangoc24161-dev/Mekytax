@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "EXE=%~dp0PhanTichHoSoThue\PhanTichHoSoThue.exe"
if not exist "%EXE%" set "EXE=%~dp0PhanTichHoSoThue.exe"
if not exist "%EXE%" (
  echo Khong thay PhanTichHoSoThue.exe
  pause
  exit /b 1
)
if "%~1"=="" (
  "%EXE%"
) else (
  "%EXE%" --input-dir "%~1" --out "%~1\ket-qua"
)
echo.
pause
