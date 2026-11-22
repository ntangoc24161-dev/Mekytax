@echo off
chcp 65001 > nul
echo ======================================================
echo    TIEN TRINH DONG BO SO BAO CAO CONG VAN (EXCEL COM)
echo ======================================================
echo.
python "%~dp0sync_dashboard_windows.py"
echo.
pause
