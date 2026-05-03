@echo off
setlocal
chcp 65001 >nul
title Happ Routing Fix - process check
set "ROOT=%~dp0"
set "SCRIPT=%ROOT%scripts\happ-routing.ps1"

powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Action process-monitor
set "ERR=%errorlevel%"

if not "%ERR%"=="0" (
    echo.
    echo Happ Routing Fix process check closed with an error.
    echo Check happ-routing.log for details.
    pause
)

endlocal & exit /b %ERR%
