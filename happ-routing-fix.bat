@echo off
setlocal
chcp 65001 >nul
title Happ Routing Fix by gulenok91
set "HAPP_ROUTING_SIGNATURE=HRF_GULENOK91"
set "HAPP_ROUTING_LAUNCHER=%~f0"
set "ROOT=%~dp0"
set "SCRIPT=%ROOT%scripts\happ-routing.ps1"

if /i "%~1"=="/background" (
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Action ensure-defaults >nul 2>nul
    powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Action start-agent-if-enabled >nul 2>nul
    endlocal & exit /b %errorlevel%
)

powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Action ui
set "ERR=%errorlevel%"

if not "%ERR%"=="0" (
    echo.
    echo Happ Routing Fix closed with an error.
    echo Check happ-routing.log for details.
    pause
)

endlocal & exit /b %ERR%
