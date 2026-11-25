@echo off
rem ============================================================================
rem  Remove-MassLynx - double-click this file
rem  Copies the tool to a folder every account can read
rem  (C:\ProgramData\MassLynxOnderhoud\app) and then starts it with
rem  administrator rights. Windows may ask for an admin account.
rem ============================================================================
title MassLynx verwijderen
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Start-MassLynxRemoval.ps1"

if errorlevel 2 (
  echo.
  echo Het programma is niet gestart. Zie de melding in het venster hierboven.
  pause
)
exit /b
