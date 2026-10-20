@echo off
rem Double-click launcher. With no arguments it runs a full sync (import, then export).
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ClaudeSessions.ps1" -Mode Sync
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-ClaudeSessions.ps1" %*
)
pause
