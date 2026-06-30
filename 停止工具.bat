@echo off
setlocal

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5174" ^| findstr "LISTENING"') do (
  taskkill /PID %%a /F >nul 2>nul
)

echo Local Office document generator has been stopped if it was running.
pause
