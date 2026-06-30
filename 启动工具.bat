@echo off
setlocal

cd /d "%~dp0"

set "LOG_DIR=%~dp0logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set "LOG_FILE=%LOG_DIR%\start.log"

echo Starting Office document generator...
echo Working directory: %CD%
echo Log file: %LOG_FILE%
echo.

set "NODE_EXE=%~dp0runtime\node\node.exe"
if exist "%NODE_EXE%" goto run

set "NODE_EXE=node"

:run
echo Node command: %NODE_EXE% > "%LOG_FILE%"
echo Starting... >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

"%NODE_EXE%" "%~dp0start-local.js" >> "%LOG_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo The tool has stopped. Exit code: %EXIT_CODE%
echo.
echo If it did not open, please send this log file to the maintainer:
echo %LOG_FILE%
echo.
type "%LOG_FILE%"
echo.
pause
