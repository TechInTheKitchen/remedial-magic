@echo off
setlocal
title Obsidian GitHub Web Reader
cd /d "%~dp0.."

where node.exe >nul 2>nul
if not errorlevel 1 (
  node.exe "%~dp0local-server.cjs"
  goto :done
)

set "SYSTEM_NODE=%ProgramFiles%\nodejs\node.exe"
if exist "%SYSTEM_NODE%" (
  "%SYSTEM_NODE%" "%~dp0local-server.cjs"
  goto :done
)

set "SYSTEM_NODE_X86=%ProgramFiles(x86)%\nodejs\node.exe"
if exist "%SYSTEM_NODE_X86%" (
  "%SYSTEM_NODE_X86%" "%~dp0local-server.cjs"
  goto :done
)

echo.
echo Node.js could not be found. Install Node.js or open this folder through another local web server.
echo Opening index.html directly will not work because browsers block Markdown fetches from file URLs.
echo.
pause
exit /b 1

:done
if errorlevel 1 (
  echo.
  echo The local reader stopped because of the error shown above.
  echo.
  pause
)
