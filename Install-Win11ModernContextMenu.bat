@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "LOG_DIR=%SCRIPT_DIR%artifacts"
set "LOG_FILE=%LOG_DIR%\install.log"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%" >nul 2>&1

net session >nul 2>&1
if not "%ERRORLEVEL%"=="0" (
  echo Requesting administrator rights...
  if "%~1"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath 'cmd.exe' -ArgumentList '/k ""%~f0""' -Verb RunAs"
  ) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath 'cmd.exe' -ArgumentList '/k ""%~f0"" %*' -Verb RunAs"
  )
  if not "%ERRORLEVEL%"=="0" (
    echo.
    echo Failed to start elevated installer.
    pause
  )
  exit /b 0
)

echo Installing Windows 11 modern Codex context menu...
echo Log: %LOG_FILE%
echo.

if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { param([string]$ScriptPath, [string]$LogPath) try { & $ScriptPath 2>&1 | Tee-Object -FilePath $LogPath; if ($LASTEXITCODE -is [int]) { exit $LASTEXITCODE }; exit 0 } catch { $_ | Out-String | Tee-Object -FilePath $LogPath -Append; exit 1 } }" "%SCRIPT_DIR%scripts\Install-Win11ModernContextMenu.ps1" "%LOG_FILE%"
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { param([string]$ScriptPath, [string]$LogPath, [Parameter(ValueFromRemainingArguments=$true)][string[]]$Rest) try { & $ScriptPath @Rest 2>&1 | Tee-Object -FilePath $LogPath; if ($LASTEXITCODE -is [int]) { exit $LASTEXITCODE }; exit 0 } catch { $_ | Out-String | Tee-Object -FilePath $LogPath -Append; exit 1 } }" "%SCRIPT_DIR%scripts\Install-Win11ModernContextMenu.ps1" "%LOG_FILE%" %*
)
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo Install completed successfully.
) else (
  echo Install failed with exit code %EXITCODE%.
)
echo Log: %LOG_FILE%
echo.
pause

exit /b %EXITCODE%
