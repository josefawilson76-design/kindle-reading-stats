@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_kindle_reader.ps1" %*
set CODE=%ERRORLEVEL%
if not "%CODE%"=="0" (
  echo.
  echo Run failed. Error code: %CODE%
  pause
  exit /b %CODE%
)

echo.
echo Run completed.
echo Output paths follow settings.json.
echo.

:menu
choice /c RO /n /m "Press R to reopen the default report, or O to exit: "
if errorlevel 2 goto end
if exist "%~dp0dist\kindle_reading_dashboard.html" (
  start "" "%~dp0dist\kindle_reading_dashboard.html"
) else (
  start "" "%~dp0kindle_reading_dashboard.html"
)
goto menu

:end
endlocal
exit /b 0
