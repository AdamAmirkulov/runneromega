@echo off
rem Plain Chrome for portal-sot.kz with a debug port (no chromedriver launch,
rem no --enable-automation -> navigator.webdriver = false). Scripts attach to it.
rem Usage: chrome_portal_sot.bat "<profile dir>" <port>
rem Can be run by hand too, e.g.:
rem   chrome_portal_sot.bat "C:\Users\User\Documents\ChromePortalSot" 9301
if "%~1"=="" (echo Usage: %~nx0 "profile dir" port & exit /b 1)
if "%~2"=="" (echo Usage: %~nx0 "profile dir" port & exit /b 1)

set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" (echo chrome.exe not found & exit /b 2)

rem Without the backgrounding flags a covered/minimized window becomes
rem visibilityState=hidden and Selenium's clicks stop reaching the page.
start "" "%CHROME%" --remote-debugging-port=%~2 --user-data-dir="%~1" --no-first-run --no-default-browser-check --start-maximized --disable-notifications --disable-backgrounding-occluded-windows --disable-renderer-backgrounding --disable-background-timer-throttling https://portal-sot.kz/
