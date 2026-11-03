@echo off
rem Excel 100x Pro - one-click local install for ONE PC (small business / trial).
rem 1) trusts a localhost HTTPS certificate (Windows will ask you to confirm)
rem 2) registers the add-in with Excel (HKCU only, no admin rights)
rem 3) starts the local server - keep this window open while using Excel.
setlocal
cd /d "%~dp0"
where node >nul 2>nul || (echo Node.js is required: https://nodejs.org  & pause & exit /b 1)
echo [1/3] Installing trusted localhost certificate...
call npx --yes office-addin-dev-certs install || (echo Certificate step failed & pause & exit /b 1)
echo [2/3] Registering add-in with Excel...
call npx --yes office-addin-dev-settings sideload dist\manifest.xml excel || (echo Register step failed & pause & exit /b 1)
echo [3/3] Starting server. Open Excel: Home ^> Add-ins ^> Developer Add-ins ^> Excel 100x Pro
echo       (tab "100x Pro" appears on the ribbon).
node server\server.js
