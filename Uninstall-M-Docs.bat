@echo off
setlocal
title Go cai dat M-Docs
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "TARGET=%STARTUP_DIR%\M-Docs.dotm"
set "LOCKFILE=%STARTUP_DIR%\~$M-Docs.dotm"
set "OLD_TARGET=%STARTUP_DIR%\MhaleMath.dotm"
set "OLD_LOCK=%STARTUP_DIR%\~$aleMath.dotm"

echo ============================================================
echo   GO CAI DAT M-DOCS
echo ============================================================
echo.

rem 1) Go nap add-in khoi Word dang mo (qua COM) va tat cac tien trinh Word chay ngam
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); foreach ($a in @($w.AddIns)) { if ($a.Name -like '*M-Docs*' -or $a.Name -like '*MhaleMath*') { $a.Installed = $false } } } catch {}; " ^
  "Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force -ErrorAction SilentlyContinue" 2>nul

rem 2) Xoa file add-in va file khoa tam
if exist "%LOCKFILE%" del /F /Q /A "%LOCKFILE%" 2>nul
if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
if exist "%~dp0~$M-Docs.dotm" del /F /Q /A "%~dp0~$M-Docs.dotm" 2>nul
if exist "%~dp0~$aleMath.dotm" del /F /Q /A "%~dp0~$aleMath.dotm" 2>nul

if exist "%TARGET%" del /F /Q /A "%TARGET%" 2>nul
if exist "%OLD_TARGET%" del /F /Q /A "%OLD_TARGET%" 2>nul

rem 3) Neu van con do Word dang giu khoa file, dong Word de xoa triet de
if exist "%TARGET%" (
    echo Word dang giu file add-in. Dang dong Word de hoan tat go cai dat...
    taskkill /F /IM WINWORD.EXE >nul 2>nul
    timeout /t 1 /nobreak >nul
    if exist "%LOCKFILE%" del /F /Q /A "%LOCKFILE%" 2>nul
    if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
    del /F /Q /A "%TARGET%" 2>nul
    if exist "%OLD_TARGET%" del /F /Q /A "%OLD_TARGET%" 2>nul
)
if exist "%OLD_TARGET%" (
    echo Word dang giu file add-in cu. Dang dong Word de hoan tat go cai dat...
    taskkill /F /IM WINWORD.EXE >nul 2>nul
    timeout /t 1 /nobreak >nul
    if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
    del /F /Q /A "%OLD_TARGET%" 2>nul
)

if exist "%TARGET%" (
    echo [LOI] Chua xoa duoc file: %TARGET%
    echo       Hay dong tat ca cua so Word roi chay lai file nay.
) else (
    echo Da go cai dat M-Docs thanh cong!
    echo Tab "M-Docs" da duoc xoa khoi Word.
)
echo.
pause
endlocal
