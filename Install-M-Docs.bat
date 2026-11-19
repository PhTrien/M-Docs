@echo off
setlocal
title Cai dat M-Docs
set "SRC=%~dp0M-Docs.dotm"
set "DEST=%APPDATA%\Microsoft\Word\STARTUP"
set "TARGET=%DEST%\M-Docs.dotm"
set "LOCKFILE=%DEST%\~$M-Docs.dotm"
set "OLD_TARGET=%DEST%\MhaleMath.dotm"
set "OLD_LOCK=%DEST%\~$aleMath.dotm"

echo ============================================================
echo   CAI DAT M-DOCS CHO MICROSOFT WORD
echo ============================================================
echo.

if not exist "%SRC%" (
    echo [LOI] Khong tim thay file M-Docs.dotm
    echo       Hay de file .bat nay CUNG THU MUC voi M-Docs.dotm roi chay lai.
    echo.
    pause
    exit /b 1
)

if not exist "%DEST%" mkdir "%DEST%"

rem Tam go nap add-in neu Word dang mo va tat cac tien trinh Word chay ngam
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); foreach ($a in @($w.AddIns)) { if ($a.Name -like '*M-Docs*' -or $a.Name -like '*MhaleMath*') { $a.Installed = $false } } } catch {}; " ^
  "Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | Stop-Process -Force -ErrorAction SilentlyContinue" 2>nul

if exist "%LOCKFILE%" del /F /Q /A "%LOCKFILE%" 2>nul
if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
if exist "%OLD_TARGET%" del /F /Q /A "%OLD_TARGET%" 2>nul

copy /Y "%SRC%" "%TARGET%" >nul 2>nul
if errorlevel 1 (
    echo Word dang giu file add-in cu. Dang dong Word de cap nhat...
    taskkill /F /IM WINWORD.EXE >nul 2>nul
    timeout /t 1 /nobreak >nul
    if exist "%LOCKFILE%" del /F /Q /A "%LOCKFILE%" 2>nul
    if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
    if exist "%OLD_TARGET%" del /F /Q /A "%OLD_TARGET%" 2>nul
    copy /Y "%SRC%" "%TARGET%" >nul
    if errorlevel 1 (
        echo [LOI] Khong sao chep duoc file vao thu muc STARTUP.
        echo       Hay tat Word roi chay lai file cai dat nay.
        echo.
        pause
        exit /b 1
    )
)

if exist "%OLD_TARGET%" (
    taskkill /F /IM WINWORD.EXE >nul 2>nul
    timeout /t 1 /nobreak >nul
    if exist "%OLD_LOCK%" del /F /Q /A "%OLD_LOCK%" 2>nul
    del /F /Q /A "%OLD_TARGET%" 2>nul
)

rem Bo danh dau "file tu internet" (Mark of the Web) de Word khong chan macro
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Unblock-File -LiteralPath '%TARGET%' -ErrorAction SilentlyContinue; " ^
  "try { $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); $ai = $w.AddIns.Add('%TARGET%', $true); $ai.Installed = $true } catch {}" 2>nul

echo Da cai xong: %TARGET%
echo.
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
    echo LUU Y: Neu chua thay cap nhat tren thanh cong cu, hay TAT va MO LAI Word.
) else (
    echo Bay gio mo Word, tab "M-Docs" se xuat hien tren thanh ribbon.
)
echo.
pause
endlocal
