@echo off
chcp 65001 >nul
title WIM Driver Manager - 系統需求檢查

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║       WIM Driver Manager - 系統需求檢查                       ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

REM 檢查 Windows 版本
echo [1/4] 檢查 Windows 版本...
for /f "tokens=4-5 delims=. " %%i in ('ver') do set VERSION=%%i.%%j
echo       系統版本: %VERSION%
if "%VERSION%"=="10.0" (
    echo       ✓ Windows 10/11 - 支援
) else if "%VERSION%"=="6.3" (
    echo       ✓ Windows 8.1 - 支援
) else if "%VERSION%"=="6.2" (
    echo       ⚠ Windows 8 - 可能支援
) else (
    echo       ✗ 不支援的 Windows 版本
)
echo.

REM 檢查系統架構
echo [2/4] 檢查系統架構...
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo       ✓ 64 位元系統 - 支援
) else (
    echo       ✗ 32 位元系統 - 不支援
    echo       此程式僅支援 64 位元 Windows
)
echo.

REM 檢查 VC++ Redistributable
echo [3/4] 檢查 Visual C++ Redistributable...
reg query "HKLM\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" /v Version >nul 2>&1
if %errorlevel%==0 (
    for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" /v Version 2^>nul ^| findstr Version') do (
        echo       ✓ 已安裝: %%a
    )
) else (
    echo       ✗ 未安裝 Visual C++ Redistributable
    echo.
    echo       請下載並安裝:
    echo       https://aka.ms/vs/17/release/vc_redist.x64.exe
)
echo.

REM 檢查管理員權限
echo [4/4] 檢查管理員權限...
net session >nul 2>&1
if %errorlevel%==0 (
    echo       ✓ 已具有管理員權限
) else (
    echo       ⚠ 尚未取得管理員權限
    echo       執行程式時需要以管理員身分執行
)
echo.

echo ════════════════════════════════════════════════════════════════
echo.
echo 系統需求總結:
echo   • Windows 10/11 64 位元
echo   • Visual C++ Redistributable 2015-2022 (x64)
echo   • 管理員權限 (用於 DISM 操作)
echo.
pause
