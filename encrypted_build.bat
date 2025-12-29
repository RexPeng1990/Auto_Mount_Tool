@echo off
chcp 65001 >nul
title WIM Driver Manager 加密打包工具

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║         WIM Driver Manager 加密打包工具                       ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo 加密級別說明:
echo   [1] 基本加密 - XOR + zlib 壓縮 (快速，基本保護)
echo   [2] 標準加密 - RC4 + 位元組打散 + 壓縮 (推薦)
echo   [3] 進階加密 - 多層加密 + 反調試 (最高保護)
echo.
echo ════════════════════════════════════════════════════════════════
echo.

set /p LEVEL="請選擇加密級別 [1-3] (預設: 3): "
if "%LEVEL%"=="" set LEVEL=3

echo.
echo 正在執行 Level %LEVEL% 加密打包...
echo.

python encrypted_build.py -l %LEVEL%

echo.
pause
