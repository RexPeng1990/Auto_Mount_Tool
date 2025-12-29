@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

title WIM Driver Manager - 打包系統

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║      WIM Driver Manager - 統一打包系統              ║
echo ╠══════════════════════════════════════════════════════╣
echo ║  1. 直接打包 (Direct)   - 無保護，最高穩定性        ║
echo ║  2. 簡單保護 (Simple)   - 字符串混淆，推薦使用      ║
echo ║  3. 進階保護 (Advanced) - 多層加密，適合發布        ║
echo ║  4. 終極保護 (Ultimate) - 加密+反調試，最高安全     ║
echo ║  5. 構建全部版本                                    ║
echo ║  6. 清理並重新構建全部                              ║
echo ║  0. 退出                                            ║
echo ╚══════════════════════════════════════════════════════╝
echo.

set /p choice=請選擇打包選項 (0-6): 

if "%choice%"=="1" (
    echo.
    echo 正在構建 Direct 版本...
    python build_all.py -l direct
    goto end
)
if "%choice%"=="2" (
    echo.
    echo 正在構建 Simple 版本...
    python build_all.py -l simple
    goto end
)
if "%choice%"=="3" (
    echo.
    echo 正在構建 Advanced 版本...
    python build_all.py -l advanced
    goto end
)
if "%choice%"=="4" (
    echo.
    echo 正在構建 Ultimate 版本...
    python build_all.py -l ultimate
    goto end
)
if "%choice%"=="5" (
    echo.
    echo 正在構建全部版本...
    python build_all.py -l all
    goto end
)
if "%choice%"=="6" (
    echo.
    echo 正在清理並重新構建全部版本...
    python build_all.py -l all -c
    goto end
)
if "%choice%"=="0" (
    echo.
    echo 已退出
    goto exit
)

echo.
echo 無效選項，請重新選擇

:end
echo.
echo ════════════════════════════════════════════════════════
echo 構建完成！
echo ════════════════════════════════════════════════════════
echo.
pause

:exit
