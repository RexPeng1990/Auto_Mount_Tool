@echo off
chcp 65001 > nul
echo ======================================================================
echo           WIM/Driver 管理工具 - 單元測試套件
echo ======================================================================
echo.
echo 執行測試中...
echo.

cd /d "%~dp0.."
python -m tests.run_tests %*

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [成功] 所有測試通過！
) else (
    echo.
    echo [失敗] 部分測試未通過，請檢查上方輸出。
)

pause
