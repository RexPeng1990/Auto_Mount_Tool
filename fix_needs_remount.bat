@echo off
echo ========================================
echo  修復 "Needs Remount" 狀態工具
echo ========================================
echo.

echo 1. 檢查當前掛載狀態...
dism /Get-MountedWimInfo
echo.

echo 2. 嘗試重新掛載 (Remount)...
dism /Remount-Wim /MountDir:C:\ss\scratch
if %errorlevel% equ 0 (
    echo ✓ Remount 成功！
    goto check_final
) else (
    echo ✗ Remount 失敗，嘗試下一步...
)
echo.

echo 3. 嘗試提交並卸載...
dism /Unmount-Wim /MountDir:C:\ss\scratch /Commit
if %errorlevel% equ 0 (
    echo ✓ 提交卸載成功！
    goto check_final
) else (
    echo ✗ 提交卸載失敗，嘗試丟棄...
)
echo.

echo 4. 嘗試丟棄卸載...
dism /Unmount-Wim /MountDir:C:\ss\scratch /Discard
if %errorlevel% equ 0 (
    echo ✓ 丟棄卸載成功！
    goto check_final
) else (
    echo ✗ 丟棄卸載失敗，執行清理...
)
echo.

echo 5. 執行 WIM 清理...
dism /Cleanup-Wim
echo.

echo 6. 清理掛載點...
dism /Cleanup-Mountpoints
echo.

:check_final
echo 7. 檢查最終狀態...
dism /Get-MountedWimInfo
echo.

echo 修復完成！檢查上面的結果。
echo.
pause
