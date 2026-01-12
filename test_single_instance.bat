@echo off
echo ==========================================
echo   測試單一實例功能
echo ==========================================
echo.

echo [測試 1] 啟動第一個實例...
start "" "bin\Debug\net8.0-windows\win-x64\DesktopAnnouncement.exe"
timeout /t 3 /nobreak >nul 2>&1
echo 完成！
echo.

echo [測試 2] 檢查運行中的實例數量...
for /f %%i in ('tasklist ^| find /c "DesktopAnnouncement.exe"') do set count1=%%i
echo 目前運行中的實例數量：%count1%
echo.

echo [測試 3] 啟動第二個實例（應該會關閉第一個）...
start "" "bin\Debug\net8.0-windows\win-x64\DesktopAnnouncement.exe"
timeout /t 3 /nobreak >nul 2>&1
echo 完成！
echo.

echo [測試 4] 再次檢查運行中的實例數量...
for /f %%i in ('tasklist ^| find /c "DesktopAnnouncement.exe"') do set count2=%%i
echo 目前運行中的實例數量：%count2%
echo.

echo ==========================================
echo   測試結果
echo ==========================================
echo 第一次執行後實例數量：%count1%
echo 第二次執行後實例數量：%count2%
echo.
echo 預期結果：兩次都應該只有 1 個實例
echo.

if "%count2%"=="1" (
    echo ✓ 測試通過！單一實例功能正常運作
) else (
    echo ✗ 測試失敗！發現多個實例
)
echo.

echo 按任意鍵關閉所有實例並退出...
pause >nul

echo 正在關閉所有實例...
taskkill /F /IM DesktopAnnouncement.exe >nul 2>&1
echo 完成！
