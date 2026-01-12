@echo off
chcp 65001 >nul
echo ==========================================
echo   桌面公告小工具 - 自動打包腳本
echo ==========================================
echo.

echo [1/4] 清理舊的發布檔案...
if exist "DesktopAnnouncement_發布版" rmdir /s /q "DesktopAnnouncement_發布版" 2>nul
if exist "bin\Release" rmdir /s /q "bin\Release" 2>nul
echo 完成！
echo.

echo [2/4] 執行發布指令...
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
if errorlevel 1 (
    echo.
    echo [錯誤] 發布失敗！
    echo.
    pause
    exit /b 1
)
echo 完成！
echo.

echo [3/4] 建立發布資料夾...
if not exist "DesktopAnnouncement_發布版" mkdir "DesktopAnnouncement_發布版"
if exist "bin\Release\net8.0-windows\win-x64\publish\DesktopAnnouncement.exe" (
    copy /Y "bin\Release\net8.0-windows\win-x64\publish\DesktopAnnouncement.exe" "DesktopAnnouncement_發布版\" >nul
    echo - DesktopAnnouncement.exe ... OK
) else (
    echo [錯誤] 找不到 DesktopAnnouncement.exe
)

if exist "bin\Release\net8.0-windows\win-x64\publish\config.txt" (
    copy /Y "bin\Release\net8.0-windows\win-x64\publish\config.txt" "DesktopAnnouncement_發布版\" >nul
    echo - config.txt ... OK
) else (
    echo [警告] 找不到 config.txt，將使用預設設定
    echo 2025/12/01,我們的光 > "DesktopAnnouncement_發布版\config.txt"
)
echo 完成！
echo.

echo [4/4] 建立使用說明...
(
echo ==========================================
echo   桌面公告小工具 - 使用說明
echo ==========================================
echo.
echo 雙擊 DesktopAnnouncement.exe 即可啟動程式
echo.
echo 修改 config.txt 可以變更顯示內容：
echo 格式：YYYY/MM/DD,標題文字
echo 範例：2025/12/01,我們的光
echo.
echo 程式會在桌面右下角顯示
echo 切換到其他應用程式時會自動隱藏
echo 點擊右上角的關閉按鈕可以退出
echo.
echo 功能特色：
echo - 可拖動視窗到任意位置，自動記憶
echo - 重複執行會自動關閉舊實例
echo - 刪除 position.txt 可重設位置到右下角
echo.
echo ==========================================
) > "DesktopAnnouncement_發布版\使用說明.txt"
echo 完成！
echo.

echo ==========================================
echo   打包完成！
echo ==========================================
echo.
echo 發布位置：%CD%\DesktopAnnouncement_發布版\
echo.

if exist "DesktopAnnouncement_發布版\DesktopAnnouncement.exe" (
    for %%A in ("DesktopAnnouncement_發布版\DesktopAnnouncement.exe") do (
        set size=%%~zA
        set /a sizeMB=%%~zA/1048576
    )
    echo 檔案大小：!sizeMB! MB
)
echo.
echo 正在開啟資料夾...
start "" "DesktopAnnouncement_發布版"
echo.
pause
