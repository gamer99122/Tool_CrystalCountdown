@echo off
echo 正在打包程式...
echo.

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

if errorlevel 1 (
    echo.
    echo 打包失敗！
    pause
    exit /b 1
)

echo.
echo 打包完成！
echo.
echo 檔案位置：
echo %CD%\bin\Release\net8.0-windows\win-x64\publish\DesktopAnnouncement.exe
echo.

pause
start "" "bin\Release\net8.0-windows\win-x64\publish"
