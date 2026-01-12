@echo off
echo ==========================================
echo   Publishing DesktopAnnouncement
echo ==========================================
echo.

echo [1/3] Cleaning old files...
if exist "Release_Package" rmdir /s /q "Release_Package" 2>nul
if exist "bin\Release" rmdir /s /q "bin\Release" 2>nul
echo Done!
echo.

echo [2/3] Building Release...
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed!
    echo.
    pause
    exit /b 1
)
echo Done!
echo.

echo [3/3] Copying files to Release_Package...
if not exist "Release_Package" mkdir "Release_Package"
if exist "bin\Release\net8.0-windows\win-x64\publish\DesktopAnnouncement.exe" (
    copy /Y "bin\Release\net8.0-windows\win-x64\publish\DesktopAnnouncement.exe" "Release_Package\" >nul
    echo - DesktopAnnouncement.exe ... OK
) else (
    echo [ERROR] DesktopAnnouncement.exe not found
)

if exist "bin\Release\net8.0-windows\win-x64\publish\config.txt" (
    copy /Y "bin\Release\net8.0-windows\win-x64\publish\config.txt" "Release_Package\" >nul
    echo - config.txt ... OK
) else (
    echo [WARNING] config.txt not found, creating default
    echo 2025/12/01,Our Light > "Release_Package\config.txt"
)
echo Done!
echo.

echo ==========================================
echo   Build Complete!
echo ==========================================
echo.
echo Location: %CD%\Release_Package\
echo.
echo Opening folder...
start "" "Release_Package"
echo.
pause
