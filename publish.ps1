# 桌面公告小工具 - 自動打包腳本 (PowerShell 版本)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  桌面公告小工具 - 自動打包腳本" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 步驟 1: 清理舊檔案
Write-Host "[1/4] 清理舊的發布檔案..." -ForegroundColor Yellow
if (Test-Path "DesktopAnnouncement_發布版") {
    Remove-Item -Recurse -Force "DesktopAnnouncement_發布版"
}
if (Test-Path "bin\Release") {
    Remove-Item -Recurse -Force "bin\Release"
}
Write-Host "完成！" -ForegroundColor Green
Write-Host ""

# 步驟 2: 執行發布
Write-Host "[2/4] 執行發布指令..." -ForegroundColor Yellow
$publishCmd = "dotnet publish -c Release -r win-x64 --self-contained true " +
              "-p:PublishSingleFile=true " +
              "-p:IncludeNativeLibrariesForSelfExtract=true " +
              "-p:EnableCompressionInSingleFile=true"

Invoke-Expression $publishCmd

if ($LASTEXITCODE -ne 0) {
    Write-Host "發布失敗！" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出"
    exit 1
}
Write-Host "完成！" -ForegroundColor Green
Write-Host ""

# 步驟 3: 建立發布資料夾
Write-Host "[3/4] 建立發布資料夾..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path "DesktopAnnouncement_發布版" -Force | Out-Null
Copy-Item "bin\Release\net8.0-windows\win-x64\publish\DesktopAnnouncement.exe" "DesktopAnnouncement_發布版\"
Copy-Item "bin\Release\net8.0-windows\win-x64\publish\config.txt" "DesktopAnnouncement_發布版\"
Write-Host "完成！" -ForegroundColor Green
Write-Host ""

# 步驟 4: 建立使用說明
Write-Host "[4/4] 建立使用說明..." -ForegroundColor Yellow
$readme = @"
==========================================
  桌面公告小工具 - 使用說明
==========================================

雙擊 DesktopAnnouncement.exe 即可啟動程式

修改 config.txt 可以變更顯示內容：
格式：YYYY/MM/DD,標題文字
範例：2025/12/01,我們的光

程式會在桌面右下角顯示
切換到其他應用程式時會自動隱藏
點擊右上角的關閉按鈕可以退出

==========================================
版本：1.0
打包日期：$(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')
==========================================
"@
$readme | Out-File "DesktopAnnouncement_發布版\使用說明.txt" -Encoding UTF8
Write-Host "完成！" -ForegroundColor Green
Write-Host ""

# 顯示結果
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  打包完成！" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "發布位置：" -NoNewline
Write-Host "$PWD\DesktopAnnouncement_發布版\" -ForegroundColor Green
Write-Host ""

# 顯示檔案資訊
$exePath = "DesktopAnnouncement_發布版\DesktopAnnouncement.exe"
if (Test-Path $exePath) {
    $fileSize = (Get-Item $exePath).Length / 1MB
    Write-Host "檔案大小：" -NoNewline
    Write-Host ("{0:N2} MB" -f $fileSize) -ForegroundColor Yellow
    Write-Host ""
}

# 開啟資料夾
Write-Host "正在開啟資料夾..." -ForegroundColor Yellow
Start-Process "DesktopAnnouncement_發布版"
Write-Host ""
Read-Host "按 Enter 鍵退出"
