# DesktopAnnouncement - 桌面公告小工具

## 專案概述

DesktopAnnouncement 是一個使用 C# / WPF (.NET 8) 開發的桌面公告小工具，具備以下特色：

- **智慧型顯示/隱藏**：僅在 Windows 桌面顯示時自動顯現，切換到其他應用程式時自動隱藏
- **水晶玻璃視覺效果**：現代化的半透明玻璃材質設計，不影響桌面操作
- **日期倒數計時**：從設定檔讀取目標日期，自動計算已經過的天數
- **視窗可拖拉移動**：✨ 可自由拖動視窗到任意位置，自動記憶下次啟動位置
- **單一實例保護**：🆕 重複執行時自動關閉舊實例，避免多個視窗同時運行
- **低資源耗用**：使用高效能的 Win32 API 監控，對系統影響極小

## 專案結構

```
DesktopAnnouncement/
├── 核心程式碼
│   ├── NativeMethods.cs          # Win32 API 封裝（GetForegroundWindow, GetClassName 等）
│   ├── MainWindow.xaml           # 主視窗 UI 設計（水晶玻璃風格）
│   ├── MainWindow.xaml.cs        # 主視窗邏輯（監控、日期計算、位置管理）
│   ├── App.xaml                  # 應用程式入口點
│   ├── App.xaml.cs               # 應用程式生命週期管理
│   └── DesktopAnnouncement.csproj # 專案檔（.NET 8）
├── 設定檔
│   ├── config.txt                # 設定檔（日期與標題）
│   └── position.txt              # 視窗位置儲存檔（自動生成）
├── 編譯與發佈腳本
│   ├── build.bat                 # 編譯腳本
│   ├── publish.bat               # 發佈腳本（標準版）
│   ├── publish_en.bat            # 發佈腳本（英文版本）
│   ├── publish_simple.bat        # 簡單發佈腳本
│   └── publish.ps1               # PowerShell 發佈腳本
├── 文件與說明
│   ├── README.md                 # 本說明文件
│   ├── 單一實例說明.md            # 單一實例機制說明
│   ├── 更新說明_v1.1.md          # 版本更新日誌
│   ├── 打包說明.md               # 打包編譯說明
│   ├── 打包說明_最新.txt         # 最新打包說明
│   └── 打包問題排除.md           # 常見打包問題解決方案
└── bin/, obj/                    # 編譯生成的檔案（未版控）
```

## 技術架構

### 1. Win32 API 監控機制

- **使用 API**：`GetForegroundWindow()`, `GetClassName()`, `GetWindowThreadProcessId()`
- **監控頻率**：500ms 輪詢（平衡即時性與效能）
- **偵測目標**：
  - 桌面類別名稱：`Progman`, `WorkerW`
  - 檔案總管桌面：`explorer.exe` + 特定類別名稱
- **資源管理**：正確使用 `CloseHandle()` 釋放 Win32 Handle

### 2. UI 視覺設計

- **視窗屬性**：
  - `WindowStyle="None"`：無邊框
  - `AllowsTransparency="True"`：支援透明度
  - `Background="Transparent"`：背景透明
  - `ShowInTaskbar="False"`：不顯示在工作列
  - `IsHitTestVisible="False"`：滑鼠穿透

- **水晶玻璃效果**：
  - 漸層半透明背景（Alpha: 0.2-0.4）
  - `DropShadowEffect` 光暈效果
  - `BlurEffect` 模糊背景層
  - 圓角邊框（CornerRadius: 20）

### 3. 資料處理

- **設定檔格式**：`YYYY/MM/DD,標題文字`
- **日期計算**：`DateTime.Today - targetDate = 經過天數`
- **錯誤處理**：
  - 檔案不存在檢查
  - 日期格式驗證
  - 空白內容處理

## 文件說明

### 快速參考
- **單一實例說明.md** - 了解如何實現單一實例保護機制
- **更新說明_v1.1.md** - 查看版本更新內容與改進事項
- **打包說明.md** - 學習如何打包成可執行檔案
- **打包說明_最新.txt** - 最新的打包與編譯相關建議
- **打包問題排除.md** - 排查常見的打包與發佈問題

## 編譯與執行

### 前置需求

- **.NET 8 SDK** 或更高版本
- Windows 10/11 作業系統
- Visual Studio 2022（建議）或 Visual Studio Code + C# Dev Kit

### 編譯指令

```bash
# 開發環境編譯
dotnet build

# 發布為單一執行檔（含 .NET Runtime）
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

### 執行方式

1. 確保 `config.txt` 與執行檔在同一目錄
2. 執行 `DesktopAnnouncement.exe`
3. 程式將常駐背景，在桌面顯示時自動顯現

## config.txt 設定說明

**格式**：`日期,標題文字`

**範例**：
```
2024/01/01,與你相遇的第
```

**結果**：
- 如果今天是 2026/01/09，則顯示：「與你相遇的第 **374** 天」

**注意事項**：
- 日期格式必須為 `YYYY/MM/DD`
- 逗號為分隔符號
- 標題文字不可為空

## 視窗位置功能 ✨

### 如何使用
1. **拖動視窗**：點擊視窗任意位置（除關閉按鈕外），按住滑鼠左鍵拖動
2. **自動儲存**：放開滑鼠後，位置會自動儲存到 `position.txt`
3. **自動載入**：下次啟動時會在上次的位置顯示
4. **重設位置**：刪除 `position.txt` 即可恢復預設右下角位置

### position.txt 格式
```
1420  ← X 座標（距離螢幕左側的像素）
850   ← Y 座標（距離螢幕頂部的像素）
```

**特色**：
- ✅ 螢幕邊界自動檢查
- ✅ 工作列自動避讓
- ✅ 位置無效時使用預設位置
- ✅ 支援任意位置拖放

## 核心功能說明

### 1. 桌面偵測邏輯 (NativeMethods.cs:152)

```csharp
internal static bool IsForegroundWindowDesktop()
{
    string className = GetForegroundWindowClassName();

    // 檢查桌面類別名稱
    if (className.Equals("Progman") || className.Equals("WorkerW"))
        return true;

    // 檢查是否為 Explorer.exe
    string processPath = GetForegroundWindowProcessPath();
    if (Path.GetFileName(processPath).Equals("explorer.exe"))
        return className.Equals("Shell_TrayWnd");

    return false;
}
```

### 2. 日期計算邏輯 (MainWindow.xaml.cs:99)

```csharp
private void UpdateAnnouncementDisplay()
{
    DateTime today = DateTime.Today;
    TimeSpan difference = today - _targetDate;
    int daysPassed = (int)difference.TotalDays;

    DaysTextBlock.Text = daysPassed.ToString("N0");
}
```

### 3. 顯示/隱藏控制 (MainWindow.xaml.cs:144)

```csharp
private void UpdateVisibilityByDesktopState()
{
    bool isDesktop = NativeMethods.IsForegroundWindowDesktop();

    if (isDesktop != _lastDesktopState)
    {
        _lastDesktopState = isDesktop;
        this.Opacity = isDesktop ? 1.0 : 0.0;
    }
}
```

## 效能特性

- **CPU 使用率**：< 0.5%（閒置時）
- **記憶體佔用**：約 15-20 MB
- **啟動時間**：< 1 秒
- **切換反應時間**：< 500ms

## 限制與注意事項

1. **不支援多螢幕**：預設顯示在主螢幕中央
2. **不可調整大小**：固定尺寸 480x200
3. **不支援自訂主題**：需修改 XAML 原始碼
4. **僅支援 Windows**：使用 Win32 API，無法跨平台

## 安全性考量

- ✅ 正確釋放 Win32 Handle（使用 `CloseHandle()`）
- ✅ 例外處理機制（避免應用程式崩潰）
- ✅ 無網路存取（完全本地運作）
- ✅ 無敏感權限需求（僅讀取視窗資訊）

## 授權資訊

本專案為示範用途，可自由修改與使用。

## 技術支援

如有問題，請檢查以下項目：

1. **.NET 8 是否正確安裝**：`dotnet --version`
2. **config.txt 格式是否正確**：確認日期格式為 `YYYY/MM/DD`
3. **檔案權限是否正常**：確保 config.txt 可讀取
4. **防毒軟體是否阻擋**：部分防毒軟體會誤判 Win32 API 調用

## 未來改進方向

- [ ] 支援自訂視窗位置與大小
- [ ] 新增多主題切換功能
- [ ] 支援多個公告輪播
- [ ] 加入淡入/淡出動畫效果
- [ ] 提供系統托盤圖示與設定介面
