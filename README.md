# DesktopAnnouncement - 桌面公告小工具

## 🎯 專案概述

DesktopAnnouncement 是一個使用 C# / WPF (.NET 8) 開發的桌面公告小工具，具備以下特色：

- **智慧型桌面整合**：視窗固定於桌面底層（Z-Order 最底端），僅在回到桌面時顯現，不干擾其他程式運行。
- **水晶玻璃視覺效果**：現代化的半透明玻璃材質設計，提供優雅的 UI 體驗。
- **日期倒數計時**：自動計算目標日期至今的天數。
- **視窗可拖拉移動**：✨ 可自由拖動視窗到任意位置，並自動記憶下次啟動位置。
- **單一實例保護**：🆕 重複執行時自動關閉舊實例，確保系統中永遠只有一個視窗。
- **低資源耗用**：採用 Win32 視窗樣式優化與底層置放技術，對系統負擔極小。

---

## 🚀 快速開始

### 前置需求
- **Windows 10/11** 作業系統
- **.NET 8 Runtime** (如果使用非獨立發布版本)

### 執行方式
1. 下載並解壓縮發布檔案。
2. 確保 `config.txt` 與 `DesktopAnnouncement.exe` 在同一目錄。
3. 雙擊 `DesktopAnnouncement.exe` 即可啟動。

---

## ⚙️ 設定說明

### 1. `config.txt` - 公告內容設定
**格式**：`日期,標題文字`
**範例**：
```text
2025/12/01,我們的光
```
*   **日期格式**：必須為 `YYYY/MM/DD`。
*   **逗號**：為半形分隔符號。

### 2. `position.txt` - 位置儲存檔
此檔案由程式自動生成，紀錄視窗座標。
*   **重設位置**：若想恢復預設位置（右下角），直接刪除此檔案後重啟程式即可。

---

## 📦 打包與發布 (Publishing)

如果您是開發者，可以使用以下方式打包程式：

### 推薦方法：使用批次檔
執行專案根目錄下的 `publish_en.bat`，它會自動完成清理、編譯與檔案整理。
*   **輸出位置**：`Release_Package/` 資料夾。
*   **內容物**：`DesktopAnnouncement.exe` 與 `config.txt`。

### 手動打包指令
在終端機執行：
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```

### 常見打包問題排除
- **'dotnet' 不是內部或外部命令**：請安裝 [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)。
- **檔案被佔用**：請先關閉正在運行的 `DesktopAnnouncement.exe`。
- **中文亂碼**：Windows 批次檔若有編碼問題，請優先使用 `publish_en.bat`。

---

## 🛠️ 技術實作細節

### 1. 單一實例機制 (Single Instance)
程式使用 **Mutex (互斥鎖)** 確保唯一性：
- 啟動時嘗試建立全域 Mutex。
- 若 Mutex 已存在，程式會先尋找並強制關閉舊的進程 (`process.Kill()`)，然後啟動新實例。
- 這樣設計的好處是：重新執行 .exe 即可完成「重啟並載入新設定」的效果。

### 2. 桌面置底與不佔焦點技術
程式透過 Win32 API 確保視窗像壁紙一樣固定在桌面：
- **HWND_BOTTOM**：使用 `SetWindowPos` 將視窗強制置於 Z-Order 最底層。
- **WS_EX_NOACTIVATE**：設定擴展樣式，使視窗點擊時不會搶奪當前程式的焦點（Focus）。
- **自動校正**：透過 `Activated` 事件與定時器，確保視窗在被意外喚醒時能立即回到最底層。

---

## 📝 更新日誌

### v1.2 (2026/01/16)
- 🧹 整理專案結構，移除冗餘的發布腳本。
- 📝 整合所有說明文件至 README.md。

### v1.1 (2026/01/09)
- ✨ 新增視窗拖拉功能與位置記憶。
- ✨ 新增單一實例保護機制。
- 🐛 修正視窗初始位置為右下角。

### v1.0 (2026/01/09)
- 🎉 首次發布。
- ✨ 水晶玻璃效果、桌面智慧顯示/隱藏。

---

## 📂 專案結構
- `App.xaml.cs`: 應用程式入口，包含單一實例檢查。
- `MainWindow.xaml.cs`: UI 互動、桌面監控邏輯、日期計算。
- `NativeMethods.cs`: Win32 API 封裝。
- `publish_en.bat`: 核心發布腳本。

---

## 授權與支援
本專案為示範用途，可自由修改。如有問題，請確認 `config.txt` 編碼為 UTF-8 且格式正確。