# 2526 Fall Midterm 班級報告生成器

自動化生成期中考試班級報告的 Google Apps Script 專案。從 Google Sheets 讀取學生與班級資料，批次生成格式化的 Google Docs 報告。

## 🎯 功能特色

- ✅ 自動按學號排序學生名單
- ✅ 支援最多 21 位學生（動態字體調整）
- ✅ 每個班級生成獨立的 2 頁報告
- ✅ 6 欄表格：學號、班級、中文名、英文名、出席、簽收
- ✅ 批次處理，含錯誤處理與速率限制
- ✅ 自動檔案命名：`{班級}_{老師}_2526_Fall_Midterm_{時間戳}`

## 📋 快速開始

### 1. 前置準備

**必要資料**：
- Google Sheets 包含兩個工作表：
  - `Students`：學生資料（必要欄位：`English Class`, `ID`, `Home Room`, `Chinese Name`, `English Name`）
  - `Class`：班級資料（必要欄位：`ClassName`, `Tea`）

**模板文件**：
- 使用 A4 橫式 Google Docs 模板
- 第一頁：考試監考指引（固定內容）
- 第二頁：班級資訊表 + 學生名單表（右側表格為填充目標）

### 2. 設定專案

#### 2.1 配置 Google Apps Script
1. 開啟你的 Google Sheets
2. 點擊 **Extensions** > **Apps Script**
3. 使用 clasp 將本專案推送至 Apps Script：
```bash
clasp push
```

#### 2.2 更新設定
編輯 `主程式.js` 中的 CONFIG 物件：
```javascript
const CONFIG = {
  // 你的模板文件 ID
  templateId: '請替換為你的模板 ID',

  // 輸出資料夾 ID
  outputFolderId: '請替換為你的資料夾 ID',

  // 其他設定保持預設即可
  studentTableIndex: 1,
  columnWidths: [90, 100, 140, 140, 80, 120],
  fontSize: { large: 11, medium: 10, small: 9 },
  semester: '2526_Fall_Midterm',
  delayMs: 1000
};
```

**如何取得 ID？**
- **模板 ID**：開啟 Google Docs 模板，從網址複製 ID
  - 網址格式：`https://docs.google.com/document/d/{模板ID}/edit`
- **資料夾 ID**：開啟 Google Drive 資料夾，從網址複製 ID
  - 網址格式：`https://drive.google.com/drive/folders/{資料夾ID}`

### 3. 執行生成

1. 重新整理 Google Sheets 頁面
2. 點擊選單列的 **班級報告** > **生成 2526 Fall Midterm 報告**
3. 等待處理完成（會顯示成功/失敗訊息）
4. 前往輸出資料夾檢視生成的報告

## 📊 資料格式要求

### Students 工作表
| 欄位名稱 | 必要 | 說明 | 範例 |
|---------|------|-----|------|
| English Class | ✅ | 英文班級名稱 | A1 |
| ID | ✅ | 學號 | 2026001 |
| Home Room | ✅ | 導師班 | 7A |
| Chinese Name | ✅ | 中文姓名 | 王小明 |
| English Name | ✅ | 英文姓名 | Ming Wang |

### Class 工作表
| 欄位名稱 | 必要 | 說明 | 範例 |
|---------|------|-----|------|
| ClassName | ✅ | 班級名稱 | A1 |
| Tea | ✅ | 老師姓名 | John Doe |

## 🛠️ 輔助工具

### 分析模板結構
需要檢查模板格式時使用：
```javascript
// 在 Apps Script 編輯器執行：
輔助工具/分析模板.js -> analyzeTemplateStructure()
```
輸出資訊：頁面尺寸、表格數量、欄位結構、佔位符等

### 批次轉換 PDF
生成報告後，若需 PDF 格式：
1. 編輯 `輔助工具/轉檔.js` 中的資料夾 ID
2. 在 Apps Script 編輯器執行：`batchDownloadDocsToPDF()`
3. PDF 檔案會儲存到指定資料夾

## ⚙️ 技術細節

### 核心邏輯流程
```
讀取資料 → 依班級分組 → 按學號排序 → 複製模板 → 填充表格 → 格式化 → 儲存
```

### 學生排序
使用數值排序確保學號正確排列：
```javascript
students.sort((a, b) => {
  return idA.localeCompare(idB, undefined, { numeric: true });
});
```
範例：`2026001` → `2026010` → `2026050` → `2026100`

### 動態字體大小
根據學生人數自動調整字體：
- **19-21 人**：9pt（緊密排版）
- **13-18 人**：10pt（平衡）
- **1-12 人**：11pt（舒適閱讀）

### 表格填充方式
**不創建新表格**，而是填充模板現有表格：
1. 定位右側學生表格（`tables[1]`）
2. 清除現有資料行（保留標題）
3. 逐行填入排序後的學生資料
4. 套用格式（對齊、字體、邊框）
5. 設定欄寬

## 🔍 常見問題

### Q: 執行時出現「找不到 Students 或 Class 工作表」錯誤？
**A**: 確認 Google Sheets 中的工作表名稱完全一致（區分大小寫）。

### Q: 學生名單沒有按學號排序？
**A**: 檢查 `Students` 工作表的 `ID` 欄位是否為文字格式（避免前置零被移除）。

### Q: 報告超過 2 頁？
**A**: 確認每班學生數 ≤ 21 人。若超過，請拆分班級或調整模板表格大小。

### Q: 部分班級生成失敗？
**A**: 檢查執行記錄（Apps Script 編輯器 > 查看 > 記錄檔），通常是因為：
- 該班級沒有學生資料
- 學生資料缺少必要欄位
- 模板文件權限不足

### Q: 如何查看詳細執行記錄？
**A**: 使用 clasp 指令：
```bash
clasp logs
```
或在 Apps Script 編輯器 > **Executions** 查看執行歷史。

## 📦 專案結構

```
Mid-Final Assessment Form-template/
├── 主程式.js              # 主要生成邏輯（247 行）
├── 輔助工具/
│   ├── 分析模板.js        # 模板結構分析工具
│   └── 轉檔.js            # PDF 批次轉換工具
├── appsscript.json        # Apps Script 設定檔
├── CLAUDE.md              # 開發者文件
└── README.md              # 本說明文件
```

## 🚀 開發指令

### 推送程式碼到 Apps Script
```bash
clasp push
```

### 從 Apps Script 拉取最新程式碼
```bash
clasp pull
```

### 在瀏覽器開啟專案
```bash
clasp open
```

### 查看執行記錄
```bash
clasp logs
```

## 🎓 版本資訊

**版本**: 2.0 (簡化版)
**學期**: 2526 Fall Midterm
**最後更新**: 2025-10-13

**重大變更**：
- ✨ 程式碼減少 45%（450 → 247 行）
- ✨ 移除頁面格式設定邏輯（模板預設）
- ✨ 移除佔位符替換邏輯（第一頁無佔位符）
- ✨ 集中配置到單一 CONFIG 物件
- ✨ 簡化為 8 個核心函數（原 15+ 個）
- ✨ 減少批次延遲至 1 秒（原 3 秒）

## 📝 授權

本專案供內部教學使用。

---

**需要幫助？**
- 📖 查看 [CLAUDE.md](CLAUDE.md) 了解技術細節
- 🔧 檢查 Apps Script 執行記錄
- 📧 聯繫專案維護者
