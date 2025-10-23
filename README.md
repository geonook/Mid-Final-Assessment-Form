# 2526 Fall Midterm 班級報告生成器 (v2.1)

自動化生成期中考試班級報告的 Google Apps Script 專案。從 Google Sheets 讀取學生與班級資料，批次生成格式化的 Google Docs 報告，並自動按 GradeBand 合併為 PDF。

## 🎯 功能特色

- ✅ **自動佔位符替換** - 14 個考試資訊欄位自動填入
- ✅ **按學號排序** - 學生名單自動數值排序
- ✅ **固定字體 10pt** - 一致的格式化輸出
- ✅ **GradeBand 分組** - 自動建立子資料夾分類
- ✅ **PDF 自動合併** - 每個 GradeBand 合併為單一 PDF
- ✅ **兩種執行方式** - Google Sheets 選單或 Apps Script 編輯器直接執行
- ✅ **兩階段處理** - 可分別生成文件和合併 PDF，或一鍵完成
- ✅ 6 欄表格：學號、班級、中文名、英文名、出席、簽收
- ✅ 批次處理，含錯誤處理與超時防護
- ✅ 自動檔案命名：`{班級}_{老師}_2526_Fall_Midterm_v2_{時間戳}`

## 📋 快速開始

### 1. 前置準備

**必要資料**：
- Google Sheets 包含兩個工作表：
  - `Students`：學生資料（必要欄位：`English Class`, `ID`, `Home Room`, `Chinese Name`, `English Name`）
  - `Class`：班級資料（必要欄位：`ClassName`, `Teacher`, `GradeBand`, `Duration`, `Periods`, `Self-Study`, `Preparation`, `ExamTime`, `Level`, `Classroom`, `Proctor`, `Subject`, `Count`, `Students`）

**模板文件**：
- 使用 A4 橫式 Google Docs 模板
- 第一頁：考試監考指引（包含 14 個佔位符 {{...}}）
- 第二頁：考試資訊表 + 班級資訊表 + 學生名單表（第三個表格為填充目標）

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
  // Google Sheets 試算表 ID（用於 Apps Script 編輯器直接執行）
  spreadsheetId: '請替換為你的試算表 ID',

  // 你的模板文件 ID
  templateId: '請替換為你的模板 ID',

  // 輸出資料夾 ID
  outputFolderId: '請替換為你的資料夾 ID',

  // 其他設定保持預設即可
  studentTableIndex: 2,  // 第三個表格（學生名單）
  columnWidths: [75, 75, 95, 100, 85, 155],  // 總計 585pt
  fontSize: { large: 10, medium: 10, small: 10 },  // 固定 10pt
  semester: '2526_Fall_Midterm_v2',
  delayMs: 1000
};
```

**如何取得 ID？**
- **試算表 ID**：開啟 Google Sheets，從網址複製 ID
  - 網址格式：`https://docs.google.com/spreadsheets/d/{試算表ID}/edit`
- **模板 ID**：開啟 Google Docs 模板，從網址複製 ID
  - 網址格式：`https://docs.google.com/document/d/{模板ID}/edit`
- **資料夾 ID**：開啟 Google Drive 資料夾，從網址複製 ID
  - 網址格式：`https://drive.google.com/drive/folders/{資料夾ID}`

### 3. 執行生成

#### 方式 1：Google Sheets 選單（推薦給一般使用者）

1. 重新整理 Google Sheets 頁面
2. 點擊選單列的 **班級報告**
3. 選擇執行模式：
   - **🚀 一鍵執行（Docs + PDF）**：完整自動化流程（推薦）
   - **步驟 1: 生成所有 Google Docs**：僅生成文件（想在合併前檢視）
   - **步驟 2: 合併為 PDF（按 GradeBand）**：僅合併 PDF（需先完成步驟 1）
   - **🧪 測試單一班級（v2）**：測試單一班級（快速驗證）
4. 等待處理完成（會顯示成功/失敗訊息）
5. 前往輸出資料夾檢視生成的報告

#### 方式 2：Apps Script 編輯器直接執行（推薦給開發者）

1. 開啟 Apps Script 專案：**Extensions** → **Apps Script**
2. 從函數下拉選單選擇：
   - **`RUN_FULL_BATCH()`**：完整批次處理（Docs + PDF）- **推薦**
   - **`RUN_DOCS_ONLY()`**：僅生成 Google Docs
   - **`RUN_PDF_ONLY()`**：僅合併 PDF（需先執行 `RUN_DOCS_ONLY()`）
3. 點擊執行 ▶ 按鈕
4. 在控制台即時監控執行記錄
5. 前往輸出資料夾檢視結果

**執行時間估計**（以 168 個班級為例）：
- 僅生成 Docs：約 8-10 分鐘
- 僅合併 PDF：約 2-3 分鐘
- 完整批次：約 10-15 分鐘
- 測試單一班級：約 5-10 秒

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
| Teacher | ✅ | 授課教師 | John Doe |
| GradeBand | ✅ | 年級段 | G1 LT's |
| Duration | ✅ | 考試時長 | 50 mins |
| Periods | ✅ | 節次 | 3-4 |
| Self-Study | ✅ | 自習時間 | 2:25 - 2:45 pm |
| Preparation | ✅ | 準備時間 | 2:45 - 2:55 pm |
| ExamTime | ✅ | 考試時間 | 2:55 - 3:45 pm |
| Level | ✅ | 等級 | Intermediate |
| Classroom | ✅ | 教室 | A101 |
| Proctor | ✅ | 監考教師 | Jane Smith |
| Subject | ✅ | 科目 | English |
| Count | ✅ | 試卷數量 | 25 |
| Students | ✅ | 學生總數 | 20 |

## 📁 輸出結構

執行完成後，檔案會按 GradeBand 組織在子資料夾中：

```
輸出資料夾 (CONFIG.outputFolderId)
├── G1_LTs/
│   ├── A1_JohnDoe_2526_Fall_Midterm_v2_20251023_143052.docx
│   ├── A2_JaneSmith_2526_Fall_Midterm_v2_20251023_143055.docx
│   ├── ...
│   └── G1 LT's_2526_Fall_Midterm.pdf  ← 合併後的 PDF
├── G1_ITs/
│   ├── B1_...docx
│   └── G1 IT's_2526_Fall_Midterm.pdf
├── G2_LTs/
│   └── G2 LT's_2526_Fall_Midterm.pdf
└── ...
```

**檔案命名規則**：
- **個別 Docs**：`{班級}_{老師}_{學期}_{時間戳}.docx`
- **合併 PDF**：`{GradeBand}_{學期}.pdf`

## 🛠️ 輔助工具

### 分析模板結構
需要檢查模板格式時使用：
```javascript
// 在 Apps Script 編輯器執行：
輔助工具/分析模板.js -> analyzeTemplateStructure()
```
輸出資訊：頁面尺寸、表格數量、欄位結構、佔位符等

### 批次轉換 PDF（舊版，不建議）
**注意**：v2.1 已內建 PDF 合併功能，不需使用此工具。

若需要將個別 Docs 轉換為個別 PDF（而非 GradeBand 合併）：
1. 編輯 `輔助工具/轉檔.js` 中的資料夾 ID
2. 在 Apps Script 編輯器執行：`batchDownloadDocsToPDF()`
3. PDF 檔案會儲存到指定資料夾

## ⚙️ 技術細節

### 核心邏輯流程

**階段 1：文件生成**
```
讀取資料 → 班級排序(GradeBand→ClassName) → 學生分組 → 學生排序(ID)
→ 複製模板 → 替換佔位符 → 填充表格 → 格式化 → 儲存至 GradeBand 子資料夾
```

**階段 2：PDF 合併（可選）**
```
掃描 GradeBand 子資料夾 → 讀取所有 Docs → 排序(ClassName)
→ 建立臨時合併 Doc → 逐一複製元素 → 插入分頁符
→ 匯出為 PDF → 刪除臨時 Doc
```

### 班級排序（NEW）
兩層排序確保一致性：
```javascript
classList.sort((a, b) => {
  const gradeBandCompare = a.GradeBand.localeCompare(b.GradeBand);
  if (gradeBandCompare !== 0) return gradeBandCompare;
  return a.ClassName.localeCompare(b.ClassName);
});
```
範例順序：`G1 IT's (A1, A2, ...)` → `G1 LT's (B1, B2, ...)` → `G2 IT's (...)`

### 學生排序
使用數值排序確保學號正確排列：
```javascript
students.sort((a, b) => {
  return idA.localeCompare(idB, undefined, { numeric: true });
});
```
範例：`2026001` → `2026010` → `2026050` → `2026100`

### 佔位符替換（NEW）
自動替換 14 個考試資訊欄位：
```javascript
replacePlaceholders(body, classData);
// {{GradeBand}} → "G1 LT's"
// {{Duration}} → "50 mins"
// {{Periods}} → "3-4"
// ... 等 14 個欄位
```

### 固定字體大小（NEW）
不再根據學生人數調整，統一使用 10pt：
- **所有內容**：10pt
- **標題**：10pt 粗體
- **欄寬優化**：585pt 總寬度，標題完整顯示不需換行

### 表格填充方式
**不創建新表格**，而是填充模板現有表格：
1. 定位學生表格（`tables[2]` - 第三個表格）
2. 覆寫現有資料行（Solution C 方法避免 0-cell bug）
3. 逐行填入排序後的學生資料
4. 套用格式（對齊、字體、邊框）
5. 設定欄寬（總計 585pt）

### PDF 合併架構（NEW）
元素逐一複製保留模板格式：
1. 建立臨時合併文件
2. 對每個班級 Doc：
   - 開啟來源文件
   - 複製所有元素（段落、表格、清單）使用 `.copy()`
   - 附加至合併文件主體
   - 插入分頁符（班級之間）
3. 匯出合併 Doc 為 PDF
4. 刪除臨時 Doc
- **優點**：保留所有模板格式與結構，無硬編碼內容

## 🔍 常見問題

### Q: 執行時出現「找不到 Students 或 Class 工作表」錯誤？
**A**: 確認 Google Sheets 中的工作表名稱完全一致（區分大小寫）。

### Q: 學生名單沒有按學號排序？
**A**: 檢查 `Students` 工作表的 `ID` 欄位是否為文字格式（避免前置零被移除）。

### Q: 佔位符沒有被替換？
**A**: 確認：
1. Class 工作表包含所有必要欄位（GradeBand, Duration, Periods 等）
2. 欄位值不是空白或 `undefined`
3. 模板中的佔位符格式正確：`{{FieldName}}`（大小寫需一致）

### Q: PDF 合併失敗或超時？
**A**:
- 確保已先執行文件生成（階段 1）
- 檢查 GradeBand 子資料夾是否包含 Google Docs 檔案
- 若班級數量過多（>100），考慮分批處理
- 查看控制台記錄找出失敗的 GradeBand

### Q: 部分班級生成失敗？
**A**: 檢查執行記錄（Apps Script 編輯器 > 查看 > 記錄檔），通常是因為：
- 該班級沒有學生資料
- Class 工作表缺少必要欄位
- 模板文件權限不足

### Q: 執行時間過長會自動停止？
**A**: Apps Script 有 6 分鐘執行時間限制。本專案已實作：
- 策略性暫停（`Utilities.sleep()`）避免超時
- 錯誤恢復機制（個別失敗不影響整體）
- 若仍超時，可改用兩階段處理（分別執行 Docs 和 PDF）

### Q: 如何查看詳細執行記錄？
**A**:
- **方法 1**：Apps Script 編輯器 → **Executions** 查看執行歷史
- **方法 2**：使用 clasp 指令：
  ```bash
  clasp logs
  ```
- **方法 3**：執行時即時監控控制台輸出（Apps Script 編輯器執行時）

## 📦 專案結構

```
Mid-Final Assessment Form-template/
├── 主程式.js              # 主要生成邏輯（1,376 行）
│                          # - 文件生成（階段 1）
│                          # - PDF 合併（階段 2）
│                          # - Apps Script 直接執行函數
├── 輔助工具/
│   ├── 分析模板.js        # 模板結構分析工具
│   └── 轉檔.js            # PDF 批次轉換工具（舊版）
├── appsscript.json        # Apps Script 設定檔
├── .clasp.json            # Clasp 配置
├── CLAUDE.md              # 開發者文件（Claude Code 專用）
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

**版本**: 2.1 (PDF 合併 + 直接執行)
**學期**: 2526 Fall Midterm
**最後更新**: 2025-10-23

**v2.1 重大變更** (2025-10-23)：
- ✨ PDF 自動合併功能（按 GradeBand）
- ✨ Apps Script 編輯器直接執行（`RUN_FULL_BATCH()`, `RUN_DOCS_ONLY()`, `RUN_PDF_ONLY()`）
- ✨ 兩階段處理架構（文件生成 + PDF 合併）
- ✨ GradeBand 子資料夾自動組織
- ✨ 固定字體 10pt（取代動態字體）
- ✨ 欄寬優化至 585pt（標題完整顯示）
- ✨ 佔位符替換功能（14 個考試資訊欄位）
- ✨ 班級兩層排序（GradeBand → ClassName）
- ✨ 超時防護與錯誤恢復機制

**v2.0 變更** (2025-10-13)：
- 程式碼簡化（450 → 414 行，v2.1 擴充至 1,376 行）
- 集中配置到單一 CONFIG 物件
- 模板複製填充方法（避免 appendTableRow 0-cell bug）

## 📝 授權

本專案供內部教學使用。

---

**需要幫助？**
- 📖 查看 [CLAUDE.md](CLAUDE.md) 了解技術細節
- 🔧 檢查 Apps Script 執行記錄
- 📧 聯繫專案維護者
