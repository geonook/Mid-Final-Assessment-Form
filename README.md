# 2526 Fall Midterm 班級報告生成器 (v2.2-beta)

自動化生成期中考試班級報告的 Google Apps Script 專案。從 Google Sheets 讀取學生與班級資料，批次生成格式化的 Google Docs 報告。

> ⚠️ **v2.2-beta 注意事項**：Google Docs 合併功能目前處於測試階段，合併後的文件有格式問題（第一個班級正確，後續班級版面跑版）。建議使用個別班級文件，或等待功能改進。

## 🎯 功能特色

- ✅ **自動佔位符替換** - 14 個考試資訊欄位自動填入
- ✅ **按學號排序** - 學生名單自動數值排序
- ✅ **固定字體 10pt** - 一致的格式化輸出
- ✅ **GradeBand 分組** - 自動建立子資料夾分類
- ⚠️ **Google Docs 合併（測試中）** - 每個 GradeBand 合併為單一 Google Docs（有格式問題）
- ✅ **兩種執行方式** - Google Sheets 選單或 Apps Script 編輯器直接執行
- ✅ **兩階段處理** - 可分別生成文件和合併，或一鍵完成
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

> ⚠️ **建議**：由於 Google Docs 合併功能有格式問題，建議僅執行「步驟 1」生成個別班級文件

#### 方式 1：Google Sheets 選單（推薦給一般使用者）

1. 重新整理 Google Sheets 頁面
2. 點擊選單列的 **班級報告**
3. 選擇執行模式：
   - **步驟 1: 生成所有 Google Docs**：僅生成文件（✅ 推薦）
   - **步驟 2: 合併為 Google Docs（按 GradeBand）**：僅合併（⚠️ 有格式問題）
   - **🚀 一鍵執行（Docs + 合併）**：完整流程（⚠️ 合併有問題）
   - **🧪 測試單一班級（v2）**：測試單一班級（快速驗證）
4. 等待處理完成（會顯示成功/失敗訊息）
5. 前往輸出資料夾檢視生成的報告

#### 方式 2：Apps Script 編輯器直接執行（推薦給開發者）

1. 開啟 Apps Script 專案：**Extensions** → **Apps Script**
2. 從函數下拉選單選擇：
   - **`RUN_DOCS_ONLY()`**：僅生成 Google Docs（✅ 推薦）
   - **`RUN_MERGE_ONLY()`**：僅合併為 Google Docs（⚠️ 有格式問題）
   - **`RUN_FULL_BATCH()`**：完整批次處理（Docs + 合併）（⚠️ 合併有問題）
3. 點擊執行 ▶ 按鈕
4. 在控制台即時監控執行記錄
5. 前往輸出資料夾檢視結果

**執行時間估計**（以 168 個班級為例）：
- 僅生成 Docs：約 8-10 分鐘（✅ 推薦）
- 僅合併 Google Docs：約 2-3 分鐘（⚠️ 有格式問題）
- 完整批次：約 10-15 分鐘（⚠️ 合併有問題）
- 轉換為 PDF：約 6-9 分鐘（✅ 可選，見「輔助工具」章節）
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
│   ├── A1_JohnDoe_2526_Fall_Midterm_v2_20251023_143052.pdf  ← 可選：轉換後的 PDF
│   ├── A2_JaneSmith_2526_Fall_Midterm_v2_20251023_143055.docx
│   ├── A2_JaneSmith_2526_Fall_Midterm_v2_20251023_143055.pdf
│   └── ... (所有個別班級文件)
├── G1_ITs/
│   ├── B1_...docx
│   ├── B1_...pdf  ← 可選：轉換後的 PDF
│   └── ... (所有個別班級文件)
├── G2_LTs/
│   └── ... (所有個別班級文件 + PDF)
└── Merged/  ← ⚠️ 測試版：合併文件（有格式問題）
    ├── G1 LT's_2526_Fall_Midterm_Merged.docx
    ├── G1 IT's_2526_Fall_Midterm_Merged.docx
    └── G2 LT's_2526_Fall_Midterm_Merged.docx
```

**檔案命名規則**：
- **個別 Docs**：`{班級}_{老師}_{學期}_{時間戳}.docx`
- **個別 PDF**：`{班級}_{老師}_{學期}_{時間戳}.pdf`（可選，需額外執行轉換）
- **合併 Docs**：`{GradeBand}_{學期}_Merged.docx`（⚠️ 有格式問題）

## 🛠️ 輔助工具

### 批次轉換個別 PDF（✅ 推薦）
將所有班級 Google Docs 轉換為個別 PDF 檔案：

**執行步驟**：
1. 開啟 Apps Script 編輯器：**Extensions** → **Apps Script**
2. 從函數下拉選單選擇：**`convertAllSubfoldersDocsToPDF`**
3. 點擊執行 ▶ 按鈕
4. 等待處理完成（168 個班級約 6-9 分鐘）
5. PDF 檔案會儲存在與 Google Docs 相同的子資料夾內

**功能特色**：
- ✅ 自動處理所有 GradeBand 子資料夾
- ✅ 自動跳過 Merged 資料夾
- ✅ 自動跳過已存在的 PDF（不重複轉換）
- ✅ 顯示詳細處理進度和統計結果
- ✅ 超時防護（每 10 個檔案休息 1 秒）

**輸出結果**：
```
G1_LTs/
├── A1_Teacher_xxx.docx         ← Google Docs
├── A1_Teacher_xxx.pdf          ← 新增 PDF
├── A2_Teacher_xxx.docx
├── A2_Teacher_xxx.pdf
└── ...
```

### 分析模板結構
需要檢查模板格式時使用：
```javascript
// 在 Apps Script 編輯器執行：
輔助工具/分析模板.js -> analyzeTemplateStructure()
```
輸出資訊：頁面尺寸、表格數量、欄位結構、佔位符等

## ⚙️ 技術細節

### 核心邏輯流程

**階段 1：文件生成（✅ 穩定）**
```
讀取資料 → 班級排序(GradeBand→ClassName) → 學生分組 → 學生排序(ID)
→ 複製模板 → 替換佔位符 → 填充表格 → 格式化 → 儲存至 GradeBand 子資料夾
```

**階段 2：Google Docs 合併（⚠️ 測試中，有格式問題）**
```
掃描 GradeBand 子資料夾 → 讀取所有 Docs → 排序(ClassName)
→ 複製第一個 Doc 作為基底 (makeCopy) → 附加其他 Docs 元素 → 插入分頁符
→ 儲存到 Merged 資料夾
```

**⚠️ 已知問題**：第一個班級格式正確，後續班級版面跑版（元素逐一複製無法保留複雜的多欄版面配置）

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

### Google Docs 合併架構（⚠️ 測試中）
使用 `makeCopy()` 複製第一個文件 + 元素附加：
1. 複製第一個班級 Doc 作為基底（`DriveApp.makeCopy()`）
2. 對剩餘班級 Doc：
   - 插入分頁符
   - 開啟來源文件
   - 複製所有元素（段落、表格、清單）使用 `.copy()`
   - 附加至合併文件主體
3. 儲存到 Merged 資料夾

**⚠️ 已知問題**：
- 第一個班級格式正確（完整複製）
- 後續班級格式跑版（逐元素複製無法保留複雜版面）
- 根本原因：元素逐一複製不支援多欄版面、絕對定位、精確間距

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

### Q: 合併後的 Google Docs 格式跑版？（⚠️ v2.2-beta 已知問題）
**A**: 這是目前版本的已知限制：
- **原因**：元素逐一複製無法保留複雜的多欄版面配置
- **影響**：第一個班級正確，後續班級版面錯亂
- **暫時解決方案**：
  1. 僅使用個別班級文件（執行 `RUN_DOCS_ONLY()`）
  2. 手動合併或使用第三方工具
  3. 等待未來版本改進（可能使用 Google Docs API）
- **狀態**：正在研究更好的合併方法

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
├── 主程式.js              # 主要生成邏輯（~1,350 行）
│                          # - 文件生成（階段 1）✅
│                          # - Google Docs 合併（階段 2）⚠️
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

**版本**: 2.2-beta (Google Docs 合併測試版)
**學期**: 2526 Fall Midterm
**最後更新**: 2025-10-23

**v2.2-beta 變更** (2025-10-23)：
- ⚠️ **Google Docs 合併功能（測試中）**：改為合併到 Google Docs 而非 PDF
  - 使用 `makeCopy()` + 元素附加方式
  - **已知問題**：第一個班級正確，後續班級格式跑版
  - 合併檔案儲存至 Merged 資料夾
- 📝 更新所有文件說明反映目前狀態
- 🔄 函式重新命名：`mergeDocsToPDF*` → `mergeDocs*`
- 📁 新增 Merged 資料夾用於存放合併文件
- ⚠️ 建議僅使用階段 1（文件生成），跳過合併功能

**v2.1 變更** (2025-10-23)：
- ✨ PDF 自動合併功能（按 GradeBand）
- ✨ Apps Script 編輯器直接執行（`RUN_FULL_BATCH()`, `RUN_DOCS_ONLY()`, `RUN_MERGE_ONLY()`）
- ✨ 兩階段處理架構（文件生成 + 合併）
- ✨ GradeBand 子資料夾自動組織
- ✨ 固定字體 10pt（取代動態字體）
- ✨ 欄寬優化至 585pt（標題完整顯示）
- ✨ 佔位符替換功能（14 個考試資訊欄位）
- ✨ 班級兩層排序（GradeBand → ClassName）
- ✨ 超時防護與錯誤恢復機制

**v2.0 變更** (2025-10-13)：
- 程式碼簡化（450 → 414 行，v2.1 擴充至 ~1,350 行）
- 集中配置到單一 CONFIG 物件
- 模板複製填充方法（避免 appendTableRow 0-cell bug）

## 📝 授權

本專案供內部教學使用。

---

**需要幫助？**
- 📖 查看 [CLAUDE.md](CLAUDE.md) 了解技術細節
- 🔧 檢查 Apps Script 執行記錄
- 📧 聯繫專案維護者
