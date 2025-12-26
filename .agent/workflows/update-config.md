---
description: 更新專案配置（學期、模板 ID 等）
---

更新配置時，需同時修改兩個檔案以保持同步：

1. 修改 `config.json`（AI 可讀寫）
2. 同步更新 `主程式.js` 中的 CONFIG 物件

## 常用配置項目

### 更換學期

修改 `semester` 欄位，例如：

- `2526_Fall_Midterm_v2` → `2526_Fall_Final`

### 更換資料來源

修改 `spreadsheetId` 為新的 Google Sheets ID

### 更換模板

修改 `templateId` 為新的 Google Docs 模板 ID

### 更換輸出資料夾

修改 `outputFolderId` 為新的 Google Drive 資料夾 ID

## 驗證配置

// turbo

更新完成後，執行測試驗證：

```bash
clasp push && clasp run testSingleClass
```
