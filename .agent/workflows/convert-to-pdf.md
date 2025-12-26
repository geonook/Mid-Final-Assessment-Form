---
description: 將所有 Google Docs 批次轉換為 PDF
---

// turbo-all

1. 推送程式碼

```bash
clasp push
```

2. 執行 PDF 轉換

```bash
clasp run convertAllSubfoldersDocsToPDF
```

3. 檢查結果

```bash
clasp logs
```

注意事項：

- 168 個班級約需 6-9 分鐘
- PDF 會儲存在與 Docs 相同的子資料夾
- 自動跳過已存在的 PDF
