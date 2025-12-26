---
description: 生成所有班級的 Google Docs 報告
---

// turbo-all

1. 推送最新程式碼到 Apps Script

```bash
clasp push
```

2. 執行報告生成函數

```bash
clasp run RUN_DOCS_ONLY
```

3. 檢查執行結果

```bash
clasp logs
```

4. 開啟輸出資料夾確認結果

- 輸出資料夾: https://drive.google.com/drive/folders/1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi
- 檔案會按 GradeBand 分類在子資料夾中
