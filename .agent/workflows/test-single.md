---
description: 測試生成單一班級報告（快速驗證）
---

// turbo-all

1. 推送程式碼

```bash
clasp push
```

2. 執行單一班級測試

```bash
clasp run testSingleClass
```

3. 檢查結果

```bash
clasp logs
```

預期結果：生成第一個班級的報告，用於驗證模板填充和佔位符替換是否正常。
