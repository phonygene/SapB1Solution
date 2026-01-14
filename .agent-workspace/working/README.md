# Agent 工作副本區

此目錄存放 Agent 在有衝突風險時複製的代碼副本。

## 目錄結構

```
working/
├── A1/
│   └── {task-id}/
│       └── {複製的檔案}
├── A2/
│   └── {task-id}/
│       └── {複製的檔案}
└── ...
```

## 使用時機

當藍圖中標註多個 Agent 需要修改同一檔案時，各 Agent 會：
1. 複製檔案到自己的 working/{agent-id}/{task-id}/ 目錄
2. 在副本上進行修改
3. Commit 時副本和主代碼一起提交
4. 由 Super Agent 執行 /integrate 時智能合併

## 生命週期

1. `/claim` 時建立（如果藍圖標註需要副本）
2. Agent 完成任務後 commit（副本保留）
3. `/integrate` 成功後刪除

## 注意事項

- 副本必須 commit，整合時才看得到
- 發布到 IIS 時會被排除（透過 Publish Profile）
- 不要手動刪除，由 /integrate 統一清理
