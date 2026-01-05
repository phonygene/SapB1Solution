---
description: 手動觸發反思，識別問題模式 (project)
argument-hint: [-w|-m|-a]
---

# Reflect - 手動觸發反思

手動觸發工作反思，識別問題模式並更新經驗庫。

## 參數

- 無參數：反思最近 7 天的 work-logs
- `-w` 或 `--week`：本週反思
- `-m` 或 `--month`：本月反思
- `-a` 或 `--all`：全部反思（謹慎使用，Token 消耗大）

## 執行步驟

1. **讀取指定範圍的工作紀錄**
   - `work-logs/daily/YYYY-MM/YYYY-MM-DD.md`

2. **分析問題模式**
   - 識別重複出現的問題
   - 統計問題類型分布
   - 找出失敗/部分完成的案例

3. **更新 Insights**
   - 更新 `work-logs/insights/patterns.md`
   - 如有新模式，提出 `skills/` 優化建議

4. **產出反思報告**
   - 統計摘要
   - 重複問題清單
   - 建議的 skills 更新
   - 下一步行動項目

## 報告格式

```markdown
# 反思報告 - YYYY-MM-DD

## 統計
- 分析範圍：YYYY-MM-DD ~ YYYY-MM-DD
- 任務總數：N
- 成功：N | 部分完成：N | 失敗：N

## 重複問題模式
| 模式 | 出現次數 | 相關任務 |
|------|----------|----------|
| ... | ... | ... |

## 建議更新的 Skills
- [ ] skills/backend-checklist.md：新增 XXX 檢查項
- [ ] skills/patterns.md：新增 [P00X] 模式

## 下一步
1. ...
2. ...
```

## 使用時機

- 感覺最近頻繁遇到類似問題
- 完成一個階段性目標後
- 每週/每月定期執行
