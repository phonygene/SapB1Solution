---
description: 初始化專案配置與符號連結 (project)
---

# 專案初始化

初始化專案的 claude-config 整合，包括符號連結設定和結構驗證。

## 執行步驟

### 1. 驗證 claude-config 存在

檢查 `C:\Projects\claude-config\` 是否存在：

```powershell
Test-Path "C:\Projects\claude-config"
```

若不存在，提示用戶執行：
```bash
git clone <claude-config-repo-url> "C:\Projects\claude-config"
```

### 2. 檢查符號連結狀態

檢查以下通用指令的符號連結：

| 連結位置 | 目標 |
|----------|------|
| `.claude/commands/lc.md` | `C:\Projects\claude-config\commands\lc.md` |
| `.claude/commands/lcp.md` | `C:\Projects\claude-config\commands\lcp.md` |
| `.claude/commands/reflect.md` | `C:\Projects\claude-config\commands\reflect.md` |
| `.claude/commands/backend.md` | `C:\Projects\claude-config\commands\backend.md` |
| `.claude/commands/manager.md` | `C:\Projects\claude-config\commands\manager.md` |
| `.claude/commands/super.md` | `C:\Projects\claude-config\commands\super.md` |
| `.claude/commands/ui-ux.md` | `C:\Projects\claude-config\commands\ui-ux.md` |
| `.claude/commands/init.md` | `C:\Projects\claude-config\commands\init.md` |

### 3. 建立缺失的符號連結

若符號連結不存在或損壞，執行建立：

```powershell
# 需要管理員權限或開發者模式
# 先刪除舊檔案（如果存在）
Remove-Item ".claude\commands\lc.md" -ErrorAction SilentlyContinue

# 建立符號連結
New-Item -ItemType SymbolicLink -Path ".claude\commands\lc.md" -Target "C:\Projects\claude-config\commands\lc.md"
```

**注意**：Windows 符號連結需要：
- 管理員權限執行 PowerShell，或
- 啟用開發者模式（設定 > 更新與安全性 > 開發人員專用）

### 4. 驗證專案結構

檢查以下目錄/檔案是否存在：

| 路徑 | 必要性 | 說明 |
|------|--------|------|
| `CLAUDE.md` | 必要 | 專案配置 |
| `.claude/` | 必要 | Claude Code 設定 |
| `.claude/commands/` | 必要 | Slash commands |
| `work-logs/` | 建議 | 工作日誌 |
| `work-logs/daily/` | 建議 | 每日日誌 |
| `work-logs/TODO.md` | 建議 | 待辦事項 |

若缺失建議項目，詢問是否建立。

### 5. 產出狀態報告

```
┌─────────────────────────────────────────────────────────┐
│  🔧 專案初始化狀態                                      │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  claude-config: ✅ 已連結                               │
│  位置: C:\Projects\claude-config\                       │
│                                                         │
│  符號連結狀態:                                          │
│  ├─ lc.md      ✅                                       │
│  ├─ lcp.md     ✅                                       │
│  ├─ reflect.md ✅                                       │
│  ├─ backend.md ✅                                       │
│  ├─ manager.md ✅                                       │
│  ├─ super.md   ✅                                       │
│  ├─ ui-ux.md   ✅                                       │
│  └─ init.md    ✅                                       │
│                                                         │
│  專案結構:                                              │
│  ├─ CLAUDE.md      ✅                                   │
│  ├─ .claude/       ✅                                   │
│  ├─ work-logs/     ✅                                   │
│  └─ VERSION        ✅                                   │
│                                                         │
│  ✅ 初始化完成                                          │
└─────────────────────────────────────────────────────────┘
```

## 故障排除

### 符號連結建立失敗

如果出現「權限不足」錯誤：

1. **方法一：啟用開發者模式**
   - 開啟「設定」
   - 「更新與安全性」→「開發人員專用」
   - 啟用「開發人員模式」

2. **方法二：以管理員身分執行**
   - 右鍵點擊 PowerShell
   - 選擇「以系統管理員身分執行」
   - 再次執行 `/init`

### claude-config 位置不同

如果 claude-config 不在 `C:\Projects\`，需要：
1. 修改符號連結的目標路徑
2. 或將 claude-config 移動到標準位置
