# 通用開發參考

> 🔗 **核心規則已移至**：`claude-config/core/`
> 本檔案僅保留專案特定的參考資料和錯誤案例

---

## UTF-8 BOM 操作步驟

> 🔴 **觸發規則**：ENCODE-001（修改 .aspx/.vb 後必須執行）

### 驗證 BOM
```powershell
[System.IO.File]::ReadAllBytes("檔案路徑")[0..2] -join ','
# 應輸出 239,187,191
```

### 修復 BOM
```powershell
$path = "檔案路徑"
$content = Get-Content -Path $path -Raw -Encoding UTF8
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)
```

---

## Windows 環境下的 Bash 執行規範

### MSBuild 執行

**正確寫法**：
```bash
"/c/Program Files (x86)/Microsoft Visual Studio/2019/Professional/MSBuild/Current/Bin/MSBuild.exe" \
  "C:/Projects/SapB1Solution/MgmSP/MgmSP.vbproj" \
  -t:Build \
  -p:Configuration=Debug \
  -verbosity:minimal
```

**常見錯誤**：
| 錯誤 | 說明 |
|------|------|
| `2019/Community/...` | 本機是 `Professional` |
| `/t:Build` | Git Bash 會解析為路徑，應用 `-t:Build` |

---

## 從錯誤中學習

### 2026-01-07: UTF-8 BOM 問題
- **問題**：新建的 .aspx 缺少 BOM，導致中文變亂碼
- **根因**：Claude Code 的 Write 工具預設不添加 BOM
- **解法**：修改後必須執行 BOM 驗證和修復

### 2026-01-14: UTF-8 BOM 問題（修改現有檔案）
- **問題**：編碼從 UTF-8 變成 Big5，Unicode 符號變成 ??
- **解法**：修改後也必須驗證編碼

### 2026-01-08: MSBuild 執行失敗
- **問題**：使用錯誤的 VS 路徑和參數格式
- **解法**：Git Bash 中參數用 `-` 而非 `/`
