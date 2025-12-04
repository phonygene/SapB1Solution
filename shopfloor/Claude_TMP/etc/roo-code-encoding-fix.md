# Roo Code 執行指令：解決 apply_diff 編碼失敗問題

## 問題描述

`apply_diff` 失敗且相似度極高 (93%-99%)，原因是檔案編碼 (BOM) 與換行符號 (CRLF/LF) 的隱形差異。

## 解決方案

統一專案檔案格式為：**UTF-8 with BOM + LF**

---

## 步驟 0：備份現有檔案（必須先執行）

### 0.1 建立備份資料夾並複製所有目標檔案

```powershell
# 設定變數
$projectDir = "."  # 改成你的專案路徑
$backupDir = ".\_encoding_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")

# 建立備份資料夾
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
Write-Host "備份資料夾: $backupDir" -ForegroundColor Cyan

# 備份檔案（保留目錄結構）
$count = 0
foreach ($ext in $extensions) {
    Get-ChildItem -Path $projectDir -Filter $ext -Recurse | ForEach-Object {
        $relativePath = $_.FullName.Substring((Get-Item $projectDir).FullName.Length + 1)
        $destPath = Join-Path $backupDir $relativePath
        $destDir = Split-Path $destPath -Parent
        
        if (!(Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        
        Copy-Item $_.FullName -Destination $destPath
        $count++
    }
}

Write-Host "已備份 $count 個檔案到 $backupDir" -ForegroundColor Green

# 產生備份清單
Get-ChildItem -Path $backupDir -Recurse -File | 
    Select-Object FullName, Length, LastWriteTime |
    Export-Csv -Path "$backupDir\_backup_manifest.csv" -NoTypeInformation -Encoding UTF8

Write-Host "備份清單已儲存: $backupDir\_backup_manifest.csv" -ForegroundColor Green
```

### 0.2 驗證備份完整性

```powershell
# 確認備份檔案數量
$backupDir = ".\_encoding_backup_*" | Get-Item | Sort-Object Name -Descending | Select-Object -First 1
$backupCount = (Get-ChildItem -Path $backupDir -Recurse -File | Where-Object { $_.Name -ne "_backup_manifest.csv" }).Count
Write-Host "備份檔案數量: $backupCount" -ForegroundColor Cyan

# 確認備份資料夾存在且不為空才繼續
if ($backupCount -gt 0) {
    Write-Host "✓ 備份完成，可以繼續執行標準化" -ForegroundColor Green
} else {
    Write-Host "✗ 備份失敗，請勿繼續執行！" -ForegroundColor Red
    exit 1
}
```

---

## 步驟 1：建立 .editorconfig（專案根目錄）

```ini
# EditorConfig - 統一編碼格式
root = true

[*]
indent_style = space
indent_size = 4
end_of_line = lf
trim_trailing_whitespace = true
insert_final_newline = true

[*.{vb,aspx,ascx,asax,master,config}]
charset = utf-8-bom
end_of_line = lf
```

---

## 步驟 2：建立 .gitattributes（專案根目錄）

```
# 防止 Git 自動轉換換行符
* text=auto eol=lf

# ASP.NET 專案檔案
*.vb text eol=lf
*.aspx text eol=lf
*.ascx text eol=lf
*.asax text eol=lf
*.master text eol=lf
*.config text eol=lf
*.cs text eol=lf
```

---

## 步驟 3：一次性標準化現有檔案

執行以下 PowerShell 腳本，將所有目標檔案轉換為「UTF-8 with BOM + LF」：

```powershell
# ============================================
# 標準化腳本：UTF-8 BOM + LF
# ============================================

$extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")
$targetDir = "."  # 改成你的專案路徑

$successCount = 0
$errorCount = 0
$errorFiles = @()

foreach ($ext in $extensions) {
    Get-ChildItem -Path $targetDir -Filter $ext -Recurse | ForEach-Object {
        $file = $_.FullName
        
        try {
            # 讀取內容（自動處理各種編碼）
            $content = [System.IO.File]::ReadAllText($file)
            
            # 統一換行符為 LF
            $content = $content -replace "`r`n", "`n"
            $content = $content -replace "`r", "`n"
            
            # 寫入 UTF-8 with BOM
            $utf8Bom = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllText($file, $content, $utf8Bom)
            
            Write-Host "✓ $file" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "✗ $file - $($_.Exception.Message)" -ForegroundColor Red
            $errorCount++
            $errorFiles += $file
        }
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "轉換完成！" -ForegroundColor Cyan
Write-Host "成功: $successCount 個檔案" -ForegroundColor Green
Write-Host "失敗: $errorCount 個檔案" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })

if ($errorFiles.Count -gt 0) {
    Write-Host "`n失敗的檔案:" -ForegroundColor Red
    $errorFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}
```

---

## 步驟 4：驗證轉換結果

```powershell
# ============================================
# 驗證腳本：檢查所有檔案格式
# ============================================

$extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")
$targetDir = "."

$correctFiles = 0
$incorrectFiles = @()

foreach ($ext in $extensions) {
    Get-ChildItem -Path $targetDir -Filter $ext -Recurse | ForEach-Object {
        $file = $_.FullName
        $bytes = [System.IO.File]::ReadAllBytes($file)
        
        # 檢查 BOM
        $hasBom = ($bytes.Length -ge 3) -and 
                  ($bytes[0] -eq 0xEF) -and 
                  ($bytes[1] -eq 0xBB) -and 
                  ($bytes[2] -eq 0xBF)
        
        # 檢查換行符（跳過 BOM 後的內容）
        $content = [System.IO.File]::ReadAllText($file)
        $hasCRLF = $content.Contains("`r`n")
        
        if ($hasBom -and !$hasCRLF) {
            $correctFiles++
        } else {
            $issues = @()
            if (!$hasBom) { $issues += "缺少 BOM" }
            if ($hasCRLF) { $issues += "含有 CRLF" }
            $incorrectFiles += [PSCustomObject]@{
                File = $file
                Issues = $issues -join ", "
            }
        }
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "驗證結果" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "正確格式: $correctFiles 個檔案" -ForegroundColor Green

if ($incorrectFiles.Count -gt 0) {
    Write-Host "格式錯誤: $($incorrectFiles.Count) 個檔案" -ForegroundColor Red
    Write-Host "`n問題檔案:" -ForegroundColor Yellow
    $incorrectFiles | ForEach-Object {
        Write-Host "  - $($_.File)" -ForegroundColor Red
        Write-Host "    問題: $($_.Issues)" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n✓ 所有檔案格式正確！" -ForegroundColor Green
}
```

---

## 步驟 5：提交變更到 Git

```bash
git add .editorconfig .gitattributes
git add -A
git commit -m "chore: 標準化檔案編碼為 UTF-8 BOM + LF"
```

---

## Rollback：還原到備份狀態

如果轉換後出現問題（亂碼、編譯錯誤等），執行以下腳本還原：

```powershell
# ============================================
# Rollback 腳本：從備份還原
# ============================================

# 找到最新的備份資料夾
$backupDir = Get-ChildItem -Path "." -Directory -Filter "_encoding_backup_*" | 
             Sort-Object Name -Descending | 
             Select-Object -First 1

if ($null -eq $backupDir) {
    Write-Host "✗ 找不到備份資料夾！" -ForegroundColor Red
    exit 1
}

Write-Host "找到備份: $($backupDir.FullName)" -ForegroundColor Cyan
Write-Host "準備還原..." -ForegroundColor Yellow

# 確認還原
$confirm = Read-Host "確定要還原嗎？這會覆蓋目前的檔案 (y/N)"
if ($confirm -ne "y" -and $confirm -ne "Y") {
    Write-Host "已取消還原" -ForegroundColor Yellow
    exit 0
}

# 執行還原
$restoreCount = 0
$errorCount = 0

Get-ChildItem -Path $backupDir.FullName -Recurse -File | 
    Where-Object { $_.Name -ne "_backup_manifest.csv" } | 
    ForEach-Object {
        $relativePath = $_.FullName.Substring($backupDir.FullName.Length + 1)
        $destPath = Join-Path "." $relativePath
        
        try {
            # 確保目標目錄存在
            $destDir = Split-Path $destPath -Parent
            if (!(Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            
            Copy-Item $_.FullName -Destination $destPath -Force
            Write-Host "✓ 還原: $relativePath" -ForegroundColor Green
            $restoreCount++
        }
        catch {
            Write-Host "✗ 還原失敗: $relativePath - $($_.Exception.Message)" -ForegroundColor Red
            $errorCount++
        }
    }

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "還原完成！" -ForegroundColor Cyan
Write-Host "成功: $restoreCount 個檔案" -ForegroundColor Green
Write-Host "失敗: $errorCount 個檔案" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })

# 提醒刪除設定檔
Write-Host "`n如果要完全還原，請手動刪除以下檔案:" -ForegroundColor Yellow
Write-Host "  - .editorconfig" -ForegroundColor Yellow
Write-Host "  - .gitattributes" -ForegroundColor Yellow
```

---

## 後續 apply_diff 規則

完成上述設定後，在 Roo Code 的 Custom Instructions 或 `.roo/rules` 中加入：

```markdown
## 專案編碼規範

此專案的 .vb / .aspx / .ascx 等檔案格式為：
- 編碼：UTF-8 with BOM
- 換行符：LF

執行 apply_diff 時：
1. 確保輸出內容的換行符使用 LF（\n），不要使用 CRLF（\r\n）
2. 不要移除檔案開頭的 BOM (EF BB BF)
```

---

## 執行順序檢查清單

- [ ] 步驟 0：執行備份並確認備份完整
- [ ] 步驟 1：建立 .editorconfig
- [ ] 步驟 2：建立 .gitattributes  
- [ ] 步驟 3：執行標準化腳本
- [ ] 步驟 4：執行驗證腳本確認全部正確
- [ ] 測試：部署到測試環境確認沒有亂碼
- [ ] 步驟 5：提交到 Git
- [ ] （如有問題）執行 Rollback 腳本還原
