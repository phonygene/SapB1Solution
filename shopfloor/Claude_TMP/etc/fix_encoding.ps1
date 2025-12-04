# ============================================
# 標準化腳本：UTF-8 BOM + LF
# ============================================
# 預防中文亂碼
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

$extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")
$targetDir = "."

$successCount = 0
$errorCount = 0
$errorFiles = @()

foreach ($ext in $extensions) {
    Get-ChildItem -Path $targetDir -Filter $ext -Recurse | ForEach-Object {
        $file = $_.FullName
        
        try {
            # 讀取內容
            $content = [System.IO.File]::ReadAllText($file)
            
            # 統一換行符為 LF
            $content = $content -replace "`r`n", "`n"
            $content = $content -replace "`r", "`n"
            
            # 寫入 UTF-8 with BOM
            $utf8Bom = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllText($file, $content, $utf8Bom)
            
            Write-Host "V $file" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "X $file - $($_.Exception.Message)" -ForegroundColor Red
            $errorCount++
            $errorFiles += $file
        }
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "轉換完成！" -ForegroundColor Cyan
Write-Host "成功: $successCount 個檔案" -ForegroundColor Green
if ($errorCount -gt 0) {
    Write-Host "失敗: $errorCount 個檔案" -ForegroundColor Red
} else {
    Write-Host "失敗: 0 個檔案" -ForegroundColor Green
}

if ($errorFiles.Count -gt 0) {
    Write-Host "`n失敗的檔案:" -ForegroundColor Red
    foreach ($errFile in $errorFiles) {
        Write-Host "  - $errFile" -ForegroundColor Red
    }
}