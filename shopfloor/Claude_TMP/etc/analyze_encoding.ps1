param(
    [string]$filePath
)

if (-not (Test-Path $filePath)) {
    Write-Host "File not found: $filePath"
    exit 1
}

$bytes = [System.IO.File]::ReadAllBytes($filePath)
$hasBOM = $false
$encoding = "Unknown"

# Check for UTF-8 BOM (EF BB BF)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $hasBOM = $true
    $encoding = "UTF-8 with BOM"
}
# Check for UTF-16 LE BOM (FF FE)
elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    $hasBOM = $true
    $encoding = "UTF-16 LE BOM"
}
# Check for UTF-16 BE BOM (FE FF)
elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
    $hasBOM = $true
    $encoding = "UTF-16 BE BOM"
}
else {
    $encoding = "No BOM (likely UTF-8 or ANSI)"
}

# Check Line Endings
$text = [System.IO.File]::ReadAllText($filePath)
$hasCR = $text.Contains("`r")
$hasLF = $text.Contains("`n")
$hasCRLF = $text.Contains("`r`n")

$lineEnding = "Unknown"
if ($hasCRLF) {
    $lineEnding = "CRLF (Windows)"
    # Check if mixed
    if ($text -match "[^`r]`n" -or $text -match "`r[^`n]") {
         $lineEnding = "Mixed (CRLF dominant)"
    }
} elseif ($hasLF) {
    $lineEnding = "LF (Unix/Linux)"
} elseif ($hasCR) {
    $lineEnding = "CR (Old Mac)"
} else {
    $lineEnding = "None (Single Line or Empty)"
}

Write-Host "File: $filePath"
Write-Host "Encoding: $encoding"
Write-Host "BOM Present: $hasBOM"
Write-Host "Line Endings: $lineEnding"