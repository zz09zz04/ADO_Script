param(
    [Parameter(Mandatory = $true)]
    [string]$Path,                        # Excel 檔案路徑

    [Parameter(Mandatory = $true)]
    [string]$StatusColumnName             # 狀態欄名稱，例如 "Status"
)

# Excel 常數
$xlExpression = 2  # xlExpression

# 開啟 Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# 取得完整路徑，避免相對路徑問題
$fullPath = (Resolve-Path $Path).Path
Write-Host "Open file: $fullPath"
$workbook = $excel.Workbooks.Open($fullPath)
$sheet = $workbook.Sheets.Item(1)

# 找出狀態欄位置
$headerRow = 1
$colCount  = $sheet.UsedRange.Columns.Count
$rowCount  = $sheet.UsedRange.Rows.Count

$statusColIndex = $null
for ($c = 1; $c -le $colCount; $c++) {
    $value = $sheet.Cells.Item($headerRow, $c).Text
    if ($value -eq $StatusColumnName) {
        $statusColIndex = $c
        break
    }
}

if (-not $statusColIndex) {
    throw "❌ 找不到欄位 '$StatusColumnName'"
}

# 套用範圍
$range = $sheet.Range(
    $sheet.Cells.Item($headerRow+1, 1),
    $sheet.Cells.Item($rowCount, $colCount)
)
$range.FormatConditions.Delete()

# 計算狀態欄字母 (A, B, C...)
$colLetter = [char](64 + $statusColIndex)

# ✅ 設定條件式格式
$done  = $range.FormatConditions.Add($xlExpression, $null, "=$${colLetter}2=`"Done`"")
$done.Interior.Color = 0x00FF00   # 綠色

$late  = $range.FormatConditions.Add($xlExpression, $null, "=$${colLetter}2=`"Will Not Execute`"")
$late.Interior.Color = 0xCEC7FF   # 紫紅色

$doing = $range.FormatConditions.Add($xlExpression, $null, "=$${colLetter}2=`"Not Started`"")
$doing.Interior.Color = 0x9CEBFF  # 淺藍

# 儲存、關閉
$workbook.Save()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "✅ Conditional Formatting Applied Successfully"
