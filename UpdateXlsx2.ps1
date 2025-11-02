param(
    [Parameter(Mandatory = $true)]
    [string]$Path,                        # Excel 檔案路徑

    [Parameter(Mandatory = $true)]
    [string]$StatusColumnName,            # 狀態欄名稱，例如 "Status"

    [hashtable]$ColorMap = @{             # 狀態 → 顏色對照表
        "Done"            = 0x00FF00    # 綠
        "Not Started"     = 0x9CEBFF    # 淺藍
        "Will Not Execute"= 0xCEC7FF    # 紫紅
    }
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

Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlExpression = [Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression

# 計算狀態欄字母 (A, B, C...)
$colLetter = [char](64 + $statusColIndex)
Write-Host "Formula used: " "=`$${colLetter}2=`"Done`""

# ✅ 根據 hashtable 自動建立條件式格式
foreach ($status in $ColorMap.Keys) {
    $color = $ColorMap[$status]
    $formula = "=`$${colLetter}2=`"$status`""
    Write-Host "新增條件: $formula (顏色: $color)"
    $rule = $range.FormatConditions.Add($xlExpression, $null, $formula)
    $rule.Interior.Color = $color
}

# 儲存、關閉
$workbook.Save()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "✅ Conditional Formatting Applied Successfully"
