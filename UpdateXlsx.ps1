param(
    [Parameter(Mandatory = $true)]
    [string]$Path,                        # Excel 檔案路徑

    [Parameter(Mandatory = $true)]
    [string]$StatusColumnName,            # 狀態欄名稱，例如 "Status"

    [hashtable]$ColorMap = @{             # 狀態→顏色對照表，可自訂
        "Done"            = [System.Drawing.Color]::FromArgb(198,239,206)   # 綠
        "Not Started"     = [System.Drawing.Color]::FromArgb(255,242,204)   # 黃
        "Will Not Execute"= [System.Drawing.Color]::FromArgb(244,204,204)   # 紅
    }
)

Import-Module ImportExcel -ErrorAction Stop

if (-not (Test-Path $Path)) {
    throw "Excel file not found: $Path"
}

# 打開 Excel 檔
$pkg = Open-ExcelPackage -Path $Path
$ws  = $pkg.Workbook.Worksheets[1]

# 找出 Status 欄的實際位置
$headerRow = 1
$colCount  = $ws.Dimension.Columns
$rowCount  = $ws.Dimension.Rows

$statusColIndex = $null
for ($c = 1; $c -le $colCount; $c++) {
    if ($ws.Cells[$headerRow, $c].Text -eq $StatusColumnName) {
        $statusColIndex = $c
        break
    }
}

if (-not $statusColIndex) {
    throw "Column '$StatusColumnName' not found in Excel file."
}

# 根據狀態上色整列
for ($r = 2; $r -le $rowCount; $r++) {
    $status = $ws.Cells[$r, $statusColIndex].Text
    if ($ColorMap.ContainsKey($status)) {
        $color = $ColorMap[$status]
        for ($c = 1; $c -le $colCount; $c++) {
            $ws.Cells[$r, $c].Style.Fill.PatternType = 'Solid'
            $ws.Cells[$r, $c].Style.Fill.BackgroundColor.SetColor($color)
        }
    }
}

Close-ExcelPackage $pkg
Write-Host "✅ Formatting complete for $Path"
