function Set-ExcelConditionalFormat {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        # 多欄位 → 多條件
        [hashtable]$ColumnRules = @{
            "State" = @{
                "Done"              = 0xC6EFCE
                "Not Started"       = 0x9CEBFF
                "Will Not Execute"  = 0xCEC7FF
            }

            "Owner" = @{
                "Ken Tsai"          = 0xFFFF99
                "Amy"               = 0xFFCCFF
            }

            "Iteration" = @{
                "__CONTAINS:CY25Q4\2Wk\2Wk03 (Nov 02 - Nov 15)"  = 0x0000FF
                "__CONTAINS:CY25Q4\2Wk\2Wk04 (Nov 16 - Nov 29)"  = 0x83A9F1
                "__CONTAINS:CY25Q4\2Wk\2Wk05 (Nov 30 - Dec 13)"  = 0x00C0FF
                "__CONTAINS:CY25Q4\2Wk\2Wk06 (Dec 14 - Dec 27)"  = 0x00FFFF
                "__CONTAINS:CY26Q1\2Wk\2Wk07 (Dec 28 - Jan 10)"  = 0xECC9A6
                "__CONTAINS:CY26Q1\2Wk\2Wk08 (Jan 11 - Jan 24)"  = 0xDD9EE4
            }

            "RobotName" = @{
                "__NOTEMPTY__"      = 0xD0F2DA
            }
        },

        # ⭐新增：固定 Title 欄位寬度（預設 40）
        [int]$TitleColumnWidth = 160,

        [switch]$Visible
    )

    try {
        # Excel COM
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $Visible.IsPresent

        $fullPath = (Resolve-Path $Path).Path
        $workbook = $excel.Workbooks.Open($fullPath)

        Add-Type -AssemblyName Microsoft.Office.Interop.Excel
        $xlExpression = [Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression

        foreach ($sheet in $workbook.Sheets) {

            $headerRow = 1
            $colCount  = $sheet.UsedRange.Columns.Count
            $rowCount  = $sheet.UsedRange.Rows.Count

            if ($rowCount -le 1) { continue }

            # 找所有欄位位置
            $columnIndexMap = @{}
            $titleColumnIndex = $null   # ⭐新增

            for ($c = 1; $c -le $colCount; $c++) {
                $value = $sheet.Cells.Item($headerRow, $c).Text

                # 用於 Conditional Format 的欄位
                if ($ColumnRules.ContainsKey($value)) {
                    $columnIndexMap[$value] = $c
                }

                # ⭐偵測 Title 欄位
                if ($value -eq "Title") {
                    $titleColumnIndex = $c
                }
            }

            # ⭐如果找到 Title 欄 → 固定欄寬
            if ($titleColumnIndex) {
                $sheet.Columns.Item($titleColumnIndex).ColumnWidth = $TitleColumnWidth
            }

            if ($columnIndexMap.Count -eq 0) { continue }

            # 套用範圍
            $range = $sheet.Range(
                $sheet.Cells.Item(2, 1),
                $sheet.Cells.Item($rowCount, $colCount)
            )
            $range.FormatConditions.Delete()

            foreach ($colName in $columnIndexMap.Keys) {

                $colIndex = $columnIndexMap[$colName]
                $rules = $ColumnRules[$colName]

                # 欄位轉字母
                $colLetter = ""
                $index = $colIndex
                while ($index -gt 0) {
                    $remainder = ($index - 1) % 26
                    $colLetter = [char](65 + $remainder) + $colLetter
                    $index = [math]::Floor(($index - 1) / 26)
                }

                foreach ($key in $rules.Keys) {

                    $color = $rules[$key]

                    if ($key -eq "__EMPTY__") {
                        $formula = "=LEN(`$${colLetter}2)=0"
                        $rule = $range.FormatConditions.Add($xlExpression, $null, $formula)
                        $rule.Interior.ColorIndex = -4142
                        continue
                    }

                    if ($key -eq "__NOTEMPTY__") {
                        $formula = "=LEN(`$${colLetter}2)>0"
                        $rule = $range.FormatConditions.Add($xlExpression, $null, $formula)
                        $rule.Interior.Color = $color
                        continue
                    }

                    # ⭐⭐⭐ 新增：字串包含比對規則 __CONTAINS:keyword
                    if ($key -like "__CONTAINS:*") {
                        $keyword = $key.Split(":", 2)[1]
                        $formula = "=ISNUMBER(SEARCH(""$keyword"", `$${colLetter}2))"
                        $rule = $range.FormatConditions.Add($xlExpression, $null, $formula)
                        $rule.Interior.Color = $color
                        continue
                    }

                    # 一般比對
                    $formula = "=`$${colLetter}2=`"$key`""
                    $rule = $range.FormatConditions.Add($xlExpression, $null, $formula)
                    $rule.Interior.Color = $color
                }
            }
        }

        $workbook.Save()
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    Write-Host "✅ Excel conditional formatting & Title column width applied successfully"
}

Export-ModuleMember -Function Set-ExcelConditionalFormat
