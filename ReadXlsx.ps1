$data = Import-Excel -Path ".\WorkItems.xlsx" -WorksheetName "WorkItems"

# 查看內容
$data | Format-Table
