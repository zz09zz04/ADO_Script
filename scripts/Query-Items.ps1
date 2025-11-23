$DebugPreference = "Continue" # Continue, SilentlyContinue and Stop
Write-Host "Check DebugPreference: $DebugPreference" -ForegroundColor Red

# 找到 module 路徑
$modulesRoot = Join-Path $PSScriptRoot "..\modules"

# 載入 Config module
Import-Module (Join-Path $modulesRoot "GetConfig") -Force

Import-Module (Join-Path $modulesRoot "GetWorkItem") -Force

$config = Get-Config

$Env:AZURE_DEVOPS_EXT_PAT = $config.AZURE_DEVOPS_EXT_PAT
Write-Debug $Env:AZURE_DEVOPS_EXT_PAT

$outputDir = ".\output"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Write-Host `n`tConstruct Query Rule... -ForegroundColor Green
$selectFields = ($config.SelectItems | ForEach-Object { "[$_]" }) -join ", "
$tagConditions = ($config.Tags | ForEach-Object { "[System.Tags] CONTAINS '$_'" }) -join " AND "
$wiqlQuery = @"
SELECT $selectFields
FROM WorkItems
WHERE [System.TeamProject] = '$($config.Project)'
  AND [System.WorkItemType] = 'Test Case'
  AND ($tagConditions)
ORDER BY [System.Id] DESC
"@
Write-Host $wiqlQuery -ForegroundColor Gray
$flatWiqlQuery = $wiqlQuery -replace '\s+', ' '
Write-Host $flatWiqlQuery -ForegroundColor Cyan

Write-Host `n`tStarting Query... -ForegroundColor Green
$queryResultJson=$(echo $config.AZURE_DEVOPS_EXT_PAT | az boards query `
  --wiql $flatWiqlQuery `
  --organization "https://dev.azure.com/$($config.Organization)")

$outputFile = Join-Path $outputDir "QueryResult.json"
$queryResultJson | ConvertFrom-Json | ConvertTo-Json -Depth 10 | Out-File $outputFile -Encoding utf8

#Write-Host "Query Result: $queryResultJson"  -ForegroundColor Yellow
Write-Debug "Query Result: $queryResultJson"

$queryResult = $queryResultJson | ConvertFrom-Json

$queryResult | ForEach-Object { $_.fields."System.Title" }

Write-Host `n`tList Query Result... -ForegroundColor Green
$selectedItems = $queryResult | Select-Object `
    @{Name="ID"; Expression={$_.id}},
    @{Name="State"; Expression={$_.fields."System.State"}},
    @{Name="Title"; Expression={$_.fields."System.Title"}},
    @{Name="Owner"; Expression={$_.fields."AzureCSI-V1.2-RequirementsTest.Owner".displayName}},
    @{Name="URL"; Expression={$_.url}}
Write-Debug ($selectedItems | ConvertTo-Json -Depth 99)

Write-Host `n`tCreate Excel... -ForegroundColor Green
# Purpose: Collect system information and export to Excel
# Requirement: ImportExcel module
# =========================================

# Check and install ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

if ($selectedItems.Count -gt 0) {
    try {
        $selectedItems | Export-Csv -Path ".\WorkItems.csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        $selectedItems | Export-Excel -Path ".\WorkItems.xlsx" -WorksheetName "WorkItems" -AutoSize -BoldTopRow -FreezeTopRow
        Write-Host "`n`t✅ Export succeeded" -ForegroundColor Green
    }
    catch {
        Write-Host "`n`t❌ Export failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "`n`t⚠️ No data to export" -ForegroundColor Yellow
}

### Extended: Output all query items detail by ID
if ($true) {
	if ($null -ne $queryResult) {
    
        Write-Host "Found $($queryResult.Count) work items. Fetching details one by one..."

        # Create a timestamp-based output folder
        $rootDir = ".\output\workitems"
        $timestampFolder = Get-Date -Format "yyyyMMdd_HHmmss"
        $outputDir = Join-Path $rootDir $timestampFolder
        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
        }
		
        # Iterate each work item
        $index = 0
		foreach ($item in $queryResult) {
            $index++
			
			$workItemId = $item.id 
			
            Write-Host "[$index/$($queryResult.Count)] --- Fetching details for Work Item $workItemId ---"

            # Generate filename
            $fileName = Join-Path $outputDir "$workItemId.json"
			
			$workItemDetails = Get-WorkItem `
								-WorkItemId $workItemId `
								-Organization $config.Organization `
								-Project $config.Project `
								-Pat $config.AZURE_DEVOPS_EXT_PAT `
								-OutputPath $fileName

            # Print basic info
			Write-Host "  ID: $($workItemDetails.id)"
			Write-Host "  Title: $($workItemDetails.fields.'System.Title')"
			Write-Host "  State: $($workItemDetails.fields.'System.State')"
			Write-Host "  Created By: $($workItemDetails.fields.'System.CreatedBy'.displayName)"
            Write-Host "-------------------------------------------------`n"
		}

	} else {
        Write-Host "No work items found based on your query criteria."
	}
}