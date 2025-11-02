#$DebugPreference = "Continue" # Continue, SilentlyContinue and Stop
Write-Host "Check DebugPreference: $DebugPreference" -ForegroundColor Red

Write-Host `n`tCheck Config... -ForegroundColor Green
$configPath = ".\_config.json"
if (-not (Test-Path -Path $configPath)) {
    Write-Error "Error: $configPath can not be found!"
    return 
}
$config = Get-Content -Path ".\_config.json" | ConvertFrom-Json
Write-Debug $config
if ([string]::IsNullOrEmpty($config.AZURE_DEVOPS_EXT_PAT) -or `
    [string]::IsNullOrEmpty($config.Organization) -or `
    [string]::IsNullOrEmpty($config.Project)) {
    
    return
}

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
#    @{Name="Tag"; Expression={$_.fields."System.Tags"}},
    @{Name="URL"; Expression={$_.url}}
Write-Debug "Selected Items: $selectedItems"

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

#& ".\UpdateXlsx.ps1" -Path ".\WorkItems.xlsx" -StatusColumnName "State"
& ".\UpdateXlsx2.ps1" -Path "WorkItems.xlsx" -StatusColumnName "State"

### Extended: Output all query items detail by ID
if ($false) {
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
			
            # Call az boards work-item show
			$workItemDetailsJson = (echo $Env:AZURE_DEVOPS_EXT_PAT | az boards work-item show `
				--id $workItemId `
				--organization "https://dev.azure.com/$($config.Organization)")

            if (-not $workItemDetailsJson) {
                Write-Host "⚠️ Failed to get details for ID $workItemId" -ForegroundColor Yellow
                continue
            }

			#Write-Host $workItemDetailsJson        
			Write-Debug ($workItemDetailsJson -join "`n")

            # Generate filename
            $fileName = Join-Path $outputDir "$workItemId.json"
			
            # Write JSON to file
            $workItemDetailsJson | Out-File -FilePath $fileName -Encoding UTF8

            # Convert to PowerShell object
			$workItemDetails = $workItemDetailsJson | ConvertFrom-Json
			
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