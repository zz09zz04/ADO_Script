
$modulesRoot = Join-Path $PSScriptRoot "..\modules"
Import-Module (Join-Path $modulesRoot "GetConfig") -Force

$config = Get-Config
# config parameters
$testplanname = $config.TestPlanName
$organization = $config.Organization
$project      = $config.Project
$planId       = $config.TestPlanId
$pat          = $config.AZURE_DEVOPS_EXT_PAT
$outputPath   = ".\output\TestPlan_$testplanname.xlsx"

# Generate Base64 authentication string
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))
$EncodedOrg     = [System.Uri]::EscapeDataString($organization)
$EncodedProject = [System.Uri]::EscapeDataString($project)

$fieldMap = $config.FieldNames
Write-Debug $fieldMap
foreach ($prop in $fieldMap.PSObject.Properties) {
    $key = $prop.Name
    $value = $prop.Value
    Write-Host "$key -> $value"
}

# Get all Test Suites under Test Plan
$uriSuites = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/testplan/Plans/$planId/suites?api-version=7.2-preview.1"
$content = Invoke-RestMethod -Uri $uriSuites -Headers @{Authorization = "Basic $base64AuthInfo"} -Method Get

$suites = $content.value
if (-not $suites) {
    Write-Host "❌ No test suites found under Test Plan $planId"
    exit
}

# Import Excel Module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# Remove old Excel file
if (Test-Path $outputPath) { Remove-Item $outputPath }

foreach ($suite in $suites) {
	$suiteId = $suite.id
	$suiteName = $suite.name -replace '[\\\/\:\*\?\[\]]','_'  # 清理非法字元

	Write-Host "📋 Fetching Test Cases for Suite [$suiteName] (ID=$suiteId)..."

	$uriCases = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/testplan/Plans/$planId/Suites/$suiteId/TestCase?api-version=7.2-preview.3"
	$casesContent = Invoke-RestMethod -Uri $uriCases -Headers @{Authorization = "Basic $base64AuthInfo"} -Method Get
	#$casesContent | ConvertTo-Json -Depth 10 | Write-Output
	Write-Host $casesContent -ForegroundColor Yellow

	if (-not $casesContent.value) {
		Write-Host "⚠️ No test cases found for this suite."
		continue
	}

#    $cases = $casesContent.value
#    $testCaseIds = $casesContent.value.workItem.id | Sort-Object -Unique
	$testCaseIds = @( ($casesContent.value.workItem.id | Sort-Object -Unique) )
	Write-Host $testCaseIds -ForegroundColor Cyan

	$body = @{
		ids    = $testCaseIds
		fields = $fieldMap.PSObject.Properties.Value
	} | ConvertTo-Json -Depth 10
	Write-Host $body -ForegroundColor DarkYellow

	$batchUrl = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/wit/workitemsbatch?api-version=7.2-preview.1"
	$allRows = @()
	$chunkSize = 200
	for ($i = 0; $i -lt $testCaseIds.Count; $i += $chunkSize) {
		$chunk = $testCaseIds[$i..([Math]::Min($i + $chunkSize - 1, $testCaseIds.Count - 1))]

		$batchResp = Invoke-RestMethod -Uri $batchUrl -Headers @{Authorization = "Basic $base64AuthInfo"} -Method Post -Body $body -ContentType "application/json"
		$output = $batchResp.value | ConvertTo-Json
		#Write-Host $output -ForegroundColor White

		# === 整理成輸出物件 ===
		foreach ($item in $batchResp.value) {
			$fields = $item.fields
			# According to config setting to create output items
			$row = [ordered]@{}
			foreach ($prop in $fieldMap.PSObject.Properties) {
				$key = $prop.Name
				$value = $prop.Value
				$testItemField = $value
#				$row[$key] = $fields.$testItemField
				Write-Debug "[Debug] $fields.$testItemField"
				if ($fields.$testItemField -is [psobject] -and $fields.$testItemField.PSObject.Properties.Name -contains 'displayName') {
					$row[$key] = $fields.$testItemField.displayName
				} elseif ($value -eq "AzureCSI-V1.1.MergedTags") {
					$robotName = $fields.$testItemField -replace '<[^>]+>', '' # Filter HTML style
					$row[$key] = $robotName
				} elseif ($value -eq "url") {
					$row[$key] = "https://azurecsi.visualstudio.com/$EncodedProject/_workitems/edit/$($fields.'System.Id')"
				} else {
					$row[$key] = $fields.$testItemField
				}
			}
			#$row["URL"] = "https://azurecsi.visualstudio.com/$EncodedProject/_workitems/edit/$($fields.'System.Id')"
			$allRows += [PSCustomObject]$row

			### Reserved for reference - start ###
			#$allRows += [PSCustomObject]@{
			#    "ID"        = $fields.'System.Id'
			#    "Title"     = $fields.'System.Title'
			#    "State"     = $fields.'System.State'
			#    "Owner"     = $fields.'AzureCSI-V1.2-RequirementsTest.Owner'.displayName
			#    "AssignedTo"= $fields.'System.AssignedTo'.displayName
			#    "Robot Name"= $fields.'AzureCSI-V1.1.MergedTags'
			#    "URL" = "https://azurecsi.visualstudio.com/$EncodedProject/_workitems/edit/$($fields.'System.Id')"
			#}
			#Write-Host $allRows -ForegroundColor White
			### Reserved for reference - end ###
		}
		$allRows | Export-Excel -Path $outputPath `
		    -WorksheetName $suiteName `
			-TableName ("Suite_" + $suiteId) `
			-AutoSize `
		    -BoldTopRow `
			-FreezeTopRow `
			-TableStyle None
	}
}

Write-Host "✅ Export complete: $outputPath"
