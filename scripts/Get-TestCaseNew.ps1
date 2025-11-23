function Get-AdoTestPlanExport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [string]$Organization,

        [Parameter(Mandatory)]
        [string]$Project,

        [Parameter(Mandatory)]
        [string]$Pat,

        [Parameter(Mandatory)]
        [string]$PlanId,

        [Parameter(Mandatory)]
        [object]$FieldMap
    )

    Write-Verbose "Exporting ADO Test Plan: $PlanId"
	
    # Encode PAT
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    $EncodedOrg     = [System.Uri]::EscapeDataString($Organization)
    $EncodedProject = [System.Uri]::EscapeDataString($Project)

    # Display field map
    foreach ($prop in $FieldMap.PSObject.Properties) {
        Write-Verbose "$($prop.Name) -> $($prop.Value)"
    }

    # Retrieve Suites
    $uriSuites = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/testplan/Plans/$PlanId/suites?api-version=7.2-preview.1"
    $content = Invoke-RestMethod -Uri $uriSuites -Headers @{Authorization = "Basic $base64AuthInfo"}

    $suites = $content.value
    if (-not $suites) {
        Write-Host "❌ No test suites found under Test Plan $PlanId"
        return
    }

    # Import Excel Module
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel -Force

    # Remove previous export
    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    foreach ($suite in $suites) {

        $suiteId = $suite.id
        $suiteName = $suite.name -replace '[\\\/\:\*\?\[\]]','_'

        Write-Host "📋 Fetching Test Cases for Suite [$suiteName] (ID=$suiteId)..."

        $uriCases = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/testplan/Plans/$PlanId/Suites/$suiteId/TestCase?api-version=7.2-preview.3"
        $casesContent = Invoke-RestMethod -Uri $uriCases -Headers @{Authorization = "Basic $base64AuthInfo"}
        Write-Host $casesContent -ForegroundColor Yellow

        if (-not $casesContent.value) {
            Write-Host "⚠ No test cases in suite $suiteName"
            continue
        }

        $testCaseIds = @(($casesContent.value.workItem.id | Sort-Object -Unique))
        Write-Host $testCaseIds -ForegroundColor Cyan
        
        $body = @{
            ids    = $testCaseIds
            fields = $FieldMap.PSObject.Properties.Value
        } | ConvertTo-Json -Depth 10
        Write-Host $body -ForegroundColor DarkYellow

        $batchUrl = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/wit/workitemsbatch?api-version=7.2-preview.1"

        $allRows = @()

        $batchResp = Invoke-RestMethod -Uri $batchUrl `
            -Headers @{Authorization = "Basic $base64AuthInfo"} `
            -Method Post `
            -Body $body `
            -ContentType "application/json"

        foreach ($item in $batchResp.value) {

            $fields = $item.fields
            # According to config setting to create output items
            $row = [ordered]@{}

            foreach ($prop in $FieldMap.PSObject.Properties) {
                $key   = $prop.Name
                $value = $prop.Value

                if ($fields.$value -is [psobject] -and $fields.$value.PSObject.Properties.Name -contains 'displayName') {
                    $row[$key] = $fields.$value.displayName
                }
                elseif ($value -eq "AzureCSI-V1.1.MergedTags") {
                    $row[$key] = ($fields.$value -replace '<[^>]+>', '') # Filter HTML style
                }
                elseif ($value -eq "url") {
                    $row[$key] = "https://azurecsi.visualstudio.com/$EncodedProject/_workitems/edit/$($fields.'System.Id')"
                }
                else {
                    $row[$key] = $fields.$value
                }
            }

            $allRows += [PSCustomObject]$row
        }

        $allRows | Export-Excel -Path $OutputPath `
            -WorksheetName $suiteName `
            -TableName ("Suite_" + $suiteId) `
            -AutoSize `
	    -BoldTopRow `
	    -FreezeTopRow `
	    -TableStyle None
    }

    Write-Host "✅ Export completed: $OutputPath"
}

# ======================================================
# ⭐ MAIN MODE：只有直接執行這個 .ps1 才會進入這裡
# ======================================================
Write-Host "🛠 $PSCommandPath"
Write-Host "🛠 $($MyInvocation.ScriptName)"
"MyInvocation.InvocationName = $($MyInvocation.InvocationName)"
"MyInvocation.MyCommand.Name = $($MyInvocation.MyCommand.Name)"

if ($PSCommandPath -eq $MyInvocation.InvocationName) {
    Write-Host "🛠 Test"

    $modulesRoot = Join-Path $PSScriptRoot "..\modules"
    Import-Module (Join-Path $modulesRoot "GetConfig") -Force

    $config = Get-Config

    $params = @{
        OutputPath   = ".\output\TestPlan_$($config.TestPlanName).xlsx"
        Organization = $config.Organization
        Project      = $config.Project
        PlanId       = $config.TestPlanId
        Pat          = $config.AZURE_DEVOPS_EXT_PAT
        FieldMap     = $config.FieldNames
    }
    
    Get-AdoTestPlanExport @params
}
