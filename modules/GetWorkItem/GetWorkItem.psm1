
function Get-WorkItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$WorkItemId,

        [Parameter(Mandatory)]
        [string]$Organization,

        [Parameter(Mandatory)]
        [string]$Project,

        [Parameter(Mandatory)]
        [string]$Pat,

        [string]$OutputPath
    )

    # Create auth header
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    $headers = @{
        Authorization = "Basic $base64AuthInfo"
    }
    $EncodedOrg     = [System.Uri]::EscapeDataString($Organization)
    $EncodedProject = [System.Uri]::EscapeDataString($Project)
	
    # API URL
    $url = "https://dev.azure.com/$EncodedOrg/$EncodedProject/_apis/wit/workitems/${WorkItemId}?api-version=7.1-preview.3"
    Write-Debug "Calling: $url"

    try {
        $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        Write-Debug ("Raw JSON from API:`n$response")
    }
    catch {
        Write-Host "❌ Failed to fetch work item ID: $WorkItemId" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
        return $null
    }

    # --- Default output path ---
    if (-not $OutputPath) {
        $timestamp = (Get-Date).ToString("yyyyMMdd_HH")
        $OutputPath = ".\output\workitems\$timestamp\$WorkItemId.json"
        Write-Debug "OutputPath not provided — using default: $OutputPath"
    }

    # --- Ensure directory exists ---
    $dir = Split-Path $OutputPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    # --- Save JSON file ---
    try {
        $jsonString = $response | ConvertTo-Json -Depth 99
        $jsonString | Out-File -FilePath $OutputPath -Encoding utf8

        Write-Host "📄 JSON saved to: $OutputPath" -ForegroundColor Green

        return $response
    }
    catch {
        Write-Host "❌ Failed to write JSON file: $OutputPath" -ForegroundColor Red
        Write-Host $_.Exception.Message
        return $null
    }
}

function AzCliGet-WorkItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$WorkItemId,

        [Parameter(Mandatory)]
        [string]$Organization,

        [Parameter(Mandatory)]
        [string]$Project,

        [Parameter(Mandatory)]
        [string]$Pat,

        [string]$OutputPath
    )

    # Set PAT
#    $Env:AZURE_DEVOPS_EXT_PAT = $Pat
    $EncodedOrg = [System.Uri]::EscapeDataString($Organization)

    Write-Debug "Getting work item $WorkItemId via az CLI..."

    $workItemDetailsJson = (echo $Pat | az boards work-item show `
        --id $WorkItemId `
        --organization "https://dev.azure.com/$EncodedOrg")

    if (-not $workItemDetailsJson) {
        Write-Host "⚠️ Failed to get details for ID $WorkItemId" -ForegroundColor Yellow
        return $null
    }
    Write-Debug ("Raw JSON from az:`n$workItemDetailsJson")

    # 轉換 JSON
    try {
        $response = $workItemDetailsJson | ConvertFrom-Json
    }
    catch {
        Write-Host "❌ Failed to parse JSON!" -ForegroundColor Red
        Write-Host $_.Exception.Message
        return $null
    }

    # ---- OutputPath 處理 ----
    if (-not $OutputPath) {
        $timestamp = (Get-Date).ToString("yyyyMMdd_HH")
        $OutputPath = ".\output\workitems\$timestamp\${WorkItemId}_az.json"
        Write-Debug "OutputPath not provided → using default: $OutputPath"
    }

    # 自動建立資料夾
    $dir = Split-Path $OutputPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    # 寫入 JSON 檔案
    try {
        $jsonStr = $response | ConvertTo-Json -Depth 99
        $jsonStr | Out-File -FilePath $OutputPath -Encoding utf8

        Write-Host "📄 JSON saved to: $OutputPath" -ForegroundColor Green
        return $response
    }
    catch {
        Write-Host "❌ Failed to write JSON file!" -ForegroundColor Red
        Write-Host $_.Exception.Message
        return $null
    }
}

Export-ModuleMember -Function Get-WorkItem, AzCliGet-WorkItem
