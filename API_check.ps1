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

$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($config.AZURE_DEVOPS_EXT_PAT)"))
}
$url = "https://dev.azure.com/$($config.Organization)/$($config.Project)/_apis/wit/fields?api-version=6.0"
$response = Invoke-RestMethod -Uri $url -Headers $headers
Write-Host `n`tShow API list... -ForegroundColor Green
$response.value | Select-Object -Property name, referenceName