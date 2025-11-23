
$modulesRoot = Join-Path $PSScriptRoot "..\modules"
Import-Module (Join-Path $modulesRoot "GetConfig") -Force

$config = Get-Config
	
$headers = @{
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($config.AZURE_DEVOPS_EXT_PAT)"))
}
$url = "https://dev.azure.com/$($config.Organization)/$($config.Project)/_apis/wit/fields?api-version=6.0"
$response = Invoke-RestMethod -Uri $url -Headers $headers
Write-Host Show API list... -ForegroundColor Green
$response.value | Select-Object -Property name, referenceName