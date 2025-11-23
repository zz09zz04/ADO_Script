# 找到 module 路徑
$modulesRoot = Join-Path $PSScriptRoot "..\modules"

# 載入 Config module
Import-Module (Join-Path $modulesRoot "GetConfig") -Force

$config = Get-Config

Write-Host $config

. "$PSScriptRoot\Get-TestCaseNew.ps1"

$params = @{
	OutputPath   = ".\output\TestPlan_$($config.TestPlanName).xlsx"
	Organization = $config.Organization
	Project      = $config.Project
	PlanId       = $config.TestPlanId
	Pat          = $config.AZURE_DEVOPS_EXT_PAT
	FieldMap     = $config.FieldNames
}

Get-AdoTestPlanExport @params

Import-Module (Join-Path $modulesRoot "SetExcelFormat") -Force

Set-ExcelConditionalFormat -Path ".\output\TestPlan_$($config.TestPlanName).xlsx"
