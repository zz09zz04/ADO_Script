function Get-Config {
    [CmdletBinding()]
    param(
        [string]$ConfigName = "_config.json"
    )

    # 取得 module 所在路徑
    $moduleRoot = Split-Path -Parent $PSCommandPath

    # config 路徑 (module/../../config)
    $configPath = Join-Path $moduleRoot "..\..\config\$ConfigName"
    $fullPath = (Resolve-Path $configPath).Path

    if (-not (Test-Path $fullPath)) {
        throw "Config file not found: $fullPath"
    }

    # 讀取 JSON
    $json = Get-Content $fullPath -Raw | ConvertFrom-Json
	Write-Host $(Get-Content $fullPath -Raw) -ForegroundColor Cyan
    return $json
}

Export-ModuleMember -Function Get-Config
