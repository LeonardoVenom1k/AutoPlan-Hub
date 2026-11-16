$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$versionMatch = Select-String -Path (Join-Path $projectRoot 'app_info.py') -Pattern "APP_VERSION = '([0-9]+\.[0-9]+\.[0-9]+)'"
if (-not $versionMatch) {
    throw 'Não foi possível identificar APP_VERSION.'
}
$version = $versionMatch.Matches[0].Groups[1].Value
$name = "AutoPlan-Hub-v$version"
$outputDir = Join-Path $projectRoot 'executavel'
$workDir = Join-Path $env:TEMP 'autoplan-hub-build'

python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $name `
    --icon (Join-Path $projectRoot 'assets\autoplan-hub.ico') `
    --add-data "$(Join-Path $projectRoot 'assets');assets" `
    --add-data "$(Join-Path $projectRoot 'sounds');sounds" `
    --hidden-import win32com.client `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --distpath $outputDir `
    --workpath $workDir `
    --specpath $workDir `
    (Join-Path $projectRoot 'main.py')

$executable = Join-Path $outputDir "$name.exe"
if (-not (Test-Path -LiteralPath $executable)) {
    throw "O executável não foi criado: $executable"
}

$hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $executable).Hash.ToLowerInvariant()
$manifest = [ordered]@{
    version = $version
    filename = "$name.exe"
    download_url = "https://raw.githubusercontent.com/LeonardoVenom1k/AutoPlan-Hub/main/executavel/$name.exe"
    sha256 = $hash
}
$manifestJson = $manifest | ConvertTo-Json
$utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $outputDir 'manifest.json'), $manifestJson, $utf8WithoutBom)
Write-Output $executable
Write-Output "SHA-256: $hash"

