#requires -Version 5.1
[CmdletBinding()]
param([string]$ManifestPath = (Join-Path $PSScriptRoot 'manifest.xml'))
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "manifest.xml が見つかりません: $ManifestPath"
}
[xml]$manifest = [System.IO.File]::ReadAllText($ManifestPath)
$id = [string]$manifest.OfficeApp.Id
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
Remove-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue
Write-Host "開発用アドインの登録を解除しました: $id"
