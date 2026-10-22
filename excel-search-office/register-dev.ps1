#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ManifestPath = (Join-Path $PSScriptRoot 'manifest.xml')
)
$ErrorActionPreference = 'Stop'
$ManifestPath = [System.IO.Path]::GetFullPath($ManifestPath)
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "manifest.xml が見つかりません。先に configure.ps1 を実行してください: $ManifestPath"
}
[xml]$manifest = [System.IO.File]::ReadAllText($ManifestPath)
$id = [string]$manifest.OfficeApp.Id
$checkedGuid = [guid]::Empty
if (-not [guid]::TryParse($id, [ref]$checkedGuid)) {
    throw "manifest.xml のIdがGUID形式ではありません: $id"
}
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name $id -PropertyType String -Value $ManifestPath -Force | Out-Null
Write-Host "開発用アドインを登録しました: $id"
Write-Host 'Excelを一度終了して開き直してください。'
Write-Host "解除: .\unregister-dev.ps1"
