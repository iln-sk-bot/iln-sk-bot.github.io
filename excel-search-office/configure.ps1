#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PagesUrl
)
$ErrorActionPreference = 'Stop'
$uri = [Uri]$PagesUrl
if (-not $uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.Query -or $uri.Fragment) {
    throw 'PagesUrl は HTTPS の公開先ディレクトリ URL にしてください。例: https://USER.github.io/REPO/'
}
$base = $uri.AbsoluteUri.TrimEnd('/') + '/'
$root = 'https://example.github.io/excel-search-office-taskpane/'
$template = Join-Path $PSScriptRoot 'manifest.template.xml'
$output = Join-Path $PSScriptRoot 'manifest.xml'
$content = [System.IO.File]::ReadAllText($template)
$content = $content.Replace($root, $base)
$content = $content.Replace('https://example.github.io</AppDomain>', "$($uri.GetLeftPart([System.UriPartial]::Authority))</AppDomain>")
[xml]$checked = $content
[System.IO.File]::WriteAllText($output, $content, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "manifest.xml を生成しました: $output"
Write-Host "Officeアドイン画面: $($base)taskpane.html"
Write-Host 'docs/ を GitHub Pages に配置し、manifest.xml はExcel側にサイドロードしてください。'
