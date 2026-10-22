#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PagesUrl
)

$ErrorActionPreference = 'Stop'
$uri = [Uri]$PagesUrl
if ($uri.Scheme -ne 'https' -or -not $uri.IsAbsoluteUri -or $uri.Query -or $uri.Fragment) {
    throw 'PagesUrl は HTTPS の公開先ディレクトリ URL にしてください。例: https://USER.github.io/REPO/'
}
$base = $uri.AbsoluteUri.TrimEnd('/') + '/'
$template = Join-Path $PSScriptRoot 'manifest.template.xml'
$output = Join-Path $PSScriptRoot 'manifest.xml'
$oldBase = 'https://example.github.io/excel-search-office-popup-poc/'
$content = [System.IO.File]::ReadAllText($template)
$content = $content.Replace($oldBase, $base)
$content = $content.Replace('https://example.github.io</AppDomain>', "$($uri.GetLeftPart([System.UriPartial]::Authority))</AppDomain>")
[xml]$validated = $content
[System.IO.File]::WriteAllText($output, $content, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "manifest.xml を生成しました: $output"
Write-Host "Web アプリ: $($base)host.html"
Write-Host "デモ画面: $($base)dialog.html?demo=1"
Write-Host 'manifest.xml のみを Windows 上でサイドロードしてください。docs/ は GitHub Pages で公開します。'
