<#
  Register this spreadsheet-only add-in with the local client (type=et).
  ASCII-only on purpose: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.
  Usage:
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/install.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/install.ps1 -Url "http://127.0.0.1:3890/addon/"
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/install.ps1 -Uninstall
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/install.ps1 -PublishPath C:\temp\publish.xml
  If an existing publish.xml is malformed (multiple root nodes, repeated
  </jsplugins> tags written by other installers), this script rebuilds it and
  keeps every recoverable jsplugin / jspluginonline entry. The broken original
  is saved as publish.xml.broken-<timestamp>, plus the usual publish.xml.bak.
#>
param(
  [string]$Url = 'file://',
  [switch]$Uninstall,
  [string]$PublishPath = ''
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$addonDir = Join-Path $projectRoot 'addon'
$jsaddons = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
if (-not [string]::IsNullOrWhiteSpace($PublishPath)) {
  $jsaddons = Split-Path -Parent $PublishPath
  $publish = $PublishPath
} else {
  $publish = Join-Path $jsaddons 'publish.xml'
}
$version = '1.0.0'
$names = @('SMART_SHEET')
$types = @('et')
$legacyNames = @('WPS_AI_ET', 'WPS_AI_WPS')

function Test-OwnPluginName {
  param([string]$Name)
  return ($Name -like 'SMART_SHEET*' -or $Name -like 'WPS_AI_*')
}

function New-PublishDocument {
  $d = New-Object System.Xml.XmlDocument
  $r = $d.CreateElement('jsplugins')
  [void]$d.AppendChild($r)
  return $d
}

function Restore-PluginNodes {
  param([string]$Raw, [System.Xml.XmlDocument]$Document)
  $root = $Document.DocumentElement
  $tagPattern = '<(jspluginonline|jsplugin)\b([^>]*?)/?>'
  $attrPattern = "([\w:-]+)\s*=\s*(?:""([^""]*)""|'([^']*)')"
  $seen = @{}
  $count = 0
  foreach ($match in [regex]::Matches($Raw, $tagPattern, 'IgnoreCase')) {
    $tag = $match.Groups[1].Value
    $attrs = @{}
    foreach ($attrMatch in [regex]::Matches($match.Groups[2].Value, $attrPattern)) {
      $value = $attrMatch.Groups[2].Value
      if (-not $attrMatch.Groups[2].Success) { $value = $attrMatch.Groups[3].Value }
      $key = $attrMatch.Groups[1].Value
      if ($key -eq 'url' -or $key -eq 'version' -or $key -eq 'name' -or $key -eq 'type') {
        $value = $value -replace '\s+', ''
      }
      $attrs[$key] = $value
    }
    if (-not $attrs.ContainsKey('name')) { continue }
    if (Test-OwnPluginName $attrs['name']) { continue }
    $dedup = $attrs['name'] + '|' + $attrs['type']
    if ($seen.ContainsKey($dedup)) { continue }
    $seen[$dedup] = $true
    $node = $Document.CreateElement($tag)
    foreach ($k in $attrs.Keys) { $node.SetAttribute($k, [string]$attrs[$k]) }
    [void]$root.AppendChild($node)
    $count++
  }
  return $count
}

if (-not (Test-Path $addonDir)) { throw "Add-in folder not found: $addonDir" }
if (-not (Test-Path $jsaddons)) { New-Item -ItemType Directory -Path $jsaddons -Force | Out-Null }

$doc = $null
$repaired = $false
$recovered = 0
$brokenBackup = ''
if (Test-Path $publish) {
  Copy-Item -Path $publish -Destination ($publish + '.bak') -Force
  $raw = Get-Content -Path $publish -Raw -Encoding UTF8
  try {
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($raw)
    if ($doc.DocumentElement -eq $null -or $doc.DocumentElement.LocalName -ne 'jsplugins') {
      throw 'Unexpected root node'
    }
  } catch {
    $repaired = $true
    $brokenBackup = $publish + '.broken-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    Copy-Item -Path $publish -Destination $brokenBackup -Force
    $doc = New-PublishDocument
    $recovered = Restore-PluginNodes -Raw $raw -Document $doc
  }
} else {
  $doc = New-PublishDocument
}

$rootNode = $doc.DocumentElement

foreach ($existing in @($doc.SelectNodes('/jsplugins/*'))) {
  if (Test-OwnPluginName $existing.GetAttribute('name')) {
    [void]$rootNode.RemoveChild($existing)
  }
}

# Remove folders left by older installs so only one tab remains in the ribbon.
foreach ($old in $legacyNames) {
  $oldDest = Join-Path $jsaddons ($old + '_' + $version)
  if (Test-Path $oldDest) { Remove-Item -Path $oldDest -Recurse -Force }
}

if (-not $Uninstall) {
  for ($i = 0; $i -lt $names.Count; $i++) {
    $dest = Join-Path $jsaddons ($names[$i] + '_' + $version)
    if (Test-Path $dest) { Remove-Item -Path $dest -Recurse -Force }
    Copy-Item -Path $addonDir -Destination $dest -Recurse -Force

    $node = $doc.CreateElement('jsplugin')
    $node.SetAttribute('name', $names[$i])
    $node.SetAttribute('enable', 'enable_dev')
    $node.SetAttribute('url', $Url)
    $node.SetAttribute('type', $types[$i])
    $node.SetAttribute('version', $version)
    [void]$rootNode.AppendChild($node)
    Write-Host ("Installed {0}: {1}" -f $types[$i], $dest)
  }
} else {
  foreach ($name in $names) {
    $dest = Join-Path $jsaddons ($name + '_' + $version)
    if (Test-Path $dest) { Remove-Item -Path $dest -Recurse -Force }
  }
  Write-Host 'Removed add-in folders.'
}

foreach ($node in @($doc.ChildNodes)) {
  if ($node -is [System.Xml.XmlDeclaration]) { [void]$doc.RemoveChild($node) }
}

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Indent = $true
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$settings.OmitXmlDeclaration = $false
$writer = [System.Xml.XmlWriter]::Create($publish, $settings)
try {
  $doc.Save($writer)
} finally {
  $writer.Close()
}
Write-Host ("Updated {0} (backup: publish.xml.bak)" -f $publish)
if ($repaired) {
  Write-Host ("WARNING: original publish.xml was malformed; rebuilt and recovered {0} third-party entries." -f $recovered)
  Write-Host ("Broken original saved as: {0}" -f $brokenBackup)
}

if (-not $Uninstall) {
  Write-Host ''
  Write-Host 'Done. Fully quit and restart the spreadsheet app.'
  Write-Host 'Then open the Smart Sheet tab in the ribbon and click Open Sidebar.'
}
