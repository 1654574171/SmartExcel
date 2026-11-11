<#
  Isolated regression test for tools/install.ps1 (Windows only).
  It never touches the real jsaddons folder: -PublishPath redirects all
  output to a temp folder, which is removed at the end.
  Run: npm run test:install
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$tmp = Join-Path $env:TEMP 'smartsheet-publish-test'
if (Test-Path $tmp) { Remove-Item -Path $tmp -Recurse -Force }
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$publishPath = Join-Path $tmp 'publish.xml'
$installer = Join-Path $root 'tools/install.ps1'
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Write-Publish([string]$text) {
  [System.IO.File]::WriteAllText($publishPath, $text, $utf8)
}

function Assert-Publish {
  param([string[]]$Expected, [string]$Label)
  $text = [System.IO.File]::ReadAllText($publishPath)
  if (-not $text.TrimStart().StartsWith('<?xml')) { throw "$Label : missing XML declaration" }
  if (([regex]::Matches($text, '<\?xml')).Count -ne 1) { throw "$Label : duplicate XML declaration" }
  [xml]$x = $text
  $nodes = @($x.DocumentElement.ChildNodes)
  $names = @($nodes | ForEach-Object { $_.GetAttribute('name') })
  if ($x.DocumentElement.LocalName -ne 'jsplugins' -or $nodes.Count -ne $Expected.Count) {
    throw "$Label : unexpected node count $($nodes.Count)"
  }
  foreach ($e in $Expected) { if ($names -notcontains $e) { throw "$Label : missing $e" } }
  foreach ($legacy in @('WPS_AI_ET', 'WPS_AI_WPS')) {
    if ($names -contains $legacy) { throw "$Label : legacy entry $legacy was not removed" }
  }
  # Third-party entries may still target other components; only our own must be et-only.
  foreach ($node in $nodes) {
    $n = $node.GetAttribute('name')
    if ($n -like 'SMART_SHEET*' -or $n -like 'WPS_AI_*') {
      if ($node.GetAttribute('type') -ne 'et') { throw "$Label : own plugin registered for a non-spreadsheet component" }
    }
  }
  $et = @($nodes | Where-Object { $_.GetAttribute('type') -eq 'et' -and $_.GetAttribute('name') -eq 'SMART_SHEET' })
  if ($et.Count -gt 1) { throw "$Label : duplicate spreadsheet entry" }
  Write-Output ("$Label : OK (" + $nodes.Count + " nodes)")
}

$broken = @'
<?xml version="1.0" encoding="UTF-8"?>
  <jspluginonline name="officeai-wps-personal" type="wps" url="http://127.0.0.1:25682/officeai-wps/pages/plugin/" enable="enable" install="null"/>
  <jspluginonline name="officeai-wps-personal-et" type="et" url="http://127.0.0.1:25682/officeai-wps/pages/plugin-et/" enable="enable" install="null"/>
</jsplugins>
  <jspluginonline name="officeai-wps-personal" type="wps" url="http://127.0.0.1:25682/officeai-wps/pages/plugin/" enable="enable" install="null"/>
  <jspluginonline name="officeai-wps-personal-et" type="et" url="http://127.0.0.1:25682/officeai-wps/pages/plugin-et/" enable="enable" install="null"/>
</jsplugins>
<jsplugins><jsplugin url="http://127.0.0.1:3889/" type="wps" enable="enable_dev" install="null" version="1.0.0" name="wps-zotero"/>
  <jspluginonline name="officeai-wps-personal" type="wps" url="http://127.0.0.1:25682/officeai-wps/pages/plugin/" enable="enable" install="null"/>
  <jspluginonline name="officeai-wps-personal-et" type="et" url="http://127.0.0.1:25682/officeai-wps/pages/plugin-et/" enable="enable" install="null"/>
</jsplugins>
'@
Write-Publish $broken
& powershell -NoProfile -ExecutionPolicy Bypass -File $installer -PublishPath $publishPath | Out-Null
Assert-Publish @('officeai-wps-personal', 'officeai-wps-personal-et', 'wps-zotero', 'SMART_SHEET') 'SCENE 1 malformed rebuild'

$valid = @'
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jspluginonline name="officeai-wps-personal" type="wps" url="http://127.0.0.1:25682/officeai-wps/pages/plugin/" enable="enable" install="null"/>
  <jsplugin name="wps-zotero" type="wps" url="http://127.0.0.1:3889/" enable="enable_dev" install="null" version="1.0.0"/>
  <jsplugin name="WPS_AI_ET" type="et" url="file://" enable="enable_dev" version="1.0.0"/>
  <jsplugin name="WPS_AI_WPS" type="wps" url="file://" enable="enable_dev" version="1.0.0"/>
</jsplugins>
'@
Write-Publish $valid
& powershell -NoProfile -ExecutionPolicy Bypass -File $installer -PublishPath $publishPath | Out-Null
Assert-Publish @('officeai-wps-personal', 'wps-zotero', 'SMART_SHEET') 'SCENE 2 valid publish.xml'

& powershell -NoProfile -ExecutionPolicy Bypass -File $installer -PublishPath $publishPath | Out-Null
Assert-Publish @('officeai-wps-personal', 'wps-zotero', 'SMART_SHEET') 'SCENE 3 reinstall idempotent'

& powershell -NoProfile -ExecutionPolicy Bypass -File $installer -PublishPath $publishPath -Uninstall | Out-Null
Assert-Publish @('officeai-wps-personal', 'wps-zotero') 'SCENE 4 uninstall'

Remove-Item -Path $tmp -Recurse -Force
Write-Output 'ALL INSTALLER TESTS PASSED'
