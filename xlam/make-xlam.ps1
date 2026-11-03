# Builds Excel-100x-Pro.xlam with your installed Excel.
# One-time requirement: Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings >
#   tick "Trust access to the VBA project object model".   Then:  powershell -ExecutionPolicy Bypass -File xlam\make-xlam.ps1
param([string]$Out = (Join-Path (Split-Path $PSScriptRoot) 'Excel-100x-Pro.xlam'))
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
node (Join-Path $PSScriptRoot 'build-xlam.js'); if ($LASTEXITCODE) { throw 'ribbon build failed' }

# VBA import wants ANSI text with CRLF
$tmp = Join-Path $env:TEMP 'pro100x-bas'; Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue; New-Item $tmp -ItemType Directory | Out-Null
$mods = @()
Get-ChildItem $PSScriptRoot -Filter 'Pro*.bas' | ForEach-Object {
  $t = ((Get-Content $_.FullName -Raw) -replace "`r?`n", "`r`n"); $p = Join-Path $tmp $_.Name
  [IO.File]::WriteAllText($p, $t, [Text.Encoding]::GetEncoding(1252)); $mods += $p }
$mods += (Join-Path $PSScriptRoot 'out\Dispatch.bas')

Remove-Item $Out -ErrorAction SilentlyContinue
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  try { if (-not $wb.VBProject) { throw 'x' }; $null = $wb.VBProject.VBComponents.Count }
  catch { throw 'Excel blocks VBA access. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > "Trust access to the VBA project object model", then run again.' }
  foreach ($m in $mods) { $null = $wb.VBProject.VBComponents.Import($m) }
  $wb.SaveAs($Out, 55)      # xlOpenXMLAddIn
  $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

python (Join-Path $PSScriptRoot 'pack_xlam.py') $Out (Join-Path $PSScriptRoot 'out'); if ($LASTEXITCODE) { throw 'pack failed' }
Write-Host "Built $Out"
