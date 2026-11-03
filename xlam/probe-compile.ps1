# Finds VBA compile errors: calls one procedure per module, dismisses the VBE error dialog and reports its text + the offending line.
param([string]$Xlam = (Join-Path (Split-Path $PSScriptRoot) 'Excel-100x-Pro.xlam'))
$poll = Join-Path $env:TEMP 'pro100x-poll.ps1'; $log = Join-Path $env:TEMP 'pro100x-dialogs.txt'; Remove-Item $log -ErrorAction SilentlyContinue
@'
Add-Type -AssemblyName UIAutomationClient
$root = [Windows.Automation.AutomationElement]::RootElement
while ($true) {
  $c = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, 'Microsoft Visual Basic for Applications')
  $w = $root.FindFirst([Windows.Automation.TreeScope]::Children, $c)
  if ($w) {
    $t = $w.FindAll([Windows.Automation.TreeScope]::Descendants, (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Text))) | % { $_.Current.Name }
    ($t -join ' ') | Out-File -Append $args[0]
    $b = $w.FindFirst([Windows.Automation.TreeScope]::Descendants, (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Button)))
    if ($b) { try { $b.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke() } catch {} }
    Start-Sleep -Milliseconds 400
  }
  Start-Sleep -Milliseconds 200
}
'@ | Set-Content $poll
$p = Start-Process powershell -ArgumentList '-NoProfile', '-File', $poll, $log -PassThru -WindowStyle Hidden
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
try {
  $null = $xl.Workbooks.Open($Xlam); $null = $xl.Workbooks.Add()
  $n = (Split-Path $Xlam -Leaf)
  foreach ($m in 'ClearBar', 'ToggleGrid', 'HeatMap', 'MakePayslip', 'SelfTest', 'DataForm', 'RunTool') {
    try { $xl.Run("'$n'!$m", 'zzz') | Out-Null } catch { try { $xl.Run("'$n'!$m") | Out-Null } catch { } }
    $line = ''
    try { $cp = $xl.VBE.ActiveCodePane; $s = 0; $e = 0; $a = 0; $b = 0; $cp.GetSelection([ref]$s, [ref]$e, [ref]$a, [ref]$b); $line = "$($cp.CodeModule.Name):$s  " + $cp.CodeModule.Lines($s, 1) } catch { }
    "$m -> $line"
  }
} finally { foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} }; $xl.Quit(); Stop-Process -Id $p.Id -Force }
'--- dialogs:'; if (Test-Path $log) { Get-Content $log }
