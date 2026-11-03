# Runs every non-interactive tool against sample data in two selection modes and reports failures.
param([string]$Xlam = (Join-Path (Split-Path $PSScriptRoot) 'Excel-100x-Pro.xlam'))
$skip = 'split merge prefix suffix zeros findrep roundv topn botn t_loan pdfsel pdfall pdfeach protectall unprotectall zoom sheetnav copysheet insights forecast about mask allvalues delempty web breaklinks backup backup2 scrub compilebooks foldersheets listfiles splitbook comparesheets lockformulas dropdown calendar payslips emailpdf dataform pdfactive' -split ' '
$ids = (node -e "const {TABS}=require('$($PSScriptRoot -replace '\\','/')/xlam-tools.js');console.log(TABS.flatMap(t=>t.groups.map(g=>g[2])).join(' '))") -split ' '
$poll = Join-Path $env:TEMP 'pro100x-poll.ps1'; $log = Join-Path $env:TEMP 'pro100x-dialogs.txt'; Remove-Item $log -ErrorAction SilentlyContinue
@'
Add-Type -AssemblyName UIAutomationClient
$root = [Windows.Automation.AutomationElement]::RootElement
while ($true) {
  $w = $null
  foreach ($nm in 'Microsoft Visual Basic for Applications', 'Excel 100x Pro', 'Microsoft Excel') {
    $c = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, $nm)
    $w = $root.FindFirst([Windows.Automation.TreeScope]::Children, $c); if ($w) { break }
  }
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
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
$fail = @(); $ok = 0
try {
  $null = $xl.Workbooks.Open($Xlam)
  foreach ($mode in 'range', 'cell') {
    foreach ($id in $ids) {
      if ($skip -contains $id) { continue }
      $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
      $a = New-Object 'object[,]' 61, 5
      $h = 'Name', 'Region', 'Amount', 'Date', 'Contact'
      for ($j = 0; $j -lt 5; $j++) { $a[0, $j] = $h[$j] }
      $reg = 'North', 'South', 'East', 'West'
      for ($i = 1; $i -le 60; $i++) {
        $a[$i, 0] = if ($i % 7 -eq 0) { "  Item $($i % 5)  " } else { "item $($i % 9)" }
        $a[$i, 1] = $reg[$i % 4]
        $a[$i, 2] = if ($i % 11 -eq 0) { "1,2$i" } else { [double](100 + ($i * 37) % 500) }
        $a[$i, 3] = [double](46000 + $i * 3)
        $a[$i, 4] = "user$i@mail.com 98765432$($i % 10)0"
      }
      $ws.Range('A1:E61').Value2 = $a
      $ws.Range('D2:D61').NumberFormat = 'dd-mmm-yyyy'
      if ($mode -eq 'range') { $ws.Range('A1:E61').Select() | Out-Null } else { $ws.Range('C3').Select() | Out-Null }
      "== [$mode] $id" | Out-File -Append $log
      try { $r = $xl.Run("'$((Split-Path $Xlam -Leaf))'!TestTool", $id); if ($r -eq 'ok') { $ok++ } else { $fail += "[$mode] $id : $r" } }
      catch { $fail += "[$mode] $id : COM " + $_.Exception.Message }
      foreach ($w in @($xl.Workbooks)) { if ($w.Name -notlike 'Excel-100x-Pro*') { $w.Close($false) } }
    }
  }
} finally { Stop-Process -Id $p.Id -Force; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
"OK: $ok   FAILED: $($fail.Count)"
$fail

'--- dialog log:'; if (Test-Path $log) { Get-Content $log | Select-String -Context 0,1 'Visual Basic|error|Error|Run-time|Compile' | Select -First 60 }
