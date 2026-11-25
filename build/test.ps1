# End-to-end test of dist\TrackChanges.xlam in a real (hidden) Excel instance.
# Usage: powershell -ExecutionPolicy Bypass -File build\test.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xlam = Join-Path $root 'dist\TrackChanges.xlam'
if (-not (Test-Path $xlam)) { throw "Build first: $xlam not found" }

$tmp = Join-Path $env:TEMP ('tc_test_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$script:fail = 0
function Check($name, $cond) {
    if ($cond) { Write-Host "  PASS  $name" } else { Write-Host "  FAIL  $name" -ForegroundColor Red; $script:fail++ }
}

$BLUE = 255 * 65536      # RGB(0,0,255) as BGR long
$WHITE = 16777215

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = New-Object -ComObject Excel.Application
$xlPid = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $add = $xl.Workbooks.Open($xlam)
    $an = "'" + $add.Name + "'!"
    function Run($macro) { $xl.Run($an + $macro) }
    function RunA($macro, $a, $b) { $xl.Run($an + $macro, $a, $b) }
    $xl.Run($an + 'TC_SetSilent', $true) | Out-Null
    RunA 'TC_SetStyle' 0 0 | Out-Null

    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Bob's data"
    $ws.Range('A1').Value2 = 'abc'
    $ws.Range('A2').Formula = '=""'
    $ws.Range('A3').Value2 = 5
    $ws.Range('A4').Formula = '=NA()'
    $ws.Range('A5').NumberFormat = '@'; $ws.Range('A5').Value2 = '001'
    $ws.Range('A6').Value2 = $true
    $ws.Range('A7').Formula = '=A3*2'
    $ws.Range('B1').Font.Color = 255; $ws.Range('B1').Interior.Color = 65535; $ws.Range('B1').Value2 = 'mine'
    $ws.Range('A10').Value2 = 'hdr'; 11..20 | % { $ws.Range("A$_").Value2 = $_ }
    $ws.Range('A10:A20').AutoFilter(1, '>15') | Out-Null       # hide some rows
    $ws2 = $wb.Worksheets.Add([Type]::Missing, $ws); $ws2.Name = 'Résumé 2'
    $ws2.Range('C3').Formula = "='Bob''s data'!A3+1"

    function Font($sh, $a) { $sh.Range($a).DisplayFormat.Font.Color }
    function Fill($sh, $a) { $sh.Range($a).DisplayFormat.Interior.Color }

    Write-Host 'Start tracking'
    Run 'TC_Start' | Out-Null
    $snaps = @($wb.Worksheets | ? { $_.Name -like '_tc_*' })
    Check 'two snapshot sheets created' ($snaps.Count -eq 2)
    Check 'snapshot sheets are very hidden' (@($snaps | ? { $_.Visible -eq 2 }).Count -eq 2)
    Check 'active sheet restored' ($wb.ActiveSheet.Name -eq 'Résumé 2')
    Check 'one rule per sheet' ($ws.Cells.FormatConditions.Count -eq 1 -and $ws2.Cells.FormatConditions.Count -eq 1)
    foreach ($a in 'A1','A2','A3','A4','A5','A6','A7','A11','A20','D5') { Check "unchanged $a not highlighted" ((Font $ws $a) -ne $BLUE) }
    Check 'user format visible' ((Font $ws 'B1') -eq 255 -and (Fill $ws 'B1') -eq 65535)

    Write-Host 'Edit cells'
    $ws.Range('A1').Value2 = 'ABC'
    $ws.Range('A2').Value2 = 0
    $ws.Range('A3').Value2 = 6
    $ws.Range('B1').Value2 = 'changed'
    $ws.Range('D5').Value2 = 'new'
    $ws.Range('A12').Value2 = 99        # hidden by filter
    Check 'case change highlighted' ((Font $ws 'A1') -eq $BLUE)
    Check 'empty text -> 0 highlighted' ((Font $ws 'A2') -eq $BLUE)
    Check 'number change highlighted' ((Font $ws 'A3') -eq $BLUE)
    Check 'dependent formula highlighted' ((Font $ws 'A7') -eq $BLUE)
    Check 'other-sheet formula highlighted' ((Font $ws2 'C3') -eq $BLUE)
    Check 'new cell highlighted' ((Font $ws 'D5') -eq $BLUE)
    Check 'user fill kept under font highlight' ((Fill $ws 'B1') -eq 65535 -and (Font $ws 'B1') -eq $BLUE)
    Check 'error cell untouched' ((Font $ws 'A4') -ne $BLUE)

    Write-Host 'Revert'
    $ws.Range('A3').Value2 = 5; $ws.Range('B1').Value2 = 'mine'; $ws.Range('D5').ClearContents() | Out-Null
    Check 'reverted number not highlighted' ((Font $ws 'A3') -ne $BLUE)
    Check 'dependent formula back' ((Font $ws 'A7') -ne $BLUE)
    Check 'user colors back' ((Font $ws 'B1') -eq 255 -and (Fill $ws 'B1') -eq 65535)
    Check 'cleared new cell not highlighted' ((Font $ws 'D5') -ne $BLUE)

    Write-Host 'Style: fill, green'
    RunA 'TC_SetStyle' 1 2 | Out-Null
    $GREENFILL = 206 * 65536 + 239 * 256 + 198
    Check 'fill style applied' ((Fill $ws 'A1') -eq $GREENFILL -and (Font $ws 'A1') -ne $BLUE)
    Check 'still one rule' ($ws.Cells.FormatConditions.Count -eq 1)
    RunA 'TC_SetStyle' 0 0 | Out-Null

    Write-Host 'Accept selection'
    $ws.Activate(); $ws.Range('A1:A2').Select() | Out-Null
    Run 'TC_AcceptSelection' | Out-Null
    Check 'accepted cells not highlighted' ((Font $ws 'A1') -ne $BLUE -and (Font $ws 'A2') -ne $BLUE)
    Check 'other changes still highlighted' ((Font $ws 'A12') -eq $BLUE)

    Write-Host 'Paste with formatting + Refresh'
    $ws.Range('Z1').Value2 = 'x'; $ws.Range('Z1').Copy() | Out-Null; $ws.Range('F1:F3').PasteSpecial(-4104) | Out-Null
    $ws3 = $wb.Worksheets.Add(); $ws3.Name = 'Added'; $ws3.Range('A1').Value2 = 1
    Run 'TC_Refresh' | Out-Null
    Check 'refresh: one rule again' ($ws.Cells.FormatConditions.Count -eq 1)
    Check 'refresh: pasted cell highlighted' ((Font $ws 'F2') -eq $BLUE)
    Check 'refresh: new sheet tracked' ($ws3.Cells.FormatConditions.Count -eq 1 -and (Font $ws3 'A1') -ne $BLUE)
    $ws3.Range('A1').Value2 = 2
    Check 'new sheet change highlighted' ((Font $ws3 'A1') -eq $BLUE)

    Write-Host 'Save as xlsx and reopen'
    $path = Join-Path $tmp 'book.xlsx'
    $wb.SaveAs($path, 51); $wb.Close($false)
    $wb = $xl.Workbooks.Open($path); $ws = $wb.Worksheets.Item("Bob's data")
    Check 'highlight survives reopen' ((Font $ws 'A12') -eq $BLUE -and (Font $ws 'A3') -ne $BLUE)

    Write-Host 'Copy to another workbook'
    $other = $xl.Workbooks.Add()
    $ws.Range('A1:A12').Copy() | Out-Null
    $other.Worksheets.Item(1).Range('A1').PasteSpecial(-4104) | Out-Null
    $links = $null
    try { $links = $other.LinkSources(1) } catch { }
    $linkText = 'none'
    if ($links -is [array] -and $links.Length -gt 0) { $linkText = ($links | % { "$_" }) -join ', ' }
    Write-Host ("  INFO  external links after paste: " + $linkText)
    Write-Host ("  INFO  names in other workbook: " + (@($other.Names | % { $_.Name + ' = ' + $_.RefersTo }) -join '; '))
    Write-Host ("  INFO  rules in other workbook: " + $other.Worksheets.Item(1).Cells.FormatConditions.Count)
    $other.Activate()
    Run 'TC_Stop' | Out-Null
    $links = $null
    try { $links = $other.LinkSources(1) } catch { }
    Check 'turn off cleans copied rule' ($other.Worksheets.Item(1).Cells.FormatConditions.Count -eq 0)
    Check 'turn off cleans copied names' (@($other.Names).Count -eq 0)
    Check 'turn off removes external link' (-not ($links -is [array] -and $links.Length -gt 0))
    $other.Close($false)
    $wb.Activate()

    Write-Host 'Accept all'
    Run 'TC_AcceptAll' | Out-Null
    Check 'nothing highlighted after accept all' ((Font $ws 'A12') -ne $BLUE -and (Font $ws 'F2') -ne $BLUE)

    Write-Host 'Turn off'
    Run 'TC_Stop' | Out-Null
    Check 'snapshot sheets removed' (@($wb.Worksheets | ? { $_.Name -like '_tc_*' }).Count -eq 0)
    Check 'rules removed' ($ws.Cells.FormatConditions.Count -eq 0)
    $names = @($wb.Names | ? { $_.Name -like '*_TC_*' })
    Check 'hidden names removed' ($names.Count -eq 0)
    Check 'user format intact' ((Font $ws 'B1') -eq 255 -and (Fill $ws 'B1') -eq 65535)
    Check 'data intact' ($ws.Range('A12').Value2 -eq 99 -and $ws.Range('A5').Text -eq '001')
    Write-Host ('  INFO  last message: ' + $xl.Run($an + 'TC_LastMessage'))

    $wb.Close($false)
} finally {
    $xl.Run("'TrackChanges.xlam'!TC_SetStyle", 0, 0) 2>$null | Out-Null
    $xl.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # make sure our own hidden Excel instance is gone (never touches other Excel windows)
    foreach ($id in $xlPid) { Wait-Process -Id $id -Timeout 10 -ErrorAction SilentlyContinue; Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
if ($script:fail) { Write-Host "$($script:fail) check(s) failed" -ForegroundColor Red; exit 1 } else { Write-Host 'All checks passed' }
